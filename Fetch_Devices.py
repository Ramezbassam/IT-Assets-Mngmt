"""
Device Inventory
================

Concurrent Windows device inventory using WMI/DCOM.

Features
--------
- Reads device names from an Excel input file.
- Queries remote Windows computers concurrently.
- Collects:
    ComputerName
    Model
    SerialNumber
    CPU
    RAM_GB
    DiskSize_GB
    DiskType
    IPAddress
    MACAddress
    LoggedUser
    Status

- Existing successful records are skipped.
- Existing successful records missing disk information are repaired.
- Failed records are retried on subsequent executions.
- Existing records are updated instead of duplicated.
- RAM and disk sizes are rounded.
- Results are written to Excel.
- Logging is written to console and a log file.

Requirements
------------
pip install pywin32 openpyxl
"""

from __future__ import annotations

import logging
import math
import os
import re
import tempfile
import time
from concurrent.futures import (
    ThreadPoolExecutor,
    as_completed,
    Future,
)
from dataclasses import dataclass
from pathlib import Path
from threading import Lock
from typing import Any

import pythoncom
import win32com.client
from openpyxl import Workbook, load_workbook


# ============================================================
# CONFIGURATION
# ============================================================

INPUT_XLSX = Path(r"D:\Scripts\devices.xlsx")
OUTPUT_XLSX = Path(r"D:\Scripts\Inventory_Report.xlsx")
LOG_FILE = Path(r"D:\Scripts\Device-Inventory.log")

DEVICE_NAME_COLUMN = "ComputerName"

# Number of remote devices queried simultaneously.
#
# Start with 10.
# If the environment is stable, try 15 or 20 and measure.
MAX_WORKERS = 1

# Timeout for a single WMI operation.
WMI_TIMEOUT_SECONDS = 30

# Retry count for transient WMI failures.
MAX_RETRIES = 1

# Delay between retries.
RETRY_DELAY_SECONDS = 2

# Rounding configuration.
RAM_ROUND_STEP_GB = 1
DISK_ROUND_STEP_GB = 50

# Excel worksheet.
WORKSHEET_NAME = "Inventory"

# Excel output headers.
HEADERS = [
    "ComputerName",
    "Model",
    "SerialNumber",
    "CPU",
    "RAM_GB",
    "DiskSize_GB",
    "DiskType",
    "IPAddress",
    "MACAddress",
    "LoggedUser",
    "Status",
]


# ============================================================
# LOGGING
# ============================================================

def configure_logging() -> logging.Logger:
    """
    Configure logging to both console and file.
    """

    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("device_inventory")
    logger.setLevel(logging.INFO)

    # Avoid duplicate handlers if the function is called twice.
    if logger.handlers:
        return logger

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    file_handler = logging.FileHandler(
        LOG_FILE,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)

    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    return logger


logger = configure_logging()


# ============================================================
# DATA STRUCTURES
# ============================================================

@dataclass
class InventoryResult:
    """
    Result returned by one inventory operation.
    """

    computer_name: str
    record: dict[str, Any]
    success: bool
    disk_only: bool = False
    skipped: bool = False
    error: str | None = None


# ============================================================
# GENERAL HELPERS
# ============================================================

def normalize_device_name(value: Any) -> str:
    """
    Convert a device name to a normalized string.
    """

    if value is None:
        return ""

    return str(value).strip()


def is_blank(value: Any) -> bool:
    """
    Determine whether a value is empty/null.
    """

    if value is None:
        return True

    return str(value).strip() == ""


def round_size(value: float, step: int) -> int:
    """
    Round a numeric value to the nearest multiple of step.

    Uses half-away-from-zero behavior similar to the
    PowerShell implementation.
    """

    if step <= 0:
        step = 1

    quotient = value / step

    if quotient >= 0:
        rounded = math.floor(quotient + 0.5)
    else:
        rounded = math.ceil(quotient - 0.5)

    return int(rounded * step)


def safe_value(value: Any, default: str = "N/A") -> Any:
    """
    Return default for None/empty values.
    """

    if value is None:
        return default

    if isinstance(value, str) and not value.strip():
        return default

    return value


def is_ipv4(address: str) -> bool:
    """
    Basic IPv4 validation.
    """

    pattern = r"^\d{1,3}(\.\d{1,3}){3}$"

    if not re.match(pattern, address):
        return False

    try:
        return all(0 <= int(part) <= 255 for part in address.split("."))
    except ValueError:
        return False


def extract_ipv4(addresses: Any) -> str:
    """
    Extract the first IPv4 address from a WMI IPAddress value.
    """

    if addresses is None:
        return "N/A"

    if isinstance(addresses, str):
        addresses = [addresses]

    try:
        for address in addresses:
            address = str(address).strip()

            if is_ipv4(address):
                return address
    except TypeError:
        pass

    return "N/A"


# ============================================================
# WMI CONNECTION
# ============================================================

class WMIClient:
    """
    Small wrapper around WbemScripting.SWbemLocator.

    A new client is created inside each worker thread.

    This is intentional:
    COM/WMI objects should not be shared between worker threads.
    """

    def __init__(self, computer_name: str):
        self.computer_name = computer_name
        self.locator = None
        self.services = None

    def connect(self) -> None:
        """
        Connect to the remote WMI service.
        """

        self.locator = win32com.client.Dispatch(
            "WbemScripting.SWbemLocator"
        )

        self.services = self.locator.ConnectServer(
            self.computer_name,
            "root\\cimv2",
        )

    def query(
        self,
        wql: str,
        namespace: str = "root\\cimv2",
    ) -> list[Any]:
        """
        Execute a WQL query.
        """

        if namespace.lower() == "root\\cimv2":
            services = self.services
        else:
            services = self.locator.ConnectServer(
                self.computer_name,
                namespace,
            )

        result = services.ExecQuery(
            wql,
            "WQL",
            0x30,  # RETURN_IMMEDIATELY | FORWARD_ONLY
        )

        return list(result)


# ============================================================
# REMOTE QUERY FUNCTIONS
# ============================================================

def query_system(client: WMIClient) -> Any:
    """
    Query Win32_ComputerSystem.
    """

    rows = client.query(
        """
        SELECT
            Model,
            TotalPhysicalMemory,
            UserName
        FROM Win32_ComputerSystem
        """
    )

    if not rows:
        raise RuntimeError(
            "Win32_ComputerSystem returned no results."
        )

    return rows[0]


def query_bios(client: WMIClient) -> Any:
    """
    Query Win32_BIOS.
    """

    rows = client.query(
        """
        SELECT
            SerialNumber
        FROM Win32_BIOS
        """
    )

    if not rows:
        raise RuntimeError(
            "Win32_BIOS returned no results."
        )

    return rows[0]


def query_cpu(client: WMIClient) -> Any:
    """
    Query Win32_Processor.
    """

    rows = client.query(
        """
        SELECT
            Name
        FROM Win32_Processor
        """
    )

    if not rows:
        raise RuntimeError(
            "Win32_Processor returned no results."
        )

    return rows[0]


def query_network(client: WMIClient) -> Any:
    """
    Query enabled network adapters.
    """

    rows = client.query(
        """
        SELECT
            IPAddress,
            MACAddress
        FROM Win32_NetworkAdapterConfiguration
        WHERE IPEnabled = TRUE
        """
    )

    if not rows:
        raise RuntimeError(
            "No IP-enabled network adapter was found."
        )

    return rows[0]


def query_disk(client: WMIClient) -> Any:
    """
    Query the first physical disk.
    """

    rows = client.query(
        """
        SELECT
            Size,
            Model
        FROM Win32_DiskDrive
        """
    )

    if not rows:
        raise RuntimeError(
            "No physical disk was found."
        )

    return rows[0]


def query_physical_disk_type(client: WMIClient) -> Any | None:
    """
    Query MSFT_PhysicalDisk from the Storage namespace.

    MediaType:
        3 = HDD
        4 = SSD
        5 = SCM
    """

    try:
        rows = client.query(
            """
            SELECT
                MediaType,
                FriendlyName
            FROM MSFT_PhysicalDisk
            """,
            namespace="root\\Microsoft\\Windows\\Storage",
        )

        if rows:
            return rows[0]

    except Exception:
        # This query is optional.
        # We can fall back to model-based detection.
        pass

    return None


# ============================================================
# DISK DETECTION
# ============================================================

def determine_disk_type(
    media_type: Any,
    model: str,
) -> str:
    """
    Determine HDD/SSD/SCM.

    First uses MSFT_PhysicalDisk.MediaType.
    Falls back to disk model.
    """

    try:
        media_type = int(media_type)
    except (TypeError, ValueError):
        media_type = None

    if media_type == 3:
        return "HDD"

    if media_type == 4:
        return "SSD"

    if media_type == 5:
        return "SCM"

    model = model or ""

    if re.search(
        r"SSD|NVMe|Solid|Hynix|Samsung|Micron|Kingston|"
        r"WD.*Green|WD.*Blue SN|Crucial",
        model,
        re.IGNORECASE,
    ):
        return "SSD"

    return "HDD"


def get_disk_info(client: WMIClient) -> dict[str, Any]:
    """
    Retrieve disk size and disk type.
    """

    disk = query_disk(client)

    if disk.Size is None:
        raise RuntimeError(
            "Physical disk returned no size."
        )

    size_gb = float(disk.Size) / (1024 ** 3)

    disk_size = round_size(
        size_gb,
        DISK_ROUND_STEP_GB,
    )

    model = safe_value(
        getattr(disk, "Model", None),
        "",
    )

    physical_disk = query_physical_disk_type(client)

    media_type = None

    if physical_disk is not None:
        media_type = getattr(
            physical_disk,
            "MediaType",
            None,
        )

    disk_type = determine_disk_type(
        media_type,
        str(model),
    )

    return {
        "DiskSize_GB": disk_size,
        "DiskType": disk_type,
    }


# ============================================================
# RECORD BUILDERS
# ============================================================

def create_failed_record(
    computer_name: str,
) -> dict[str, Any]:
    """
    Create a failed inventory record.
    """

    return {
        "ComputerName": computer_name,
        "Model": "N/A",
        "SerialNumber": "N/A",
        "CPU": "N/A",
        "RAM_GB": "N/A",
        "DiskSize_GB": "N/A",
        "DiskType": "N/A",
        "IPAddress": "N/A",
        "MACAddress": "N/A",
        "LoggedUser": "N/A",
        "Status": "Failed or Offline",
    }


def create_success_record(
    computer_name: str,
    system: Any,
    bios: Any,
    cpu: Any,
    network: Any,
    disk_info: dict[str, Any],
) -> dict[str, Any]:
    """
    Build the final successful inventory record.
    """

    total_memory = getattr(
        system,
        "TotalPhysicalMemory",
        None,
    )

    if total_memory is None:
        ram_gb = "N/A"
    else:
        ram_gb = round_size(
            float(total_memory) / (1024 ** 3),
            RAM_ROUND_STEP_GB,
        )

    cpu_name = safe_value(
        getattr(cpu, "Name", None)
    )

    serial_number = safe_value(
        getattr(bios, "SerialNumber", None)
    )

    model = safe_value(
        getattr(system, "Model", None)
    )

    logged_user = safe_value(
        getattr(system, "UserName", None)
    )

    ip_address = extract_ipv4(
        getattr(network, "IPAddress", None)
    )

    mac_address = safe_value(
        getattr(network, "MACAddress", None)
    )

    return {
        "ComputerName": computer_name,
        "Model": model,
        "SerialNumber": serial_number,
        "CPU": cpu_name,
        "RAM_GB": ram_gb,
        "DiskSize_GB": disk_info["DiskSize_GB"],
        "DiskType": disk_info["DiskType"],
        "IPAddress": ip_address,
        "MACAddress": mac_address,
        "LoggedUser": logged_user,
        "Status": "Success",
    }


# ============================================================
# INVENTORY OPERATIONS
# ============================================================

def perform_full_inventory(
    computer_name: str,
) -> InventoryResult:
    """
    Perform a complete inventory for one computer.

    This function runs inside a worker thread.
    """

    pythoncom.CoInitialize()

    try:

        last_error = None

        for attempt in range(MAX_RETRIES + 1):

            try:

                client = WMIClient(computer_name)
                client.connect()

                system = query_system(client)
                bios = query_bios(client)
                cpu = query_cpu(client)
                network = query_network(client)
                disk_info = get_disk_info(client)

                record = create_success_record(
                    computer_name,
                    system,
                    bios,
                    cpu,
                    network,
                    disk_info,
                )

                return InventoryResult(
                    computer_name=computer_name,
                    record=record,
                    success=True,
                )

            except Exception as exc:

                last_error = str(exc)

                if attempt < MAX_RETRIES:
                    time.sleep(RETRY_DELAY_SECONDS)

        return InventoryResult(
            computer_name=computer_name,
            record=create_failed_record(computer_name),
            success=False,
            error=last_error,
        )

    finally:
        pythoncom.CoUninitialize()


def repair_disk_info(
    computer_name: str,
    existing_record: dict[str, Any],
) -> InventoryResult:
    """
    Repair disk information for an already-successful record.

    Only disk information is queried.
    """

    pythoncom.CoInitialize()

    try:

        last_error = None

        for attempt in range(MAX_RETRIES + 1):

            try:

                client = WMIClient(computer_name)
                client.connect()

                disk_info = get_disk_info(client)

                updated_record = dict(existing_record)

                updated_record["ComputerName"] = computer_name
                updated_record["DiskSize_GB"] = disk_info[
                    "DiskSize_GB"
                ]
                updated_record["DiskType"] = disk_info[
                    "DiskType"
                ]
                updated_record["Status"] = "Success"

                return InventoryResult(
                    computer_name=computer_name,
                    record=updated_record,
                    success=True,
                    disk_only=True,
                )

            except Exception as exc:

                last_error = str(exc)

                if attempt < MAX_RETRIES:
                    time.sleep(RETRY_DELAY_SECONDS)

        return InventoryResult(
            computer_name=computer_name,
            record=existing_record,
            success=False,
            disk_only=True,
            error=last_error,
        )

    finally:
        pythoncom.CoUninitialize()


# ============================================================
# EXISTING RECORD LOGIC
# ============================================================

def inventory_is_complete(
    record: dict[str, Any] | None,
) -> bool:
    """
    Determine whether an existing record is complete.
    """

    if not record:
        return False

    if record.get("Status") != "Success":
        return False

    disk_size = record.get("DiskSize_GB")

    if is_blank(disk_size):
        return False

    if str(disk_size).strip() == "N/A":
        return False

    return True


def needs_disk_repair(
    record: dict[str, Any] | None,
) -> bool:
    """
    Determine whether an existing successful record
    needs disk information repaired.
    """

    if not record:
        return False

    if record.get("Status") != "Success":
        return False

    disk_size = record.get("DiskSize_GB")

    return (
        is_blank(disk_size)
        or str(disk_size).strip() == "N/A"
    )


# ============================================================
# EXCEL FUNCTIONS
# ============================================================

def load_input_devices(
    path: Path,
) -> list[str]:
    """
    Read device names from the input workbook.
    """

    if not path.exists():
        raise FileNotFoundError(
            f"Input file does not exist: {path}"
        )

    workbook = load_workbook(
        filename=path,
        read_only=True,
        data_only=True,
    )

    try:

        worksheet = workbook.active

        headers = [
            cell.value
            for cell in worksheet[1]
        ]

        if DEVICE_NAME_COLUMN not in headers:
            raise ValueError(
                f"Column '{DEVICE_NAME_COLUMN}' was not found. "
                f"Available columns: {headers}"
            )

        column_index = headers.index(
            DEVICE_NAME_COLUMN
        )

        devices: list[str] = []

        for row in worksheet.iter_rows(
            min_row=2,
            values_only=True,
        ):

            if column_index >= len(row):
                continue

            device_name = normalize_device_name(
                row[column_index]
            )

            if device_name:
                devices.append(device_name)

        return devices

    finally:
        workbook.close()


def load_existing_inventory(
    path: Path,
) -> list[dict[str, Any]]:
    """
    Load existing inventory records.

    Returns an empty list if:
    - The report does not exist.
    - The worksheet doesn't exist.
    - The worksheet is empty.
    """

    if not path.exists():
        return []

    try:

        workbook = load_workbook(
            filename=path,
            read_only=True,
            data_only=True,
        )

    except Exception as exc:

        logger.warning(
            "Could not open existing inventory: %s",
            exc,
        )

        return []

    try:

        if WORKSHEET_NAME not in workbook.sheetnames:
            logger.warning(
                "Worksheet '%s' does not exist. "
                "Starting with empty inventory.",
                WORKSHEET_NAME,
            )

            return []

        worksheet = workbook[WORKSHEET_NAME]

        rows = worksheet.iter_rows(
            values_only=True
        )

        try:
            headers = list(next(rows))
        except StopIteration:
            return []

        records: list[dict[str, Any]] = []

        for row in rows:

            record = {}

            for index, header in enumerate(headers):

                if header is None:
                    continue

                value = (
                    row[index]
                    if index < len(row)
                    else None
                )

                record[str(header)] = value

            if any(
                value is not None
                for value in record.values()
            ):
                records.append(record)

        return records

    finally:
        workbook.close()


def save_inventory(
    path: Path,
    records: list[dict[str, Any]],
) -> None:
    """
    Save inventory records to Excel.

    Writes to a temporary file first and replaces the
    destination after successful save.
    """

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    workbook = Workbook()

    try:

        worksheet = workbook.active
        worksheet.title = WORKSHEET_NAME

        # Headers
        for column, header in enumerate(
            HEADERS,
            start=1,
        ):
            worksheet.cell(
                row=1,
                column=column,
                value=header,
            )

        # Records
        for row_number, record in enumerate(
            records,
            start=2,
        ):

            for column, header in enumerate(
                HEADERS,
                start=1,
            ):

                worksheet.cell(
                    row=row_number,
                    column=column,
                    value=record.get(header),
                )

        # Freeze first row.
        worksheet.freeze_panes = "A2"

        # Bold header.
        for cell in worksheet[1]:
            cell.font = cell.font.copy(
                bold=True
            )

        # Auto-size columns.
        for column_cells in worksheet.columns:

            max_length = 0

            for cell in column_cells:

                if cell.value is None:
                    continue

                length = len(str(cell.value))

                if length > max_length:
                    max_length = length

            column_letter = column_cells[0].column_letter

            worksheet.column_dimensions[
                column_letter
            ].width = min(
                max(max_length + 2, 10),
                60,
            )

        # Save to a temporary file in the same directory.
        fd, temporary_name = tempfile.mkstemp(
            suffix=".xlsx",
            prefix=".inventory_",
            dir=path.parent,
        )

        os.close(fd)

        temporary_path = Path(temporary_name)

        try:

            workbook.save(temporary_path)

            # Replace existing report.
            os.replace(
                temporary_path,
                path,
            )

        finally:

            if temporary_path.exists():
                try:
                    temporary_path.unlink()
                except OSError:
                    pass

    finally:
        workbook.close()


# ============================================================
# RESULT INDEX
# ============================================================

def build_result_index(
    records: list[dict[str, Any]],
) -> dict[str, int]:
    """
    Build ComputerName -> record index.
    """

    index: dict[str, int] = {}

    for position, record in enumerate(records):

        computer_name = normalize_device_name(
            record.get("ComputerName")
        )

        if not computer_name:
            continue

        if computer_name not in index:
            index[computer_name] = position

    return index


def update_result(
    records: list[dict[str, Any]],
    index: dict[str, int],
    record: dict[str, Any],
) -> None:
    """
    Add or replace a record.
    """

    computer_name = normalize_device_name(
        record.get("ComputerName")
    )

    if not computer_name:
        raise ValueError(
            "Cannot store inventory record without ComputerName."
        )

    if computer_name in index:

        records[index[computer_name]] = record

    else:

        index[computer_name] = len(records)
        records.append(record)


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    start_time = time.perf_counter()

    logger.info("=" * 70)
    logger.info("Starting Device Inventory")
    logger.info("=" * 70)

    # --------------------------------------------------------
    # Load devices
    # --------------------------------------------------------

    try:

        devices = load_input_devices(
            INPUT_XLSX
        )

    except Exception as exc:

        logger.error(
            "Failed to load input devices: %s",
            exc,
        )

        return 1

    if not devices:

        logger.error(
            "Input workbook contains no devices."
        )

        return 1

    logger.info(
        "Loaded %d devices from input workbook.",
        len(devices),
    )

    # --------------------------------------------------------
    # Load existing results
    # --------------------------------------------------------

    records = load_existing_inventory(
        OUTPUT_XLSX
    )

    result_index = build_result_index(
        records
    )

    logger.info(
        "Loaded %d existing inventory records.",
        len(records),
    )

    logger.info(
        "Using %d worker threads.",
        MAX_WORKERS,
    )

    # --------------------------------------------------------
    # Determine work
    # --------------------------------------------------------

    full_inventory_devices: list[str] = []
    disk_repair_devices: list[
        tuple[str, dict[str, Any]]
    ] = []

    skipped = 0

    for device_name in devices:

        if device_name in result_index:

            existing_record = records[
                result_index[device_name]
            ]

            if inventory_is_complete(
                existing_record
            ):

                logger.info(
                    "%s -> already complete; skipping.",
                    device_name,
                )

                skipped += 1

                continue

            if needs_disk_repair(
                existing_record
            ):

                logger.info(
                    "%s -> successful record "
                    "missing disk information; "
                    "scheduling disk repair.",
                    device_name,
                )

                disk_repair_devices.append(
                    (
                        device_name,
                        existing_record,
                    )
                )

                continue

        # New device or previous failure.
        full_inventory_devices.append(
            device_name
        )

    logger.info(
        "Full inventory jobs : %d",
        len(full_inventory_devices),
    )

    logger.info(
        "Disk repair jobs     : %d",
        len(disk_repair_devices),
    )

    logger.info(
        "Skipped jobs         : %d",
        skipped,
    )

    # --------------------------------------------------------
    # Run concurrent jobs
    # --------------------------------------------------------

    successful = 0
    failed = 0
    disk_updated = 0

    futures: dict[
        Future[InventoryResult],
        str,
    ] = {}

    with ThreadPoolExecutor(
        max_workers=MAX_WORKERS,
        thread_name_prefix="inventory",
    ) as executor:

        # Full inventory jobs.
        for device_name in full_inventory_devices:

            future = executor.submit(
                perform_full_inventory,
                device_name,
            )

            futures[future] = device_name

        # Disk repair jobs.
        for (
            device_name,
            existing_record,
        ) in disk_repair_devices:

            future = executor.submit(
                repair_disk_info,
                device_name,
                existing_record,
            )

            futures[future] = device_name

        # Process results as soon as each worker finishes.
        for future in as_completed(futures):

            device_name = futures[future]

            try:

                result = future.result()

                if result.success:

                    update_result(
                        records,
                        result_index,
                        result.record,
                    )

                    if result.disk_only:

                        disk_updated += 1

                        logger.info(
                            "%s -> disk information updated.",
                            device_name,
                        )

                    else:

                        successful += 1

                        logger.info(
                            "%s -> inventory successful.",
                            device_name,
                        )

                else:

                    failed += 1

                    # For a failed disk repair, preserve
                    # the existing successful record.
                    #
                    # For a full inventory failure, update
                    # the record with the failed status.
                    if not result.disk_only:

                        update_result(
                            records,
                            result_index,
                            result.record,
                        )

                    logger.error(
                        "%s -> failed: %s",
                        device_name,
                        result.error or "Unknown error",
                    )

            except Exception as exc:

                failed += 1

                logger.exception(
                    "%s -> unexpected worker error: %s",
                    device_name,
                    exc,
                )

    # --------------------------------------------------------
    # Save results
    # --------------------------------------------------------

    logger.info(
        "Saving %d records to %s",
        len(records),
        OUTPUT_XLSX,
    )

    try:

        save_inventory(
            OUTPUT_XLSX,
            records,
        )

    except Exception as exc:

        logger.error(
            "Failed to save inventory report: %s",
            exc,
        )

        return 1

    # --------------------------------------------------------
    # Summary
    # --------------------------------------------------------

    elapsed = time.perf_counter() - start_time

    logger.info("")
    logger.info("=" * 70)
    logger.info("INVENTORY COMPLETE")
    logger.info("=" * 70)

    logger.info(
        "Total input devices : %d",
        len(devices),
    )

    logger.info(
        "Successful          : %d",
        successful,
    )

    logger.info(
        "Disk repaired       : %d",
        disk_updated,
    )

    logger.info(
        "Failed              : %d",
        failed,
    )

    logger.info(
        "Skipped             : %d",
        skipped,
    )

    logger.info(
        "Records in report   : %d",
        len(records),
    )

    logger.info(
        "Execution time      : %.2f seconds",
        elapsed,
    )

    if elapsed > 0:
        logger.info(
            "Average throughput  : %.2f devices/sec",
            len(devices) / elapsed,
        )

    logger.info(
        "Report saved to     : %s",
        OUTPUT_XLSX,
    )

    logger.info(
        "Log saved to        : %s",
        LOG_FILE,
    )

    return 0


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    raise SystemExit(main())