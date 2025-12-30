"""
USB Refresher
=============

How it works:
- Verifies Administrator privileges on Windows.
- Locates adb.exe and devcon.exe (using provided paths or PATH).
- Runs `adb devices` to determine whether the target device is healthy.
- If unhealthy, performs a soft reset (restart ADB server + reconnect) and rechecks.
- If still unhealthy, locates the USB device via devcon by name or VID/PID,
  performs a disable/enable cycle, then restarts ADB and polls until healthy.

How to run:
    python usb_refresher.py --adb-path C:\\Android\\platform-tools\\adb.exe --devcon-path C:\\devcon.exe

Required privileges:
- Windows Administrator privileges are required to use devcon to disable/enable devices.
  The script exits with code 3 if it is not running as Administrator.
"""

import argparse
import ctypes
import logging
import os
import re
import shutil
import subprocess
import sys
import time


ADB_HEALTHY_STATE = "device"
ADB_SOFT_RESET_COMMANDS = [
    ["kill-server"],
    ["start-server"],
    ["reconnect"],
]
ANDROID_DEVICE_NAME = "Android Composite ADB Interface"
COMMON_ANDROID_VIDS = {
    "18D1",  # Google
    "0BB4",  # HTC
    "12D1",  # Huawei
    "04E8",  # Samsung
    "22B8",  # Motorola
    "2A70",  # OnePlus
    "0FCE",  # Sony
    "0502",  # Acer
    "05C6",  # Qualcomm
}


class CommandError(Exception):
    pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Refresh ADB USB devices using devcon.")
    parser.add_argument("--adb-path", default="adb", help="Path to adb.exe (or adb on PATH).")
    parser.add_argument("--devcon-path", default="devcon", help="Path to devcon.exe (or devcon on PATH).")
    parser.add_argument("--timeout", type=int, default=30, help="Seconds to wait for recovery per phase.")
    parser.add_argument("--serial", help="ADB device serial to target.")
    parser.add_argument("--dry-run", action="store_true", help="Log actions without executing devcon changes.")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose logging.")
    return parser.parse_args()


def configure_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(level=level, format="%(levelname)s: %(message)s")


def is_windows_admin() -> bool:
    if sys.platform != "win32":
        return True
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except OSError:
        return False


def resolve_executable(path: str) -> str | None:
    if os.path.sep in path or (os.path.altsep and os.path.altsep in path):
        if os.path.isfile(path):
            return path
        return None
    resolved = shutil.which(path)
    if resolved:
        return resolved
    return None


def run_command(command: list[str], timeout: int | None = None) -> subprocess.CompletedProcess:
    logging.debug("Running command: %s", " ".join(command))
    try:
        return subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise CommandError(f"Command timed out: {' '.join(command)}") from exc


def adb_command(adb_path: str, args: list[str], timeout: int | None = None) -> subprocess.CompletedProcess:
    return run_command([adb_path, *args], timeout=timeout)


def parse_adb_devices(output: str, serial: str | None) -> str | None:
    lines = [line.strip() for line in output.splitlines() if line.strip()]
    if not lines:
        return None
    for line in lines:
        if line.lower().startswith("list of devices"):
            continue
        parts = line.split()
        if len(parts) < 2:
            continue
        device_serial, state = parts[0], parts[1]
        if serial and device_serial != serial:
            continue
        return state
    return None


def is_adb_healthy(adb_path: str, serial: str | None) -> bool:
    result = adb_command(adb_path, ["devices"])
    if result.returncode != 0:
        logging.warning("adb devices failed: %s", result.stderr.strip())
        return False
    state = parse_adb_devices(result.stdout, serial)
    if state is None:
        logging.info("No ADB device found.")
        return False
    if state == ADB_HEALTHY_STATE:
        logging.info("ADB device is healthy (%s).", state)
        return True
    logging.warning("ADB device unhealthy (%s).", state)
    return False


def soft_reset(adb_path: str) -> None:
    for args in ADB_SOFT_RESET_COMMANDS:
        result = adb_command(adb_path, args)
        if result.returncode != 0:
            logging.warning("adb %s failed: %s", " ".join(args), result.stderr.strip())


def parse_devcon_findall(output: str) -> list[tuple[str, str]]:
    devices = []
    for line in output.splitlines():
        if ":" not in line:
            continue
        instance_id, name = line.split(":", 1)
        devices.append((instance_id.strip(), name.strip()))
    return devices


def parse_devcon_hwids(output: str) -> list[dict[str, object]]:
    devices: list[dict[str, object]] = []
    current: dict[str, object] | None = None
    for line in output.splitlines():
        if not line.strip():
            if current:
                devices.append(current)
                current = None
            continue
        if ":" in line and not line.startswith(" "):
            instance_id, name = line.split(":", 1)
            current = {"id": instance_id.strip(), "name": name.strip(), "hwids": []}
            continue
        if current is not None:
            match = re.search(r"USB\\VID_[0-9A-Fa-f]{4}&PID_[0-9A-Fa-f]{4}", line)
            if match:
                current["hwids"].append(match.group(0).upper())
    if current:
        devices.append(current)
    return devices


def find_devcon_device(devcon_path: str) -> str | None:
    findall = run_command([devcon_path, "findall", "=usb"])
    if findall.returncode == 0:
        devices = parse_devcon_findall(findall.stdout)
        for instance_id, name in devices:
            if ANDROID_DEVICE_NAME.lower() in name.lower():
                logging.info("Matched device by name: %s", name)
                return instance_id

    hwids_result = run_command([devcon_path, "hwids", "=usb"])
    if hwids_result.returncode != 0:
        logging.error("devcon hwids failed: %s", hwids_result.stderr.strip())
        return None
    for device in parse_devcon_hwids(hwids_result.stdout):
        for hwid in device.get("hwids", []):
            vid_match = re.search(r"VID_([0-9A-F]{4})", hwid)
            if vid_match and vid_match.group(1) in COMMON_ANDROID_VIDS:
                logging.info("Matched device by VID/PID: %s", hwid)
                return str(device["id"])
    return None


def hard_reset(devcon_path: str, instance_id: str, dry_run: bool) -> bool:
    if dry_run:
        logging.info("Dry run: would disable %s", instance_id)
        logging.info("Dry run: would enable %s", instance_id)
        return True
    disable = run_command([devcon_path, "disable", instance_id])
    if disable.returncode != 0:
        logging.error("devcon disable failed: %s", disable.stderr.strip())
        return False
    time.sleep(2)
    enable = run_command([devcon_path, "enable", instance_id])
    if enable.returncode != 0:
        logging.error("devcon enable failed: %s", enable.stderr.strip())
        return False
    return True


def poll_until_healthy(adb_path: str, serial: str | None, timeout: int) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        if is_adb_healthy(adb_path, serial):
            return True
        time.sleep(2)
    return False


def main() -> int:
    args = parse_args()
    configure_logging(args.verbose)

    if not is_windows_admin():
        logging.error("Administrator privileges are required on Windows.")
        return 3

    adb_path = resolve_executable(args.adb_path)
    devcon_path = resolve_executable(args.devcon_path)
    if not adb_path:
        logging.error("adb.exe not found at %s", args.adb_path)
        return 2
    if not devcon_path:
        logging.error("devcon.exe not found at %s", args.devcon_path)
        return 2

    if is_adb_healthy(adb_path, args.serial):
        return 0

    logging.info("Attempting soft reset of ADB server.")
    soft_reset(adb_path)
    if poll_until_healthy(adb_path, args.serial, args.timeout):
        return 0

    logging.warning("Soft reset did not recover device; attempting hard reset.")
    instance_id = find_devcon_device(devcon_path)
    if not instance_id:
        logging.error("Unable to locate Android USB device for hard reset.")
        return 1

    if not hard_reset(devcon_path, instance_id, args.dry_run):
        return 1

    soft_reset(adb_path)
    if poll_until_healthy(adb_path, args.serial, args.timeout):
        return 0

    logging.error("Device did not recover before timeout.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
