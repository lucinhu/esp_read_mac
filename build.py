import datetime
import os
import shutil
import subprocess
import sys


def run(cmd: list[str]) -> None:
    completed = subprocess.run(cmd)
    if completed.returncode != 0:
        raise SystemExit(completed.returncode)


def read_version(project_root: str) -> str:
    path = os.path.join(project_root, "VERSION")
    try:
        with open(path, "r", encoding="utf-8") as handle:
            version = handle.readline().strip()
            return version or "0.0.0"
    except OSError:
        return "0.0.0"


def main() -> None:
    project_root = os.path.dirname(os.path.abspath(__file__))
    version = read_version(project_root)
    build_root = os.path.join(project_root, "build", version)
    dist_dir = os.path.join(build_root, "dist")
    work_dir = os.path.join(build_root, "work")
    spec_dir = os.path.join(build_root, "spec")
    bin_dir = os.path.join(project_root, "bin")

    if os.path.isdir(build_root):
        shutil.rmtree(build_root)
    os.makedirs(bin_dir, exist_ok=True)

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--clean",
        "--noconfirm",
        "--name",
        "esp32_mac_monitor",
        "--windowed",
        "--onefile",
        "--distpath",
        dist_dir,
        "--workpath",
        work_dir,
        "--specpath",
        spec_dir,
        "--hidden-import",
        "serial.tools.list_ports",
        "--collect-all",
        "openpyxl",
        "--collect-all",
        "esptool",
        os.path.join(project_root, "main.py"),
    ]

    run(cmd)

    suffix = ".exe" if sys.platform.startswith("win") else ""
    src = os.path.join(dist_dir, f"esp32_mac_monitor{suffix}")
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(
        bin_dir, f"esp32_mac_monitor_v{version}_{timestamp}{suffix}"
    )
    shutil.copy2(src, dst)


if __name__ == "__main__":
    main()
