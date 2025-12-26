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

def add_data_arg(source: str, target: str) -> str:
    separator = ";" if sys.platform.startswith("win") else ":"
    return f"{source}{separator}{target}"


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

    exclude_modules = [
        # wxPython optional modules and extras
        "wx.py",
        "wx.svg",
        "wx.lib.agw",
        "wx.lib.pubsub",
        "wx.lib.mixins",
        "wx.lib.sized_controls",
        "wx.lib.gizmos",
        "wx.lib.plot",
        "wx.lib.floatcanvas",
    ]

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
        "--strip",
        "--distpath",
        dist_dir,
        "--workpath",
        work_dir,
        "--specpath",
        spec_dir,
        "--hidden-import",
        "serial.tools.list_ports",
        "--collect-submodules",
        "openpyxl",
        "--collect-submodules",
        "esptool",
        "--collect-data",
        "wx",
        "--collect-binaries",
        "wx",
        "--collect-data",
        "openpyxl",
        "--collect-data",
        "esptool",
        "--exclude-module",
        "openpyxl.tests",
        "--exclude-module",
        "esptool.tests",
        "--add-data",
        add_data_arg(os.path.join(project_root, "VERSION"), "."),
        os.path.join(project_root, "main.py"),
    ]
    for module in exclude_modules:
        cmd.extend(["--exclude-module", module])

    run(cmd)

    suffix = ".exe" if sys.platform.startswith("win") else ""
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    src = os.path.join(dist_dir, f"esp32_mac_monitor{suffix}")
    dst = os.path.join(
        bin_dir, f"esp32_mac_monitor_v{version}_{timestamp}{suffix}"
    )
    if os.path.isfile(dst):
        os.remove(dst)
    shutil.copy2(src, dst)


if __name__ == "__main__":
    main()
