import datetime
import os
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import wx
from serial.tools import list_ports


def ensure_gtk_resources() -> None:
    if not sys.platform.startswith("linux"):
        return
    if not getattr(sys, "frozen", False):
        return
    meipass = getattr(sys, "_MEIPASS", "")
    if not meipass:
        return
    share_dir = os.path.join(meipass, "share")
    if os.path.isdir(share_dir):
        existing = os.environ.get("XDG_DATA_DIRS", "")
        if existing:
            os.environ["XDG_DATA_DIRS"] = f"{share_dir}:{existing}"
        else:
            os.environ["XDG_DATA_DIRS"] = share_dir


def get_config_path() -> Path:
    if sys.platform.startswith("win"):
        root = os.environ.get("APPDATA", str(Path.home()))
        return Path(root) / "esp32-mac-monitor" / "config.toml"
    if sys.platform == "darwin":
        return (
            Path.home()
            / "Library"
            / "Application Support"
            / "esp32-mac-monitor"
            / "config.toml"
        )
    return Path.home() / ".config" / "esp32-mac-monitor" / "config.toml"


def load_config() -> dict:
    path = get_config_path()
    try:
        import tomllib
    except Exception:
        return {}
    try:
        with path.open("rb") as handle:
            return tomllib.load(handle)
    except FileNotFoundError:
        return {}
    except Exception:
        return {}


def save_config(data: dict) -> None:
    path = get_config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = []
    for key, value in data.items():
        if isinstance(value, bool):
            lines.append(f"{key} = {'true' if value else 'false'}")
        elif isinstance(value, int):
            lines.append(f"{key} = {value}")
        elif isinstance(value, str):
            escaped = value.replace("\\", "\\\\").replace('"', '\\"')
            lines.append(f'{key} = "{escaped}"')
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def default_max_workers() -> int:
    count = os.cpu_count() or 4
    count = max(2, count)
    gil_enabled = getattr(sys, "_is_gil_enabled", None)
    if callable(gil_enabled) and not gil_enabled():
        return min(16, count * 2)
    return min(8, count)


def make_check_bitmap(size: int, checked: bool) -> wx.Bitmap:
    bmp = wx.Bitmap(size, size)
    dc = wx.MemoryDC(bmp)
    bg = wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW)
    dc.SetBackground(wx.Brush(bg))
    dc.Clear()

    border = wx.Colour(30, 136, 229) if checked else wx.Colour(156, 163, 175)
    fill = wx.Colour(30, 136, 229) if checked else bg
    dc.SetPen(wx.Pen(border, 1))
    dc.SetBrush(wx.Brush(fill))
    dc.DrawRoundedRectangle(1, 1, size - 2, size - 2, 2)

    if checked:
        dc.SetPen(wx.Pen(wx.Colour(255, 255, 255), 2))
        dc.DrawLine(3, size // 2, size // 2, size - 4)
        dc.DrawLine(size // 2 - 1, size - 4, size - 3, 3)

    dc.SelectObject(wx.NullBitmap)
    return bmp


def make_arrow_bitmap(size: int) -> wx.Bitmap:
    bmp = wx.Bitmap(size, size)
    dc = wx.MemoryDC(bmp)
    bg = wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW)
    dc.SetBackground(wx.Brush(bg))
    dc.Clear()

    dc.SetPen(wx.Pen(wx.Colour(107, 114, 128), 1))
    dc.SetBrush(wx.Brush(wx.Colour(107, 114, 128)))
    center = size // 2
    points = [
        (center - 3, center - 1),
        (center + 3, center - 1),
        (center, center + 3),
    ]
    dc.DrawPolygon(points)

    dc.SelectObject(wx.NullBitmap)
    return bmp


def format_mac(value: object) -> str:
    if isinstance(value, (bytes, bytearray)):
        raw = value
        return ":".join(f"{byte:02x}" for byte in raw)
    if isinstance(value, (list, tuple)):
        try:
            raw = bytes(value)
            return ":".join(f"{byte:02x}" for byte in raw)
        except Exception:
            return str(value).lower()
    if isinstance(value, str):
        text = value.strip().lower()
        if len(text) == 12 and all(c in "0123456789abcdef" for c in text):
            return ":".join(text[i : i + 2] for i in range(0, 12, 2))
        return text
    return str(value).lower()


def close_esp_port(esp: object) -> None:
    port = getattr(esp, "_port", None)
    if port is None:
        return
    try:
        port.close()
    except Exception:
        pass


def read_mac_via_esptool(port: str) -> tuple[str, str]:
    try:
        import esptool
    except Exception as exc:
        return "", f"import error: {exc}"

    try:
        if hasattr(esptool, "detect_chip"):
            esp = esptool.detect_chip(port=port, baud=115200)
        elif hasattr(esptool, "ESPLoader") and hasattr(esptool.ESPLoader, "detect_chip"):
            esp = esptool.ESPLoader.detect_chip(port=port, baud=115200)
        else:
            return "", "esptool api not found"

        if hasattr(esp, "connect"):
            esp.connect()

        mac_raw = esp.read_mac()
        mac = format_mac(mac_raw)
        close_esp_port(esp)

        if not mac:
            return "", "mac not found"
        return mac, "ok"
    except Exception as exc:
        return "", f"error: {exc}"


class MainFrame(wx.Frame):
    def __init__(self, version: str) -> None:
        title = f"ESP32 MAC 监测工具 v{version}"
        super().__init__(None, title=title, size=(860, 520))
        panel = wx.Panel(self)
        self.version = version
        self.config = load_config()

        self.start_button = wx.Button(panel, label="开始")
        self.stop_button = wx.Button(panel, label="停止")
        self.export_button = wx.Button(panel, label="导出 Excel")
        self.clear_button = wx.Button(panel, label="清除所有")
        self.remove_failed_button = wx.Button(panel, label="清除无用数据")
        self.dedup_button = wx.Button(panel, label="清除重复")

        self.search_input = wx.SearchCtrl(panel, style=wx.TE_PROCESS_ENTER)
        self.search_input.SetHint("搜索：串口 / MAC / 状态")
        self.status_filter_value = "全部"
        self.status_filter = wx.Button(panel, label=self.status_filter_label())
        self.status_filter.SetMinSize((-1, -1))
        self.status_bar = self.CreateStatusBar(2)
        self.status_bar.SetStatusWidths([-1, 140])
        self.status_bar.SetStatusText("空闲", 0)
        self.status_bar.SetStatusText(f"v{version}", 1)

        self.list_ctrl = wx.ListCtrl(
            panel,
            style=wx.LC_REPORT | wx.BORDER_SUNKEN,
        )
        self.list_ctrl.InsertColumn(0, "时间", width=160)
        self.list_ctrl.InsertColumn(1, "串口", width=120)
        self.list_ctrl.InsertColumn(2, "MAC", width=200)
        self.list_ctrl.InsertColumn(3, "状态", width=220)

        monitor_box = wx.StaticBox(panel, label="监测")
        monitor_sizer = wx.StaticBoxSizer(monitor_box, wx.HORIZONTAL)
        monitor_sizer.Add(self.start_button, 0, wx.ALL, 6)
        monitor_sizer.Add(self.stop_button, 0, wx.ALL, 6)

        export_box = wx.StaticBox(panel, label="导出")
        export_sizer = wx.StaticBoxSizer(export_box, wx.HORIZONTAL)
        export_sizer.Add(self.export_button, 0, wx.ALL, 6)
        self.export_mac_only_toggle = wx.ToggleButton(panel, label="仅导出 MAC")
        export_sizer.Add(
            self.export_mac_only_toggle, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 6
        )

        clean_box = wx.StaticBox(panel, label="清理")
        clean_sizer = wx.StaticBoxSizer(clean_box, wx.HORIZONTAL)
        clean_sizer.Add(self.clear_button, 0, wx.ALL, 6)
        clean_sizer.Add(self.remove_failed_button, 0, wx.ALL, 6)
        clean_sizer.Add(self.dedup_button, 0, wx.ALL, 6)

        group_row = wx.BoxSizer(wx.HORIZONTAL)
        group_row.Add(monitor_sizer, 0, wx.RIGHT, 10)
        group_row.Add(export_sizer, 0, wx.RIGHT, 10)
        group_row.Add(clean_sizer, 0, wx.RIGHT, 10)
        group_row.AddStretchSpacer(1)

        self.search_input.SetMinSize((260, -1))
        filter_row = wx.FlexGridSizer(rows=1, cols=3, vgap=0, hgap=8)
        filter_row.Add(wx.StaticText(panel, label="过滤"), 0, wx.ALIGN_CENTER_VERTICAL)
        filter_row.Add(self.search_input, 1, wx.EXPAND)
        filter_row.Add(self.status_filter, 0, wx.EXPAND)
        filter_row.AddGrowableCol(1, 1)

        main_sizer = wx.BoxSizer(wx.VERTICAL)
        main_sizer.Add(group_row, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, 12)
        main_sizer.Add(filter_row, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, 8)
        main_sizer.Add(self.list_ctrl, 1, wx.EXPAND | wx.ALL, 12)
        panel.SetSizer(main_sizer)
        main_sizer.Fit(self)
        self.SetMinSize(self.GetSize())
        self.SetSize((860, 520))
        self.Layout()
        wx.CallAfter(self.SendSizeEvent)

        self.stop_button.Disable()

        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.on_timer, self.timer)

        self.executor = None
        self.scan_inflight = False
        self.known_ports: set[str] = set()
        self.pending_ports: set[str] = set()
        self.rows: list[dict[str, str]] = []
        self.export_mac_only = bool(self.config.get("export_mac_only", False))
        self.export_mac_only_toggle.SetValue(self.export_mac_only)
        self.update_export_toggle_label()
        self.restore_status_filter()
        self.init_custom_icons()

        self.start_button.Bind(wx.EVT_BUTTON, self.start_monitoring)
        self.stop_button.Bind(wx.EVT_BUTTON, self.stop_monitoring)
        self.export_button.Bind(wx.EVT_BUTTON, self.export_excel)
        self.status_filter.Bind(wx.EVT_BUTTON, self.show_status_menu)
        self.export_mac_only_toggle.Bind(
            wx.EVT_TOGGLEBUTTON, self.on_export_mac_only_toggle
        )
        self.clear_button.Bind(wx.EVT_BUTTON, self.clear_table)
        self.remove_failed_button.Bind(wx.EVT_BUTTON, self.remove_failed_rows)
        self.dedup_button.Bind(wx.EVT_BUTTON, self.remove_duplicate_rows)
        self.search_input.Bind(wx.EVT_TEXT, self.apply_filters)
        self.Bind(wx.EVT_CLOSE, self.on_close)

    def start_monitoring(self, _event: wx.CommandEvent) -> None:
        self.ensure_executor()
        self.status_bar.SetStatusText("监测中...", 0)
        self.start_button.Disable()
        self.stop_button.Enable()
        self.timer.Start(1000)

    def stop_monitoring(self, _event: wx.CommandEvent) -> None:
        self.timer.Stop()
        self.status_bar.SetStatusText("已停止", 0)
        self.start_button.Enable()
        self.stop_button.Disable()

    def on_timer(self, _event: wx.TimerEvent) -> None:
        if self.scan_inflight:
            return
        self.scan_inflight = True
        self.ensure_executor()
        future = self.executor.submit(self.scan_ports)
        future.add_done_callback(lambda fut: wx.CallAfter(self.on_scan_result, fut))

    def scan_ports(self) -> set[str]:
        return {port.device for port in list_ports.comports()}

    def on_scan_result(self, future) -> None:
        self.scan_inflight = False
        try:
            current_ports = future.result()
        except Exception:
            current_ports = set()

        removed = self.known_ports - current_ports
        for port in removed:
            self.known_ports.discard(port)
            self.pending_ports.discard(port)

        new_ports = current_ports - self.known_ports
        for port in sorted(new_ports):
            self.known_ports.add(port)
            if port in self.pending_ports:
                continue
            self.pending_ports.add(port)
            self.ensure_executor()
            task = self.executor.submit(read_mac_via_esptool, port)
            task.add_done_callback(lambda fut, p=port: self.on_mac_result(p, fut))

    def on_mac_result(self, port: str, future) -> None:
        try:
            mac, status = future.result()
        except Exception as exc:
            mac, status = "", f"error: {exc}"

        def update_ui() -> None:
            self.pending_ports.discard(port)
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            row_data = {
                "time": timestamp,
                "port": port,
                "mac": mac,
                "status": status,
            }
            self.rows.append(row_data)
            self.apply_filters()
            count = self.list_ctrl.GetItemCount()
            if count > 0:
                self.list_ctrl.EnsureVisible(count - 1)

        wx.CallAfter(update_ui)

    def apply_filters(self, _event: wx.CommandEvent | None = None) -> None:
        query = self.search_input.GetValue().strip().lower()
        status_choice = self.status_filter_value
        self.config["status_filter"] = status_choice
        save_config(self.config)

        self.list_ctrl.Freeze()
        self.list_ctrl.DeleteAllItems()

        for row in self.rows:
            values = [row["time"], row["port"], row["mac"], row["status"]]
            row_text = " ".join(values).lower()
            if query and query not in row_text:
                continue

            if status_choice == "成功" and row["status"] != "ok":
                continue
            if status_choice == "失败" and row["status"] == "ok":
                continue

            index = self.list_ctrl.InsertItem(self.list_ctrl.GetItemCount(), row["time"])
            self.list_ctrl.SetItem(index, 1, row["port"])
            self.list_ctrl.SetItem(index, 2, row["mac"])
            self.list_ctrl.SetItem(index, 3, row["status"])

        self.list_ctrl.Thaw()

    def clear_table(self, _event: wx.CommandEvent) -> None:
        self.rows.clear()
        self.apply_filters()

    def remove_failed_rows(self, _event: wx.CommandEvent) -> None:
        self.rows = [row for row in self.rows if row["status"] == "ok"]
        self.apply_filters()

    def remove_duplicate_rows(self, _event: wx.CommandEvent) -> None:
        seen: set[str] = set()
        deduped: list[dict[str, str]] = []
        for row in self.rows:
            mac = row.get("mac", "")
            if not mac:
                deduped.append(row)
                continue
            if mac in seen:
                continue
            seen.add(mac)
            deduped.append(row)
        self.rows = deduped
        self.apply_filters()

    def export_excel(self, _event: wx.CommandEvent) -> None:
        if not self.rows:
            wx.MessageBox("没有可导出的数据。", "导出", wx.OK | wx.ICON_INFORMATION)
            return

        mac_only = self.export_mac_only_toggle.GetValue()
        self.export_mac_only = mac_only
        self.config["export_mac_only"] = self.export_mac_only
        save_config(self.config)

        dialog = wx.FileDialog(
            self,
            message="保存 Excel",
            wildcard="Excel Files (*.xlsx)|*.xlsx",
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
        )
        if dialog.ShowModal() != wx.ID_OK:
            return
        path = dialog.GetPath()
        if not path.lower().endswith(".xlsx"):
            path = f"{path}.xlsx"

        try:
            import openpyxl
        except Exception as exc:
            wx.MessageBox(f"openpyxl 导入失败: {exc}", "导出", wx.OK | wx.ICON_WARNING)
            return

        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "ESP32 MAC"
        if mac_only:
            for row in self.rows:
                mac_value = row.get("mac", "")
                if not mac_value:
                    continue
                sheet.append([mac_value])
        else:
            sheet.append(["时间", "串口", "MAC", "状态"])
            for row in self.rows:
                sheet.append([row["time"], row["port"], row["mac"], row["status"]])

        try:
            workbook.save(path)
        except Exception as exc:
            wx.MessageBox(f"保存失败: {exc}", "导出", wx.OK | wx.ICON_ERROR)
            return

        wx.MessageBox(f"已保存到: {path}", "导出", wx.OK | wx.ICON_INFORMATION)

    def on_export_mac_only_toggle(self, _event: wx.CommandEvent) -> None:
        self.export_mac_only = self.export_mac_only_toggle.GetValue()
        self.config["export_mac_only"] = self.export_mac_only
        save_config(self.config)
        self.update_export_toggle_label()

    def restore_status_filter(self) -> None:
        value = self.config.get("status_filter")
        if value in ("全部", "成功", "失败"):
            self.status_filter_value = value
            self.status_filter.SetLabel(self.status_filter_label())

    def status_filter_label(self) -> str:
        return f"状态: {self.status_filter_value}"

    def show_status_menu(self, _event: wx.CommandEvent) -> None:
        menu = wx.Menu()
        for choice in ("全部", "成功", "失败"):
            item = menu.AppendRadioItem(wx.ID_ANY, choice)
            if choice == self.status_filter_value:
                item.Check(True)
            self.Bind(
                wx.EVT_MENU,
                lambda event, value=choice: self.set_status_filter(value),
                item,
            )
        self.PopupMenu(menu)
        menu.Destroy()

    def set_status_filter(self, value: str) -> None:
        self.status_filter_value = value
        self.status_filter.SetLabel(self.status_filter_label())
        self.config["status_filter"] = self.status_filter_value
        save_config(self.config)
        self.apply_filters()

    def update_export_toggle_label(self) -> None:
        self.export_mac_only_toggle.SetLabel("仅导出 MAC")
        if hasattr(self, "check_on_bmp"):
            bmp = self.check_on_bmp if self.export_mac_only_toggle.GetValue() else self.check_off_bmp
            self.export_mac_only_toggle.SetBitmap(bmp)
            self.export_mac_only_toggle.SetBitmapPosition(wx.LEFT)

    def init_custom_icons(self) -> None:
        self.check_on_bmp = make_check_bitmap(14, True)
        self.check_off_bmp = make_check_bitmap(14, False)
        self.arrow_bmp = make_arrow_bitmap(12)
        self.status_filter.SetBitmap(self.arrow_bmp)
        self.status_filter.SetBitmapPosition(wx.RIGHT)
        self.update_export_toggle_label()

    def ensure_executor(self) -> None:
        if self.executor is None:
            self.executor = ThreadPoolExecutor(
                max_workers=default_max_workers(),
                thread_name_prefix="esp32-mac",
            )

    def on_close(self, event: wx.CloseEvent) -> None:
        save_config(self.config)
        self.Destroy()

    def Destroy(self) -> bool:  # noqa: N802
        if self.timer.IsRunning():
            self.timer.Stop()
        if self.executor is not None:
            self.executor.shutdown(wait=False)
        return super().Destroy()


def load_version() -> str:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    meipass = getattr(sys, "_MEIPASS", None)
    candidates = [base_dir]
    if meipass:
        candidates.insert(0, meipass)
    try:
        for root in candidates:
            version_path = os.path.join(root, "VERSION")
            if not os.path.isfile(version_path):
                continue
            with open(version_path, "r", encoding="utf-8") as handle:
                version = handle.readline().strip()
                if version:
                    return version
    except OSError:
        pass
    return "0.0.0"


def main() -> None:
    version = load_version()
    ensure_gtk_resources()
    app = wx.App()
    frame = MainFrame(version)
    frame.Show()
    app.MainLoop()


if __name__ == "__main__":
    main()
