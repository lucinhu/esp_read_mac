import datetime
import os
import sys
from concurrent.futures import ThreadPoolExecutor

import wx
from serial.tools import list_ports


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

        self.start_button = wx.Button(panel, label="开始")
        self.stop_button = wx.Button(panel, label="停止")
        self.export_button = wx.Button(panel, label="导出 Excel")
        self.clear_button = wx.Button(panel, label="清除所有")
        self.remove_failed_button = wx.Button(panel, label="清除无用数据")

        self.search_input = wx.SearchCtrl(panel, style=wx.TE_PROCESS_ENTER)
        self.search_input.SetHint("搜索：串口 / MAC / 状态")
        self.status_filter = wx.Choice(panel, choices=["全部", "成功", "失败"])
        self.status_filter.SetSelection(0)
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

        clean_box = wx.StaticBox(panel, label="清理")
        clean_sizer = wx.StaticBoxSizer(clean_box, wx.HORIZONTAL)
        clean_sizer.Add(self.clear_button, 0, wx.ALL, 6)
        clean_sizer.Add(self.remove_failed_button, 0, wx.ALL, 6)

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

        self.executor = ThreadPoolExecutor(max_workers=4)
        self.known_ports: set[str] = set()
        self.pending_ports: set[str] = set()
        self.rows: list[dict[str, str]] = []

        self.start_button.Bind(wx.EVT_BUTTON, self.start_monitoring)
        self.stop_button.Bind(wx.EVT_BUTTON, self.stop_monitoring)
        self.export_button.Bind(wx.EVT_BUTTON, self.export_excel)
        self.clear_button.Bind(wx.EVT_BUTTON, self.clear_table)
        self.remove_failed_button.Bind(wx.EVT_BUTTON, self.remove_failed_rows)
        self.search_input.Bind(wx.EVT_TEXT, self.apply_filters)
        self.status_filter.Bind(wx.EVT_CHOICE, self.apply_filters)

    def start_monitoring(self, _event: wx.CommandEvent) -> None:
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
        current_ports = {port.device for port in list_ports.comports()}

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
            future = self.executor.submit(read_mac_via_esptool, port)
            future.add_done_callback(lambda fut, p=port: self.on_mac_result(p, fut))

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
        status_choice = self.status_filter.GetStringSelection()

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

    def export_excel(self, _event: wx.CommandEvent) -> None:
        if not self.rows:
            wx.MessageBox("没有可导出的数据。", "导出", wx.OK | wx.ICON_INFORMATION)
            return

        dialog = wx.FileDialog(
            self,
            message="保存 Excel",
            wildcard="Excel Files (*.xlsx)|*.xlsx",
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
        )
        if dialog.ShowModal() != wx.ID_OK:
            return
        path = dialog.GetPath()

        try:
            import openpyxl
        except Exception as exc:
            wx.MessageBox(f"openpyxl 导入失败: {exc}", "导出", wx.OK | wx.ICON_WARNING)
            return

        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "ESP32 MAC"
        sheet.append(["时间", "串口", "MAC", "状态"])
        for row in self.rows:
            sheet.append([row["time"], row["port"], row["mac"], row["status"]])

        try:
            workbook.save(path)
        except Exception as exc:
            wx.MessageBox(f"保存失败: {exc}", "导出", wx.OK | wx.ICON_ERROR)
            return

        wx.MessageBox(f"已保存到: {path}", "导出", wx.OK | wx.ICON_INFORMATION)

    def Destroy(self) -> bool:  # noqa: N802
        if self.timer.IsRunning():
            self.timer.Stop()
        self.executor.shutdown(wait=False)
        return super().Destroy()


def load_version() -> str:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    version_path = os.path.join(base_dir, "VERSION")
    try:
        with open(version_path, "r", encoding="utf-8") as handle:
            version = handle.readline().strip()
            return version or "0.0.0"
    except OSError:
        return "0.0.0"


def main() -> None:
    version = load_version()
    app = wx.App()
    frame = MainFrame(version)
    frame.Show()
    app.MainLoop()


if __name__ == "__main__":
    main()
