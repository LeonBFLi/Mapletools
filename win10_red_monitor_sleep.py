import ctypes
import threading
import time
import tkinter as tk
from dataclasses import dataclass
from tkinter import messagebox, ttk

try:
    from PIL import ImageGrab
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "缺少依赖 Pillow，请先执行: pip install pillow"
    ) from exc


@dataclass
class Region:
    left: int
    top: int
    right: int
    bottom: int

    @property
    def width(self) -> int:
        return self.right - self.left

    @property
    def height(self) -> int:
        return self.bottom - self.top


class RegionSelector(tk.Toplevel):
    """全屏覆盖层：拖拽鼠标选择监控区域。"""

    def __init__(self, master: tk.Tk, callback):
        super().__init__(master)
        self.callback = callback
        self.start_x = 0
        self.start_y = 0
        self.rect_id = None

        self.attributes("-fullscreen", True)
        self.attributes("-alpha", 0.25)
        self.attributes("-topmost", True)
        self.configure(bg="black")

        self.canvas = tk.Canvas(self, cursor="cross", bg="gray20", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.bind("<Escape>", lambda _: self.destroy())

    def on_press(self, event):
        self.start_x, self.start_y = event.x, event.y
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = self.canvas.create_rectangle(
            self.start_x,
            self.start_y,
            self.start_x,
            self.start_y,
            outline="red",
            width=2,
        )

    def on_drag(self, event):
        if self.rect_id:
            self.canvas.coords(self.rect_id, self.start_x, self.start_y, event.x, event.y)

    def on_release(self, event):
        left = min(self.start_x, event.x)
        top = min(self.start_y, event.y)
        right = max(self.start_x, event.x)
        bottom = max(self.start_y, event.y)

        if right - left < 5 or bottom - top < 5:
            messagebox.showwarning("区域太小", "请至少选择 5x5 像素区域")
            return

        self.callback(Region(left, top, right, bottom))
        self.destroy()


class RedMonitorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Win10 红点监控自动休眠")
        self.root.geometry("520x290")

        self.region = None
        self.running = False
        self.worker = None
        self.last_trigger = 0.0

        self.red_threshold = tk.IntVar(value=220)
        self.delta_threshold = tk.IntVar(value=35)
        self.check_interval = tk.IntVar(value=120)
        self.cooldown = tk.IntVar(value=8)
        self.pixel_ratio_threshold = tk.DoubleVar(value=0.001)

        self._build_ui()

    def _build_ui(self):
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Button(container, text="选择监控区域", command=self.select_region).grid(
            row=0, column=0, sticky="w"
        )
        self.region_label = ttk.Label(container, text="尚未选择区域")
        self.region_label.grid(row=0, column=1, columnspan=3, sticky="w", padx=8)

        ttk.Label(container, text="红色阈值 (R >=)").grid(row=1, column=0, sticky="w", pady=8)
        ttk.Entry(container, textvariable=self.red_threshold, width=8).grid(
            row=1, column=1, sticky="w"
        )

        ttk.Label(container, text="红色优势 (R-G/B >=)").grid(row=1, column=2, sticky="w")
        ttk.Entry(container, textvariable=self.delta_threshold, width=8).grid(
            row=1, column=3, sticky="w"
        )

        ttk.Label(container, text="红点比例阈值").grid(row=2, column=0, sticky="w", pady=8)
        ttk.Entry(container, textvariable=self.pixel_ratio_threshold, width=8).grid(
            row=2, column=1, sticky="w"
        )

        ttk.Label(container, text="检测间隔(ms)").grid(row=2, column=2, sticky="w")
        ttk.Entry(container, textvariable=self.check_interval, width=8).grid(
            row=2, column=3, sticky="w"
        )

        ttk.Label(container, text="触发冷却(秒)").grid(row=3, column=0, sticky="w", pady=8)
        ttk.Entry(container, textvariable=self.cooldown, width=8).grid(
            row=3, column=1, sticky="w"
        )

        btn_row = ttk.Frame(container)
        btn_row.grid(row=4, column=0, columnspan=4, sticky="w", pady=10)
        ttk.Button(btn_row, text="开始监控", command=self.start).pack(side=tk.LEFT)
        ttk.Button(btn_row, text="停止监控", command=self.stop).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_row, text="立即测试休眠", command=self.trigger_sleep).pack(side=tk.LEFT)

        self.status = ttk.Label(container, text="状态：待机")
        self.status.grid(row=5, column=0, columnspan=4, sticky="w", pady=(10, 0))

        tips = (
            "说明：当监控区域中检测到红色像素比例突然变化（跳动）时，"
            "程序将发送休眠键并尝试关闭显示器。"
        )
        ttk.Label(container, text=tips, foreground="gray40", wraplength=490).grid(
            row=6, column=0, columnspan=4, sticky="w", pady=(8, 0)
        )

    def select_region(self):
        RegionSelector(self.root, self._on_region_selected)

    def _on_region_selected(self, region: Region):
        self.region = region
        self.region_label.config(
            text=f"({region.left},{region.top})-({region.right},{region.bottom}) "
            f"{region.width}x{region.height}"
        )

    def start(self):
        if self.running:
            return
        if not self.region:
            messagebox.showwarning("提示", "请先选择监控区域")
            return

        self.running = True
        self.status.config(text="状态：监控中")
        self.worker = threading.Thread(target=self.monitor_loop, daemon=True)
        self.worker.start()

    def stop(self):
        self.running = False
        self.status.config(text="状态：已停止")

    def monitor_loop(self):
        prev_ratio = None

        while self.running:
            ratio = self._get_red_ratio(self.region)
            if prev_ratio is not None:
                changed = abs(ratio - prev_ratio) >= self.pixel_ratio_threshold.get()
                has_red = ratio > 0
                now = time.time()
                cooling = (now - self.last_trigger) < self.cooldown.get()
                if changed and has_red and not cooling:
                    self.last_trigger = now
                    self.root.after(0, lambda: self.status.config(text=f"状态：触发休眠，红点比例 {ratio:.4f}"))
                    self.trigger_sleep()
            prev_ratio = ratio
            time.sleep(max(0.02, self.check_interval.get() / 1000.0))

    def _get_red_ratio(self, region: Region) -> float:
        img = ImageGrab.grab(bbox=(region.left, region.top, region.right, region.bottom))
        pixels = img.convert("RGB").getdata()

        red_th = self.red_threshold.get()
        delta = self.delta_threshold.get()

        red_count = 0
        total = region.width * region.height
        for r, g, b in pixels:
            if r >= red_th and (r - g) >= delta and (r - b) >= delta:
                red_count += 1

        return red_count / total if total else 0.0

    def trigger_sleep(self):
        self._send_sleep_key()
        self._turn_off_monitor()

    @staticmethod
    def _send_sleep_key():
        VK_SLEEP = 0x5F
        KEYEVENTF_KEYUP = 0x0002
        user32 = ctypes.windll.user32
        user32.keybd_event(VK_SLEEP, 0, 0, 0)
        user32.keybd_event(VK_SLEEP, 0, KEYEVENTF_KEYUP, 0)

    @staticmethod
    def _turn_off_monitor():
        HWND_BROADCAST = 0xFFFF
        WM_SYSCOMMAND = 0x0112
        SC_MONITORPOWER = 0xF170
        MONITOR_OFF = 2
        user32 = ctypes.windll.user32
        user32.SendMessageW(HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_OFF)


def main():
    root = tk.Tk()
    app = RedMonitorApp(root)
    root.protocol("WM_DELETE_WINDOW", app.stop)
    root.mainloop()


if __name__ == "__main__":
    main()
