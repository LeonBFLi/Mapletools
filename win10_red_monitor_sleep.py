import threading
import time
import tkinter as tk
from dataclasses import dataclass
from tkinter import messagebox, ttk
from typing import Optional

try:
    from PIL import ImageGrab
except ImportError:  # pragma: no cover
    ImageGrab = None


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
        self.root.title("Win10 红点监控自动最小化")
        self.root.geometry("520x290")

        self.region = None
        self.running = False
        self.worker = None
        self.last_trigger = 0.0

        self.red_threshold = tk.IntVar(value=220)
        self.delta_threshold = tk.IntVar(value=35)
        self.check_interval = tk.IntVar(value=120)
        self.cooldown = tk.IntVar(value=8)
        self.min_red_pixels = tk.IntVar(value=1)
        self.alert_showing = False

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

        ttk.Label(container, text="最少红像素数量").grid(row=2, column=0, sticky="w", pady=8)
        ttk.Entry(container, textvariable=self.min_red_pixels, width=8).grid(
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
        ttk.Button(btn_row, text="立即测试弹窗", command=self.trigger_minimize).pack(side=tk.LEFT)

        self.status = ttk.Label(container, text="状态：待机")
        self.status.grid(row=5, column=0, columnspan=4, sticky="w", pady=(10, 0))

        tips = (
            "说明：只要监控区域中出现满足阈值的红色像素，"
            "程序会弹出最高优先级提示框（含“确定”按钮）。"
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
        while self.running:
            red_count, ratio = self._get_red_stats(self.region)
            now = time.time()
            cooling = (now - self.last_trigger) < self.cooldown.get()
            has_enough_red = red_count >= max(1, self.min_red_pixels.get())

            if has_enough_red and not cooling:
                self.last_trigger = now
                self.root.after(
                    0,
                    lambda current_ratio=ratio, current_count=red_count: self._on_detected(
                        current_count, current_ratio
                    ),
                )
            time.sleep(max(0.02, self.check_interval.get() / 1000.0))

    def _on_detected(self, red_count: int, ratio: float):
        self.status.config(text=f"状态：检测到红点 {red_count}px，占比 {ratio:.4f}")
        self.trigger_alert()

    def _get_red_stats(self, region: Region) -> tuple[int, float]:
        if ImageGrab is None:
            raise RuntimeError("缺少依赖 Pillow，请先执行: pip install pillow")
        img = ImageGrab.grab(bbox=(region.left, region.top, region.right, region.bottom))
        pixels = img.convert("RGB").getdata()

        red_th = self.red_threshold.get()
        delta = self.delta_threshold.get()

        red_count = 0
        total = region.width * region.height
        for r, g, b in pixels:
            if r >= red_th and (r - g) >= delta and (r - b) >= delta:
                red_count += 1

        return red_count, (red_count / total if total else 0.0)

    def trigger_minimize(self):
        self.trigger_alert()

    def trigger_alert(self):
        if self.alert_showing:
            return

        self.alert_showing = True
        dialog = tk.Toplevel(self.root)
        dialog.title("红点提醒")
        dialog.geometry("340x160")
        dialog.attributes("-topmost", True)
        dialog.transient(self.root)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding=18)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            frame,
            text="检测到红点，请确认。",
            anchor="center",
            justify="center",
            font=("Microsoft YaHei UI", 11),
        ).pack(fill=tk.X, pady=(8, 18))

        def on_close():
            self.alert_showing = False
            dialog.grab_release()
            dialog.destroy()

        confirm_btn = ttk.Button(frame, text="确定", command=on_close)
        confirm_btn.pack(anchor="center")

        dialog.protocol("WM_DELETE_WINDOW", on_close)
        dialog.bind("<Return>", lambda _: on_close())
        dialog.bind("<Escape>", lambda _: on_close())
        dialog.lift()
        dialog.focus_force()
        confirm_btn.focus_set()


def main():
    root: Optional[tk.Tk] = None
    try:
        if ImageGrab is None:
            raise RuntimeError("缺少依赖 Pillow，请先执行: pip install pillow")

        root = tk.Tk()
        app = RedMonitorApp(root)
        root.protocol("WM_DELETE_WINDOW", app.stop)
        root.mainloop()
    except Exception as exc:  # pragma: no cover
        _show_startup_error(str(exc), root)


def _show_startup_error(message: str, root: Optional[tk.Tk]):
    try:
        if root is None:
            temp_root = tk.Tk()
            temp_root.withdraw()
            messagebox.showerror("程序启动失败", f"{message}\n\n按回车可退出。")
            temp_root.destroy()
        else:
            messagebox.showerror("程序启动失败", f"{message}\n\n按回车可退出。")
    except Exception:
        pass

    print("\n程序启动失败：")
    print(message)
    print("\n常见原因：")
    print("1) 没有安装 Pillow")
    print("2) 不是在 Windows 图形桌面环境运行")
    input("\n按回车键退出...")


if __name__ == "__main__":
    main()
