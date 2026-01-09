import sys
import os

if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

import tkinter as tk
from tkinter import ttk
import speedtest
import pandas as pd
from datetime import datetime
import pystray
from PIL import Image, ImageDraw
import threading
import time


INTERVAL = 15 * 60


class SpeedTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("NetSpeed")
        self.root.geometry("600x550")
        self.root.resizable(False, False)

        try:
            if hasattr(sys, "_MEIPASS"):
                icon_path = os.path.join(sys._MEIPASS, "icon.ico")
            else:
                icon_path = "icon.ico"
            self.root.iconbitmap(icon_path)
        except:
            pass

        self.running = False
        self.data = []
        self.start_time = None

        self.create_widgets()
        self.center_window()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def create_widgets(self):
        main_frame = tk.Frame(self.root)
        main_frame.pack(expand=True, fill="both", padx=20, pady=20)
        main_frame.columnconfigure(0, weight=1)

        button_frame = tk.Frame(main_frame)
        button_frame.grid(row=0, column=0, pady=10, sticky="ew")
        button_frame.columnconfigure(0, weight=1)

        self.start_btn = tk.Button(
            button_frame,
            text="Запустить",
            command=self.start,
            bg="#d4edda",
            width=25,
            height=2,
        )
        self.start_btn.grid(row=0, column=0, pady=5, padx=(0, 18))

        self.stop_btn = tk.Button(
            button_frame,
            text="Остановить",
            command=self.stop,
            state=tk.DISABLED,
            bg="#ffcccc",
            width=25,
            height=2,
        )
        self.stop_btn.grid(row=1, column=0, pady=5, padx=(0, 18))

        tree_frame = tk.Frame(main_frame)
        tree_frame.grid(row=2, column=0, sticky="nsew", pady=10)

        columns = ("date", "time", "speed")
        self.tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings", height=10
        )
        scrollbar = ttk.Scrollbar(
            tree_frame, orient="vertical", command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", expand=True, fill="both")

        col_w = 150
        self.tree.heading("date", text="Дата")
        self.tree.heading("time", text="Время")
        self.tree.heading("speed", text="Скорость (Мбит/с)")
        self.tree.column("date", anchor="center", width=col_w)
        self.tree.column("time", anchor="center", width=col_w)
        self.tree.column("speed", anchor="center", width=col_w)

        self.tree.tag_configure("red", background="#ffcccc")
        self.tree.tag_configure("yellow", background="#fff3cd")
        self.tree.tag_configure("green", background="#d4edda")

        self.status_label = tk.Label(
            main_frame, text="Ожидание", fg="blue", font=("Arial", 10, "bold")
        )
        self.status_label.grid(row=3, column=0, pady=(10, 2), padx=(0, 18))

        self.info_label = tk.Label(main_frame, text="", font=("Arial", 9, "italic"))
        self.info_label.grid(row=4, column=0, pady=2, padx=(0, 18))

        self.folder_btn = tk.Button(
            main_frame,
            text="Открыть папку с отчетами",
            command=self.open_reports_folder,
            bg="#e2e3e5",
            font=("Arial", 9),
            padx=10,
        )
        self.folder_btn.grid(row=5, column=0, pady=10, padx=(0, 18))
        self.folder_btn.grid_remove()

    def open_reports_folder(self):
        docs_folder = os.path.join(os.path.expanduser("~"), "Documents", "NetSpeed")
        if not os.path.exists(docs_folder):
            os.makedirs(docs_folder)
        os.startfile(docs_folder)

    def get_speed_tag(self, speed):
        if speed < 30:
            return "red"
        elif speed < 60:
            return "yellow"
        else:
            return "green"

    def start(self):
        self.running = True
        self.start_time = datetime.now()
        self.data.clear()
        self.tree.delete(*self.tree.get_children())

        self.info_label.config(text="")
        self.folder_btn.grid_remove()

        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.status_label.config(text="Работает", fg="green")
        threading.Thread(target=self.run_speedtest, daemon=True).start()

    def stop(self):
        if not self.running:
            return
        self.running = False
        self.stop_btn.config(state=tk.DISABLED)
        self.start_btn.config(state=tk.NORMAL)
        self.status_label.config(text="Остановка...", fg="orange")

        end_time = datetime.now()
        elapsed = end_time - self.start_time
        total_seconds = int(elapsed.total_seconds())

        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)

        time_parts = []
        if hours > 0:
            time_parts.append(f"{hours}ч")
        if minutes > 0:
            time_parts.append(f"{minutes}мин")
        if seconds > 0 or not time_parts:
            time_parts.append(f"{seconds}сек")
        elapsed_str = " ".join(time_parts)

        if self.data:
            avg_speed = round(sum(d["speed"] for d in self.data) / len(self.data))
            self.info_label.config(
                text=f"Проработано: {elapsed_str} | Средняя скорость: {avg_speed} Мбит/с"
            )
            self.save_csv(self.start_time, end_time)
            self.status_label.config(text="Отчет сохранен", fg="blue")
            self.folder_btn.grid()
        else:
            self.info_label.config(text=f"Проработано: {elapsed_str} | Замеров не было")
            self.status_label.config(text="Завершено без отчета", fg="grey")
            self.folder_btn.grid_remove()

    def run_speedtest(self):
        while self.running:
            try:
                self.root.after(
                    0,
                    lambda: self.status_label.config(
                        text="Выполняется замер...", fg="red"
                    ),
                )

                st = speedtest.Speedtest(secure=True)
                st.get_best_server()
                speed = round(st.download() / 1_000_000)

                now = datetime.now()
                date_str, time_str = now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")
                self.data.append({"date": date_str, "time": time_str, "speed": speed})
                tag = self.get_speed_tag(speed)

                def update_tree(d, t, s, tg):
                    item = self.tree.insert("", tk.END, values=(d, t, s), tags=(tg,))
                    self.tree.see(item)

                self.root.after(0, lambda: update_tree(date_str, time_str, speed, tag))

            except Exception as e:
                self.root.after(
                    0,
                    lambda err=e: self.status_label.config(
                        text=f"Ошибка: {err}", fg="red"
                    ),
                )

            for i in range(INTERVAL):
                if not self.running:
                    return

                remaining = INTERVAL - i
                time_countdown = self.format_time(remaining)

                self.root.after(
                    0,
                    lambda t=time_countdown: self.status_label.config(
                        text=f"Ожидание замера {t}", fg="green"
                    ),
                )

                time.sleep(1)

    def save_csv(self, start, end):
        docs_folder = os.path.join(os.path.expanduser("~"), "Documents", "NetSpeed")
        if not os.path.exists(docs_folder):
            os.makedirs(docs_folder)
        filename = f"Report_{start.strftime('%Y-%m-%d_%H-%M-%S')}.csv"
        df = pd.DataFrame(
            [
                {"Дата": d["date"], "Время": d["time"], "Скорость (Мбит/с)": d["speed"]}
                for d in self.data
            ]
        )
        df.to_csv(
            os.path.join(docs_folder, filename),
            index=False,
            sep=";",
            encoding="utf-8-sig",
        )

    def create_tray_icon(self):
        try:
            icon_path = (
                os.path.join(sys._MEIPASS, "icon.ico")
                if hasattr(sys, "_MEIPASS")
                else "icon.ico"
            )
            return Image.open(icon_path)
        except:
            img = Image.new("RGB", (64, 64), "blue")
            draw = ImageDraw.Draw(img)
            draw.rectangle((16, 16, 48, 48), fill="white")
            return img

    def show_window(self, icon, item):
        self.root.after(0, self.root.deiconify)
        icon.stop()

    def quit_app(self, icon, item):
        self.running = False
        icon.stop()
        self.root.after(0, self.root.destroy)

    def run_tray(self):
        menu = pystray.Menu(
            pystray.MenuItem("Показать", self.show_window),
            pystray.MenuItem("Выход", self.quit_app),
        )
        self.tray_icon = pystray.Icon(
            "speedtest", self.create_tray_icon(), "NetSpeed", menu
        )
        self.tray_icon.run()

    def on_close(self):
        if self.running:
            self.root.withdraw()
            threading.Thread(target=self.run_tray, daemon=True).start()
        else:
            self.root.destroy()

    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def format_time(self, seconds):
        minutes = seconds // 60
        secs = seconds % 60
        return f"{minutes:02d}:{secs:02d}"


if __name__ == "__main__":
    root = tk.Tk()
    app = SpeedTestApp(root)
    root.mainloop()