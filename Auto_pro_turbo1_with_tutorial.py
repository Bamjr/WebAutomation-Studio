import csv
import io
import json
import threading
import time
import traceback
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import Frame, Label

import ttkbootstrap as ttk
from ttkbootstrap.constants import *

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import platform


@dataclass
class ActionBlock:
    kind: str
    widgets: dict
    frame: object


class TabbedApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WebAutomation Studio")
        self.root.geometry("1280x820")

        self.browser_config = {
            "part_profile": None,
            "chrome_path": None,
            "driver_path": None,
            "version": None,
        }

        self.driver = None
        self.actions: list[ActionBlock] = []
        self.dataset_rows: list[dict] = []
        self.is_running = False
        self.speed_mode = tk.StringVar(value="fast")

        self.notebook = ttk.Notebook(root, bootstyle="primary")
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.add_tab(self.chrome_setup_page(), "Custom Chrome Setup")
        self.add_tab(self.studio_page(), "Automation Studio")
        self.add_tab(self.tutorial_page(), "คู่มือการใช้งาน")
        self.add_tab(self.about_page(), "About")

    def add_tab(self, frame, title):
        self.notebook.add(frame, text=title)

    # =========================
    # Browser setup
    # =========================
    def chrome_setup_page(self):
        frame = Frame(self.notebook, bg="#1f1f2e")

        ttk.Label(frame, text="Setup Chrome for dev", font=("Helvetica", 20), bootstyle="info").pack(pady=(12, 0))
        ttk.Label(frame, text="ตั้งค่า Chrome และ Chromedriver ก่อนเริ่ม", bootstyle="secondary").pack(pady=(0, 10))

        ttk.Label(frame, text="ตำแหน่ง Profile Folder", bootstyle="info").pack(pady=(10, 0))
        self.profile_entry = ttk.Entry(frame, width=80)
        self.profile_entry.pack(pady=(0, 10))
        ttk.Button(frame, text="เลือกโฟลเดอร์ Profile", command=self.select_profile, bootstyle="outline-info").pack(pady=(0, 10))

        ttk.Label(frame, text="ตำแหน่ง Chrome.exe", bootstyle="info").pack(pady=(10, 0))
        self.chrome_entry = ttk.Entry(frame, width=80)
        self.chrome_entry.pack(pady=(0, 5))
        chrome_row = ttk.Frame(frame)
        chrome_row.pack(pady=(0, 10))
        ttk.Button(chrome_row, text="เลือก Chrome.exe", command=self.select_chrome, bootstyle="outline-info").pack(side="left", padx=(0, 5))
        ttk.Button(chrome_row, text="📥 Download", command=self.download_chrome, bootstyle="secondary-outline").pack(side="left")

        ttk.Label(frame, text="ตำแหน่ง Chromedriver.exe", bootstyle="info").pack(pady=(10, 0))
        self.driver_entry = ttk.Entry(frame, width=80)
        self.driver_entry.pack(pady=(0, 5))
        driver_row = ttk.Frame(frame)
        driver_row.pack(pady=(0, 10))
        ttk.Button(driver_row, text="เลือก Chromedriver.exe", command=self.select_driver, bootstyle="outline-info").pack(side="left", padx=(0, 5))
        ttk.Button(driver_row, text="📥 Download", command=self.download_chrome, bootstyle="secondary-outline").pack(side="left")

        ttk.Label(frame, text="เวอร์ชัน Chrome (optional)", bootstyle="info").pack(pady=(10, 0))
        self.version_entry = ttk.Entry(frame, width=25)
        self.version_entry.pack(pady=(0, 10))

        btn_row = ttk.Frame(frame)
        btn_row.pack(pady=12)
        ttk.Button(btn_row, text="เริ่ม Browser", command=self.start_browser, bootstyle="success").pack(side="left", padx=5)
        ttk.Button(btn_row, text="เปิด Google", command=lambda: self.goto_url_from_gui("https://www.google.com"), bootstyle="outline-info").pack(side="left", padx=5)
        ttk.Button(btn_row, text="ปิด Browser", command=self.stop_browser, bootstyle="danger-outline").pack(side="left", padx=5)

        return frame

    def download_chrome(self):
        webbrowser.open("https://googlechromelabs.github.io/chrome-for-testing/")

    def select_profile(self):
        path = filedialog.askdirectory()
        if path:
            self.profile_entry.delete(0, "end")
            self.profile_entry.insert(0, path)
            self.browser_config["part_profile"] = path

    def select_chrome(self):
        path = filedialog.askopenfilename(filetypes=[("Chrome Executable", "chrome.exe")])
        if path:
            self.chrome_entry.delete(0, "end")
            self.chrome_entry.insert(0, path)
            self.browser_config["chrome_path"] = path

    def select_driver(self):
        path = filedialog.askopenfilename(filetypes=[("Chromedriver Executable", "chromedriver.exe")])
        if path:
            self.driver_entry.delete(0, "end")
            self.driver_entry.insert(0, path)
            self.browser_config["driver_path"] = path

    def start_browser(self):
        self.browser_config["part_profile"] = self.profile_entry.get().strip() or None
        self.browser_config["chrome_path"] = self.chrome_entry.get().strip() or None
        self.browser_config["driver_path"] = self.driver_entry.get().strip() or None
        self.browser_config["version"] = self.version_entry.get().strip()

        if not all([self.browser_config["part_profile"], self.browser_config["chrome_path"], self.browser_config["driver_path"]]):
            messagebox.showerror("Error", "กรุณาเลือก path ให้ครบก่อนเริ่มทำงาน")
            return

        try:
            if self.driver:
                try:
                    self.driver.quit()
                except Exception:
                    pass
                self.driver = None

            chrome_options = Options()
            chrome_options.add_argument(f"--user-data-dir={self.browser_config['part_profile']}")
            chrome_options.add_argument("--profile-directory=Default")
            chrome_options.binary_location = self.browser_config["chrome_path"]

            service = Service(self.browser_config["driver_path"])
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.driver.get("https://google.com")
            self.log("เปิด Chrome สำเร็จแล้ว")
            messagebox.showinfo("Success", "เปิด Chrome สำเร็จแล้ว")
        except Exception as e:
            self.driver = None
            messagebox.showerror("Error", f"เปิด Chrome ไม่ได้:\n{e}")

    def stop_browser(self):
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None
            self.log("ปิด browser แล้ว")
            messagebox.showinfo("Info", "ปิด browser แล้ว")

    def goto_url_from_gui(self, url: str):
        if not self.ensure_driver():
            return
        self.driver.get(url)
        self.log(f"เปิด URL: {url}")

    # =========================
    # Main Studio page
    # =========================
    def studio_page(self):
        outer = Frame(self.notebook, bg="#1f1f2e")

        top = ttk.Frame(outer, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="Automation Studio", font=("Helvetica", 18, "bold"), bootstyle="info").pack(side="left")

        self.mode_var = tk.StringVar(value="builder")
        ttk.Radiobutton(top, text="Builder Mode", variable=self.mode_var, value="builder", command=self.switch_mode, bootstyle="info-toolbutton").pack(side="left", padx=(15, 5))
        ttk.Radiobutton(top, text="Python Mode", variable=self.mode_var, value="python", command=self.switch_mode, bootstyle="warning-toolbutton").pack(side="left", padx=5)

        ttk.Label(top, text="Speed", bootstyle="secondary").pack(side="left", padx=(18, 4))
        speed_box = ttk.Combobox(top, textvariable=self.speed_mode, values=["safe", "normal", "fast", "turbo"], state="readonly", width=10)
        speed_box.pack(side="left")

        body = ttk.Panedwindow(outer, orient=tk.HORIZONTAL)
        body.pack(fill="both", expand=True, padx=8, pady=8)

        # left panel
        self.left_panel = ttk.Frame(body, padding=8)
        body.add(self.left_panel, weight=3)

        # right panel
        self.right_panel = ttk.Frame(body, padding=8)
        body.add(self.right_panel, weight=2)

        self.build_builder_mode()
        self.build_right_panel()
        self.switch_mode()

        return outer

    def build_builder_mode(self):
        self.builder_wrap = ttk.Frame(self.left_panel)
        self.builder_wrap.pack(fill="both", expand=True)

        toolbar = ttk.Frame(self.builder_wrap)
        toolbar.pack(fill="x", pady=(0, 8))
        ttk.Label(toolbar, text="Add Action", bootstyle="secondary").pack(side="left", padx=(0, 8))

        button_defs = [
            ("Goto URL", self.add_goto_block),
            ("Fill by Label", self.add_fill_label_block),
            ("Fill by Selector", self.add_fill_selector_block),
            ("Click Text", self.add_click_text_block),
            ("Click Selector", self.add_click_selector_block),
            ("Wait", self.add_wait_block),
            ("Press Key", self.add_press_key_block),
        ]
        for title, cmd in button_defs:
            ttk.Button(toolbar, text=title, command=cmd, bootstyle="outline-info").pack(side="left", padx=3)

        self.canvas = tk.Canvas(self.builder_wrap, bg="#fafafa", highlightthickness=0)
        scroll = ttk.Scrollbar(self.builder_wrap, orient="vertical", command=self.canvas.yview)
        self.command_container = ttk.Frame(self.canvas)
        self.command_container.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.command_container, anchor="nw")
        self.canvas.configure(yscrollcommand=scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        run_row = ttk.Frame(self.builder_wrap)
        run_row.pack(fill="x", pady=(8, 0))
        ttk.Button(run_row, text="🚀 Execute Builder Actions", command=self.execute_actions_threaded, bootstyle="success").pack(side="left", padx=4)
        ttk.Button(run_row, text="Clear All Actions", command=self.clear_actions, bootstyle="danger-outline").pack(side="left", padx=4)
        ttk.Button(run_row, text="Import Builder JSON", command=self.import_actions_json, bootstyle="outline-secondary").pack(side="left", padx=4)
        ttk.Button(run_row, text="Export Builder JSON", command=self.export_actions_json, bootstyle="secondary-outline").pack(side="left", padx=4)
        self.python_wrap = ttk.Frame(self.left_panel)

        code_toolbar = ttk.Frame(self.python_wrap)
        code_toolbar.pack(fill="x", pady=(0, 8))
        ttk.Label(code_toolbar, text="Python Editor", font=("Helvetica", 14, "bold"), bootstyle="warning").pack(side="left")
        ttk.Button(code_toolbar, text="Run Python", command=self.run_python_threaded, bootstyle="warning").pack(side="left", padx=8)
        ttk.Button(code_toolbar, text="Load .py", command=self.load_python_file, bootstyle="outline-warning").pack(side="left", padx=4)
        ttk.Button(code_toolbar, text="Save .py", command=self.save_python_file, bootstyle="outline-warning").pack(side="left", padx=4)
        ttk.Button(code_toolbar, text="Insert Example", command=self.insert_python_template, bootstyle="outline-secondary").pack(side="left", padx=4)

        self.code_text = tk.Text(self.python_wrap, wrap="none", font=("Consolas", 11), undo=True, bg="#111827", fg="#f9fafb", insertbackground="white")
        code_scroll_y = ttk.Scrollbar(self.python_wrap, orient="vertical", command=self.code_text.yview)
        code_scroll_x = ttk.Scrollbar(self.python_wrap, orient="horizontal", command=self.code_text.xview)
        self.code_text.configure(yscrollcommand=code_scroll_y.set, xscrollcommand=code_scroll_x.set)
        self.code_text.pack(side="top", fill="both", expand=True)
        code_scroll_y.pack(side="right", fill="y")
        code_scroll_x.pack(side="bottom", fill="x")
        self.insert_python_template()

    def build_right_panel(self):
        # data area
        data_card = ttk.Labelframe(self.right_panel, text="Dataset / Rows", padding=8)
        data_card.pack(fill="both", expand=True, pady=(0, 8))

        toolbar = ttk.Frame(data_card)
        toolbar.pack(fill="x", pady=(0, 6))
        ttk.Button(toolbar, text="Load CSV", command=self.load_csv, bootstyle="outline-info").pack(side="left", padx=3)
        ttk.Button(toolbar, text="Load JSON", command=self.load_json_rows, bootstyle="outline-info").pack(side="left", padx=3)
        ttk.Button(toolbar, text="Apply Text Rows", command=self.apply_dataset_text, bootstyle="secondary-outline").pack(side="left", padx=3)
        ttk.Button(toolbar, text="Example Rows", command=self.insert_dataset_template, bootstyle="secondary-outline").pack(side="left", padx=3)

        ttk.Label(data_card, text="ใส่ JSON array ของข้อมูล เช่น [{'ชื่อ':'แบม','อีเมล':'a@b.com'}]", bootstyle="secondary").pack(anchor="w")
        self.dataset_text = tk.Text(data_card, height=12, wrap="none", font=("Consolas", 10))
        self.dataset_text.pack(fill="both", expand=True, pady=(4, 6))

        mapping_card = ttk.Labelframe(self.right_panel, text="Logs", padding=8)
        mapping_card.pack(fill="both", expand=True)
        self.log_text = tk.Text(mapping_card, height=18, wrap="word", font=("Consolas", 10), state="disabled", bg="#0f172a", fg="#dbeafe")
        self.log_text.pack(fill="both", expand=True)

    def switch_mode(self):
        mode = self.mode_var.get()
        self.builder_wrap.pack_forget()
        self.python_wrap.pack_forget()
        if mode == "builder":
            self.builder_wrap.pack(fill="both", expand=True)
        else:
            self.python_wrap.pack(fill="both", expand=True)

    # =========================
    # Action block builders
    # =========================
    def _new_action_card(self, title: str, kind: str):
        card = ttk.Frame(self.command_container,padding=10,relief="solid",borderwidth=1)
        card.pack(fill="x", padx=6, pady=6)
        header = ttk.Frame(card)
        header.pack(fill="x", pady=(0, 8))
        ttk.Label(header, text=title, font=("Helvetica", 11, "bold"), bootstyle="info").pack(side="left")
        ttk.Button(header, text="ลบ", command=lambda c=card: self.remove_action_card(c), bootstyle="danger-link").pack(side="right")
        return card

    def remove_action_card(self, frame):
        self.actions = [a for a in self.actions if a.frame != frame]
        frame.destroy()
        self.log("ลบ action block แล้ว")

    def clear_actions(self):
        for action in list(self.actions):
            try:
                action.frame.destroy()
            except Exception:
                pass
        self.actions.clear()
        self.log("ล้าง action ทั้งหมดแล้ว")

    def add_goto_block(self):
        card = self._new_action_card("Goto URL", "goto")
        ttk.Label(card, text="URL").pack(anchor="w")
        url = ttk.Entry(card, width=80)
        url.pack(fill="x", pady=(2, 6))
        ttk.Label(card, text="Wait after load (sec)").pack(anchor="w")
        wait = ttk.Entry(card, width=10)
        wait.insert(0, "0.3")
        wait.pack(anchor="w")
        self.actions.append(ActionBlock("goto", {"url": url, "wait": wait}, card))

    def add_fill_label_block(self):
        card = self._new_action_card("Fill by Label", "fill_label")
        ttk.Label(card, text="Label text / question text").pack(anchor="w")
        label_entry = ttk.Entry(card, width=60)
        label_entry.pack(fill="x", pady=(2, 6))
        ttk.Label(card, text="Value or {{column_name}} from dataset").pack(anchor="w")
        value_entry = ttk.Entry(card, width=60)
        value_entry.pack(fill="x", pady=(2, 6))

        row = ttk.Frame(card)
        row.pack(fill="x")
        ttk.Label(row, text="Field type").pack(side="left")
        field_type = ttk.Combobox(row, values=["auto", "input", "textarea"], state="readonly", width=12)
        field_type.set("auto")
        field_type.pack(side="left", padx=(6, 12))
        ttk.Label(row, text="Index").pack(side="left")
        index_entry = ttk.Entry(row, width=6)
        index_entry.insert(0, "0")
        index_entry.pack(side="left", padx=(6, 12))
        ttk.Label(row, text="Timeout").pack(side="left")
        timeout = ttk.Entry(row, width=8)
        timeout.insert(0, "2")
        timeout.pack(side="left", padx=(6, 0))

        self.actions.append(ActionBlock("fill_label", {
            "label": label_entry,
            "value": value_entry,
            "field_type": field_type,
            "index": index_entry,
            "timeout": timeout,
        }, card))

    def add_fill_selector_block(self):
        card = self._new_action_card("Fill by Selector", "fill_selector")
        top = ttk.Frame(card)
        top.pack(fill="x", pady=(0, 6))
        ttk.Label(top, text="Selector type").pack(side="left")
        selector_type = ttk.Combobox(top, values=["xpath", "css", "name", "id", "class_name", "tag_name"], state="readonly", width=14)
        selector_type.set("xpath")
        selector_type.pack(side="left", padx=(6, 12))
        ttk.Label(top, text="Timeout").pack(side="left")
        timeout = ttk.Entry(top, width=8)
        timeout.insert(0, "2")
        timeout.pack(side="left", padx=6)

        ttk.Label(card, text="Selector").pack(anchor="w")
        selector = ttk.Entry(card, width=80)
        selector.pack(fill="x", pady=(2, 6))
        ttk.Label(card, text="Value or {{column_name}} from dataset").pack(anchor="w")
        value = ttk.Entry(card, width=80)
        value.pack(fill="x", pady=(2, 6))
        self.actions.append(ActionBlock("fill_selector", {
            "selector_type": selector_type,
            "selector": selector,
            "value": value,
            "timeout": timeout,
        }, card))

    def add_click_text_block(self):
        card = self._new_action_card("Click by Text", "click_text")
        ttk.Label(card, text="Text to click").pack(anchor="w")
        text_entry = ttk.Entry(card, width=60)
        text_entry.pack(fill="x", pady=(2, 6))
        ttk.Label(card, text="Timeout").pack(anchor="w")
        timeout = ttk.Entry(card, width=8)
        timeout.insert(0, "2")
        timeout.pack(anchor="w")
        self.actions.append(ActionBlock("click_text", {"text": text_entry, "timeout": timeout}, card))

    def add_click_selector_block(self):
        card = self._new_action_card("Click by Selector", "click_selector")
        top = ttk.Frame(card)
        top.pack(fill="x", pady=(0, 6))
        ttk.Label(top, text="Selector type").pack(side="left")
        selector_type = ttk.Combobox(top, values=["xpath", "css", "name", "id", "class_name", "tag_name"], state="readonly", width=14)
        selector_type.set("xpath")
        selector_type.pack(side="left", padx=(6, 12))
        ttk.Label(top, text="Timeout").pack(side="left")
        timeout = ttk.Entry(top, width=8)
        timeout.insert(0, "2")
        timeout.pack(side="left", padx=6)
        ttk.Label(card, text="Selector").pack(anchor="w")
        selector = ttk.Entry(card, width=80)
        selector.pack(fill="x", pady=(2, 6))
        self.actions.append(ActionBlock("click_selector", {
            "selector_type": selector_type,
            "selector": selector,
            "timeout": timeout,
        }, card))

    def add_wait_block(self):
        card = self._new_action_card("Wait", "wait")
        ttk.Label(card, text="Seconds").pack(anchor="w")
        seconds = ttk.Entry(card, width=12)
        seconds.insert(0, "1")
        seconds.pack(anchor="w")
        self.actions.append(ActionBlock("wait", {"seconds": seconds}, card))

    def add_press_key_block(self):
        card = self._new_action_card("Press Key", "press_key")
        ttk.Label(card, text="Selector type").pack(anchor="w")
        selector_type = ttk.Combobox(card, values=["xpath", "css", "name", "id", "class_name", "tag_name"], state="readonly", width=14)
        selector_type.set("xpath")
        selector_type.pack(anchor="w", pady=(2, 6))
        ttk.Label(card, text="Selector").pack(anchor="w")
        selector = ttk.Entry(card, width=80)
        selector.pack(fill="x", pady=(2, 6))
        ttk.Label(card, text="Key name (ENTER, TAB, ESCAPE, SPACE)").pack(anchor="w")
        key_name = ttk.Entry(card, width=20)
        key_name.insert(0, "ENTER")
        key_name.pack(anchor="w", pady=(2, 6))
        ttk.Label(card, text="Timeout").pack(anchor="w")
        timeout = ttk.Entry(card, width=8)
        timeout.insert(0, "2")
        timeout.pack(anchor="w")
        self.actions.append(ActionBlock("press_key", {
            "selector_type": selector_type,
            "selector": selector,
            "key_name": key_name,
            "timeout": timeout,
        }, card))

    # =========================
    # Dataset handling
    # =========================
    def insert_dataset_template(self):
        example = [
            {"ชื่อ": "แบม", "อีเมล": "bam@example.com", "ที่อยู่": "Bangkok", "หมายเลขโทรศัพท์": "0812345678", "ความคิดเห็น": "ทดสอบชุดที่ 1"},
            {"ชื่อ": "Bam 2", "อีเมล": "bam2@example.com", "ที่อยู่": "Chiang Mai", "หมายเลขโทรศัพท์": "0899999999", "ความคิดเห็น": "ทดสอบชุดที่ 2"}
        ]
        self.dataset_text.delete("1.0", "end")
        self.dataset_text.insert("1.0", json.dumps(example, ensure_ascii=False, indent=2))

    def apply_dataset_text(self):
        raw = self.dataset_text.get("1.0", "end").strip()
        if not raw:
            self.dataset_rows = []
            self.log("dataset ว่าง จะรันแบบไม่มี rows")
            return
        try:
            data = json.loads(raw)
            if isinstance(data, dict):
                data = [data]
            if not isinstance(data, list):
                raise ValueError("dataset ต้องเป็น list หรือ dict")
            normalized = []
            for item in data:
                if not isinstance(item, dict):
                    raise ValueError("ทุก row ต้องเป็น object/dict")
                normalized.append(item)
            self.dataset_rows = normalized
            self.log(f"โหลด dataset สำเร็จ {len(self.dataset_rows)} row")
        except Exception as e:
            messagebox.showerror("Dataset Error", f"อ่าน dataset ไม่ได้:\n{e}")

    def load_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f)
                rows = list(reader)
            self.dataset_rows = rows
            self.dataset_text.delete("1.0", "end")
            self.dataset_text.insert("1.0", json.dumps(rows, ensure_ascii=False, indent=2))
            self.log(f"โหลด CSV สำเร็จ {len(rows)} row")
        except Exception as e:
            messagebox.showerror("CSV Error", str(e))

    def load_json_rows(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                data = [data]
            self.dataset_rows = data
            self.dataset_text.delete("1.0", "end")
            self.dataset_text.insert("1.0", json.dumps(data, ensure_ascii=False, indent=2))
            self.log(f"โหลด JSON สำเร็จ {len(data)} row")
        except Exception as e:
            messagebox.showerror("JSON Error", str(e))

    # =========================
    # Logging
    # =========================
    def log(self, message: str):
        ts = time.strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{ts}] {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.root.update_idletasks()

    # =========================
    # Common helpers
    # =========================
    def ensure_driver(self):
        if not self.driver:
            messagebox.showerror("Error", "กรุณาเปิด Chrome ก่อน")
            return False
        return True

    def resolve_value(self, template: str, row: dict | None):
        value = template
        if row:
            for key, raw in row.items():
                value = value.replace(f"{{{{{key}}}}}", str(raw))
        return value

    def selector_to_by(self, selector_type: str):
        mapping = {
            "xpath": By.XPATH,
            "css": By.CSS_SELECTOR,
            "name": By.NAME,
            "id": By.ID,
            "class_name": By.CLASS_NAME,
            "tag_name": By.TAG_NAME,
        }
        return mapping[selector_type]

    def get_effective_timeout(self, timeout: float = 10):
        mode = self.speed_mode.get() if hasattr(self, "speed_mode") else "fast"
        if mode == "turbo":
            return min(timeout, 1.2)
        if mode == "fast":
            return min(timeout, 2.0)
        if mode == "normal":
            return min(timeout, 4.0)
        return timeout

    def should_scroll(self):
        mode = self.speed_mode.get() if hasattr(self, "speed_mode") else "fast"
        return mode in ("safe", "normal")

    def should_click_before_fill(self):
        mode = self.speed_mode.get() if hasattr(self, "speed_mode") else "fast"
        return mode in ("safe", "normal")

    def should_use_js_fill(self):
        mode = self.speed_mode.get() if hasattr(self, "speed_mode") else "fast"
        return mode in ("fast", "turbo")

    def fast_find_elements(self, xpath: str, timeout: float = 10):
        effective_timeout = self.get_effective_timeout(timeout)
        end_time = time.time() + effective_timeout
        last_elements = []
        while time.time() < end_time:
            try:
                elements = self.driver.find_elements(By.XPATH, xpath)
                if elements:
                    return elements
                last_elements = elements
            except Exception:
                pass
            time.sleep(0.05)
        return last_elements

    def wait_element(self, selector_type: str, selector: str, timeout: float = 10):
        by = self.selector_to_by(selector_type)
        effective_timeout = self.get_effective_timeout(timeout)
        return WebDriverWait(self.driver, effective_timeout).until(EC.presence_of_element_located((by, selector)))

    def click_selector(self, selector_type: str, selector: str, timeout: float = 10):
        by = self.selector_to_by(selector_type)
        effective_timeout = self.get_effective_timeout(timeout)
        element = WebDriverWait(self.driver, effective_timeout).until(EC.element_to_be_clickable((by, selector)))
        if self.should_scroll():
            self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
        element.click()
        return element

    def fill_selector(self, selector_type: str, selector: str, value: str, timeout: float = 10):
        element = self.wait_element(selector_type, selector, timeout)
        return self.fill_element(element, value)

    def get_select_all_key(self):
        return Keys.COMMAND if platform.system() == "Darwin" else Keys.CONTROL

    def normalize_space_xpath(self, text: str):
        return f"normalize-space(translate(., '\u00A0', ' '))={self.escape_xpath_text(text.strip())}"

    def build_label_xpath(self, label_text: str, field_type: str = "auto"):
        safe = self.escape_xpath_text(label_text.strip())
        exact_text = self.normalize_space_xpath(label_text)

        input_candidates = [
            f"//label[@for and {exact_text}]/@for",
            f"//label[{exact_text}]/following::input[1]",
            f"//label[contains(normalize-space(.), {safe})]/following::input[1]",
            f"//*[self::div or self::span][{exact_text}]/ancestor::*[self::div or self::section or self::li][1]//input[1]",
            f"//*[self::div or self::span][contains(normalize-space(.), {safe})]/ancestor::*[self::div or self::section or self::li][1]//input[1]",
            f"//input[@placeholder={safe}]",
            f"//input[@aria-label={safe}]",
            f"//input[@name={safe}]",
        ]

        textarea_candidates = [
            f"//label[{exact_text}]/following::textarea[1]",
            f"//label[contains(normalize-space(.), {safe})]/following::textarea[1]",
            f"//*[self::div or self::span][{exact_text}]/ancestor::*[self::div or self::section or self::li][1]//textarea[1]",
            f"//*[self::div or self::span][contains(normalize-space(.), {safe})]/ancestor::*[self::div or self::section or self::li][1]//textarea[1]",
            f"//textarea[@placeholder={safe}]",
            f"//textarea[@aria-label={safe}]",
            f"//textarea[@name={safe}]",
        ]

        if field_type == "input":
            return [x for x in input_candidates if '/@for' not in x]
        if field_type == "textarea":
            return textarea_candidates
        return [x for x in input_candidates if '/@for' not in x] + textarea_candidates

    def resolve_label_for_targets(self, label_text: str):
        safe = self.escape_xpath_text(label_text.strip())
        exact_text = self.normalize_space_xpath(label_text)
        ids = []
        try:
            label_nodes = self.driver.find_elements(By.XPATH, f"//label[@for and ({exact_text} or contains(normalize-space(.), {safe}))]")
            for node in label_nodes:
                target_id = node.get_attribute("for")
                if target_id:
                    ids.append(target_id)
        except Exception:
            pass
        return ids

    def fill_element(self, element, value: str):
        if self.should_scroll():
            self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)

        if self.should_use_js_fill():
            try:
                self.driver.execute_script("""
                    const el = arguments[0];
                    const val = arguments[1];
                    if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
                        el.focus();
                        el.value = '';
                        el.value = val;
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        return true;
                    }
                    return false;
                """, element, value)
                return element
            except Exception:
                pass

        if self.should_click_before_fill():
            try:
                element.click()
            except Exception:
                pass
        try:
            element.clear()
        except Exception:
            pass
        select_all_key = self.get_select_all_key()
        try:
            element.send_keys(select_all_key, "a")
        except Exception:
            pass
        element.send_keys(value)
        return element

    def pick_visible_element(self, elements, index: int = 0):
        usable = [e for e in elements if e.is_displayed() and e.is_enabled()]
        if not usable:
            usable = [e for e in elements if e.is_enabled()] or list(elements)
        if not usable:
            raise NoSuchElementException("ไม่เจอ element ที่ใช้งานได้")
        if index < 0 or index >= len(usable):
            raise NoSuchElementException(f"เจอ field แค่ {len(usable)} ช่อง แต่ขอ index {index}")
        return usable[index]

    def fill_by_label(self, label_text: str, value: str, timeout: float = 10, field_type: str = "auto", index: int = 0):
        last_error = None

        target_ids = self.resolve_label_for_targets(label_text)
        for target_id in target_ids:
            try:
                field_xpath = f"//*[@id={self.escape_xpath_text(target_id)}]"
                elements = self.fast_find_elements(field_xpath, timeout)
                if not elements:
                    raise NoSuchElementException(f"ไม่เจอ element id={target_id}")
                element = self.pick_visible_element(elements, index)
                return self.fill_element(element, value)
            except Exception as e:
                last_error = e

        for xpath in self.build_label_xpath(label_text, field_type):
            try:
                elements = self.fast_find_elements(xpath, timeout)
                if not elements:
                    raise NoSuchElementException(f"ไม่เจอ xpath: {xpath}")
                element = self.pick_visible_element(elements, index)
                return self.fill_element(element, value)
            except Exception as e:
                last_error = e
        raise NoSuchElementException(f"หา field จาก label '{label_text}' ไม่เจอ: {last_error}")

    def click_by_text(self, text: str, timeout: float = 10):
        safe = self.escape_xpath_text(text)
        xpaths = [
            f"//*[self::button or self::span or self::div or self::a][contains(normalize-space(.), {safe})]",
            f"//*[@role='button'][contains(normalize-space(.), {safe})]",
        ]
        last_error = None
        effective_timeout = self.get_effective_timeout(timeout)
        for xpath in xpaths:
            try:
                element = WebDriverWait(self.driver, effective_timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))
                if self.should_scroll():
                    self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
                element.click()
                return element
            except Exception as e:
                last_error = e
        raise NoSuchElementException(f"หา element ด้วย text '{text}' ไม่เจอ: {last_error}")

    def press_key_on_element(self, selector_type: str, selector: str, key_name: str, timeout: float = 10):
        key_map = {
            "ENTER": Keys.ENTER,
            "TAB": Keys.TAB,
            "ESCAPE": Keys.ESCAPE,
            "SPACE": Keys.SPACE,
        }
        key = key_map.get(key_name.upper())
        if not key:
            raise ValueError(f"ไม่รู้จัก key '{key_name}'")
        element = self.wait_element(selector_type, selector, timeout)
        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
        element.send_keys(key)
        return element

    def escape_xpath_text(self, text: str):
        if "'" not in text:
            return f"'{text}'"
        parts = text.split("'")
        return "concat(" + ", \"'\", ".join([f"'{p}'" for p in parts]) + ")"

    # =========================
    # Execute Builder actions
    # =========================
    def execute_actions_threaded(self):
        if self.is_running:
            messagebox.showwarning("Busy", "กำลังรันอยู่")
            return
        if not self.ensure_driver():
            return
        self.apply_dataset_text()
        thread = threading.Thread(target=self.execute_actions, daemon=True)
        thread.start()

    def execute_actions(self):
        self.is_running = True
        try:
            rows = self.dataset_rows if self.dataset_rows else [None]
            self.log(f"เริ่มรัน builder mode จำนวน {len(rows)} รอบ")
            for idx, row in enumerate(rows, start=1):
                self.log(f"--- เริ่ม row {idx}/{len(rows)} ---")
                for action in self.actions:
                    self.execute_single_action(action, row)
                self.log(f"--- จบ row {idx}/{len(rows)} ---")
            self.log("รัน builder mode สำเร็จ")
            messagebox.showinfo("Success", "Execute builder actions สำเร็จ")
        except Exception as e:
            self.log(f"ERROR: {e}")
            self.log(traceback.format_exc())
            messagebox.showerror("Execution Error", f"เกิดข้อผิดพลาด:\n{e}")
        finally:
            self.is_running = False

    def execute_single_action(self, action: ActionBlock, row: dict | None):
        w = action.widgets
        if action.kind == "goto":
            url = self.resolve_value(w["url"].get().strip(), row)
            wait_sec = float(w["wait"].get().strip() or 0)
            self.driver.get(url)
            self.log(f"goto {url}")
            if wait_sec > 0:
                time.sleep(wait_sec)
            return

        if action.kind == "fill_label":
            label_text = self.resolve_value(w["label"].get().strip(), row)
            value = self.resolve_value(w["value"].get().strip(), row)
            field_type = w["field_type"].get().strip()
            timeout = float(w["timeout"].get().strip() or 10)
            index = int((w.get("index").get().strip() if w.get("index") else "0") or 0)
            self.fill_by_label(label_text, value, timeout=timeout, field_type=field_type, index=index)
            self.log(f"fill_by_label '{label_text}'[{index}] = '{value}'")
            return

        if action.kind == "fill_selector":
            selector_type = w["selector_type"].get().strip()
            selector = self.resolve_value(w["selector"].get().strip(), row)
            value = self.resolve_value(w["value"].get().strip(), row)
            timeout = float(w["timeout"].get().strip() or 10)
            self.fill_selector(selector_type, selector, value, timeout=timeout)
            self.log(f"fill_selector {selector_type} => {selector}")
            return

        if action.kind == "click_text":
            text = self.resolve_value(w["text"].get().strip(), row)
            timeout = float(w["timeout"].get().strip() or 10)
            self.click_by_text(text, timeout=timeout)
            self.log(f"click_text '{text}'")
            return

        if action.kind == "click_selector":
            selector_type = w["selector_type"].get().strip()
            selector = self.resolve_value(w["selector"].get().strip(), row)
            timeout = float(w["timeout"].get().strip() or 10)
            self.click_selector(selector_type, selector, timeout=timeout)
            self.log(f"click_selector {selector_type} => {selector}")
            return

        if action.kind == "wait":
            seconds = float(w["seconds"].get().strip() or 0)
            self.log(f"wait {seconds} sec")
            time.sleep(seconds)
            return

        if action.kind == "press_key":
            selector_type = w["selector_type"].get().strip()
            selector = self.resolve_value(w["selector"].get().strip(), row)
            key_name = self.resolve_value(w["key_name"].get().strip(), row)
            timeout = float(w["timeout"].get().strip() or 10)
            self.press_key_on_element(selector_type, selector, key_name, timeout=timeout)
            self.log(f"press_key {key_name} on {selector_type} => {selector}")
            return

        raise ValueError(f"ยังไม่รองรับ action kind: {action.kind}")

    # =========================
    # Python mode
    # =========================
    def insert_python_template(self):
        template = '''# โหมดนี้ให้เขียน Python ควบคุม Selenium ได้เอง
# ของที่ใช้ได้ทันที:
#   driver                 -> Selenium WebDriver
#   app                    -> instance ของโปรแกรมนี้
#   rows                   -> dataset list จากด้านขวา
#   row                    -> row ปัจจุบัน ถ้ามี
# helper functions:
#   goto(url)
#   fill_by_label(label_text, value, timeout=10, field_type="auto", index=0)
#   fill_selector(selector_type, selector, value, timeout=10)
#   click_text(text, timeout=10)
#   click_selector(selector_type, selector, timeout=10)
#   press_key(selector_type, selector, key_name, timeout=10)
#   wait(seconds)
#   log(message)
# speed mode in UI: safe / normal / fast / turbo

log("เริ่ม python mode")

if not rows:
    rows = [{"ชื่อ": "แบม", "อีเมล": "bam@example.com"}]

for row in rows:
    # ตัวอย่าง Google Form
    # goto("https://docs.google.com/forms/your-form-id")
    # fill_by_label("ชื่อ", row["ชื่อ"], index=0)
    # fill_by_label("อีเมล", row["อีเมล"], index=0)
    # click_text("ส่ง")
    # wait(1)
    # click_text("ส่งคำตอบอีกครั้ง")

    log(f"current row => {row}")
'''
        self.code_text.delete("1.0", "end")
        self.code_text.insert("1.0", template)

    def load_python_file(self):
        path = filedialog.askopenfilename(filetypes=[("Python", "*.py")])
        if not path:
            return
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        self.code_text.delete("1.0", "end")
        self.code_text.insert("1.0", content)
        self.log(f"โหลดไฟล์ Python: {path}")

    def save_python_file(self):
        path = filedialog.asksaveasfilename(defaultextension=".py", filetypes=[("Python", "*.py")])
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write(self.code_text.get("1.0", "end"))
        self.log(f"บันทึกไฟล์ Python: {path}")

    def run_python_threaded(self):
        if self.is_running:
            messagebox.showwarning("Busy", "กำลังรันอยู่")
            return
        if not self.ensure_driver():
            return
        self.apply_dataset_text()
        thread = threading.Thread(target=self.run_python_code, daemon=True)
        thread.start()

    def run_python_code(self):
        self.is_running = True
        try:
            code = self.code_text.get("1.0", "end")
            local_rows = self.dataset_rows[:] if self.dataset_rows else []
            current_row = local_rows[0] if local_rows else None

            env = {
                "driver": self.driver,
                "app": self,
                "rows": local_rows,
                "row": current_row,
                "By": By,
                "Keys": Keys,
                "WebDriverWait": WebDriverWait,
                "EC": EC,
                "goto": lambda url: self.driver.get(url),
                "fill_by_label": self.fill_by_label,
                "fill_selector": self.fill_selector,
                "click_text": self.click_by_text,
                "click_selector": self.click_selector,
                "press_key": self.press_key_on_element,
                "wait": time.sleep,
                "log": self.log,
            }

            self.log("เริ่มรัน python mode")
            exec(code, env, env)
            self.log("python mode สำเร็จ")
            messagebox.showinfo("Success", "Run Python สำเร็จ")
        except Exception as e:
            self.log(f"PYTHON ERROR: {e}")
            self.log(traceback.format_exc())
            messagebox.showerror("Python Error", f"เกิดข้อผิดพลาด:\n{e}")
        finally:
            self.is_running = False

    # =========================
    # Export / About
    # =========================
    def export_actions_json(self):
        payload = []
        for action in self.actions:
            item = {"kind": action.kind, "data": {}}
            for key, widget in action.widgets.items():
                try:
                    item["data"][key] = widget.get()
                except Exception:
                    item["data"][key] = None
            payload.append(item)
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        self.log(f"export actions json แล้ว: {path}")

    def import_actions_json(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)

            if not isinstance(payload, list):
                raise ValueError("ไฟล์ JSON ต้องเป็น list ของ actions")

            self.clear_actions()

            loaded_count = 0
            for item in payload:
                if not isinstance(item, dict):
                    continue

                kind = item.get("kind")
                data = item.get("data", {})

                if not kind:
                    continue

                self.load_action_from_json(kind, data)
                loaded_count += 1

            self.log(f"import actions json สำเร็จ: {loaded_count} actions")
            messagebox.showinfo("Success", f"Import Builder JSON สำเร็จ {loaded_count} action(s)")

        except Exception as e:
            self.log(f"IMPORT ERROR: {e}")
            messagebox.showerror("Import Error", f"อ่านไฟล์ builder json ไม่ได้:\n{e}")

    def load_action_from_json(self, kind: str, data: dict):
        action_before = len(self.actions)

        if kind == "goto":
            self.add_goto_block()

        elif kind == "fill_label":
            self.add_fill_label_block()

        elif kind == "fill_selector":
            self.add_fill_selector_block()

        elif kind == "click_text":
            self.add_click_text_block()

        elif kind == "click_selector":
            self.add_click_selector_block()

        elif kind == "wait":
            self.add_wait_block()

        elif kind == "press_key":
            self.add_press_key_block()

        else:
            self.log(f"ข้าม action ที่ไม่รองรับ: {kind}")
            return

        if len(self.actions) <= action_before:
            raise RuntimeError(f"สร้าง action ไม่สำเร็จสำหรับ kind={kind}")

        action = self.actions[-1]

        for key, value in data.items():
            widget = action.widgets.get(key)
            if widget is None:
                continue

            try:
                if isinstance(widget, ttk.Combobox):
                    widget.set("" if value is None else str(value))
                else:
                    widget.delete(0, "end")
                    widget.insert(0, "" if value is None else str(value))
            except Exception:
                try:
                    widget.set("" if value is None else str(value))
                except Exception:
                    pass


    def tutorial_page(self):
        frame = Frame(self.notebook, bg="#10243f")

        ttk.Label(
            frame,
            text="คู่มือการใช้งานโปรแกรม",
            font=("Helvetica", 20, "bold"),
            bootstyle="info"
        ).pack(pady=(16, 6))

        ttk.Label(
            frame,
            text="คำแนะนำฉบับย่อสำหรับการตั้งค่า การสร้างงานอัตโนมัติ และการใช้งานอย่างถูกต้อง",
            bootstyle="secondary"
        ).pack(pady=(0, 12))

        container = ttk.Frame(frame, padding=12)
        container.pack(fill="both", expand=True, padx=14, pady=8)

        tutorial_text = tk.Text(
            container,
            wrap="word",
            font=("TH Sarabun New", 16),
            bg="#f8fafc",
            fg="#0f172a",
            padx=16,
            pady=16
        )
        tutorial_scroll = ttk.Scrollbar(container, orient="vertical", command=tutorial_text.yview)
        tutorial_text.configure(yscrollcommand=tutorial_scroll.set)

        tutorial_text.pack(side="left", fill="both", expand=True)
        tutorial_scroll.pack(side="right", fill="y")

        content = """
คู่มือการใช้งาน WebAutomation Studio

ฉบับย่อแบบเป็นทางการ

คู่มือนี้จัดทำขึ้นเพื่ออธิบายการใช้งานโปรแกรม WebAutomation Studio สำหรับเปิดหน้าเว็บ กรอกข้อมูล คลิกปุ่ม และรันงานอัตโนมัติผ่าน Selenium โดยรองรับทั้ง Builder Mode และ Python Mode รวมถึงการใช้ข้อมูลจาก JSON หรือ CSV เพื่อวนทำงานหลายรายการ

--------------------------------------------------

1. ภาพรวมของโปรแกรม

โปรแกรมประกอบด้วย 3 ส่วนหลัก

1.1 Custom Chrome Setup

ใช้กำหนดค่าการเปิดเบราว์เซอร์ ได้แก่

- Profile Folder
- Chrome.exe
- Chromedriver.exe

เมื่อกำหนดครบแล้วจึงสามารถเริ่ม Browser ได้ หากข้อมูลไม่ครบ โปรแกรมจะไม่อนุญาตให้เริ่มทำงาน

1.2 Automation Studio

เป็นพื้นที่หลักสำหรับสร้างลำดับคำสั่งอัตโนมัติ โดยมี 2 โหมด

- Builder Mode สำหรับใช้งานแบบบล็อกคำสั่ง
- Python Mode สำหรับเขียนโค้ดควบคุมเอง

1.3 About

แสดงข้อมูลสรุปของโปรแกรม

--------------------------------------------------

2. ข้อควรระวังก่อนเริ่มใช้งาน

2.1 ห้ามปิดแท็บแรกที่โปรแกรมเปิดขึ้นมา

หลังจากกดเริ่ม Browser โปรแกรมจะเปิด Chrome และเข้าสู่หน้าเริ่มต้นให้โดยอัตโนมัติ หากผู้ใช้ปิดแท็บแรกนั้นเอง อาจทำให้ลำดับการทำงานภายหลังเกิดข้อผิดพลาดได้ เนื่องจากการควบคุมของ Selenium ยังอ้างอิงกับหน้าต่างและแท็บที่ถูกเปิดไว้ตั้งแต่ต้น

คำแนะนำ
ควรปล่อยแท็บแรกไว้ตลอดการทำงาน และให้โปรแกรมเป็นผู้เปลี่ยน URL ผ่านคำสั่ง Goto URL หรือคำสั่งใน Python Mode แทน

2.2 ต้องใช้ Chrome และ Chromedriver ที่เข้ากัน

หากเวอร์ชันไม่สอดคล้องกัน อาจเปิด Browser ไม่ได้หรือเกิดข้อผิดพลาดระหว่างทำงาน

2.3 ไม่ควรเปิดใช้ Profile เดียวกันซ้อนกันหลายหน้าต่าง

อาจทำให้ Chrome เปิดไม่ขึ้น หรือ session เกิดการชนกัน

--------------------------------------------------

3. หลักการทำงานพื้นฐาน

การสั่งงานเว็บอัตโนมัติจำเป็นต้องระบุว่าโปรแกรมจะไปจับ element ตัวใดบนหน้าเว็บ เช่น

- ช่องกรอกข้อความ
- ปุ่มส่ง
- ช่องค้นหา
- กล่องข้อความยาว

โปรแกรมจึงรองรับการค้นหา element หลายรูปแบบ เช่น

- id
- name
- css
- xpath
- class_name
- tag_name

--------------------------------------------------

4. การ Inspect HTML เบื้องต้น

การ Inspect ใช้สำหรับดูโครงสร้าง HTML ของหน้าเว็บ เพื่อค้นหาว่าควรอ้างอิง element ด้วยวิธีใด

วิธีเปิด Inspect

ใน Chrome ให้
- คลิกขวาที่ช่องหรือปุ่มที่ต้องการ
- เลือก Inspect
- หรือกด F12

สิ่งที่ควรมองหา

- id
- name
- class
- placeholder
- aria-label
- ข้อความที่แสดงบนปุ่มหรือ label
- โครงสร้างพ่อแม่ลูกของ element

ตัวอย่าง

label for="email" อีเมล
input id="email" name="email" type="text"

จากตัวอย่างนี้ ช่องกรอกสามารถอ้างอิงได้ด้วย

- id = email
- name = email
- XPath ที่อ้างอิง input

--------------------------------------------------

5. ความหมายของ XPath

XPath คือรูปแบบการระบุตำแหน่ง element ภายในโครงสร้าง HTML โดยเหมาะสำหรับกรณีที่ต้องการอ้างอิงจากข้อความหรือโครงสร้าง

ตัวอย่าง

หา input จาก id
//*[@id='email']

หา input จาก name
//input[@name='email']

หาปุ่มที่มีข้อความว่า ส่ง
//button[normalize-space()='ส่ง']

หาข้อความ ส่ง ที่อยู่ใน span แล้วไต่ขึ้นไปยังปุ่ม
//span[normalize-space()='ส่ง']/ancestor::*[@role='button']

--------------------------------------------------

6. ความหมายของ CSS Selector

CSS Selector เหมาะกับการอ้างอิง element จาก id, class หรือ attribute

ตัวอย่าง

#email
input[name="email"]
button.submit-btn

หลักการเลือกใช้งาน

- หากมี id ชัดเจน ควรใช้ id
- หากมี name ชัดเจน ควรใช้ name
- หากต้องอ้างอิงจากข้อความบนหน้าจอ ควรใช้ XPath

--------------------------------------------------

7. แนวทางเลือก selector ที่เหมาะสม

ลำดับความเสถียรโดยทั่วไป

1. id
2. name
3. css
4. xpath
5. class_name
6. tag_name

หมายเหตุสำคัญ

class_name ไม่เหมาะกับ element ที่มี class ซ้ำกันจำนวนมาก และไม่ควรใส่ class หลายชื่อรวมกันเป็นสตริงเดียว

--------------------------------------------------

8. การใช้งาน Builder Mode

8.1 Goto URL

ใช้เปิดหน้าเว็บตาม URL ที่กำหนด

ค่าที่ต้องระบุ

- URL
- Wait after load (วินาที)

คำแนะนำ
ควรใช้เป็นบล็อกแรกของงานส่วนใหญ่

--------------------------------------------------

8.2 Fill by Label

ใช้กรอกข้อมูลโดยอาศัยข้อความของ label หรือคำถาม

ค่าที่ต้องระบุ

- Label text
- Value
- Field type
- Index
- Timeout

คำแนะนำ

- ใช้ auto เมื่อไม่แน่ใจ
- index เริ่มที่ 0

--------------------------------------------------

8.3 Fill by Selector

ใช้กรอกข้อมูลโดยอ้างอิง selector โดยตรง

ค่าที่ต้องระบุ

- Selector type
- Selector
- Value
- Timeout

--------------------------------------------------

8.4 Click by Text

ใช้คลิก element จากข้อความ เช่น ส่ง ถัดไป ยืนยัน

--------------------------------------------------

8.5 Click by Selector

ใช้คลิก element จาก selector โดยตรง

--------------------------------------------------

8.6 Wait

ใช้หน่วงเวลาเป็นวินาที

--------------------------------------------------

8.7 Press Key

ใช้ส่งคำสั่งคีย์บอร์ด เช่น ENTER TAB ESCAPE SPACE

--------------------------------------------------

9. การใช้ข้อมูลหลายรายการด้วย Dataset

รองรับ CSV และ JSON

ตัวอย่างข้อมูล

[
 {"ชื่อ":"แบม","อีเมล":"bam@example.com"},
 {"ชื่อ":"เมย์","อีเมล":"may@example.com"}
]

การเรียกใช้

{{ชื่อ}}
{{อีเมล}}

--------------------------------------------------

10. การใช้งาน Python Mode

มีตัวช่วยดังนี้

driver
rows
row
goto()
fill_by_label()
fill_selector()
click_text()
click_selector()
press_key()
wait()
log()

เหมาะสำหรับงานที่ซับซ้อน

--------------------------------------------------

11. Speed Mode

มี 4 ระดับ

safe
normal
fast
turbo

แนวทางเลือก

safe สำหรับเว็บช้า
normal สำหรับทั่วไป
fast สำหรับเว็บเสถียร
turbo สำหรับเว็บเร็วมาก

--------------------------------------------------

12. Import และ Export Builder JSON

สามารถบันทึกและโหลดลำดับคำสั่งได้

--------------------------------------------------

13. แนวทางแก้ปัญหาเบื้องต้น

กรณีหา field ไม่เจอ

- ตรวจสอบ label
- ตรวจสอบประเภท field
- ตรวจสอบ selector

กรณีคลิกไม่ได้

- ตรวจสอบ element จริง
- ตรวจสอบว่ามีสิ่งบังหรือไม่

กรณี timeout

- selector ผิด
- หน้าเว็บยังโหลดไม่เสร็จ
- element ยังไม่พร้อม

--------------------------------------------------

14. ข้อแนะนำในการใช้งานจริง

- ทดสอบทีละบล็อก
- ใช้ id หรือ name ก่อน
- ใช้ XPath เมื่อจำเป็น
- อย่าปิดแท็บแรก
- บันทึกงานด้วย Export JSON

--------------------------------------------------

15. สรุป

WebAutomation Studio เป็นเครื่องมือสำหรับสร้างงานอัตโนมัติบนเว็บผ่าน Selenium โดยมีทั้ง Builder Mode และ Python Mode และรองรับ dataset จาก JSON และ CSV
""".strip()

        tutorial_text.insert("1.0", content)
        tutorial_text.configure(state="disabled")

        return frame

    def about_page(self):
        frame = Frame(self.notebook, bg="#2b1e4d")
        Label(frame, text="WebAutomation Studio", font=("Helvetica", 20, "bold"), bg="#2b1e4d", fg="#a0f5c0").pack(pady=(20, 5))
        Label(
            frame,
            text=(
                "Builder Mode: ประกอบ action เป็น block\n"
                "Python Mode: เขียนโค้ด Selenium เองได้\n"
                "รองรับ dataset แบบ JSON/CSV เพื่อวนกรอกหลายชุด"
            ),
            font=("Helvetica", 12),
            bg="#2b1e4d",
            fg="#ffffff",
            justify="left",
        ).pack(pady=10)
        Label(frame, text="GitHub: https://github.com/Bamjr", font=("Helvetica", 14), bg="#2b1e4d", fg="#a0f5c0").pack(pady=10)
        Label(frame, text="GitHub: https://github.com/nooparnjnag", font=("Helvetica", 14), bg="#2b1e4d", fg="#a0f5c0").pack(pady=10)
        return frame


if __name__ == "__main__":
    app = ttk.Window(themename="vapor")
    TabbedApp(app)
    app.mainloop()
