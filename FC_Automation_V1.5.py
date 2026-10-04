import os
import sys
import time
import datetime
import threading
import traceback
import queue
import ctypes
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import load_workbook

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

TARGET_URL = "http://127.0.0.1:12580/ra/index.html#initial"

HOTFOLDER_CONFIGS = [
    {"suffix": "HotFolder", "path": r"C:\MediaHub\HotFolder"},
    {"suffix": "Missing-Files", "path": r"C:\MediaHub\Missing-Files"}
]

# Color Schemes for Modern Web-like Themes
THEMES = {
    "dark": {
        "bg": "#0f172a",
        "card_bg": "#1e293b",
        "card_border": "#334155",
        "fg_main": "#f8fafc",
        "fg_sub": "#94a3b8",
        "accent": "#0ea5e9",
        "accent_hover": "#0284c7",
        "input_bg": "#0f172a",
        "input_fg": "#38bdf8",
        "log_bg": "#0f172a",
        "log_fg": "#38bdf8",
        "tree_bg": "#0f172a",
        "tree_fg": "#f1f5f9",
        "tree_head_bg": "#1e293b",
        "tree_head_fg": "#38bdf8",
        "tree_select": "#0284c7",
        "icon_fg": "#38bdf8"
    },
    "light": {
        "bg": "#f1f5f9",
        "card_bg": "#ffffff",
        "card_border": "#cbd5e1",
        "fg_main": "#0f172a",
        "fg_sub": "#64748b",
        "accent": "#0284c7",
        "accent_hover": "#0369a1",
        "input_bg": "#f8fafc",
        "input_fg": "#0284c7",
        "log_bg": "#f8fafc",
        "log_fg": "#0369a1",
        "tree_bg": "#ffffff",
        "tree_fg": "#0f172a",
        "tree_head_bg": "#e2e8f0",
        "tree_head_fg": "#0284c7",
        "tree_select": "#38bdf8",
        "icon_fg": "#0284c7"
    }
}


def get_exe_dir():
    """Gets the exact path directory where the .exe or script is located."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def read_excel_data(file_path):
    """Reads Excel records and converts headers to standardized lowercase key names."""
    wb = load_workbook(file_path, data_only=True)
    sheet = wb.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return []
    
    headers = []
    for cell in rows[0]:
        if cell is not None:
            clean_hdr = str(cell).strip().lower().replace(".", "").replace(" ", "_").replace("/", "_")
            headers.append(clean_hdr)
        else:
            headers.append("")

    data = []
    for row in rows[1:]:
        if not any(row):
            continue
        row_dict = {}
        for idx, value in enumerate(row):
            if idx < len(headers):
                row_dict[headers[idx]] = str(value).strip() if value is not None else ""
        data.append(row_dict)
        
    return data


class HeaderIcon(tk.Canvas):
    """Custom vector emblem for application header banner using primitive shapes."""
    def __init__(self, parent, bg_color="#0f172a", fg_color="#38bdf8", size=36):
        super().__init__(parent, width=size, height=size, bg=bg_color, highlightthickness=0)
        self.size = size
        self.draw_icon(bg_color, fg_color)

    def draw_icon(self, bg_color, fg_color):
        self.delete("all")
        self.configure(bg=bg_color)
        # Server Rack 1
        self.create_rectangle(3, 4, 33, 13, outline=fg_color, width=2)
        self.create_oval(6, 7, 10, 10, fill="#22c55e", outline="")
        self.create_line(14, 8, 29, 8, fill=fg_color, width=1)

        # Server Rack 2
        self.create_rectangle(3, 17, 33, 26, outline=fg_color, width=2)
        self.create_oval(6, 20, 10, 23, fill="#22c55e", outline="")
        self.create_line(14, 21, 29, 21, fill=fg_color, width=1)

        # Foreground Bolt
        self.create_polygon(22, 1, 14, 18, 20, 18, 16, 35, 29, 15, 22, 15, fill="#f59e0b", outline="")


class FileCatalystAutomatorGUI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("FileCatalyst Enterprise TC Provisioning Portal")
        self.geometry("1400x900")
        self.minsize(1200, 750)

        self.current_theme = "dark"
        self.excel_file_path = ""
        self.is_running = False
        self.selected_record = None
        self.log_queue = queue.Queue()

        # Define non-modifiable executable log file path
        exe_folder = get_exe_dir()
        self.log_file_path = os.path.join(exe_folder, "provisioning_pipeline.log")

        self._set_window_title_icon()
        self._build_ui()
        self.apply_theme(self.current_theme)
        self._start_queue_listener()

        self.log(f"System initialized. Executable Folder: '{exe_folder}'")
        self.log(f"Automatic audit logging engaged: '{self.log_file_path}'")

    def _set_window_title_icon(self):
        """Draws the exact server-rack icon into Tkinter window title bar."""
        try:
            self.app_icon_img = tk.PhotoImage(width=16, height=16)
            icon_pixels = [
                "...............Y",
                "..............YY",
                ".CCCCCCCCCCCC.Y.",
                ".C...G...C..C.Y.",
                ".C...G...C..CY..",
                ".CCCCCCCCCCCYYY.",
                "............Y...",
                "...........YY...",
                "..........YY....",
                ".CCCCCCCCYYYCCCC",
                ".C...G..YY..C..C",
                ".C...G.YY...C..C",
                ".CCCCCCCYCCCCCCC",
                ".......YY.......",
                "......YY........",
                "......Y........."
            ]
            color_map = {'C': '#38bdf8', 'G': '#22c55e', 'Y': '#f59e0b', '.': '#0f172a'}

            for y, row in enumerate(icon_pixels):
                for x, char in enumerate(row):
                    self.app_icon_img.put(color_map[char], (x, y))

            self.iconphoto(True, self.app_icon_img)
        except Exception:
            pass

    def toggle_theme(self):
        self.current_theme = "light" if self.current_theme == "dark" else "dark"
        self.apply_theme(self.current_theme)

    def apply_theme(self, theme_key):
        t = THEMES[theme_key]
        self.configure(bg=t["bg"])

        # Main Containers
        self.header_frame.configure(bg=t["bg"])
        self.title_frame.configure(bg=t["bg"])
        self.main_container.configure(bg=t["bg"])
        self.paned.configure(bg=t["bg"])

        # Header Elements
        self.lbl_title.configure(fg=t["fg_main"], bg=t["bg"])
        self.lbl_subtitle.configure(fg=t["fg_sub"], bg=t["bg"])
        self.header_icon.draw_icon(t["bg"], t["icon_fg"])
        
        self.theme_btn.configure(
            text="☀️️ Light Mode" if theme_key == "dark" else "🌙 Dark Mode",
            bg=t["card_bg"],
            fg=t["fg_main"],
            activebackground=t["card_border"],
            activeforeground=t["fg_main"]
        )

        # Cards / LabelFrames
        for card in [self.step1_frame, self.step2_frame, self.step3_frame]:
            card.configure(bg=t["card_bg"], fg=t["accent"], bd=1, relief="solid")

        self.grid_frame.configure(bg=t["card_bg"])
        self.progress_container.configure(bg=t["card_bg"])

        # Inputs & Labels
        self.lbl_excel.configure(fg=t["fg_main"], bg=t["card_bg"])
        self.lbl_tc.configure(fg=t["fg_main"], bg=t["card_bg"])
        self.file_status_lbl.configure(bg=t["card_bg"])

        self.path_entry.configure(bg=t["input_bg"], fg=t["input_fg"], insertbackground=t["input_fg"])
        self.tc_entry.configure(bg=t["input_bg"], fg=t["input_fg"], insertbackground=t["input_fg"])

        self.browse_btn.configure(bg=t["accent"], activebackground=t["accent_hover"])
        self.fetch_btn.configure(bg=t["accent"], activebackground=t["accent_hover"])

        # Console Logs
        self.log_text.configure(bg=t["log_bg"], fg=t["log_fg"])
        self.progress_lbl.configure(bg=t["card_bg"])

        # TTK Styles for Treeview & Progressbar
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 9, "bold"),
            background=t["tree_head_bg"],
            foreground=t["tree_head_fg"],
            relief="flat",
            padding=8
        )
        style.configure(
            "Treeview",
            font=("Segoe UI", 9),
            rowheight=34,
            background=t["tree_bg"],
            fieldbackground=t["tree_bg"],
            foreground=t["tree_fg"],
            borderwidth=0
        )
        style.map("Treeview", background=[("selected", t["tree_select"])], foreground=[("selected", "#ffffff")])
        style.configure(
            "TProgressbar",
            thickness=10,
            troughcolor=t["bg"],
            background=t["accent"],
            borderwidth=0
        )

    def _build_ui(self):
        # ------------------- HEADER -------------------
        self.header_frame = tk.Frame(self, height=70)
        self.header_frame.pack(fill="x", padx=25, pady=(15, 10))

        self.header_icon = HeaderIcon(self.header_frame, size=36)
        self.header_icon.pack(side="left", padx=(0, 12))

        self.title_frame = tk.Frame(self.header_frame)
        self.title_frame.pack(side="left")

        self.lbl_title = tk.Label(
            self.title_frame,
            text="FileCatalyst Sequential Provisioning Console",
            font=("Segoe UI", 16, "bold")
        )
        self.lbl_title.pack(anchor="w")

        self.lbl_subtitle = tk.Label(
            self.title_frame,
            text="Automated Site Configuration & HotFolder Deployment Engine",
            font=("Segoe UI", 9)
        )
        self.lbl_subtitle.pack(anchor="w")

        # Theme Switcher Button
        self.theme_btn = tk.Button(
            self.header_frame,
            text="☀️ Light Mode",
            command=self.toggle_theme,
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=14,
            pady=4,
            cursor="hand2"
        )
        self.theme_btn.pack(side="right", pady=5)

        # ------------------- MAIN CONTAINER -------------------
        self.main_container = tk.Frame(self, padx=25, pady=0)
        self.main_container.pack(fill="both", expand=True)

        # ------------------- STEP 1: CARD CONTAINER -------------------
        self.step1_frame = tk.LabelFrame(
            self.main_container,
            text=" STEP 1: Target TC Selection & Query ",
            font=("Segoe UI", 10, "bold"),
            padx=18,
            pady=16
        )
        self.step1_frame.pack(fill="x", pady=(0, 12))

        self.grid_frame = tk.Frame(self.step1_frame)
        self.grid_frame.pack(fill="x")

        # File Upload Label & Input
        self.lbl_excel = tk.Label(self.grid_frame, text="Excel Template:", font=("Segoe UI", 9, "bold"))
        self.lbl_excel.grid(row=0, column=0, sticky="w", pady=4)

        self.path_entry = tk.Entry(
            self.grid_frame,
            font=("Consolas", 9, "bold"),
            bd=1,
            relief="solid",
            state="readonly"
        )
        self.path_entry.grid(row=0, column=1, sticky="ew", padx=(10, 10), pady=4)

        self.browse_btn = tk.Button(
            self.grid_frame,
            text="Browse File...",
            command=self.browse_file,
            fg="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=14,
            pady=4,
            cursor="hand2"
        )
        self.browse_btn.grid(row=0, column=2, padx=(0, 20), pady=4)

        # TC Search Code Input
        self.lbl_tc = tk.Label(self.grid_frame, text="Target TC Code:", font=("Segoe UI", 9, "bold"))
        self.lbl_tc.grid(row=0, column=3, sticky="w", pady=4)

        self.tc_entry = tk.Entry(
            self.grid_frame,
            font=("Segoe UI", 10, "bold"),
            bd=1,
            relief="solid",
            width=18
        )
        self.tc_entry.grid(row=0, column=4, sticky="w", padx=(10, 10), pady=4)
        self.tc_entry.bind("<Return>", lambda e: self.fetch_tc_details())

        self.fetch_btn = tk.Button(
            self.grid_frame,
            text="Fetch Details",
            command=self.fetch_tc_details,
            fg="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=16,
            pady=4,
            cursor="hand2"
        )
        self.fetch_btn.grid(row=0, column=5, pady=4)

        self.grid_frame.columnconfigure(1, weight=1)

        # Status Label below input field
        self.file_status_lbl = tk.Label(
            self.step1_frame,
            text="No Excel file selected.",
            font=("Segoe UI", 8, "italic"),
            fg="#64748b"
        )
        self.file_status_lbl.pack(anchor="w", pady=(4, 0))

        # Execute Pipeline Button
        self.start_btn = tk.Button(
            self.step1_frame,
            text="▶ Start Provisioning Pipeline for Selected TC",
            command=self.start_automation,
            bg="#334155",
            fg="#94a3b8",
            activebackground="#15803d",
            activeforeground="white",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            pady=8,
            cursor="hand2",
            state="disabled"
        )
        self.start_btn.pack(fill="x", pady=(10, 0))

        # ------------------- SPLIT PANED WINDOW (STEPS 2 & 3) -------------------
        self.paned = tk.PanedWindow(self.main_container, orient=tk.VERTICAL, sashwidth=8, bd=0)
        self.paned.pack(fill="both", expand=True)

        # ------------------- STEP 2: DEPLOYMENT STATUS CARD -------------------
        self.step2_frame = tk.LabelFrame(
            self.paned,
            text=" STEP 2: Selected TC Parameters & Deployment Status ",
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=14
        )
        self.paned.add(self.step2_frame, minsize=260)

        # Progress Status Container
        self.progress_container = tk.Frame(self.step2_frame)
        self.progress_container.pack(fill="x", pady=(0, 10))

        self.progress_lbl = tk.Label(
            self.progress_container,
            text="Current Progress: Idle / Ready",
            font=("Segoe UI", 9, "bold"),
            fg="#94a3b8"
        )
        self.progress_lbl.pack(anchor="w", pady=(0, 4))

        self.progress_bar = ttk.Progressbar(self.progress_container, style="TProgressbar", mode="determinate")
        self.progress_bar.pack(fill="x")

        # Table Listing details
        self.columns = (
            "Sr.No.", "TC Code", "Center Name", "City", 
            "State", "Zone", "Center Status", "Site ID", 
            "Host", "Port", "Pipeline Status"
        )
        self.tree = ttk.Treeview(self.step2_frame, columns=self.columns, show="headings")

        col_widths = {
            "Sr.No.": 60, "TC Code": 90, "Center Name": 220, "City": 110,
            "State": 110, "Zone": 80, "Center Status": 110, "Site ID": 140,
            "Host": 120, "Port": 60, "Pipeline Status": 180
        }

        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(
                col, 
                anchor="center" if col in ("Sr.No.", "TC Code", "Port", "Pipeline Status", "Zone") else "w", 
                width=col_widths.get(col, 100)
            )

        self.tree.pack(side="left", fill="both", expand=True)

        tree_scroll_y = ttk.Scrollbar(self.step2_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=tree_scroll_y.set)
        tree_scroll_y.pack(side="right", fill="y")

        # ------------------- STEP 3: EXECUTION LOGS CARD -------------------
        self.step3_frame = tk.LabelFrame(
            self.paned,
            text=" STEP 3: Real-time Pipeline Execution Logs (Read-Only) ",
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=14
        )
        self.paned.add(self.step3_frame, minsize=200)

        # Log Text Box is disabled so user cannot type or change logs in UI
        self.log_text = tk.Text(
            self.step3_frame,
            font=("Consolas", 10),
            wrap="word",
            bd=0,
            padx=10,
            pady=10,
            state="disabled"
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        console_scroll = ttk.Scrollbar(self.step3_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscroll=console_scroll.set)
        console_scroll.pack(side="right", fill="y")

    # ------------------- HELPER METHODS -------------------
    def _write_to_disk_log(self, entry):
        """Appends log entry to log file in executable directory and locks it as Read-Only."""
        try:
            # Temporarily un-protect file if it exists
            if os.path.exists(self.log_file_path):
                if sys.platform.startswith("win"):
                    ctypes.windll.kernel32.SetFileAttributesW(self.log_file_path, 0x80)  # FILE_ATTRIBUTE_NORMAL

            # Append entry
            with open(self.log_file_path, "a", encoding="utf-8") as f:
                f.write(entry)

            # Re-apply Read-Only and System attributes on Windows
            if sys.platform.startswith("win"):
                # FILE_ATTRIBUTE_READONLY (0x1) | FILE_ATTRIBUTE_SYSTEM (0x4)
                ctypes.windll.kernel32.SetFileAttributesW(self.log_file_path, 0x1 | 0x4)
        except Exception as e:
            print(f"File log error: {e}")

    def log(self, message):
        """Thread-safe logging to UI and non-modifiable disk file."""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = f"[{timestamp}] {message}\n"
        self.log_queue.put(entry)
        self._write_to_disk_log(entry)

    def _start_queue_listener(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get_nowait()
            self.log_text.config(state="normal")
            self.log_text.insert(tk.END, msg)
            self.log_text.see(tk.END)
            self.log_text.config(state="disabled")
        self.after(150, self._start_queue_listener)

    def browse_file(self):
        file_selected = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file_selected:
            self.excel_file_path = file_selected
            
            self.path_entry.config(state="normal")
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, file_selected)
            self.path_entry.config(state="readonly")
            
            filename = os.path.basename(file_selected)
            self.file_status_lbl.config(
                text=f"✓ Selected File: {filename} ({file_selected})",
                fg="#22c55e"
            )
            
            self.log(f"Excel File Loaded: {file_selected}")

    def fetch_tc_details(self):
        tc_query = self.tc_entry.get().strip()

        if not self.excel_file_path or not os.path.exists(self.excel_file_path):
            messagebox.showwarning("Missing Template", "Please browse and select a valid Excel template file first.")
            return

        if not tc_query:
            messagebox.showwarning("Input Missing", "Please enter a TC Code or Sr.No. to search.")
            return

        if not tc_query.isdigit():
            messagebox.showwarning("Invalid Input", "Please enter digits only for TC Code / Sr.No.")
            return

        try:
            records = read_excel_data(self.excel_file_path)
            matched_record = None

            for rec in records:
                tc_code_val = rec.get("tc_code", "")
                sr_no_val = rec.get("srno", "") or rec.get("sr_no", "")

                if tc_query == tc_code_val or tc_query == sr_no_val:
                    matched_record = rec
                    break

            if not matched_record:
                for item in self.tree.get_children():
                    self.tree.delete(item)
                self.selected_record = None
                self.start_btn.config(state="disabled", bg="#334155", fg="#94a3b8")
                self.progress_lbl.config(text="Status: Target TC not found", fg="#ef4444")
                
                self.log(f"SEARCH FAILED: TC Code / Sr.No. '{tc_query}' was not found in the uploaded template.")
                messagebox.showinfo(
                    "TC Not Found",
                    f"The specified TC Code / Sr.No. '{tc_query}' does not exist in the provided Excel template."
                )
                return

            self.selected_record = matched_record
            for item in self.tree.get_children():
                self.tree.delete(item)

            sr_no = matched_record.get("srno", "") or matched_record.get("sr_no", "-")
            tc_code = matched_record.get("tc_code", "-")
            center_name = matched_record.get("center_name", "-")
            city = matched_record.get("city", "-")
            state = matched_record.get("state", "-")
            zone = matched_record.get("zone", "-")
            center_status = matched_record.get("center_status", "-")
            site_id = matched_record.get("site_id", f"Site_{tc_code}")
            host = matched_record.get("host", "localhost")
            port = matched_record.get("port", "21")

            self.tree.insert(
                "",
                "end",
                iid="tc_target",
                values=(sr_no, tc_code, center_name, city, state, zone, center_status, site_id, host, port, "Ready")
            )

            self.start_btn.config(state="normal", bg="#16a34a", fg="white")
            self.progress_lbl.config(text=f"Status: Selected TC {tc_code} - Ready to execute", fg="#0ea5e9")
            self.progress_bar["value"] = 0
            self.log(f"SUCCESS: Fetched details for TC Code '{tc_code}' (Center: {center_name}, Status: {center_status}).")

        except Exception as e:
            self.log(f"ERROR reading workbook details: {e}")
            messagebox.showerror("File Error", f"Unable to read Excel workbook details:\n{e}")

    def start_automation(self):
        if not self.selected_record:
            messagebox.showwarning("No Selection", "Please fetch a valid TC Code before initiating execution.")
            return

        if self.is_running:
            return

        self.is_running = True
        self.start_btn.config(state="disabled", bg="#64748b", fg="white", text="⏳ Provisioning Target TC...")

        threading.Thread(target=self._run_selenium_process, daemon=True).start()

    def _update_step_status(self, status_val, progress_pct, status_color="#0ea5e9"):
        self.after(0, lambda: self._apply_step_update(status_val, progress_pct, status_color))

    def _apply_step_update(self, status_val, progress_pct, status_color):
        if self.tree.exists("tc_target"):
            vals = list(self.tree.item("tc_target", "values"))
            vals[10] = status_val
            self.tree.item("tc_target", values=vals)
        self.progress_lbl.config(text=f"Status: {status_val}", fg=status_color)
        self.progress_bar["value"] = progress_pct

    def _wait_for_modal_to_close(self, driver, timeout=15):
        end_time = time.time() + timeout
        while time.time() < end_time:
            modals = driver.find_elements(By.CSS_SELECTOR, "div.modal.in, div.bootbox.in, div.modal-backdrop")
            if not [m for m in modals if m.is_displayed()]:
                time.sleep(0.3)
                return True
            time.sleep(0.3)
        return False

    def _js_set_input(self, driver, element, text):
        driver.execute_script(
            "arguments[0].value = arguments[1];"
            "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));"
            "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
            element, str(text)
        )

    def _safe_send_keys(self, driver, element, text):
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        try:
            element.click()
            element.clear()
            element.send_keys(str(text))
        except Exception:
            self._js_set_input(driver, element, text)

    # --- SELENIUM WORKFLOW ---
    def _run_selenium_process(self):
        rec = self.selected_record
        site_id = rec.get("site_id") or f"Site_{rec.get('tc_code', 'Target')}"

        self.log(f"=== Initiating Execution Pipeline for TC Code: {rec.get('tc_code')} ===")

        options = webdriver.ChromeOptions()
        driver = webdriver.Chrome(options=options)
        driver.maximize_window()
        wait = WebDriverWait(driver, 12)

        try:
            # STEP 1: Site Creation
            self.log("Step 1/2: Creating Site Profile...")
            self._update_step_status("Step 1/2: Site Creation in Progress", 30, "#0ea5e9")
            
            try:
                self._process_single_site(driver, wait, rec)
                self.log("Step 1/2 Completed: Site Profile configured successfully.")
            except Exception as site_err:
                self.log(f"CRITICAL ERROR in Step 1 (Site Creation): {site_err}")
                self.log("HARD STOP: Pipeline aborted due to Step 1 failure. HotFolders will NOT be executed.")
                self._update_step_status("FAILED at Step 1 (Site)", 30, "#ef4444")
                messagebox.showerror(
                    "Pipeline Interrupted", 
                    f"Pipeline aborted: Site creation failed for TC Code '{rec.get('tc_code')}'.\n\nReason: {site_err}"
                )
                return

            # STEP 2: HotFolder Creation
            self.log("Step 2/2: Provisioning HotFolders...")
            self._update_step_status("Step 2/2: HotFolders Provisioning in Progress", 75, "#0ea5e9")

            try:
                for cfg in HOTFOLDER_CONFIGS:
                    hf_name = f"{site_id}_{cfg['suffix']}"
                    self._process_single_hotfolder(driver, wait, hf_name, cfg['path'])
                self.log("Step 2/2 Completed: HotFolders provisioned successfully.")
            except Exception as hf_err:
                self.log(f"CRITICAL ERROR in Step 2 (HotFolder Creation): {hf_err}")
                self._update_step_status("FAILED at Step 2 (HotFolder)", 75, "#ef4444")
                messagebox.showerror(
                    "Pipeline Interrupted", 
                    f"Pipeline aborted: HotFolder creation failed.\n\nReason: {hf_err}"
                )
                return

            # Completed Pipeline
            self._update_step_status("Completed: Provisioning Successful", 100, "#22c55e")
            self.log(f"=== Pipeline Successfully Completed for TC Code {rec.get('tc_code')} ===")
            messagebox.showinfo("Pipeline Success", f"Target TC Code '{rec.get('tc_code')}' provisioned successfully across all steps!")

        except Exception as e:
            self.log(f"FATAL Automation Error: {e}\n{traceback.format_exc()}")
            self._update_step_status("Fatal Automation Error", 0, "#ef4444")
            messagebox.showerror("Execution Error", f"An unexpected error occurred during execution:\n{e}")
        finally:
            driver.quit()
            self.is_running = False
            self.after(0, lambda: self.start_btn.config(state="normal", bg="#16a34a", fg="white", text="▶ Start Provisioning Pipeline for Selected TC"))

    def _process_single_site(self, driver, wait, row):
        """Site Creation with SSL handling, existence check, and warning dialog handling."""
        driver.get(TARGET_URL)
        time.sleep(1)

        site_tab = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(text(),'Sites')] | //li[contains(.,'Sites')] | //span[contains(text(),'Sites')]")
        ))
        driver.execute_script("arguments[0].click();", site_tab)
        time.sleep(0.8)

        site_id_val = row.get("site_id") or f"Site_{row.get('tc_code')}"

        existing_sites = driver.find_elements(By.XPATH, f"//td[contains(text(),'{site_id_val}')]")
        if existing_sites:
            self.log(f"Site '{site_id_val}' already exists. Skipping creation step.")
            return

        new_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[@id='btnNewSite' or contains(., 'New')]")
        ))
        driver.execute_script("arguments[0].click();", new_btn)
        time.sleep(1)

        site_id_input = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='newsite-id']")))
        self._safe_send_keys(driver, site_id_input, site_id_val)

        host_val = row.get("host") or "localhost"
        host_input = driver.find_element(By.XPATH, "//*[@id='newsite-host']")
        self._safe_send_keys(driver, host_input, host_val)

        port_val = str(int(float(row["port"]))) if row.get("port") and str(row.get("port")).replace('.','').isdigit() else "21"
        port_input = driver.find_element(By.XPATH, "//*[@id='newsite-port']")
        self._safe_send_keys(driver, port_input, port_val)

        user_name = row.get("fc_user_name") or row.get("username")
        if user_name:
            user_input = driver.find_element(By.XPATH, "//*[@id='newsite-username']")
            self._safe_send_keys(driver, user_input, str(user_name))

        password = row.get("fc_password") or row.get("password")
        if password:
            pass_input = driver.find_element(By.XPATH, "//*[@id='newsite-password']")
            self._safe_send_keys(driver, pass_input, str(password))

        # --- SSL CHECKBOX HANDLING ---
        ssl_val = row.get("ssl") if row.get("ssl") is not None else row.get("use_ssl")
        if ssl_val is not None and str(ssl_val).strip() != "":
            if isinstance(ssl_val, bool):
                target_state = ssl_val
            else:
                target_state = str(ssl_val).strip().lower() in ("true", "1", "yes", "t")

            try:
                ssl_element = wait.until(EC.presence_of_element_located((By.ID, "newsite-useSSL")))
                current_state = ssl_element.is_selected()

                self.log(f"SSL Target State: {target_state} | Current State: {current_state}")

                if current_state != target_state:
                    driver.execute_script(
                        """
                        var el = arguments[0];
                        el.click();
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        """,
                        ssl_element
                    )
                    time.sleep(0.3)

                    if ssl_element.is_selected() != target_state:
                        try:
                            label_element = driver.find_element(
                                By.XPATH, "//label[@for='newsite-useSSL'] | //input[@id='newsite-useSSL']/parent::label"
                            )
                            driver.execute_script("arguments[0].click();", label_element)
                        except Exception:
                            pass

                    self.log(f"Successfully updated SSL checkbox state to: {target_state}")
            except Exception as ssl_err:
                self.log(f"Warning: Failed to set SSL setting: {ssl_err}")

        apply_css = "button.btn.btn-primary.btn-apply.newSiteSubmitButton"
        apply_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, apply_css)))
        driver.execute_script("arguments[0].click();", apply_btn)

        time.sleep(1.5)
        try:
            confirm_btns = driver.find_elements(
                By.XPATH, 
                "//button[contains(text(),'Save') and not(contains(@class,'newSiteSubmitButton'))] | //div[contains(@class,'modal')]//button[contains(text(),'Save')]"
            )
            for btn in confirm_btns:
                if btn.is_displayed():
                    driver.execute_script("arguments[0].click();", btn)
                    self.log("Clicked confirmation 'Save anyway' button on warning dialog.")
                    break
        except Exception:
            pass

        time.sleep(1)
        if not self._wait_for_modal_to_close(driver, timeout=10):
            raise Exception("Site dialog modal failed to close after submission.")

    def _process_single_hotfolder(self, driver, wait, hf_id, path):
        """HotFolder Creation with existence check."""
        hotfolder_tab = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#hotfolders > span")))
        driver.execute_script("arguments[0].click();", hotfolder_tab)
        time.sleep(0.6)

        existing_hf = driver.find_elements(By.XPATH, f"//td[contains(text(),'{hf_id}')]")
        if existing_hf:
            self.log(f"HotFolder '{hf_id}' already exists. Skipping creation step.")
            return

        new_hf_btn = wait.until(EC.element_to_be_clickable((By.ID, "btnNewHotFolder")))
        driver.execute_script("arguments[0].click();", new_hf_btn)

        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.newHotFolderDialog.in")))
        time.sleep(0.3)

        hf_id_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.newHotFolderDialog.in #hotfolder-id")))
        self._safe_send_keys(driver, hf_id_input, hf_id)

        location_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.newHotFolderDialog.in #hotfolder-location")))
        self._safe_send_keys(driver, location_input, path)

        apply_hf_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.newHotFolderDialog.in button.hotFolderSubmitButton")))
        driver.execute_script("arguments[0].click();", apply_hf_btn)

        time.sleep(1)
        if not self._wait_for_modal_to_close(driver, timeout=10):
            raise Exception(f"HotFolder modal for '{hf_id}' failed to close after submission.")


if __name__ == "__main__":
    app = FileCatalystAutomatorGUI()
    app.mainloop()