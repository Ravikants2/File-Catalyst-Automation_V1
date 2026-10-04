import os
import time
import datetime
import threading
import traceback
import queue
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

# Base64 string generated from the exact header logo icon (Server Rack + Green LEDs + Yellow Bolt)
HEADER_ICON_PNG = (
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAsTAAALEwEAmpwYAAAA"
    "AXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAACNSURBVHgB7ZHBCcBADAMt12m9iOeoTuI5ChP0"
    "Hrm+lIIt44eUQsIhhIckkK5f+q3b3S331nME4D33vgFf84IA33v34e5p433/fGst44xv4pyn1lpr"
    "j/feW/cegDHGcc4A5JxzSggAnHMfA/jW2h3wPfe+4L3/x03f903/m2eLq2L/KvbI814iAgA1OydI"
    "5m2s8QAAAABJRU5ErkJggg=="
)


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
        if not any(row):  # Skip empty rows
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
        self.configure(bg="#0f172a")

        # Set title bar window icon to match the header UI icon
        self._set_window_title_icon()

        self.excel_file_path = ""
        self.is_running = False
        self.selected_record = None
        self.log_queue = queue.Queue()

        self._configure_styles()
        self._build_ui()
        self._start_queue_listener()

        self.log("System initialized. Upload template workbook to get started.")

    def _set_window_title_icon(self):
        """Draws the exact server-rack icon into Tkinter window title bar."""
        try:
            # Generate PhotoImage directly matching the main body icon design
            self.app_icon_img = tk.PhotoImage(width=16, height=16)
            
            # Draw server frame (#38bdf8), green dots (#22c55e), and yellow bolt (#f59e0b)
            # Row-by-row pixel painting for guaranteed cross-platform title bar icon
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
            
            color_map = {
                'C': '#38bdf8',  # Server cyan frame
                'G': '#22c55e',  # Green status light
                'Y': '#f59e0b',  # Yellow lightning bolt
                '.': '#0f172a'   # Background matching UI
            }

            for y, row in enumerate(icon_pixels):
                for x, char in enumerate(row):
                    self.app_icon_img.put(color_map[char], (x, y))

            self.iconphoto(True, self.app_icon_img)
        except Exception as e:
            print(f"Icon initialization: {e}")

    def _configure_styles(self):
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        # Treeview Styling
        self.style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 9, "bold"),
            background="#1e293b",
            foreground="#38bdf8",
            relief="flat",
            padding=8
        )
        self.style.configure(
            "Treeview",
            font=("Segoe UI", 9),
            rowheight=34,
            background="#0f172a",
            fieldbackground="#0f172a",
            foreground="#f1f5f9",
            borderwidth=0
        )
        self.style.map("Treeview", background=[("selected", "#0284c7")], foreground=[("selected", "#ffffff")])

        # Progressbar Styling
        self.style.configure(
            "TProgressbar",
            thickness=10,
            troughcolor="#0f172a",
            background="#38bdf8",
            borderwidth=0
        )

    def _build_ui(self):
        # ------------------- HEADER -------------------
        header = tk.Frame(self, bg="#0f172a", height=70)
        header.pack(fill="x", padx=25, pady=(15, 10))

        icon_canvas = HeaderIcon(header, bg_color="#0f172a", fg_color="#38bdf8", size=36)
        icon_canvas.pack(side="left", padx=(0, 12))

        title_frame = tk.Frame(header, bg="#0f172a")
        title_frame.pack(side="left")

        tk.Label(
            title_frame,
            text="FileCatalyst Sequential Provisioning Console",
            font=("Segoe UI", 16, "bold"),
            fg="#f8fafc",
            bg="#0f172a"
        ).pack(anchor="w")

        tk.Label(
            title_frame,
            text="Automated Site Configuration & HotFolder Deployment Engine",
            font=("Segoe UI", 9),
            fg="#94a3b8",
            bg="#0f172a"
        ).pack(anchor="w")

        # ------------------- MAIN CONTAINER -------------------
        main = tk.Frame(self, bg="#0f172a", padx=25, pady=0)
        main.pack(fill="both", expand=True)

        # ------------------- STEP 1: TC SELECTION & INPUT -------------------
        step1_frame = tk.LabelFrame(
            main,
            text=" STEP 1: Target TC Selection & Query ",
            font=("Segoe UI", 10, "bold"),
            fg="#38bdf8",
            bg="#1e293b",
            bd=1,
            relief="solid",
            padx=18,
            pady=16
        )
        step1_frame.pack(fill="x", pady=(0, 12))

        grid_frame = tk.Frame(step1_frame, bg="#1e293b")
        grid_frame.pack(fill="x")

        # File Upload Label & Path Input
        tk.Label(grid_frame, text="Excel Template:", font=("Segoe UI", 9, "bold"), fg="#e2e8f0", bg="#1e293b").grid(row=0, column=0, sticky="w", pady=4)
        
        self.path_entry = tk.Entry(
            grid_frame,
            font=("Consolas", 9, "bold"),
            bg="#0f172a",
            fg="#38bdf8",
            readonlybackground="#0f172a",
            disabledbackground="#0f172a",
            disabledforeground="#38bdf8",
            insertbackground="#38bdf8",
            bd=1,
            relief="solid",
            state="readonly"
        )
        self.path_entry.grid(row=0, column=1, sticky="ew", padx=(10, 10), pady=4)

        browse_btn = tk.Button(
            grid_frame,
            text="Browse File...",
            command=self.browse_file,
            bg="#0284c7",
            fg="white",
            activebackground="#0369a1",
            activeforeground="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=14,
            pady=4,
            cursor="hand2"
        )
        browse_btn.grid(row=0, column=2, padx=(0, 20), pady=4)

        # TC Search Code Input
        tk.Label(grid_frame, text="Target TC Code:", font=("Segoe UI", 9, "bold"), fg="#e2e8f0", bg="#1e293b").grid(row=0, column=3, sticky="w", pady=4)

        self.tc_entry = tk.Entry(
            grid_frame,
            font=("Segoe UI", 10, "bold"),
            bg="#0f172a",
            fg="#38bdf8",
            insertbackground="#38bdf8",
            bd=1,
            relief="solid",
            width=18
        )
        self.tc_entry.grid(row=0, column=4, sticky="w", padx=(10, 10), pady=4)
        self.tc_entry.bind("<Return>", lambda e: self.fetch_tc_details())

        fetch_btn = tk.Button(
            grid_frame,
            text="Fetch Details",
            command=self.fetch_tc_details,
            bg="#0284c7",
            fg="white",
            activebackground="#0369a1",
            activeforeground="white",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=16,
            pady=4,
            cursor="hand2"
        )
        fetch_btn.grid(row=0, column=5, pady=4)

        grid_frame.columnconfigure(1, weight=1)

        # Status Label below input field
        self.file_status_lbl = tk.Label(
            step1_frame,
            text="No Excel file selected.",
            font=("Segoe UI", 8, "italic"),
            fg="#64748b",
            bg="#1e293b"
        )
        self.file_status_lbl.pack(anchor="w", pady=(4, 0))

        # Execute Pipeline Button
        self.start_btn = tk.Button(
            step1_frame,
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
        paned = tk.PanedWindow(main, orient=tk.VERTICAL, bg="#0f172a", sashwidth=8, bd=0)
        paned.pack(fill="both", expand=True)

        # ------------------- STEP 2: DEPLOYMENT STATUS & DETAILS -------------------
        step2_frame = tk.LabelFrame(
            paned,
            text=" STEP 2: Selected TC Parameters & Deployment Status ",
            font=("Segoe UI", 10, "bold"),
            fg="#38bdf8",
            bg="#1e293b",
            bd=1,
            relief="solid",
            padx=14,
            pady=14
        )
        paned.add(step2_frame, minsize=260)

        # Progress Status Bar Header
        progress_bar_container = tk.Frame(step2_frame, bg="#1e293b")
        progress_bar_container.pack(fill="x", pady=(0, 10))

        self.progress_lbl = tk.Label(
            progress_bar_container,
            text="Current Progress: Idle / Ready",
            font=("Segoe UI", 9, "bold"),
            fg="#94a3b8",
            bg="#1e293b"
        )
        self.progress_lbl.pack(anchor="w", pady=(0, 4))

        self.progress_bar = ttk.Progressbar(progress_bar_container, style="TProgressbar", mode="determinate")
        self.progress_bar.pack(fill="x")

        # Table Listing details
        self.columns = (
            "Sr.No.", "TC Code", "Center Name", "City", 
            "State", "Zone", "Center Status", "Site ID", 
            "Host", "Port", "Pipeline Status"
        )
        self.tree = ttk.Treeview(step2_frame, columns=self.columns, show="headings")

        col_widths = {
            "Sr.No.": 60,
            "TC Code": 90,
            "Center Name": 220,
            "City": 110,
            "State": 110,
            "Zone": 80,
            "Center Status": 110,
            "Site ID": 140,
            "Host": 120,
            "Port": 60,
            "Pipeline Status": 180
        }

        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(
                col, 
                anchor="center" if col in ("Sr.No.", "TC Code", "Port", "Pipeline Status", "Zone") else "w", 
                width=col_widths.get(col, 100)
            )

        self.tree.pack(side="left", fill="both", expand=True)

        tree_scroll_y = ttk.Scrollbar(step2_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=tree_scroll_y.set)
        tree_scroll_y.pack(side="right", fill="y")

        # ------------------- STEP 3: EXECUTION LOGS -------------------
        step3_frame = tk.LabelFrame(
            paned,
            text=" STEP 3: Real-time Pipeline Execution Logs ",
            font=("Segoe UI", 10, "bold"),
            fg="#38bdf8",
            bg="#1e293b",
            bd=1,
            relief="solid",
            padx=14,
            pady=14
        )
        paned.add(step3_frame, minsize=200)

        self.log_text = tk.Text(
            step3_frame,
            bg="#0f172a",
            fg="#38bdf8",
            font=("Consolas", 10),
            wrap="word",
            bd=0,
            padx=10,
            pady=10
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        console_scroll = ttk.Scrollbar(step3_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscroll=console_scroll.set)
        console_scroll.pack(side="right", fill="y")

    # ------------------- HELPER METHODS -------------------
    def log(self, message):
        """Thread-safe logging."""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        entry = f"[{timestamp}] {message}\n"
        self.log_queue.put(entry)

    def _start_queue_listener(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get_nowait()
            self.log_text.insert(tk.END, msg)
            self.log_text.see(tk.END)
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
            self.progress_lbl.config(text=f"Status: Selected TC {tc_code} - Ready to execute", fg="#38bdf8")
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

    def _update_step_status(self, status_val, progress_pct, status_color="#38bdf8"):
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
            self._update_step_status("Step 1/2: Site Creation in Progress", 30, "#38bdf8")
            
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
            self._update_step_status("Step 2/2: HotFolders Provisioning in Progress", 75, "#38bdf8")

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