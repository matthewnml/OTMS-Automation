import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re
import glob
import time
from datetime import datetime
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException

# ------------- Helpers: data loading -------------
def load_table(path, sheet_name=None):
    ext = os.path.splitext(path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        return pd.read_excel(path, sheet_name=sheet_name)
    elif ext in [".csv"]:
        return pd.read_csv(path)
    else:
        raise ValueError("Unsupported file type. Use CSV/XLSX.")

def norm_cols(df):
    # normalize column names: lower, strip spaces, collapse whitespace
    df = df.copy()
    df.columns = (
        df.columns.str.strip()
                  .str.replace(r"\s+", " ", regex=True)
                  .str.lower()
    )
    return df

def wait_dom_idle(driver, timeout=8):
    """Basic wait for DOM to be 'complete' (helps after postbacks)."""
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

def fill_text_with_retry(driver, label, value, retries=3):
    """Type into a text/textarea paired with a label, handling stale references."""
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            el = by_label_input(driver, label, timeout=6)
            driver.execute_script("arguments[0].removeAttribute('readonly');", el)
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            el.clear()
            el.send_keys(value)
            # verify
            v = driver.execute_script("return arguments[0].value;", el)
            if str(v).strip() == str(value).strip():
                return True
        except (StaleElementReferenceException, TimeoutException) as e:
            last_err = e
            time.sleep(0.2)
            wait_dom_idle(driver, timeout=6)
            continue
        except Exception as e:
            last_err = e
            break
    print(f"[WARN] {label}: {last_err}")
    return False

def select_with_retry(driver, label, value, retries=3):
    """Pick an option by visible text with staleness retry + fuzzy match."""
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            sel = by_label_select(driver, label, timeout=6)
            try:
                sel.select_by_visible_text(value)
            except Exception:
                options = [o.text.strip() for o in sel.options]
                match = next((o for o in options if o.lower() == value.lower()), None)
                if not match:
                    match = next((o for o in options if value.lower() in o.lower()), None)
                if match:
                    sel.select_by_visible_text(match)
            # verify selection
            chosen = sel.first_selected_option.text.strip()
            if chosen:
                return True
        except (StaleElementReferenceException, TimeoutException) as e:
            last_err = e
            time.sleep(0.2)
            wait_dom_idle(driver, timeout=6)
            continue
        except Exception as e:
            last_err = e
            break
    print(f"[WARN] {label}: {last_err}")
    return False
# ------------- Selenium locator helpers -------------
def by_label_input(driver, label_text, timeout=8):
    """
    Find an <input> that is visually paired with a label text on the page.
    Uses an XPath that looks for the label text then finds the following input.
    Works with the OTMS layout in the screenshot (table-based).
    """
    L = _xpath_literal(label_text)
    xpaths = [
        f"//tr[td[normalize-space()={L}]]/td[position()=2]//input[not(@type='hidden')]",
        f"//tr[td[normalize-space()={L}]]/td[position()=2]//textarea",
        f"//*[normalize-space()={L}]/following::input[not(@type='hidden')][1]",
        f"//*[normalize-space()={L}]/following::textarea[1]",
    ]
    end_time = time.time() + timeout
    last_err = None
    for xp in xpaths:
        try:
            el = WebDriverWait(driver, max(1, int(end_time - time.time()))).until(
                EC.presence_of_element_located((By.XPATH, xp))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            return el
        except Exception as e:
            last_err = e
    raise RuntimeError(f"Could not locate input/textarea for label: {label_text} ({last_err})")


def by_label_select(driver, label_text, timeout=8):
    L = _xpath_literal(label_text)
    xpaths = [
        f"//tr[td[normalize-space()={L}]]/td[position()=2]//select",
        f"//*[normalize-space()={L}]/following::select[1]",
    ]
    end_time = time.time() + timeout
    last_err = None
    for xp in xpaths:
        try:
            el = WebDriverWait(driver, max(1, int(end_time - time.time()))).until(
                EC.presence_of_element_located((By.XPATH, xp))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            return Select(el)
        except Exception as e:
                last_err = e
    raise RuntimeError(f"Could not locate select for label: {label_text} ({last_err})")

def click_button_with_text(driver, text, timeout=10):
    xp = f"//button[normalize-space()='{text}' or @value='{text}' or @id=//label[normalize-space()='{text}']/@for]"
    try:
        btn = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, xp))
        )
        btn.click()
        return True
    except Exception:
        return False

def _norm(s: str) -> str:
    # normalize for case/space/extra punctuation tolerant matching
    return re.sub(r"\s+", " ", str(s).strip().lower())

def _xpath_literal(s: str) -> str:
    """Return an XPath string literal for arbitrary text (handles quotes)."""
    if "'" not in s:
        return f"'{s}'"
    if '"' not in s:
        return f'"{s}"'
    # both quotes present -> concat('a', "'", 'b', '"', 'c')
    parts = s.split("'")
    return "concat(" + ", \"'\", ".join([f"'{p}'" for p in parts]) + ")"

def find_person_folder(base_dir: str, person_name: str) -> str | None:
    """Return the full path to the subfolder whose name matches person_name (case-insensitive, space-insensitive)."""
    target = _norm(person_name)
    if not os.path.isdir(base_dir):
        return None
    for entry in os.listdir(base_dir):
        full = os.path.join(base_dir, entry)
        if os.path.isdir(full) and _norm(entry) == target:
            return full
    # fallback: fuzzy contains
    for entry in os.listdir(base_dir):
        full = os.path.join(base_dir, entry)
        if os.path.isdir(full) and target in _norm(entry):
            return full
    return None

def find_pdf_by_prefix(folder: str, code_prefix: str, person_name: str) -> str | None:
    """
    Look for patterns like:
      '002 AHR KHR MIN.pdf'  (exact)
      '002_AHR KHR MIN.pdf'  (underscore)
      '002-ahr khr min.pdf'  (hyphen, case-insensitive)
    """
    if not folder or not os.path.isdir(folder):
        return None

    # Build multiple glob patterns
    base = re.sub(r'[\\/:*?"<>|]', '', person_name).strip()  # remove illegal filename chars
    candidates = []
    for sep in [" ", "_", "-"]:
        candidates.append(os.path.join(folder, f"{code_prefix}{sep}{base}.pdf"))
    # also accept any pdf starting with prefix in case name has minor variance
    candidates.extend(glob.glob(os.path.join(folder, f"{code_prefix}*.pdf")))

    # Return the first case-insensitive match that also contains the name (preferred)
    lname = _norm(base)
    for p in candidates:
        if os.path.isfile(p) and p.lower().endswith(".pdf"):
            if lname in _norm(os.path.basename(p)):
                return p

    # fallback: any file with the prefix
    for p in candidates:
        if os.path.isfile(p) and p.lower().endswith(".pdf"):
            return p

    return None

# ------------- Core fill function -------------
def fill_otms_form(driver, url, row, base_dir = None, status_cb=lambda s: None, pause_check=lambda: None,stop_check=lambda: False):
    # Ensure we are at the target page
    if url:
        status_cb("Navigating to target page…")
        pause_check()
        driver.get(url)

    wait = WebDriverWait(driver, 30)

    # Map our canonical labels -> (type, source column)
    # Adjust source column names on the right to match your sheet.
    # The left keys should match the visible labels on the page.
    mapping = {
        "Name as in IC/Passport" :("text", "name in ic/passport"),
        "Sex": ("select", "sex"),
        "Country of Birth": ("select", "nationality"),
        "State/Province of Birth": ("select", "province of birth"),
        "Place of Birth": ("text", "place of township"),
        "IC Number": ("text", "ic number"),
        "Passport Number": ("text", "passport number"),
        "Date of Issue": ("text", "passport date of issue"),
        "Date of Expiry": ("text", "passport date of expiry"),
        "Date of Birth": ("text", "date of birth"),
        "Father's Name": ("text", "father name"),
        "Mother's Name": ("text", "mother name"),
        "Current Address": ("text", "current address"),
        "Awarded Institute": ("text", "school name of highest qualification"),
        "Year": ("select", "year of graduation"),
        
    }
    
    #Name from excel row
    person_name = None
    for col in ["name in ic/passport", "name in ic", "name"]:
        if col in row.index and pd.notna(row[col]) and str(row[col]).strip():
            person_name = str(row[col]).strip()
            break
    # Upload PDFs from folder structure
    if base_dir and person_name:
        status_cb("Locating person folder for PDFs…")
        pause_check()
        person_folder = find_person_folder(base_dir, person_name)
        if not person_folder:
            print(f"[WARN] No folder found for '{person_name}' under {base_dir}")
        else:
            uploads = [
                ("Upload Identity Card (IC)", "002"),   # IC = 002
                ("Upload passport", "003"),   # Passport = 003
                ("Upload highest qualification or most relevent certification", "004"),  # School Cert = 004
            ]
            for label, code in uploads:
                if stop_check():
                    status_cb("Stopped by user.")
                    return
                status_cb(f"Searching PDF: {code} for {person_name}…")
                pause_check()
                
                fpath = find_pdf_by_prefix(person_folder, code, person_name)
                if not fpath:
                    print(f"[WARN] {label}: could not find a '{code} ... .pdf' in {person_folder}")
                    continue

                try:
                    status_cb(f"Uploading: {label}")
                    pause_check()
                    file_input = WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((
                            By.XPATH,
                            f"//td[normalize-space()='{label}']/following::input[@type='file'][1]"
                        ))
                    )
                    file_input.send_keys(os.path.abspath(fpath))
                    # Click nearest upload button if present
                    try:
                        upload_btn = file_input.find_element(By.XPATH, "following::input[@type='submit' or @type='button'][1]")
                        upload_btn.click()
                    except Exception:
                        click_button_with_text(driver, "Upload")
                    print(f"[INFO] Uploaded {label}: {fpath}")
                except Exception as e:
                    print(f"[WARN] Could not upload {label}: {e}")
    else:
        print("[INFO] Skipping PDF uploads: base_dir or person_name not set.")
        
    # -------- Fill all text/selects --------
    for label, (field_type, colname) in mapping.items():
        if stop_check():
            status_cb("Stopped by user.")
            return
        if colname not in row.index:
            continue

        raw = row[colname]
        val = "" if pd.isna(raw) else str(raw).strip()
        if not val:
            continue

        status_cb(f"Filling: {label}")
        pause_check()

        try:
            if field_type == "text":
                ok = fill_text_with_retry(driver, label, value=val)
            elif field_type == "select":
                ok = select_with_retry(driver, label, value=val)
            else:
                ok = True  # unknown field type, skip

            if not ok:
                print(f"[WARN] Could not set {label} -> {val}")
        except Exception as e:
            print(f"[WARN] {label}: {e}")

            
    status_cb("All fields filled. Please review and submit on the site.")
    messagebox.showinfo("Done", "Form fields filled. Please review and submit on the site if needed.")

# ------------- Driver setup -------------
def make_driver(attach_existing, debug_port=9222):
    opts = Options()
    if attach_existing:
        # Attach to an already opened Chrome launched with --remote-debugging-port
        opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{debug_port}")
        # Use matching driver
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=opts)
    else:
        # Fresh Chrome (no login persisted; you must log in in this session)
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=opts)
        driver.maximize_window()
    return driver

# ------------- GUI -------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("OTMS Pre-Enrolment Auto-Fill")
        self.geometry("720x400")
        self.resizable(True,True)
        
        self.stop_event = threading.Event()
        self.base_dir = tk.StringVar()  # root like: C:\path\to\ESF
        self.file_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.srno = tk.StringVar()
        self.url = tk.StringVar(value="https://otms.bca.gov.sg/PreEnrolment/NewPreEnrolment.aspx")
        self.attach_mode = tk.StringVar(value="attach")  # 'attach' or 'new'
        self.debug_port = tk.StringVar(value="9222")

        self.pause_event = threading.Event()   # set = running, clear = paused
        self.pause_event.set()

        self.status_var = tk.StringVar(value="Ready.")

        self._build()

    def _build(self):
        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, padx=12, pady=14)

        # File row
        ttk.Label(frm, text="Data file (CSV/XLSX):").grid(row=0, column=0, sticky="e", **pad)
        ttk.Entry(frm, textvariable=self.file_path, width=60).grid(row=0, column=1, **pad)
        ttk.Button(frm, text="Browse", command=self.browse_file).grid(row=0, column=2, **pad)

        # PDF folder
        ttk.Label(frm, text="Base folder (ESF):").grid(row=1, column=0, sticky="e", **pad)
        ttk.Entry(frm, textvariable=self.base_dir, width=60).grid(row=1, column=1, **pad)
        ttk.Button(frm, text="Browse", command=self.browse_folder).grid(row=1, column=2, **pad)

        # Sheet name (optional)
        ttk.Label(frm, text="Sheet name (if Excel):").grid(row=2, column=0, sticky="e", **pad)
        ttk.Entry(frm, textvariable=self.sheet_name, width=30).grid(row=2, column=1, sticky="w", **pad)

        # Sr.No
        ttk.Label(frm, text="Sr.No to process:").grid(row=3, column=0, sticky="e", **pad)
        ttk.Entry(frm, textvariable=self.srno, width=20).grid(row=3, column=1, sticky="w", **pad)

        # URL
        ttk.Label(frm, text="Target URL:").grid(row=4, column=0, sticky="e", **pad)
        ttk.Entry(frm, textvariable=self.url, width=60).grid(row=4, column=1, **pad)

        # Mode
        mode_frame = ttk.Frame(frm)
        mode_frame.grid(row=5, column=1, sticky="w", **pad)
        ttk.Radiobutton(mode_frame, text="Attach to existing Chrome (9222)", variable=self.attach_mode, value="attach").pack(side="left")
        ttk.Radiobutton(mode_frame, text="Start new Chrome", variable=self.attach_mode, value="new").pack(side="left")
        ttk.Label(frm, text="Debug port:").grid(row=5, column=0, sticky="e", **pad)
        ttk.Entry(frm, textvariable=self.debug_port, width=10).grid(row=5, column=2, sticky="e", **pad)

        # Help text
        help_txt = (
            "Notes:\n"
            "• Your file must include a 'Sr.No' column to select the row.\n"
            "• Optional columns for uploads: IC_PDF_Path, Passport_PDF_Path (PDF ≤ 5MB).\n"
            "• For attach mode, start Chrome with --remote-debugging-port=9222 and log in first."
        )
        ttk.Label(frm, text=help_txt, foreground="#444").grid(row=6, column=0, columnspan=3, sticky="w", **pad)

        # Pause
        ttk.Button(frm, text="Fill Form", command=self.run).grid(row=7, column=1, sticky="e", **pad)
        self.pause_btn = ttk.Button(frm, text="Pause", command=self.on_pause_resume, state="disabled")
        self.pause_btn.grid(row=7, column=2, sticky="w", **pad)
        self.stop_btn = ttk.Button(frm, text="Stop", command=self.on_stop, state="disabled")
        self.stop_btn.grid(row=7, column=3, sticky="w", **pad)


        # Status line (full width)
        ttk.Label(frm, textvariable=self.status_var, foreground="#0066cc").grid(
            row=8, column=0, columnspan=3, sticky="w", **pad
)

        for i in range(3):
            frm.grid_columnconfigure(i, weight=1)

    def browse_folder(self):
        path = filedialog.askdirectory(title="Select ESF base folder")
        if path:
            self.base_dir.set(path)

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Select data file",
            filetypes=[("Spreadsheets", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if path:
            self.file_path.set(path)

    def run(self):
        try:
            if not self.file_path.get():
                messagebox.showerror("Error", "Please choose a CSV/XLSX file.")
                return
            if not self.srno.get().strip():
                messagebox.showerror("Error", "Please enter a Sr.No value.")
                return
            if not self.base_dir.get().strip():
                messagebox.showerror("Error", "Please choose the ESF base folder.")
                return

            # Load data
            df = load_table(self.file_path.get(), sheet_name=self.sheet_name.get() or None)
            df = norm_cols(df)

            if "sr.no" not in df.columns:
                messagebox.showerror("Error", "The file must contain a 'Sr.No' column.")
                return

            # Find the row by Sr.No (string-insensitive)
            sr_target = str(self.srno.get()).strip()
            match_df = df[df["sr.no"].astype(str).str.strip() == sr_target]
            if match_df.empty:
                messagebox.showerror("Not found", f"No row with Sr.No = {sr_target}")
                return

            row = match_df.iloc[0]

            self.pause_btn.configure(state="normal", text="Pause")
            self.stop_btn.configure(state="normal")
            self.stop_event.clear()
            self.set_status(f"Starting… Sr.No {sr_target}")

            # Launch selenium on a separate thread to keep GUI responsive
            threading.Thread(
                target=self._selenium_task,
                args=(row,self.base_dir.get().strip()),
                daemon=True
            ).start()

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _selenium_task(self, row, base_dir):
        try:
            attach = self.attach_mode.get() == "attach"
            port = int(self.debug_port.get() or "9222")
            self.set_status("Launching/attaching Chrome…")
            driver = make_driver(attach_existing=attach, debug_port=port)
            fill_otms_form(driver, self.url.get().strip(), row, base_dir=base_dir, status_cb=self.set_status,
            pause_check=self.wait_if_paused, stop_check=lambda: self.stop_event.is_set())
            
            self.set_status("Done. Review and submit on the site.")
        except Exception as e:
            self.set_status(f"Error: {e}")
            messagebox.showerror("Selenium error", str(e))
        finally:
        # ✅ Always disable pause/stop buttons at the end
            self.after(0, lambda: [
                self.pause_btn.configure(state="disabled", text="Pause"),
                self.stop_btn.configure(state="disabled")
            ])
            
    def on_pause_resume(self):
        # Toggle pause
        if self.pause_event.is_set():
            self.pause_event.clear()
            self.set_status("Paused.")
            self.pause_btn.configure(text="Resume")
        else:
            self.pause_event.set()
            self.set_status("Resumed.")
            self.pause_btn.configure(text="Pause")

    def set_status(self, msg: str):
        # Update from worker thread safely
        self.after(0, lambda: self.status_var.set(msg))

    def wait_if_paused(self):
        # Called by worker to yield while paused
        while not self.pause_event.is_set():
            time.sleep(0.1)
    
    def on_stop(self):
        self.stop_event.set()
        self.set_status("Stopping… please wait.")

if __name__ == "__main__":
    App().mainloop()

