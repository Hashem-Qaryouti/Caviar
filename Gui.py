import threading
import tkinter as tk
from tkinter import ttk, messagebox
import time
import os
import requests
from datetime import datetime

from gopro_recorder import (
    CAM_A, CAM_B, SUBJECTS, ACTIONS, ACCURACY,
    beep, setup_both, start_both, stop_both,
    get_last_file, download_file, download_both, delete_file, wait_for_camera,
    lookup_pair, is_done, mark_done, add_remark, get_save_folder, delete_recording,
    EXCEL_PATH, BASE_FOLDER,
)
from azure_sync import import_excel, export_excel, upload_trial_files


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("GoPro Trial Recorder")
        self.resizable(False, False)
        self.configure(bg="#0f1923")
        self._recording = False
        self._trial_code = None
        self._test_recording = False
        self._test_cancelled = False
        self._log_file = None
        self._record_start_time = None
        self._checklist_imported = False
        self._checklist_exported = False
        self._trials_done_this_session = 0
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ── File Logging ─────────────────────────────────────
    def _open_log_file(self, code):
        """Open a .txt log file in the trial's save folder."""
        try:
            folder = get_save_folder(code)
            os.makedirs(folder, exist_ok=True)
            log_path = os.path.join(folder, f"{code}.txt")
            self._log_file = open(log_path, "a", encoding="utf-8")
            self._log_file.write(
                f"\n{'='*60}\n"
                f"Trial: {code}\n"
                f"Session started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"{'='*60}\n"
            )
            self._log_file.flush()
        except Exception as e:
            self._log_file = None
            self.log(f"  ⚠ Could not open log file: {e}")

    def _close_log_file(self):
        if self._log_file:
            try:
                self._log_file.write("\n")
                self._log_file.close()
            except Exception:
                pass
            self._log_file = None

    def _write_marker(self, marker):
        """Write a special marker line directly to the log file (thread-safe)."""
        if self._log_file:
            try:
                self._log_file.write(f"{marker}\n")
                self._log_file.flush()
            except Exception:
                pass

    # ── UI Builder ────────────────────────────────────────
    def _build_ui(self):
        header = tk.Frame(self, bg="#0f1923")
        header.pack(fill="x", padx=20, pady=(20, 4))
        tk.Label(header, text="●  GoPro Trial Recorder",
                 font=("Helvetica Neue", 15, "bold"),
                 bg="#0f1923", fg="#e8f0fe").pack(side="left")

        card = tk.Frame(self, bg="#1a2535", bd=0, relief="flat")
        card.pack(fill="both", padx=20, pady=10, ipadx=16, ipady=12)

        def label(parent, text):
            tk.Label(parent, text=text, font=("Helvetica Neue", 10),
                     bg="#1a2535", fg="#8ba3c7").pack(anchor="w", pady=(10, 2))

        def dropdown(parent, var, options):
            cb = ttk.Combobox(parent, textvariable=var, values=options,
                              state="readonly", width=28,
                              font=("Helvetica Neue", 11))
            cb.pack(anchor="w")
            return cb

        label(card, "Nurse (Subject)  —  Camera A")
        self.var_nurse = tk.StringVar()
        self.cb_nurse = dropdown(card, self.var_nurse, SUBJECTS)
        self.cb_nurse.bind("<<ComboboxSelected>>", self._on_nurse_selected)

        label(card, "Patient (Partner)  —  Camera B")
        self.var_patient = tk.StringVar()
        self.cb_patient = dropdown(card, self.var_patient, SUBJECTS)
        self.cb_patient.bind("<<ComboboxSelected>>", self._on_selection)

        self.lbl_pair = tk.Label(card, text="Pair: —",
                                 font=("Helvetica Neue", 10, "italic"),
                                 bg="#1a2535", fg="#4fc3f7")
        self.lbl_pair.pack(anchor="w", pady=(4, 0))

        label(card, "Action")
        self.var_action = tk.StringVar()
        dropdown(card, self.var_action, list(ACTIONS.keys())).bind("<<ComboboxSelected>>", self._on_selection)

        label(card, "Accuracy")
        self.var_accuracy = tk.StringVar()
        acc_frame = tk.Frame(card, bg="#1a2535")
        acc_frame.pack(anchor="w", pady=(2, 0))
        self._acc_buttons = {}
        for val in ["Standard", "Fault"]:
            color = "#66bb6a" if val == "Standard" else "#f44336"
            btn = tk.Radiobutton(acc_frame, text=val, variable=self.var_accuracy, value=val,
                                 command=self._on_accuracy_selected,
                                 bg="#263548", fg="#e8f0fe", selectcolor=color,
                                 activebackground="#263548", activeforeground="white",
                                 relief="flat", padx=10, pady=6,
                                 font=("Helvetica Neue", 11, "bold"),
                                 indicatoron=0, bd=0, cursor="hand2")
            btn.pack(side="left", padx=(0, 8))
            self._acc_buttons[val] = (btn, color)

        label(card, "Trial Number")
        self.var_trial = tk.StringVar()
        trial_spin = ttk.Spinbox(card, from_=1, to=20, textvariable=self.var_trial,
                                 width=6, font=("Helvetica Neue", 11))
        trial_spin.pack(anchor="w")
        trial_spin.set(1)
        trial_spin.bind("<ButtonRelease>", self._on_selection)
        trial_spin.bind("<KeyRelease>",    self._on_selection)

        self.lbl_code = tk.Label(card, text="Trial Code: —",
                                 font=("Courier New", 12, "bold"),
                                 bg="#1a2535", fg="#4fc3f7")
        self.lbl_code.pack(anchor="w", pady=(10, 2))

        self.lbl_done = tk.Label(card, text="",
                                 font=("Helvetica Neue", 10),
                                 bg="#1a2535", fg="#f44336")
        self.lbl_done.pack(anchor="w")

        # Buttons
        btn_frame = tk.Frame(self, bg="#0f1923")
        btn_frame.pack(pady=(4, 0), padx=20, fill="x")

        self.btn_start = tk.Button(btn_frame, text="▶  Start Trial",
                                   font=("Helvetica Neue", 12, "bold"),
                                   bg="#424242", fg="#aaaaaa", relief="flat",
                                   padx=20, pady=10, cursor="hand2",
                                   state="disabled", command=self._start_trial)
        self.btn_start.pack(side="left", expand=True, fill="x", padx=(0, 6))

        self.btn_stop = tk.Button(btn_frame, text="■  Stop & Save",
                                  font=("Helvetica Neue", 12, "bold"),
                                  bg="#424242", fg="#aaaaaa", relief="flat",
                                  padx=20, pady=10, cursor="hand2",
                                  state="disabled", command=self._stop_trial)
        self.btn_stop.pack(side="left", expand=True, fill="x")

        self.btn_delete = tk.Button(btn_frame, text="🗑  Delete Recording",
                                    font=("Helvetica Neue", 11, "bold"),
                                    bg="#424242", fg="#aaaaaa", relief="flat",
                                    padx=20, pady=10, cursor="hand2",
                                    state="disabled", command=self._delete_recording)
        self.btn_delete.pack(side="left", expand=True, fill="x", padx=(6, 0))

        # Test row
        test_frame = tk.Frame(self, bg="#0f1923")
        test_frame.pack(pady=(6, 0), padx=20, fill="x")

        self.btn_test = tk.Button(test_frame, text="🔌  Test Cameras (no save)",
                                  font=("Helvetica Neue", 11, "bold"),
                                  bg="#424242", fg="#aaaaaa", relief="flat",
                                  padx=20, pady=8, cursor="hand2",
                                  state="disabled", command=self._test_cameras)
        self.btn_test.pack(side="left", expand=True, fill="x")

        self.btn_test_stop = tk.Button(test_frame, text="■  Stop Test",
                                       font=("Helvetica Neue", 11, "bold"),
                                       bg="#424242", fg="#aaaaaa", relief="flat",
                                       padx=20, pady=8, cursor="hand2",
                                       state="disabled", command=self._stop_test)
        self.btn_test_stop.pack(side="left", expand=True, fill="x", padx=(6, 0))

        # Azure sync row
        sync_frame = tk.Frame(self, bg="#0f1923")
        sync_frame.pack(pady=(6, 0), padx=20, fill="x")

        self.btn_import = tk.Button(sync_frame, text="📥  Import Checklist",
                                    font=("Helvetica Neue", 11, "bold"),
                                    bg="#1a3a1a", fg="#81c784", relief="flat",
                                    padx=20, pady=8, cursor="hand2",
                                    command=self._import_excel)
        self.btn_import.pack(side="left", expand=True, fill="x")

        self.btn_export = tk.Button(sync_frame, text="📤  Export Checklist",
                                    font=("Helvetica Neue", 11, "bold"),
                                    bg="#424242", fg="#aaaaaa", relief="flat",
                                    padx=20, pady=8, cursor="hand2",
                                    state="disabled", command=self._export_excel)
        self.btn_export.pack(side="left", expand=True, fill="x", padx=(6, 0))

        # Session status hint
        self.lbl_session = tk.Label(
            self, bg="#0f1923",
            text="⬇  Import the checklist before recording any trials.",
            font=("Helvetica Neue", 10, "italic"),
            fg="#f0a030"
        )
        self.lbl_session.pack(pady=(6, 0))

        # Log
        log_frame = tk.Frame(self, bg="#0f1923")
        log_frame.pack(fill="both", expand=True, padx=20, pady=(10, 20))
        tk.Label(log_frame, text="Log", font=("Helvetica Neue", 9),
                 bg="#0f1923", fg="#4a6080").pack(anchor="w")
        self.log_box = tk.Text(log_frame, height=10, bg="#0d1117", fg="#c9d1d9",
                               font=("Courier New", 9), relief="flat",
                               state="disabled", wrap="word")
        self.log_box.pack(fill="both", expand=True)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TCombobox",
                        fieldbackground="#0f1923",
                        background="#0f1923",
                        foreground="#ffffff",
                        selectbackground="#1565c0",
                        selectforeground="#ffffff",
                        arrowcolor="#4fc3f7",
                        bordercolor="#4fc3f7",
                        lightcolor="#4fc3f7",
                        darkcolor="#4fc3f7")
        style.map("TCombobox",
                  fieldbackground=[("readonly", "#0f1923")],
                  foreground=[("readonly", "#ffffff")],
                  selectbackground=[("readonly", "#1565c0")])
        style.configure("TSpinbox",
                        fieldbackground="#0f1923",
                        background="#0f1923",
                        foreground="#ffffff",
                        arrowcolor="#4fc3f7",
                        bordercolor="#4fc3f7")

    # ── Helpers ───────────────────────────────────────────
    def log(self, msg):
        """Thread-safe log: always dispatches to the main thread."""
        self.after(0, self._log_main, msg)

    def _log_main(self, msg):
        self.log_box.config(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.config(state="disabled")
        if self._log_file:
            try:
                stripped = msg.strip()
                if not stripped:
                    self._log_file.write("\n")
                else:
                    if "✗" in stripped:
                        tag = "[ERROR]  "
                    elif "⚠" in stripped:
                        tag = "[WARNING]"
                    elif "✓" in stripped or "✅" in stripped:
                        tag = "[OK]     "
                    else:
                        tag = "[INFO]   "
                    ts = datetime.now().strftime("%H:%M:%S")
                    self._log_file.write(f"[{ts}] {tag} {stripped}\n")
                self._log_file.flush()
            except Exception:
                pass

    def _build_code(self):
        nurse    = self.var_nurse.get()
        patient  = self.var_patient.get()
        action   = self.var_action.get()
        accuracy = self.var_accuracy.get()
        trial    = self.var_trial.get()
        if not all([nurse, patient, action, accuracy, trial]):
            return None, None
        if nurse == patient:
            return None, None
        pair = lookup_pair(nurse, patient)
        if not pair:
            return None, None
        code = f"{pair}_{ACTIONS[action]}_{ACCURACY[accuracy]}_T{trial}"
        return code, pair

    def _on_accuracy_selected(self):
        selected = self.var_accuracy.get()
        for val, (btn, color) in self._acc_buttons.items():
            if val == selected:
                btn.config(bg=color, fg="white")
            else:
                btn.config(bg="#263548", fg="#8ba3c7")
        self._on_selection()

    def _on_nurse_selected(self, event=None):
        nurse = self.var_nurse.get()
        filtered = [s for s in SUBJECTS if s != nurse]
        self.cb_patient.config(values=filtered)
        if self.var_patient.get() == nurse:
            self.var_patient.set("")
        self._on_selection()

    def _on_selection(self, event=None):
        code, pair = self._build_code()
        if code:
            self.lbl_code.config(text=f"Trial Code: {code}")
            self.lbl_pair.config(text=f"Pair: {pair}")
            if is_done(code):
                self.lbl_done.config(text="⚠ Already recorded", fg="#f44336")
            else:
                self.lbl_done.config(text="✓ Not yet recorded", fg="#66bb6a")
        else:
            self.lbl_code.config(text="Trial Code: —")
            self.lbl_pair.config(text="Pair: —")
            self.lbl_done.config(text="")

    def _set_controls(self, enabled):
        self.btn_start.config(state="normal" if enabled else "disabled",
                              bg="#1565c0" if enabled else "#424242",
                              fg="white" if enabled else "#aaaaaa")
        self.btn_stop.config(state="disabled" if enabled else "normal",
                             bg="#424242" if enabled else "#c62828",
                             fg="#aaaaaa" if enabled else "white")
        self.btn_test.config(state="normal" if enabled else "disabled",
                             bg="#4a3800" if enabled else "#424242",
                             fg="#ffd54f" if enabled else "#aaaaaa")

    # ── Trial Logic ───────────────────────────────────────
    def _delete_recording(self):
        code, _ = self._build_code()
        if not code:
            messagebox.showwarning("No Trial", "Please select the trial you want to delete.")
            return
        if not messagebox.askyesno("Delete Recording",
                                   f"Delete files and unmark {code} in Excel?\nThis cannot be undone."):
            return
        try:
            report = delete_recording(code)
            self.log(f"\n  🗑 Delete report for {code}:")
            labels = {"_A": "Camera A (Nurse)", "_B": "Camera B (Patient)"}
            all_ok = True
            for suffix, label in labels.items():
                status = report.get(suffix, "missing")
                if status == "deleted":
                    self.log(f"  ✓ {label}: deleted")
                elif status == "missing":
                    self.log(f"  ⚠ {label}: file not found (may never have been saved)")
                    all_ok = False
                else:
                    self.log(f"  ✗ {label}: {status}")
                    all_ok = False
            self.log(f"  ✓ Excel row cleared for {code}")
            if not all_ok:
                self.log("  ⚠ Some files were missing — Excel row cleared anyway\n")
                messagebox.showwarning("Partial Delete",
                    f"Some files for {code} were not found on disk.\n"
                    "They may have never been saved. Excel row has been cleared.")
            else:
                self.log("")
            self.btn_delete.config(state="disabled", bg="#424242", fg="#aaaaaa")
            self._on_selection()
        except Exception as e:
            self.log(f"  ✗ Delete failed: {e}")
            messagebox.showerror("Delete Failed", str(e))

    def _start_trial(self):
        code, _ = self._build_code()
        if not code:
            messagebox.showwarning("Incomplete", "Please fill in all fields and ensure nurse ≠ patient.")
            return
        if is_done(code):
            if not messagebox.askyesno("Already Done", f"{code} is already marked complete.\nRecord again?"):
                return
        self._trial_code = code
        self._set_controls(False)
        self.btn_delete.config(state="disabled", bg="#424242", fg="#aaaaaa")
        self._open_log_file(code)
        self.log(f"\n{'='*40}")
        self.log(f"  Trial: {code}")
        self.log(f"{'='*40}")
        threading.Thread(target=self._run_trial, daemon=True).start()

    def _run_trial(self):
        code = self._trial_code

        # Check both cameras
        self.log("\n[0] Checking camera connections...")
        ok_a = wait_for_camera(CAM_A, "Camera A (Nurse)",   self.log)
        ok_b = wait_for_camera(CAM_B, "Camera B (Patient)", self.log)
        if not ok_a or not ok_b:
            msg = "Camera A unreachable" if not ok_a else "" 
            msg += (" | " if msg else "") + ("Camera B unreachable" if not ok_b else "")
            self.log(f"  ✗ {msg}")
            add_remark(code, msg)
            self.after(0, lambda: self._set_controls(True))
            return
        self.log("  ✓ Both cameras reachable")

        # Setup both simultaneously
        self.log("\n[1] Setting up cameras...")
        try:
            setup_both()
            self.log("  ✓ Camera A ready (Nurse)")
            self.log("  ✓ Camera B ready (Patient)")
        except Exception as e:
            self.log(f"  ✗ Setup failed: {e}")
            add_remark(code, f"Setup failed: {e}")
            self.after(0, lambda: self._set_controls(True))
            return

        # Start both simultaneously
        self.log("\n[2] Starting recording...")
        try:
            ok_a, ok_b = start_both()
            self.log(f"  Camera A (Nurse):   {'✓ Recording' if ok_a else '✗ Failed'}")
            self.log(f"  Camera B (Patient): {'✓ Recording' if ok_b else '✗ Failed'}")
            if not ok_a or not ok_b:
                msg = "Camera A failed to start" if not ok_a else ""
                msg += (" | " if msg else "") + ("Camera B failed to start" if not ok_b else "")
                self.log(f"  ✗ {msg}")
                add_remark(code, msg)
                stop_both()
                self.after(0, lambda: self._set_controls(True))
                return
            self._recording = True
            self._record_start_time = None  # set on beep
        except Exception as e:
            self.log(f"  ✗ Error: {e}")
            self.after(0, lambda: self._set_controls(True))
            return

        time.sleep(2)
        self._record_start_time = datetime.now()
        self._write_marker(
            f">>> RECORD_START | {self._record_start_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]} | t=0.000s"
        )
        self.log("\n[3] 🔔 Action starts!")
        beep(1)
        self.log("  → Press STOP when action is complete")

    def _stop_trial(self):
        if not self._recording:
            return
        self._recording = False
        # Do NOT re-enable controls here — wait until pipeline fully completes
        self.btn_stop.config(state="disabled", bg="#424242", fg="#aaaaaa")
        threading.Thread(target=self._finish_trial, daemon=True).start()

    def _finish_trial(self):
        code = self._trial_code
        try:
            self._finish_trial_inner(code)
        finally:
            # Always restore controls no matter what happened
            self._close_log_file()
            self.after(0, lambda: self._set_controls(True))
            # Enable delete if the trial is marked done OR files exist on disk
            folder = get_save_folder(code)
            files_exist = any(
                os.path.exists(os.path.join(folder, f"{code}{s}"))
                for s in ("_A.MP4", "_B.MP4")
            )
            if files_exist or is_done(code):
                self.after(0, lambda: self.btn_delete.config(
                    state="normal", bg="#b71c1c", fg="white"))
            self.after(0, self._on_selection)

    def _finish_trial_inner(self, code):
        self.log("\n[4] 🔔 Action ends!")
        beep(2)
        time.sleep(0.5)

        # Stop both simultaneously
        self.log("\n[5] Stopping cameras...")
        stop_both()
        stop_time = datetime.now()
        if self._record_start_time:
            elapsed = (stop_time - self._record_start_time).total_seconds()
            self._write_marker(
                f">>> RECORD_STOP  | {stop_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]} | "
                f"t={elapsed:.3f}s | duration={elapsed:.3f}s"
            )
        time.sleep(2)
        self.log("  ✓ Both cameras stopped")

        # Get files
        self.log("\n[6] Fetching file info...")
        try:
            folder_a, file_a = get_last_file(CAM_A)
            folder_b, file_b = get_last_file(CAM_B)
            self.log(f"  Camera A last file: {file_a}")
            self.log(f"  Camera B last file: {file_b}")
        except Exception as e:
            self.log(f"  ✗ Could not get file list: {e}")
            add_remark(code, f"File fetch failed: {e}")
            return

        # Build save paths
        save_folder = get_save_folder(code)
        save_a = f"{save_folder}\\{code}_A.MP4"
        save_b = f"{save_folder}\\{code}_B.MP4"
        self.log(f"\n  Saving to: {save_folder}")

        # Download both simultaneously
        self.log("\n[7] Downloading both cameras simultaneously...")
        log_a = lambda msg: self.log(f"  [A] {msg}")
        log_b = lambda msg: self.log(f"  [B] {msg}")
        err_a, err_b, size_a, size_b = download_both(
            folder_a, file_a, save_a, folder_b, file_b, save_b,
            log_fn_a=log_a, log_fn_b=log_b)
        if err_a:
            self.log(f"  ✗ Camera A download failed: {err_a}")
        else:
            self.log(f"  ✓ Saved: {code}_A.MP4  ← Nurse")
        if err_b:
            self.log(f"  ✗ Camera B download failed: {err_b}")
        else:
            self.log(f"  ✓ Saved: {code}_B.MP4  ← Patient")
        if err_a or err_b:
            msg = ("Camera A download failed" if err_a else "")
            msg += (" | " if msg else "") + ("Camera B download failed" if err_b else "")
            add_remark(code, msg.strip(" | "))
            return

        # Verify file sizes match what the camera reported
        for label, size in [("Camera A", size_a), ("Camera B", size_b)]:
            if size and size[1] > 0 and size[0] != size[1]:
                actual_mb   = size[0] // (1024 * 1024)
                expected_mb = size[1] // (1024 * 1024)
                remark = f"{label} size mismatch: got {actual_mb} MB, expected {expected_mb} MB — file may be corrupt"
                self.log(f"  ⚠ {remark}")
                add_remark(code, remark)

        # Delete from both SD cards
        self.log("\n[8] Cleaning SD cards...")
        if delete_file(CAM_A, folder_a, file_a):
            self.log("  ✓ Deleted from Camera A")
        else:
            self.log("  ⚠ Could not delete from Camera A (manual cleanup needed)")
            add_remark(code, "Camera A SD delete failed - manual cleanup needed")
        if delete_file(CAM_B, folder_b, file_b):
            self.log("  ✓ Deleted from Camera B")
        else:
            self.log("  ⚠ Could not delete from Camera B (manual cleanup needed)")
            add_remark(code, "Camera B SD delete failed - manual cleanup needed")

        # Mark done in Excel
        self.log("\n[9] Updating checklist...")
        try:
            mark_done(code)
            self.log(f"  ✓ Marked {code} as done in Excel")
        except Exception as e:
            self.log(f"  ⚠ Excel update failed: {e}")
            # Can't write remark if Excel is failing, just log it

        self.log(f"\n{'='*40}")
        self.log(f"  ✅ {code}_A.MP4  ← Nurse")
        self.log(f"  ✅ {code}_B.MP4  ← Patient")
        self.log(f"  📁 {save_folder}")
        self.log(f"{'='*40}\n")



    # ── Test (no save) ────────────────────────────────────
    def _test_cameras(self):
        self._test_recording = False
        self._test_cancelled = False
        self.btn_test.config(state="disabled", bg="#424242", fg="#aaaaaa")
        self.btn_start.config(state="disabled", bg="#424242", fg="#aaaaaa")
        self.btn_test_stop.config(state="normal", bg="#e65100", fg="white")
        self.log("\n── TEST MODE (nothing will be saved) ──")
        threading.Thread(target=self._run_test, daemon=True).start()

    def _check_camera_cancellable(self, ip, label, retries=5, delay=3):
        """Like wait_for_camera but respects _test_cancelled between retries."""
        for attempt in range(1, retries + 1):
            if self._test_cancelled:
                return None  # cancelled
            try:
                r = requests.get(f"{ip}/gopro/camera/info", timeout=3)
                if r.status_code == 200:
                    return True
            except Exception:
                pass
            self.log(f"  ⚠ {label} not responding (attempt {attempt}/{retries})...")
            for _ in range(delay * 5):        # sleep in 0.2s ticks so cancel is fast
                if self._test_cancelled:
                    return None
                time.sleep(0.2)
        return False

    def _run_test(self):
        self.log("\n[0] Checking camera connections...")
        ok_a = self._check_camera_cancellable(CAM_A, "Camera A (Nurse)")
        ok_b = self._check_camera_cancellable(CAM_B, "Camera B (Patient)")
        if self._test_cancelled:
            self.log("  ⚠ Test cancelled during connection check")
            self.after(0, self._reset_test_buttons)
            return
        if not ok_a or not ok_b:
            self.log(f"  ✗ {'Camera A unreachable' if not ok_a else ''}{'  |  Camera B unreachable' if not ok_b else ''}")
            self.after(0, self._reset_test_buttons)
            return
        self.log("  ✓ Both cameras reachable")

        self.log("\n[1] Setting up cameras...")
        try:
            setup_both()
            self.log("  ✓ Camera A ready (Nurse)")
            self.log("  ✓ Camera B ready (Patient)")
        except Exception as e:
            self.log(f"  ✗ Setup failed: {e}")
            self.after(0, self._reset_test_buttons)
            return

        self.log("\n[2] Starting recording...")
        try:
            ok_a, ok_b = start_both()
            self.log(f"  Camera A: {'✓ Recording' if ok_a else '✗ Failed'}")
            self.log(f"  Camera B: {'✓ Recording' if ok_b else '✗ Failed'}")
            if not ok_a or not ok_b:
                stop_both()
                self.after(0, self._reset_test_buttons)
                return
            self._test_recording = True
        except Exception as e:
            self.log(f"  ✗ Error: {e}")
            self.after(0, self._reset_test_buttons)
            return

        beep(1)
        self.log("\n[3] 🔔 Test recording in progress — press Stop Test when done")

    def _stop_test(self):
        self._test_cancelled = True          # interrupt connection-check loop if running
        self.btn_test_stop.config(state="disabled", bg="#424242", fg="#aaaaaa")
        if not self._test_recording:
            # cancelled before recording started — just reset buttons
            self.log("  ⚠ Test cancelled")
            self.after(0, self._reset_test_buttons)
            return
        self._test_recording = False
        threading.Thread(target=self._finish_test, daemon=True).start()

    def _finish_test(self):
        beep(2)
        self.log("\n[4] Stopping cameras (test — no download, no save)...")
        stop_both()
        time.sleep(2)
        self.log("  ✓ Cameras stopped")

        # Delete the test clip from both SD cards
        self.log("\n[5] Deleting test clips from SD cards...")
        try:
            folder_a, file_a = get_last_file(CAM_A)
            if delete_file(CAM_A, folder_a, file_a):
                self.log(f"  ✓ Deleted from Camera A: {file_a}")
            else:
                self.log(f"  ⚠ Could not delete from Camera A (manual cleanup needed)")
        except Exception as e:
            self.log(f"  ⚠ Camera A cleanup failed: {e}")

        try:
            folder_b, file_b = get_last_file(CAM_B)
            if delete_file(CAM_B, folder_b, file_b):
                self.log(f"  ✓ Deleted from Camera B: {file_b}")
            else:
                self.log(f"  ⚠ Could not delete from Camera B (manual cleanup needed)")
        except Exception as e:
            self.log(f"  ⚠ Camera B cleanup failed: {e}")

        self.log("  ✓ Test complete — nothing was saved\n")
        self.after(0, self._reset_test_buttons)

    def _reset_test_buttons(self):
        self.btn_test.config(state="normal", bg="#4a3800", fg="#ffd54f")
        self.btn_test_stop.config(state="disabled", bg="#424242", fg="#aaaaaa")
        self.btn_start.config(state="normal", bg="#1565c0", fg="white")


    # ── Import Lock ───────────────────────────────────────
    def _unlock_all(self):
        """Enable all controls after a successful import."""
        self.btn_start.config(state="normal",   bg="#1565c0", fg="white")
        self.btn_test.config(state="normal",    bg="#4a3800", fg="#ffd54f")
        self.btn_export.config(state="normal",  bg="#1a3a1a", fg="#81c784")
        self.btn_import.config(state="normal",  bg="#1a3a1a", fg="#81c784")
        self.btn_stop.config(state="disabled",  bg="#424242", fg="#aaaaaa")
        self.btn_delete.config(state="disabled",bg="#424242", fg="#aaaaaa")
        self.btn_test_stop.config(state="disabled", bg="#424242", fg="#aaaaaa")
        self.lbl_session.config(
            text="✅  Checklist imported — record your trials, then export before closing.",
            fg="#81c784"
        )

    # ── Azure Sync ────────────────────────────────────────
    def _import_excel(self):
        if self._recording:
            messagebox.showwarning("Recording Active", "Cannot import checklist while a trial is in progress.")
            return
        if self._checklist_imported:
            if not messagebox.askyesno("Import Checklist",
                                       "Download the shared checklist from Azure?\n\n"
                                       "Your local copy will be replaced.\n"
                                       "A backup (.bak) will be saved first."):
                return
        self.btn_import.config(state="disabled", bg="#424242", fg="#aaaaaa")
        self.log("\n── Importing checklist from Azure...")

        def _run():
            try:
                import_excel(EXCEL_PATH, log_fn=self.log)
                self.log("  ✅ Checklist imported — Excel is now up to date\n")
                self._checklist_imported = True
                self._checklist_exported = False
                self.after(0, self._unlock_all)
            except Exception as e:
                self.log(f"  ✗ Import failed: {e}\n")
                self.after(0, lambda: messagebox.showerror("Import Failed", str(e)))
                self.after(0, lambda: self.btn_import.config(
                    state="normal", bg="#1a3a1a", fg="#81c784"))

        threading.Thread(target=_run, daemon=True).start()

    def _export_excel(self):
        if self._recording:
            messagebox.showwarning("Recording Active", "Cannot export checklist while a trial is in progress.")
            return
        if not messagebox.askyesno("Export Checklist",
                                   "Upload your local checklist to Azure?\n\n"
                                   "This will overwrite the shared copy."):
            return
        self.btn_export.config(state="disabled", bg="#424242", fg="#aaaaaa")
        self.log("\n── Exporting checklist to Azure...")

        def _run():
            try:
                export_excel(EXCEL_PATH, log_fn=self.log)
                self.log("  ✅ Checklist exported — Azure copy is now up to date\n")
                self._checklist_exported = True
                self.after(0, lambda: self.lbl_session.config(
                    text="☁️  Checklist exported — you may now close the app safely.",
                    fg="#64b5f6"
                ))
            except Exception as e:
                self.log(f"  ✗ Export failed: {e}\n")
                self.after(0, lambda: messagebox.showerror("Export Failed", str(e)))
            finally:
                self.after(0, lambda: self.btn_export.config(
                    state="normal", bg="#1a3a1a", fg="#81c784"))

        threading.Thread(target=_run, daemon=True).start()

    # ── Close Protection ──────────────────────────────────
    def _on_close(self):
        if self._recording or self._test_recording:
            if not messagebox.askyesno(
                "Recording in progress",
                "A recording is still active.\n\n"
                "Closing now will stop the cameras but the current trial\n"
                "will NOT be saved or marked done.\n\n"
                "Close anyway?"
            ):
                return
            import threading as _t
            _t.Thread(target=stop_both, daemon=True).start()
            self.destroy()
            return

        # Block closing until checklist is exported
        if self._checklist_imported and not self._checklist_exported:
            messagebox.showwarning(
                "Export Required — Cannot Close",
                "You must export the checklist before closing.\n\n"
                "Your session changes have not been saved to Azure yet.\n\n"
                "➡  Press  📤 Export Checklist  then close again."
            )
            return

        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()