import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import socket, time, threading, csv, re, subprocess, os


# Global dictionary to track each window’s ConnectivityRun
runs = {}  # key: Treeview widget, value: ConnectivityRun instance

class ConnectivityRun:
    def __init__(self, tree, tasks, source_ip=None):
        self.tree = tree
        self.tasks = tasks
        self.source_ip = source_ip
        self.index = 0
        self.stop_flag = False
        self.thread = None
        self.test_button = None

    def stop(self):
        self.stop_flag = True



def _connect_to_host(dest_ip: str, dest_port: int, source_ip: str | None = None) -> tuple[bool, int | None, str]:
    """
    Try to connect to (dest_ip, dest_port).

    Returns:
        success      – True if the socket connected.
        elapsed_ms   – Milliseconds spent in the call (None on failure).
        error_text   – Empty string on success; otherwise an error description.
    """
    start = time.perf_counter()
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(2)                    # 2‑second timeout
        if source_ip:                         # bind to a specific local interface
            sock.bind((source_ip, 0))
        code = sock.connect_ex((dest_ip, dest_port))
        sock.close()

        elapsed_ms = int((time.perf_counter() - start) * 1000)
        success   = (code == 0)

        return success, elapsed_ms, "" if success else "UNSUCCESSFUL"

    except socket.timeout:
        # Timeout: connection attempt took longer than the timeout
        return False, None, "Connection timed out"

    except ConnectionRefusedError:
        # Remote host actively refused the connection
        return False, None, "Connection refused by remote host"

    except Exception as exc:
        # On any other exception we report failure and the error message
        return False, None, f"Error: {exc}"



# ------------------------------------------------------------------
# Get the public IP that is reachable from a specific private IP.
# Uses urllib – no external libraries are required.
# ------------------------------------------------------------------
def _public_ip_for_local(local_ip: str) -> str | None:
    """
    Return the public IP address as seen from the network interface bound to `local_ip`.
    If the request fails, return None.
    """
    import urllib.request

    try:
        # Create a socket that binds to the given local IP so the outbound
        # connection originates from that interface.  This mimics the behaviour
        # of PowerShell’s Invoke‑RestMethod –Uri “https://api.ipify.org”.
        req = urllib.request.Request("https://api.ipify.org")
        opener = urllib.request.build_opener()
        opener.addheaders = [("User-Agent", "Python-ConnectivityTester/1.0")]

        # Force the socket to bind to `local_ip` by using a custom
        # socket factory (see https://bugs.python.org/issue12371).
        class BindSocketFactory(urllib.request.HTTPHandler):
            def http_open(self, req):
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(5)                     # 5‑second timeout
                sock.bind((local_ip, 0))               # bind to local IP
                conn = urllib.request.HTTPConnection(
                    host=req.host,
                    timeout=5,
                    source_address=(local_ip, 0),
                )
                return conn

        opener.add_handler(BindSocketFactory())
        with opener.open(req, timeout=5) as resp:
            return resp.read().decode("utf-8").strip()
    except Exception:
        # Any failure (timeout, DNS error, etc.) – just return None
        return None


def load_tasks(csv_path):
    rows = []
    with open(csv_path, newline='') as fh:
        for r in csv.DictReader(fh):
            rows.append((r['Description'], r['IP'], int(r['Port'])))
    return rows

def get_subfolders(base_path):
    """Return all subfolder names in base_path."""
    try:
        return [
            d for d in os.listdir(base_path)
            if os.path.isdir(os.path.join(base_path, d))
        ]
    except Exception:
        return []


def get_machine_ipv4_addresses():
    """Run ipconfig and parse out all IPv4 addresses."""
    try:
        output = subprocess.run(
            ['ipconfig'], capture_output=True, text=True
        ).stdout.splitlines()
        ips = []
        for line in output:
            if "IPv4 Address" in line:
                match = re.search(r'(\d+\.\d+\.\d+\.\d+)', line)
                if match:
                    ips.append(match.group(1))
        return ips
    except Exception:
        return []
    
def refresh_ip_dropdowns(*combos: ttk.Combobox):
    """Replace each combo’s values with the current machine IPs."""
    ips = get_machine_ipv4_addresses()
    for cb in combos:
        cur_val = cb.get()          # remember current selection
        cb['values'] = ips           # update list
        if cur_val in ips:           # keep it if still valid
            cb.set(cur_val)
        else:
            cb.set(ips[0] if ips else "")

    


def start_connectivity_check(csv_path, result_window, tree, button_name,
                             source_ip=None, local_ip_text=None, delay_ms=100):
    """
    Kick off (or restart) a connectivity run tied to this Treeview.
    If there’s an active run, stop it, wait 100ms, and restart.
    """
    # If an existing run is in progress, stop it and restart shortly
    if tree in runs and not runs[tree].stop_flag:
        runs[tree].stop()
        result_window.after(
            100,
            lambda: start_connectivity_check(
                csv_path, result_window, tree, button_name,
                source_ip, local_ip_text, delay_ms  # pass delay on restart
            )
        )
        return

    # Initialize new run
    tasks = load_tasks(csv_path)
    run = ConnectivityRun(tree, tasks, source_ip)
    runs[tree] = run

    # Clear old rows
    tree.delete(*tree.get_children())
    if local_ip_text:
        local_ip_text.delete("1.0", tk.END)

    # Preload the tree with "Testing" placeholders
    for desc, ip, port in tasks:
        tree.insert(
            "",
            tk.END,
            values=(desc, ip, port, "Testing"),
            tags=("pending",)
        )

    # Show local machine IPs if desired
    if local_ip_text:
        ips = get_machine_ipv4_addresses()
        local_ip_text.insert(tk.END, "\n".join(ips) + "\n")

    # Begin async scanning with user-selected delay
    run_task_async(tasks, tree, 0, button_name, source_ip, delay_ms=delay_ms)


# Proper disposal of window and stopping tests if user closes window while tests are running
def _on_result_window_close(tree, win):
    """Stop any running tests for this tree and close the window."""
    run = runs.get(tree)
    if run:
        # Mark remaining rows as cancelled (optional)
        for iid in tree.get_children():
            vals = tree.item(iid, "values")
            if vals[3] in ("Testing", "Not tested"):
                tree.item(
                    iid,
                    values=(vals[0], vals[1], vals[2], "Cancelled"),
                    tags=("cancelled",)
                )
        
        run.stop()          # signal the background thread to exit
    win.destroy()


def run_task_async(tasks, tree, index, button_name, source_ip=None, delay_ms=100):
    run = runs.get(tree)
    if not run or run.stop_flag or index >= len(run.tasks):
        return

    run.index = index  # track progress

    def scan_worker():
        # 1) mark row testing & Start timer
        start = time.perf_counter()

        def set_testing():
            if run.stop_flag:
                return
            iid = tree.get_children()[index]
            current = tree.item(iid, "values")
            tree.item(
                iid,
                values=(current[0], current[1], current[2], "Testing"),
                tags=("testing",)
            )
        tree.after(0, set_testing)

        # 2) do the port-connect test
        desc, ip, port = tasks[index]
        success, elapsed_ms, err_msg = _connect_to_host(ip, port, source_ip)

        status_text = f"SUCCESSFUL ({elapsed_ms} ms)" if success else err_msg
        tag = "successful" if success else "unsuccessful"

        # 3) update row with result
        def update_ui():
            if not tree.winfo_exists():
                return
            iid = tree.get_children()[index]
            current = tree.item(iid, "values")
            tree.item(
                iid,
                values=(current[0], current[1], current[2], status_text),
                tags=(tag,)
            )
            tree.see(iid)
            if index >= len(tasks) - 3:
                tree.yview_scroll(1, "units")
        tree.after(0, update_ui)

        # 4) schedule next using the configured delay
        if not run.stop_flag:
            next_delay = max(0, int(delay_ms))  # clamp to non-negative
            tree.after(
                next_delay,
                # IMPORTANT: propagate delay_ms in the recursive call
                lambda: run_task_async(
                    tasks, tree, index + 1, button_name, source_ip, delay_ms=next_delay
                )
            )

    thread = threading.Thread(target=scan_worker, daemon=True)
    thread.start()
    run.thread = thread



def open_result_window(csv_path, button_name):
    """Build a toplevel window with IP dropdown, tree, and control buttons."""
    try:
        result_window = tk.Toplevel()
        result_window.title("Connectivity Results")
        result_window.geometry("720x500")
        result_window.configure(bg="#f0f0f0")

        result_window.protocol(
            "WM_DELETE_WINDOW",
            lambda: _on_result_window_close(tree, result_window)
        )

        def _close_all_results(event=None):
            for win in result_window.master.winfo_children():
                if isinstance(win, tk.Toplevel) and win.winfo_exists():
                    win.destroy()

        result_window.bind("<Escape>", _close_all_results)

        # Title bar with button name
        title_frame = tk.Frame(result_window, bg="#f0f0f0")
        title_frame.pack(pady=10)
        tk.Label(
            title_frame,
            text="Connectivity Tester ",
            font=("Segoe UI", 12, "bold"),
            bg="#f0f0f0"
        ).pack(side=tk.LEFT)
        tk.Label(
            title_frame,
            text=f"> {button_name}",
            font=("Segoe UI", 12, "bold"),
            bg="#f0f0f0",
            fg="#2c3e50"
        ).pack(side=tk.LEFT)

        # Source-IP dropdown
        ip_frame = tk.Frame(result_window, bg="#f0f0f0")
        ip_frame.pack(pady=5)
        tk.Label(
            ip_frame,
            text="Select Local IP:",
            font=("Segoe UI", 10, "bold"),
            bg="#f0f0f0",
            fg="#2c3e50"
        ).pack(side=tk.LEFT, padx=(10, 5))
        local_ips = get_machine_ipv4_addresses()
        selected_ip = tk.StringVar(value=local_ips[0] if local_ips else "")
        combo_local_ip = ttk.Combobox(
            ip_frame,
            textvariable=selected_ip,
            values=local_ips,
            state="readonly",
            font=("Segoe UI", 9),
            width=20
        )
        combo_local_ip.pack(side=tk.LEFT, padx=5)

        # Refresh IPs button (right of the combobox)
        refresh_ip_btn = tk.Button(
            ip_frame,
            text="Refresh IPs",
            command=lambda: refresh_ip_dropdowns(combo_local_ip),
            font=("Segoe UI", 9, "bold"),
            bg="#2980b9",
            fg="#ffffff",
            activebackground="#2980b9",
            activeforeground="#ffffff",
            relief="flat",
            padx=10,
            pady=3,
            borderwidth=0
        )
        refresh_ip_btn.pack(side=tk.LEFT, padx=(5, 0))

        # Optional scrolledtext for local IP display (unused here)
        local_ip_text = None

        # Treeview + scrollbar
        table_frame = tk.Frame(result_window, bg="#f0f0f0")
        table_frame.pack(fill=tk.BOTH, expand=True,
                         padx=10, pady=(0, 10))
        tree = ttk.Treeview(
            table_frame,
            columns=("Service / Description",
                     "Destination IP / DNS",
                     "Destination Port",
                     "Result"),
            show="headings",
            style="Modern.Treeview"
        )
        for col, txt, w in [
            ("Service / Description", "Service / Description", 180),
            ("Destination IP / DNS", "Destination IP / DNS", 150),
            ("Destination Port", "Destination Port", 100),
            ("Result", "Result", 200),
        ]:
            tree.heading(col, text=txt)
            if col == "Service / Description":
                tree.column(col, width=w, anchor="w")
            else:
                tree.column(col, width=w, anchor="center")

        # preload as "Not tested"
        with open(csv_path, newline="") as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                tree.insert(
                    "",
                    tk.END,
                    values=(row["Description"],
                            row["IP"],
                            int(row["Port"]),
                            "Not tested"),
                    tags=("pending",)
                )
        vsb = ttk.Scrollbar(table_frame, orient="vertical",
                            command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # Treeview styling
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Modern.Treeview",
            background="#ffffff",
            foreground="#2c3e50",
            rowheight=25,
            fieldbackground="#ffffff"
        )
        style.map("Modern.Treeview", background=[("selected", "#3498db")])
        tree.tag_configure("pending", foreground="#000000")
        tree.tag_configure("testing",
                           background="#e0f7ff",
                           foreground="#007acc")
        tree.tag_configure("successful", foreground="#27ae60")
        tree.tag_configure("unsuccessful", foreground="#e74c3c")
        tree.tag_configure("cancelled", foreground="#f39c12")

                # --------------------  **DELAY CONFIGURATION** --------------------
        delay_frame = tk.Frame(result_window, bg="#f0f0f0")
        delay_frame.pack(pady=5)

        lbl_delay = tk.Label(
            delay_frame,
            text="Delay between tests",
            font=("Segoe UI", 9),
            bg="#f0f0f0"
        )
        lbl_delay.grid(row=0, column=0, padx=(5, 2))

        # Holds the current numeric value (default 100)
        delay_var = tk.IntVar(value=100)
        entry_delay = ttk.Entry(
            delay_frame,
            textvariable=delay_var,
            width=6,
            font=("Segoe UI", 9),
            justify="right"
        )
        entry_delay.grid(row=0, column=1, padx=(2, 2))

        lbl_ms = tk.Label(
            delay_frame,
            text="ms",
            font=("Segoe UI", 9),
            bg="#f0f0f0"
        )
        lbl_ms.grid(row=0, column=2, padx=(2, 5))
        # ----------------------------------------------------------------

        # Buttons
        btn_frame = tk.Frame(result_window, bg="#f0f0f0")
        btn_frame.pack(pady=5)

        def create_button(parent, text, command):
            return tk.Button(
                parent,
                text=text,
                command=command,
                font=("Segoe UI", 9, "bold"),
                bg="#2980b9",
                fg="#ffffff",
                activebackground="#2980b9",
                activeforeground="#ffffff",
                relief="flat",
                padx=15,
                pady=5,
                borderwidth=0
            )

        btn_local = create_button(
            btn_frame,
            "Test",
            lambda: start_connectivity_check(
                csv_path, result_window, tree,
                button_name, selected_ip.get(), local_ip_text,
                int(delay_var.get())          # <‑ use the user‑set delay
            )
        )
        btn_stop = create_button(
            btn_frame,
            "Stop",
            lambda t=tree: stop_running_tests(t)
        )
        btn_local.pack(side=tk.LEFT, padx=5)

        # ------------------------------------------------------------------
        # Status bar for the result window – mirrors the main‑window status bar.
        # ------------------------------------------------------------------
        result_status_var = tk.StringVar()
        def _update_result_status(selected_ip: str):
            pub_ip = _public_ip_for_local(selected_ip) or "Unknown"
            result_status_var.set(f"Private IP Selected: {selected_ip} (Public IP {pub_ip})")

        # Set it once with the combobox’s current value
        _update_result_status(combo_local_ip.get())

        result_status_bar = ttk.Label(
            result_window,
            textvariable=result_status_var,
            relief="sunken",
            anchor="w",
            padding=(5, 0)
        )
        result_status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        # Keep it updated on selection changes
        combo_local_ip.bind("<<ComboboxSelected>>", lambda e: _update_result_status(combo_local_ip.get()))



        btn_stop.pack(side=tk.LEFT, padx=5)

    except Exception as e:
        messagebox.showerror("Error", f"Error opening results window: {e}")



def _calc_min_height(
    title_h,
    logo_height,
    btn_rows,          # number of rows with buttons
    BTN_HEIGHT=30,
    row_sp=10,
    manual_row_h=30,
    pad_y=15,
    pad=25
):
    """
    Return the minimum height that will fit all widgets.
    Parameters are the same values you already calculate in the two places.
    """
    # Height contributed by each section
    buttons_grid_h = btn_rows * BTN_HEIGHT + (btn_rows - 1) * row_sp

    # Total minimum height
    return (
        title_h          +   # main title
        title_h          +   # CSV‑folder title
        title_h *3       +   # buffer / separator
        logo_height      +
        buttons_grid_h   +   # button grid (only once)
        manual_row_h     +   # “Manual Test” row
        pad_y * 2         +
        pad * 2           # top & bottom padding
    )



def stop_running_tests(tree):
    """Signal the ConnectivityRun for this Treeview to stop."""
    run = runs.get(tree)
    if not run:
        return

    # Mark all remaining rows as cancelled (orange)
    for iid in tree.get_children():
        current_vals = tree.item(iid, "values")
        # Only touch rows that haven’t finished
        if current_vals[3] in ("Testing", "Not tested"):
            tree.item(
                iid,
                values=(current_vals[0], current_vals[1], current_vals[2], "Cancelled"),
                tags=("cancelled",)
            )
    run.stop()



def on_button_click(button_data):
    """Handler for main window buttons: open results window."""
    csv_path = os.path.join(button_data["Folder"], "ResourcesToCheck.csv")
    open_result_window(csv_path, button_data["Name"])


def create_main_window():
    """Build the main window with one button per subfolder."""
    root = tk.Tk()
    root.title("Simple Network Connectivity Tester")
    root.configure(bg="#f0f0f0")
    folder_frame = tk.Frame(
        root,
        bg="#f0f0f0",
        bd=1,
        relief="solid"
    )
    folder_frame.pack(fill=tk.X, padx=20, pady=(10, 0))

    def refresh_folders():
        """Re‑scan for subfolders and rebuild button area."""
        # Now `folder_frame` is already defined in this outer scope
        for child in folder_frame.winfo_children():
            child.destroy()

        # 1. Scan for new folders that contain a CSV file
        new_buttons_data = [
            {"Name": f, "Folder": f}
            for f in get_subfolders(".")
            if os.path.exists(os.path.join(f, "ResourcesToCheck.csv"))
        ]

        # 2. Remove all old button widgets from the frame
        for child in folder_frame.winfo_children():
            child.destroy()

        # 3. Re‑create the title label (it was destroyed with the children)
        lbl_folder_title = tk.Label(
            folder_frame,
            text="Batch Tests",
            font=("Segoe UI", 10, "bold"),
            bg="#f0f0f0",
            fg="#2c3e50"
        )
        lbl_folder_title.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))

        # 4. Re‑populate the buttons in a grid
        cols_per_row = 2
        for i, bd in enumerate(new_buttons_data):
            r, c = divmod(i, cols_per_row)
            btn = tk.Button(
                folder_frame,
                text=bd["Name"],
                width=20,
                font=("Segoe UI", 9, "bold"),
                bg="#3498db",
                fg="#ffffff",
                activebackground="#2980b9",
                activeforeground="#ffffff",
                relief="flat",
                padx=10,
                pady=5,
                borderwidth=0,
                command=lambda data=bd: on_button_click(data)
            )
            btn.grid(row=r + 1, column=c, padx=5, pady=2, sticky="nsew")

        # 5. Adjust the bottom padding row (if needed)
        folder_frame.rowconfigure(len(new_buttons_data) // cols_per_row + 1,
                                    minsize=10)

        # ------------------------------------------------------------------
        # 6. Re‑calculate window size
        num_rows = (len(new_buttons_data) + cols_per_row - 1) // cols_per_row
        buttons_grid_h = num_rows * BTN_HEIGHT + (num_rows - 1) * row_sp

        new_min_win_h = _calc_min_height(
            title_h,
            logo_height,
            num_rows,
            BTN_HEIGHT=BTN_HEIGHT,
            row_sp=row_sp,
            manual_row_h=manual_h,
            pad_y=pad_y,
            pad=pad
        )


        root.geometry(f"{win_w}x{new_min_win_h}")
        root.minsize(win_w, new_min_win_h)



    # Logo (optional)
    logo_height = 0
    try:
        logo_image = tk.PhotoImage(file="logo.png")
        logo_label = tk.Label(root, image=logo_image, bg="#f0f0f0")
        logo_label.image = logo_image
        logo_label.pack(pady=10)
        logo_height = logo_image.height() + 20
    except tk.TclError:
        pass



    # Discover subfolders with CSV
    folders = get_subfolders(".")
    buttons_data = [
        {"Name": f, "Folder": f}
        for f in folders
        if os.path.exists(os.path.join(f, "ResourcesToCheck.csv"))
    ]

    buttons_per_row = 2
    num_rows = (len(buttons_data) + buttons_per_row - 1) // buttons_per_row
    btn_w, btn_h = 200, 30
    pad, row_sp, btn_pad = 25, 10, 5
    BTN_HEIGHT = 30
    manual_row_h    = 30          # height of the manual‑test row
    pad_y           = 15          # padding around that row
    win_w = btn_w * buttons_per_row + pad * 2 - btn_pad * 2

    # Height contributions
    title_h         = 28   # approximate height of the title label
    logo_height     = logo_image.height() + 20 if 'logo_image' in locals() else 0

    buttons_grid_h  = num_rows * BTN_HEIGHT + (num_rows - 1) * row_sp
    manual_h        = manual_row_h + pad_y

    # Total minimum height
    num_rows = (len(buttons_data) + buttons_per_row - 1) // buttons_per_row
    min_win_h = _calc_min_height(
        title_h,
        logo_height,
        num_rows,
        BTN_HEIGHT=BTN_HEIGHT,
        row_sp=row_sp,
        manual_row_h=manual_h,   # you already have this as `manual_h`
        pad_y=pad_y,
        pad=pad
    )


    root.geometry(f"{win_w}x{min_win_h}")
    root.minsize(win_w, min_win_h)

    # Create a container frame for the buttons auto‑generated from subfolders
    folder_frame = tk.Frame(
        root,
        bg="#f0f0f0",
        bd=1,                # border width
        relief="solid"       # same style as manual_frame
    )
    folder_frame.pack(fill=tk.X, padx=20, pady=(10, 0))

    cols_per_row = 2



    # Title label for the button section
    lbl_folder_title = tk.Label(
        folder_frame,
        text="Batch Tests",
        font=("Segoe UI", 10, "bold"),
        bg="#f0f0f0",
        fg="#2c3e50"
    )
    lbl_folder_title.grid(row=0, column=0, columnspan=cols_per_row,
                          sticky="ew", pady=(0, 10))


    #Configure grid weights so buttons can expand evenly
    for col in range(cols_per_row):
        folder_frame.columnconfigure(col, weight=1)

    # Place the dynamic buttons starting from row 1
    for i, bd in enumerate(buttons_data):
        r, c = divmod(i, cols_per_row)
        btn = tk.Button(
            folder_frame,
            text=bd["Name"],
            width=20,
            font=("Segoe UI", 9, "bold"),
            bg="#3498db",
            fg="#ffffff",
            activebackground="#2980b9",
            activeforeground="#ffffff",
            relief="flat",
            padx=10,
            pady=5,
            borderwidth=0,
            command=lambda data=bd: on_button_click(data)
        )
        # Use `sticky="nsew"` so the button expands to fill its cell
        btn.grid(row=r+1, column=c, padx=5, pady=2, sticky="nsew")

    #(Optional) Add a bottom padding row 
    folder_frame.rowconfigure(len(buttons_data) // cols_per_row + 1,
                             minsize=10)


    refresh_btn = tk.Button(root, text="Refresh", command=refresh_folders)
    refresh_btn.pack(pady=5)
    root.bind("<Control-r>", lambda e: refresh_btn.invoke())
    root.bind("<F5>", lambda e: refresh_btn.invoke())       


    # ---- Manual IP:Port test section ---------------------------------
    manual_frame = tk.Frame(root, bg="#f0f0f0", border=1, relief="solid")
    manual_frame.pack(fill=tk.X, padx=20, pady=20)  
    manual_frame.grid_rowconfigure(1, pad=0)


    # ----- Manual Test ----------------------------------------------------
    def _run_manual():
        """Run the manual IP:Port connectivity test."""
        ip_port = ip_port_var.get().strip()
        source_ip = source_ip_var.get()

        if not ip_port:
            messagebox.showwarning("Input required", "Please enter an IP:Port value.")
            return

        try:
            dest_ip, dest_port_str = ip_port.split(":")
            dest_port = int(dest_port_str)
        except ValueError:
            messagebox.showerror(
                "Format error",
                ("Enter in format <IP>:<port> or <DNSName>:<Port> "
                 "(e.g. 192.168.1.10:80 or example.com:443)")
            )
            return
        
        #show Testing
        result_var.set("Testing")
        result_entry.configure(foreground="#2980b9") 

        # Run the helper
        def worker():
            success, elapsed_ms, err_msg = _connect_to_host(dest_ip, dest_port, source_ip)

            if success:
                result_text = f"SUCCESSFUL ({elapsed_ms} ms)"
                color = "#27ae60"
            else:
                result_text = err_msg
                color = "#c0392b"

            # Update UI in the main thread
            root.after(0, lambda: [
                result_var.set(result_text),
                result_entry.configure(foreground=color)
            ])

        # Build the result string – only add ms if the test succeeded
            if success:
                elapsed_str = f" ({elapsed_ms} ms)"
                result_text = f"SUCCESSFUL{elapsed_str}"
            else:
                # Do **not** append the milliseconds when unsuccessful
                result_text = err_msg

            # Update the read‑only entry with color
            result_var.set(result_text)
            color = "#27ae60" if success else "#c0392b"
            result_entry.configure(foreground=color)

        threading.Thread(target=worker, daemon=True).start()







    label_manual = tk.Label(
        manual_frame,
        text="Manual Test",
        font=("Segoe UI", 10, "bold"),
        bg="#f0f0f0",
        fg="#2c3e50",
        anchor="center"  
    )
    label_manual.grid(row=0, column=0, columnspan=3, sticky="ew")


    label_dest = tk.Label(
        manual_frame,
        text="Destination",
        font=("Segoe UI", 9),
        bg="#f0f0f0",
        fg="#2c3e50",
        anchor="center"  
    )
    label_dest.grid(row=1, column=0, sticky="nsew", padx=(5, 2))

    label_local_ip = tk.Label(
        manual_frame,
        text="Local IP",
        font=("Segoe UI", 9),
        bg="#f0f0f0",
        fg="#2c3e50",
        anchor="center"  
    )
    label_local_ip.grid(row=1, column=1, sticky="nsew", padx=(5, 2))

    label_result = tk.Label(
        manual_frame,
        text="Result",
        font=("Segoe UI", 9),
        bg="#f0f0f0",
        fg="#2c3e50"
    )
    label_result.grid(row=1, column=2, sticky="nsew", padx=(5, 2))


    # Entry for “IP:Port”
    ip_port_var = tk.StringVar()
 
    # -------------  PLACEHOLDER SETUP ----------------------------------
    placeholder = "IPAddress:Port or URL:Port"
    placeholder_color = "#a0a0a0"

    def _clear_placeholder(event):
        """Remove placeholder text when the entry gets focus."""
        if entry_ipport.get() == placeholder:
            entry_ipport.delete(0, tk.END)
            entry_ipport.configure(foreground="black")   

    def _add_placeholder(event):
        """Re‑insert placeholder if the user leaves the field empty."""
        if not entry_ipport.get():
            entry_ipport.insert(0, placeholder)
            entry_ipport.configure(foreground=placeholder_color)




    # Create the ttk.Entry
    entry_ipport = ttk.Entry(
        manual_frame,
        textvariable=ip_port_var,
        width=22
    )
    entry_ipport.grid(row=2, column=0, sticky="we", padx=(5,0))
    entry_ipport.bind("<Return>", lambda e: _run_manual())

    # Insert placeholder initially
    entry_ipport.insert(0, placeholder)
    entry_ipport.configure(foreground=placeholder_color)

    # Bind focus events
    entry_ipport.bind("<FocusIn>",  _clear_placeholder)
    entry_ipport.bind("<FocusOut>", _add_placeholder)

    ip_port_var.set("google.com:443")
    entry_ipport.configure(foreground="black")


   
    # Source‑IP dropdown 
    source_ip_var = tk.StringVar()
    local_ips = get_machine_ipv4_addresses()
    if local_ips:
        source_ip_var.set(local_ips[0])   # pick first address by default

    combo_source_ip = ttk.Combobox(
        manual_frame,
        textvariable=source_ip_var,
        values=local_ips,
        font=("Segoe UI", 9),
        width=10,
        state="readonly",
        style="Modern.TCombobox",
    )
    combo_source_ip.grid(row=2, column=1, sticky="nsew", padx=(5,0))



    # Result label 
    result_var = tk.StringVar()
    result_entry = ttk.Entry(
        manual_frame,
        textvariable=result_var,
        font=("Segoe UI", 9),
        width=22,
        state="readonly",
        justify='center' 
    )
    result_entry.grid(row=2, column=2, sticky="ew", padx=(5,0))

    # Set the initial placeholder text
    initresult = "IPAddress:Port / URL:Port"
    initresult_color = "#a0a0a0"

    result_var.set("Not tested")          
    result_entry.configure(foreground=initresult_color) 


    # Test button – one‑third the width of the dropdown
    test_btn = tk.Button(
        manual_frame,
        text="Manual Test",
        command=_run_manual,  
        font=("Segoe UI", 9, "bold"),
        bg="#2980b9",
        fg="#ffffff",
        activebackground="#2980b9",
        activeforeground="#ffffff",
        relief="flat",
        padx=15,
        pady=5,
        borderwidth=0
    )
    test_btn.grid(row=3, column=0, sticky="we", padx=(5,0), pady=(10,0))

    # Refresh IPs button 
    refresh_btn = tk.Button(
        manual_frame,
        text="Refresh IPs",
        command=lambda: refresh_ip_dropdowns(combo_source_ip),
        font=("Segoe UI", 9, "bold"),
        bg="#2980b9",
        fg="#ffffff",
        activebackground="#2980b9",
        activeforeground="#ffffff",
        relief="flat",
        padx=15,
        pady=5,
        borderwidth=0
    )

    refresh_btn.grid(row=3, column=1, sticky="we", padx=(10,0), pady=(10,0))


    # --------------------------------------------------------------------
    # Layout tweaks
    manual_frame.rowconfigure(1, minsize=0)   # Title row
    manual_frame.rowconfigure(2, minsize=2)   # Labels row
    manual_frame.rowconfigure(3, minsize=10)   # Widgets row
    manual_frame.rowconfigure(4, minsize=10)   # Button row


    # ------------------------------------------------------------------
    # Status bar – shows Private IP | Public IP for the currently selected local IP.
    # ------------------------------------------------------------------
    status_var = tk.StringVar()
    status_bar = ttk.Label(
        root,
        textvariable=status_var,
        relief="sunken",
        anchor="w",
        padding=(5, 0)
    )
    status_bar.place(relx=0, rely=1.0, relwidth=1.0, anchor='sw')

    def _update_status(selected_ip: str):
        """Populate the status bar with the private & public IP for `selected_ip`."""
        pub_ip = _public_ip_for_local(selected_ip) or "Unknown"
        status_var.set(f"Private IP Selected: {selected_ip} (Public IP:{pub_ip})")

    # Initial population
    if local_ips:
        _update_status(local_ips[0])

    # Hook into the Combobox’s selection event
    combo_source_ip.bind("<<ComboboxSelected>>", lambda e: _update_status(combo_source_ip.get()))


    root.mainloop()

if __name__ == "__main__":
    create_main_window()
