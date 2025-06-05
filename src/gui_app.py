import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from dotenv import load_dotenv

# Ensure we can import main from the same directory
sys.path.append(os.path.dirname(__file__))
from main import main as cli_main

# Load existing environment variables from .env if present
load_dotenv()


def browse_directory(var: tk.StringVar) -> None:
    path = filedialog.askdirectory()
    if path:
        var.set(path)


def run_script() -> None:
    # Ensure required environment variables are provided
    if not download_var.get().strip() or not work_var.get().strip():
        messagebox.showwarning("입력 필요", "DOWNLOAD_DIR과 WORK_DIR을 모두 설정하세요.")
        return

    # Apply environment variables from the GUI fields
    os.environ["DOWNLOAD_DIR"] = download_var.get()
    os.environ["WORK_DIR"] = work_var.get()
    os.environ["MEAL_FEE"] = meal_var.get()
    os.environ["OFFICIAL_DATA_NAMES_STR"] = names_var.get()

    try:
        cli_main()
        messagebox.showinfo("완료", "처리가 완료되었습니다.")
    except Exception as exc:
        messagebox.showerror("오류", str(exc))


root = tk.Tk()
root.title("초과근무 처리 도구")

# ---- Material-like styling ----
primary_color = "#6200EE"  # Material purple 500
background_color = "#F5F5F5"  # light grey background
root.configure(bg=background_color)

style = ttk.Style(root)
style.theme_use("clam")

style.configure(
    "TFrame", background=background_color
)
style.configure(
    "TLabel",
    background=background_color,
    font=("Segoe UI", 10)
)
style.configure(
    "TEntry",
    font=("Segoe UI", 10)
)
style.configure(
    "Material.TButton",
    font=("Segoe UI", 10, "bold"),
    foreground="white",
    background=primary_color
)
style.map(
    "Material.TButton",
    background=[("active", "#3700B3"), ("!disabled", primary_color)]
)

main_frame = ttk.Frame(root, padding=10)
main_frame.grid(sticky="nsew")

# Variables with defaults from the environment
download_var = tk.StringVar(value=os.getenv("DOWNLOAD_DIR", ""))
work_var = tk.StringVar(value=os.getenv("WORK_DIR", ""))
meal_var = tk.StringVar(value=os.getenv("MEAL_FEE", "5500"))
names_var = tk.StringVar(value=os.getenv("OFFICIAL_DATA_NAMES_STR", ""))


def validate_fields(*_: str) -> None:
    """Enable run button only when required fields are set."""
    if download_var.get().strip() and work_var.get().strip():
        run_button.state(["!disabled"])
        warning_label.grid_remove()
    else:
        run_button.state(["disabled"])
        warning_label.grid()

# DOWNLOAD_DIR
ttk.Label(main_frame, text="DOWNLOAD_DIR").grid(row=0, column=0, sticky="e")
entry_download = ttk.Entry(main_frame, textvariable=download_var, width=40)
entry_download.grid(row=0, column=1, padx=5)
btn_download = ttk.Button(main_frame, text="찾기", style="Material.TButton", command=lambda: browse_directory(download_var))
btn_download.grid(row=0, column=2)

# WORK_DIR
ttk.Label(main_frame, text="WORK_DIR").grid(row=1, column=0, sticky="e")
entry_work = ttk.Entry(main_frame, textvariable=work_var, width=40)
entry_work.grid(row=1, column=1, padx=5)
btn_work = ttk.Button(main_frame, text="찾기", style="Material.TButton", command=lambda: browse_directory(work_var))
btn_work.grid(row=1, column=2)

# MEAL_FEE
ttk.Label(main_frame, text="MEAL_FEE").grid(row=2, column=0, sticky="e")
ttk.Entry(main_frame, textvariable=meal_var).grid(row=2, column=1, columnspan=2, sticky="we", padx=5)

# OFFICIAL_DATA_NAMES_STR
ttk.Label(main_frame, text="OFFICIAL_DATA_NAMES_STR").grid(row=3, column=0, sticky="e")
ttk.Entry(main_frame, textvariable=names_var).grid(row=3, column=1, columnspan=2, sticky="we", padx=5)

# Run button and warning label
run_button = ttk.Button(main_frame, text="실행", style="Material.TButton", command=run_script)
run_button.grid(row=4, column=0, columnspan=3, pady=10)

warning_label = ttk.Label(
    main_frame,
    text="DOWNLOAD_DIR과 WORK_DIR을 모두 입력하세요.",
    foreground="red",
)
warning_label.grid(row=5, column=0, columnspan=3)
warning_label.grid_remove()

# Validate fields whenever they change
download_var.trace_add("write", validate_fields)
work_var.trace_add("write", validate_fields)
validate_fields()

root.mainloop()
