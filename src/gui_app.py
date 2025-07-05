import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
import io
import subprocess
import platform

from utils.config_utils import load_config, save_config

from main import main as cli_main

# Ensure we can import main from the same directory
sys.path.append(os.path.dirname(__file__))


project_root = os.path.dirname(os.path.dirname(__file__))

# Load configuration from pickle file
config = load_config()

# Define style configuration
style_config = {
    "primary_color": "#072d6e",
    "secondary_color": "#03519c",
    "background_color": "#F5F5F5",
    "font": ("Pretendard", 12)
}


class TextRedirector(io.TextIOBase):
    """Redirect stdout to a text widget."""

    def __init__(self, widget: tk.Text) -> None:
        self.widget = widget

    def write(self, text: str) -> int:
        self.widget.configure(state="normal")
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)
        self.widget.configure(state="disabled")
        self.widget.update_idletasks()
        return len(text)

    def flush(self) -> None:  # pragma: no cover - nothing to flush
        pass


def open_folder(path: str) -> None:
    """Open folder in OS file explorer."""
    try:
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


def browse_directory(var: tk.StringVar) -> None:
    path = filedialog.askdirectory()
    if path:
        var.set(path)


def run_script() -> None:
    """
    Runs the main script with the environment variables set from the GUI.
    Displays a success or error message upon completion.
    """
    # Ensure required environment variables are provided
    if not download_var.get().strip() or not work_var.get().strip():
        messagebox.showwarning("입력 필요", "DOWNLOAD_DIR과 WORK_DIR을 모두 설정하세요.")
        return

    # Prepare arguments to pass to cli_main
    download_dir = download_var.get()
    work_dir = work_var.get()
    meal_fee = meal_var.get()
    official_data_names_str = names_var.get()

    old_stdout = sys.stdout
    try:
        run_button.state(["disabled"])
        log_text.configure(state="normal")
        log_text.delete("1.0", tk.END)
        log_text.configure(state="disabled")

        sys.stdout = TextRedirector(log_text)
        cli_main(download_dir, work_dir, meal_fee, official_data_names_str)
        messagebox.showinfo("완료", "처리가 완료되었습니다.")
        open_folder(work_dir)
    except Exception as exc:
        messagebox.showerror(
            "오류",
            f"오류가 발생했습니다: {str(exc)}\n"
            f"DOWNLOAD_DIR: {download_dir}\n"
            f"WORK_DIR: {work_dir}\n"
            f"MEAL_FEE: {meal_fee}\n"
            f"OFFICIAL_DATA_NAMES_STR: {official_data_names_str}"
        )
    finally:
        sys.stdout = old_stdout
        run_button.state(["!disabled"])


def save_and_exit():
    """Save current values to the configuration file and exit."""
    save_config(
        {
            "DOWNLOAD_DIR": download_var.get(),
            "WORK_DIR": work_var.get(),
            "MEAL_FEE": meal_var.get(),
            "OFFICIAL_DATA_NAMES_STR": names_var.get(),
        }
    )
    root.destroy()


root = tk.Tk()
root.title("초과근무 처리 도구")
root.minsize(600, 500)

# ---- Material-like styling ----
root.configure(bg=style_config["background_color"])

style = ttk.Style(root)
style.theme_use("clam")

style.configure(
    "TFrame", background=style_config["background_color"]
)
style.configure(
    "TLabel",
    background=style_config["background_color"],
    font=style_config["font"]
)
style.configure(
    "TEntry",
    font=style_config["font"]
)
style.configure(
    "Material.TButton",
    font=(style_config["font"][0], style_config["font"][1], "bold"),
    foreground="white",
    background=style_config["primary_color"]
)
style.map(
    "Material.TButton",
    background=[("active", style_config["secondary_color"]),
                ("!disabled", style_config["primary_color"])]
)

# Use generous padding to create breathing room around widgets
main_frame = ttk.Frame(root, padding=30)
main_frame.grid(sticky="nsew")

# Variables with defaults from configuration
download_var = tk.StringVar(value=config.get("DOWNLOAD_DIR", ""))
work_var = tk.StringVar(value=config.get("WORK_DIR", ""))
meal_var = tk.StringVar(value=config.get("MEAL_FEE", "5500"))
names_var = tk.StringVar(value=config.get("OFFICIAL_DATA_NAMES_STR", ""))


def validate_fields(*_: str) -> None:
    """Enable run button only when required fields are set."""
    warning_label.config(text="DOWNLOAD_DIR과 WORK_DIR을 모두 입력하세요.")
    if download_var.get().strip() and work_var.get().strip():
        run_button.state(["!disabled"])
        warning_label.config(state='disabled')
    else:
        run_button.state(["disabled"])
        warning_label.config(state='normal')


# DOWNLOAD_DIR
ttk.Label(main_frame, text="다운로드 폴더").grid(
    row=0, column=0, sticky="e", pady=12, padx=20)
entry_download = ttk.Entry(main_frame, textvariable=download_var, width=40)
entry_download.grid(row=0, column=1, padx=20, pady=12, sticky="ew")
btn_download = ttk.Button(main_frame, text="찾기", style="Material.TButton",
                          command=lambda: browse_directory(download_var))
btn_download.grid(row=0, column=2, pady=12, padx=20, sticky="ew")

# WORK_DIR
ttk.Label(main_frame, text="작업결과 저장 폴더").grid(
    row=1, column=0, sticky="e", pady=12, padx=20)
entry_work = ttk.Entry(main_frame, textvariable=work_var, width=40)
entry_work.grid(row=1, column=1, padx=20, pady=12, sticky="ew")
btn_work = ttk.Button(main_frame, text="찾기", style="Material.TButton",
                      command=lambda: browse_directory(work_var))
btn_work.grid(row=1, column=2, pady=12, padx=20, sticky="ew")

# MEAL_FEE
ttk.Label(main_frame, text="매식비 기준 금액").grid(
    row=2, column=0, sticky="e", pady=12, padx=20)
ttk.Entry(main_frame, textvariable=meal_var).grid(
    row=2, column=1, columnspan=2, sticky="ew", padx=20, pady=12)

# OFFICIAL_DATA_NAMES_STR
ttk.Label(main_frame, text="대상자 이름").grid(
    row=3, column=0, sticky="e", pady=12, padx=20)
ttk.Entry(main_frame, textvariable=names_var).grid(
    row=3, column=1, columnspan=2, sticky="ew", padx=20, pady=12)

# Run button and warning label
run_button = ttk.Button(main_frame, text="실행",
                        style="Material.TButton", command=run_script)
run_button.grid(row=4, column=0, columnspan=2,
                pady=(20, 12), padx=20, sticky="ew")

exit_button = ttk.Button(main_frame, text="종료",
                         style="Material.TButton", command=save_and_exit)
exit_button.grid(row=4, column=2, pady=(20, 12), padx=20, sticky="ew")

# Log output area
log_text = ScrolledText(main_frame, height=10, state="disabled",
                        font=style_config["font"])
log_text.grid(row=5, column=0, columnspan=3, sticky="nsew", padx=20, pady=(0, 20))

# Optional: Add tooltip for run_button when disabled


button_tooltip = {}


def show_tooltip(event):
    if "disabled" in run_button.state():
        x = run_button.winfo_rootx() + run_button.winfo_width() // 2
        y = run_button.winfo_rooty() + run_button.winfo_height() + 10
        tooltip = tk.Toplevel(run_button)
        tooltip.wm_overrideredirect(True)
        tooltip.geometry(f"+{x}+{y}")
        label = tk.Label(tooltip, text="필수 입력값을 모두 채워주세요.",
                         background="yellow", borderwidth=1)
        label.pack()
        button_tooltip["tooltip"] = tooltip


def hide_tooltip(event):
    tooltip = button_tooltip.pop("tooltip", None)
    if tooltip:
        tooltip.destroy()


run_button.bind("<Enter>", show_tooltip)
run_button.bind("<Leave>", hide_tooltip)

warning_label = ttk.Label(
    main_frame,
    text="DOWNLOAD_DIR과 WORK_DIR을 모두 입력하세요.",
    foreground="red",
)
warning_label.grid(row=6, column=0, columnspan=3, pady=(4, 0), padx=20)
warning_label.grid_remove()

# Validate fields whenever they change
download_var.trace_add("write", validate_fields)
work_var.trace_add("write", validate_fields)
validate_fields()

root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
main_frame.columnconfigure(1, weight=1)
main_frame.rowconfigure(5, weight=1)

root.mainloop()
