import os
import sys
from tkinter import Button, Label, StringVar, Tk, filedialog, messagebox, PhotoImage
import pandas as pd
from openpyxl.styles import Border, Font, PatternFill, Side


def process_csv_to_excel(input_csv_path):
    try:
        print(f"Importing data from: {input_csv_path}\n")
        df = pd.read_csv(input_csv_path, sep=";", dtype=str)
        print("Data successfully imported. Applying transformations...\n")

        if "VEHICLE-NUMBER" in df.columns:
            df["VEHICLE-NUMBER"] = df["VEHICLE-NUMBER"].str[:17]

        if "ENGINE-NUMBER" in df.columns:
            df["ENGINE-NUMBER"] = df["ENGINE-NUMBER"].str[:14]

        if "DELIVERY-NUMBER" in df.columns:
            mid_part = df["DELIVERY-NUMBER"].str[4:6].fillna("")
            right_part = df["DELIVERY-NUMBER"].str[-2:].fillna("")
            lot_number_data = mid_part + right_part
            delivery_idx = df.columns.get_loc("DELIVERY-NUMBER")
            df.insert(delivery_idx + 1, "LOT-NUMBER", lot_number_data)

        if "CODES" in df.columns:
            unnamed_cols = [
                col for col in df.columns if str(col).startswith("Unnamed")
            ]

            def combine_row_codes(row):
                segments = []
                base_code = str(row["CODES"]).strip()
                if base_code and base_code.lower() != "nan":
                    segments.append(base_code)
                for col in unnamed_cols:
                    val = str(row[col]).strip()
                    if val and val.lower() != "nan":
                        segments.append(val)
                return "-".join(segments)

            df["COMBINED-CODE"] = df.apply(combine_row_codes, axis=1)
            codes_idx = df.columns.get_loc("CODES")
            combined_col_data = df.pop("COMBINED-CODE")
            df.insert(codes_idx, "COMBINED-CODE", combined_col_data)
            df.drop(columns=unnamed_cols, inplace=True)
            df.drop(columns=["CODES"], inplace=True)

        file_dir, file_name = os.path.split(input_csv_path)
        base_name = os.path.splitext(file_name)[0]
        output_excel_path = os.path.join(file_dir, base_name + "_processed.xlsx")

        with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Processed Data")
            worksheet = writer.sheets["Processed Data"]

            thin_border = Border(
                left=Side(style="thin", color="D3D3D3"),
                right=Side(style="thin", color="D3D3D3"),
                top=Side(style="thin", color="D3D3D3"),
                bottom=Side(style="thin", color="D3D3D3"),
            )
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            header_fill = PatternFill(
                start_color="366092", end_color="366092", fill_type="solid"
            )
            highlight_fill = PatternFill(
                start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
            )
            highlight_cols = {
                "VEHICLE-NUMBER",
                "ENGINE-NUMBER",
                "DELIVERY-NUMBER",
                "LOT-NUMBER",
                "COMBINED-CODE",
            }

            for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row):
                for cell in row:
                    cell.border = thin_border
                    if cell.row == 1:
                        cell.font = header_font
                        cell.fill = header_fill
                    if cell.row > 1:
                        current_header = worksheet.cell(
                            row=1, column=cell.column
                        ).value
                        if current_header in highlight_cols:
                            cell.fill = highlight_fill

            for col in worksheet.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                worksheet.column_dimensions[col_letter].width = max_len + 3

        messagebox.showinfo(
            "Success",
            f"Data successfully processed!\n\nSaved to: {output_excel_path}",
        )
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n\n{e}")


class AppGUI:

    def __init__(self, root):
        self.root = root
        self.root.title("CSV Data Processor")
        self.root.geometry("450x300")
        self.root.resizable(False, False)
        self.root.configure(bg="#121212")

        # Color Configuration Constants
        self.bg_dark = "#121212"
        self.bg_panel = "#1E1E1E"
        self.fg_white = "#FFFFFF"
        self.fg_gray = "#A0A0A0"

        self.file_path_var = StringVar()
        self.file_path_var.set("No file selected.")

        # --- Resolve Icon Path Safely ---
        if getattr(sys, "frozen", False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        icon_filename = os.path.join(base_path, "icon.png")

        # Initialize asset reference states
        self.app_icon = None
        has_icon = False

        # --- ONE-TIME IMAGE LOADING AND ASSIGNMENT ---
        if os.path.exists(icon_filename):
            try:
                # Load exactly once into a persistent instance variable
                self.app_icon = PhotoImage(file=icon_filename)

                # Set title bar icon
                self.root.iconphoto(False, self.app_icon)
                has_icon = True
            except Exception as e:
                print(f"Image runtime initialization error: {e}")

        # 1. ICON / LOGO DISPLAY PANEL
        self.label_icon = Label(root, bg=self.bg_dark, fg=self.fg_gray)
        self.label_icon.pack(pady=(15, 2))

        if has_icon:
            # Reuse the single verified image reference object safely
            self.label_icon.config(image=self.app_icon)
        else:
            self.label_icon.config(
                text="▪ ▪ ▪  [ ICON PLACEHOLDER ]  ▪ ▪ ▪", font=("Consolas", 9)
            )

        # 2. Text Labels
        self.label_title = Label(
            root,
            text="SELECTED INPUT FILE:",
            font=("Consolas", 10, "bold"),
            bg=self.bg_dark,
            fg=self.fg_white,
        )
        self.label_title.pack(pady=(5, 2))

        self.label_path = Label(
            root,
            textvariable=self.file_path_var,
            wraplength=400,
            fg=self.fg_gray,
            bg=self.bg_dark,
            font=("Consolas", 9),
        )
        self.label_path.pack(pady=(0, 10))

        # 3. Stark Action Buttons
        self.btn_browse = Button(
            root,
            text="BROWSE FILE",
            width=18,
            font=("Consolas", 9, "bold"),
            bg=self.bg_panel,
            fg=self.fg_white,
            activebackground=self.fg_white,
            activeforeground=self.bg_dark,
            bd=1,
            relief="flat",
            command=self.browse_file,
        )
        self.btn_browse.pack(pady=4)

        self.btn_process = Button(
            root,
            text="EXECUTE PROCESS",
            width=18,
            font=("Consolas", 9, "bold"),
            bg=self.fg_white,
            fg=self.bg_dark,
            activebackground=self.bg_panel,
            activeforeground=self.fg_white,
            bd=0,
            relief="flat",
            command=self.start_processing,
        )
        self.btn_process.pack(pady=4)

    def browse_file(self):
        selected_file = filedialog.askopenfilename(
            title="Select CSV file",
            filetypes=[("CSV files", "*.csv"), ("all files", "*.*")],
        )
        if selected_file:
            self.file_path_var.set(selected_file)

    def start_processing(self):
        current_path = self.file_path_var.get()
        if current_path == "No file selected." or not current_path:
            messagebox.showwarning(
                "Warning", "Please select a valid CSV file first before executing."
            )
            return
        process_csv_to_excel(current_path)


def run_gui():
    root = Tk()
    app = AppGUI(root)
    root.mainloop()


if __name__ == "__main__":
    run_gui()