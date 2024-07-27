import pandas as pd
import json
import swatch
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES

def rgb_to_hex(rgb):
    return '#%02x%02x%02x' % (int(rgb[0] * 255), int(rgb[1] * 255), int(rgb[2] * 255))

def select_ase_file():
    file_path = filedialog.askopenfilename(
        title="Select ASE file",
        filetypes=(("ASE files", "*.ase"), ("All files", "*.*"))
    )
    ase_entry.delete(0, tk.END)
    ase_entry.insert(0, file_path)

def select_output_dir():
    dir_path = filedialog.askdirectory(
        title="Select Output Directory"
    )
    output_entry.delete(0, tk.END)
    output_entry.insert(0, dir_path)

def run_conversion():
    ase_file = ase_entry.get()
    output_dir = output_entry.get()

    if not ase_file or ase_file == "Drop ASE file here":
        messagebox.showerror("Error", "Please select an ASE file.")
        return
    if not output_dir:
        messagebox.showerror("Error", "Please select an output directory.")
        return

    try:
        colors = swatch.parse(ase_file)
        
        json_file_path = f"{output_dir}/colors.json"
        xlsx_file_path = f"{output_dir}/colors_converted.xlsx"

        # Convert `colors` to JSON and write it to a file
        with open(json_file_path, 'w') as f:
            json.dump(colors, f, indent=4)

        # Load the JSON data
        with open(json_file_path) as f:
            data = json.load(f)

        # Process data
        index = 1
        colors_data = []
        for item in data:
            name = item['name']
            color_type = item['type']
            mode = item['data']['mode']
            values = item['data']['values']
            hex_value = rgb_to_hex(values)
            colors_data.append({'index': index, 'name': name, 'hex_value': hex_value, 'type': color_type, 'mode': mode, 'rgb_value': values})
            index += 1 

        # 将处理后的数据转换为DataFrame
        df = pd.DataFrame(colors_data)

        # 写入Excel文件
        df.to_excel(xlsx_file_path, index=False)

        messagebox.showinfo("Success", f"Files saved to:\n{json_file_path}\n{xlsx_file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def on_ase_file_drop(event):
    ase_file_path = event.data.strip("{}")  # Remove curly braces if present
    ase_entry.delete(0, tk.END)
    ase_entry.insert(0, ase_file_path)

def on_entry_click(event):
    if ase_entry.get() == "Drop ASE file here":
        ase_entry.delete(0, tk.END)
        ase_entry.config(fg='black')

def on_focusout(event):
    if ase_entry.get() == '':
        ase_entry.insert(0, "Drop ASE file here")
        ase_entry.config(fg='grey')

app = TkinterDnD.Tk()
app.title("ASE to JSON and XLSX Converter")

# ASE file input
tk.Label(app, text="ASE File:").grid(row=0, column=0, padx=10, pady=10)
ase_entry = tk.Entry(app, width=50, fg='grey')
ase_entry.insert(0, "Drop ASE file here")
ase_entry.bind('<FocusIn>', on_entry_click)
ase_entry.bind('<FocusOut>', on_focusout)
ase_entry.grid(row=0, column=1, padx=10, pady=10)
ase_button = tk.Button(app, text="Browse", command=select_ase_file)
ase_button.grid(row=0, column=2, padx=10, pady=10)

# Output directory input
tk.Label(app, text="Output Directory:").grid(row=1, column=0, padx=10, pady=10)
output_entry = tk.Entry(app, width=50)
output_entry.grid(row=1, column=1, padx=10, pady=10)
output_button = tk.Button(app, text="Browse", command=select_output_dir)
output_button.grid(row=1, column=2, padx=10, pady=10)

# Run button
run_button = tk.Button(app, text="Run", command=run_conversion, width=50)
run_button.grid(row=2, column=1,padx=10, pady=20)

# Enable drag-and-drop functionality for the entire window
app.drop_target_register(DND_FILES)
app.dnd_bind('<<Drop>>', on_ase_file_drop)

app.mainloop()