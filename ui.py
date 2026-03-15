import customtkinter
from tkinter import filedialog
from datetime import datetime
import threading
from create_workbook import master_function  # For creating master workbook
from generate_files import generate_files  # For generating ready-to-upload files
from generate_files import SUBTASK_CONFIG  # Import the SUBTASK_CONFIG

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

# Create main application window
root = customtkinter.CTk()
root.title("FLEX Ultimate")
root.geometry("700x390")  # Increase width for the new frame

selected_file = None
selected_workbook = None
checkbox_values = {}

# Function to load the input file for master workbook
def load_CODE():
    global selected_file
    file_path = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        selected_file = file_path
        loaded_prompt.configure(text=f"File selected: {file_path.split('/')[-1]}")

# Function to load an existing workbook for ready-to-upload files
def load_workbook():
    global selected_workbook
    file_path = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        selected_workbook = file_path
        workbook_prompt.configure(text=f"Workbook selected: {file_path.split('/')[-1]}")

def update_progress(progress):
    progress_bar.set(progress)

# Worker thread for master workbook creation
def worker_create_workbook(output_file_path):
    try:
        master_function(selected_file, output_file_path, update_progress)
        create_prompt.configure(text=f"Workbook created successfully at {output_file_path}!")
    finally:
        progress_bar.stop()
        progress_bar.set(0.0)
        progress_bar.pack_forget()

def master_function_wrapper():
    global selected_file
    if not selected_file:
        create_prompt.configure(text="Please select an input file!")
        return

    today = datetime.now().strftime("%m%d%Y")
    default_file_name = f"Workbook_{today}.xlsx"

    output_file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        initialfile=default_file_name
    )

    if not output_file_path:
        create_prompt.configure(text="Save operation cancelled.")
        return

    progress_bar.pack(pady=(0, 0), padx=10)
    progress_bar.start()
    progress_bar.set(0.0)
    create_prompt.configure(text="")

    thread = threading.Thread(target=worker_create_workbook, args=(output_file_path,))
    thread.start()

# Function to generate ready-to-upload files
def generate_files_wrapper():
    """
    Wrapper to handle UI interaction for generating files.
    """
    global selected_workbook
    if not selected_workbook:
        workbook_prompt.configure(text="Please select a workbook!")
        return

    # Gather selected subtasks from checkboxes
    selected_subtasks = [task for task, value in checkbox_values.items() if value.get() == "yes"]
    if not selected_subtasks:
        workbook_prompt.configure(text="Please select at least one subtask to generate!")
        return

    # Choose output folder
    output_folder = filedialog.askdirectory(title="Select Output Folder")
    if not output_folder:
        generate_prompt.configure(text="Output folder selection cancelled.")
        return

    # Call the modular generate_files function
    try:
        generate_files(selected_subtasks, selected_workbook, output_folder)
        generate_prompt.configure(text="Ready-to-upload files generated successfully!")
    except Exception as e:
        generate_prompt.configure(text=f"Error: {str(e)}")

# Frame 1: Existing frame for master workbook creation
frame1 = customtkinter.CTkFrame(master=root)
frame1.pack(side="left", pady=10, padx=10, fill="both", expand=True)

label = customtkinter.CTkLabel(master=frame1, text="Create workbook from raw export:", font=('Roboto', 20))
label.pack(pady=(15, 5), padx=10)

browse_button = customtkinter.CTkButton(master=frame1, text='Browse files', command=load_CODE)
browse_button.pack(pady=(8, 5), padx=10)

loaded_prompt = customtkinter.CTkLabel(master=frame1, text="No file selected", font=('Roboto', 12), wraplength=280)
loaded_prompt.pack(pady=(5, 15), padx=10)

createbutton = customtkinter.CTkButton(master=frame1, text='Create Workbook', fg_color='Red4', command=master_function_wrapper)
createbutton.pack(pady=(20, 10), padx=10)

create_prompt = customtkinter.CTkLabel(master=frame1, text="", font=('Roboto', 12), wraplength=280)
create_prompt.pack(pady=(10, 0), padx=10)

progress_bar = customtkinter.CTkProgressBar(master=frame1)
progress_bar.configure(mode="determinate")
progress_bar.set(0.0)

# Frame 2: New frame for ready-to-upload files generation
frame2 = customtkinter.CTkFrame(master=root)
frame2.pack(side="right", pady=10, padx=10, fill="both", expand=True)

workbook_label = customtkinter.CTkLabel(master=frame2, text="Generate ready-to-upload files:", font=('Roboto', 20))
workbook_label.pack(pady=(15, 5), padx=10)

workbook_button = customtkinter.CTkButton(master=frame2, text="Browse workbook", command=load_workbook)
workbook_button.pack(pady=(8, 5), padx=10)

workbook_prompt = customtkinter.CTkLabel(master=frame2, text="No workbook selected", font=('Roboto', 12), wraplength=280)
workbook_prompt.pack(pady=(5, 8), padx=10)

# Checkboxes for subtasks - Keywords removed
checkbox_values = {
    "PD BP": customtkinter.StringVar(value="no"),
    "Attributes": customtkinter.StringVar(value="no"),
    "TCU RA": customtkinter.StringVar(value="no"),
    "Title Cleanup": customtkinter.StringVar(value="no"),
}

for task in checkbox_values.keys():
    cb = customtkinter.CTkCheckBox(master=frame2, text=task, variable=checkbox_values[task], onvalue="yes", offvalue="no")
    cb.pack(anchor="w", pady=2, padx=10)

generate_button = customtkinter.CTkButton(master=frame2, text="Generate Files", fg_color="DarkOrange3", command=generate_files_wrapper)
generate_button.pack(pady=20, padx=10)

# New label for generate files status
generate_prompt = customtkinter.CTkLabel(master=frame2, text="", font=('Roboto', 12), wraplength=280)
generate_prompt.pack(pady=(0, 0), padx=10)

# Run the application
root.mainloop()