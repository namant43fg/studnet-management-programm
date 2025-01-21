import pandas as pd
import ttkbootstrap as tb
from ttkbootstrap.constants import SUCCESS, PRIMARY, WARNING, INFO, DANGER
from tkinter import messagebox, filedialog


class Student:
    def __init__(self, roll_no, name, internal_marks, exam_marks):
        self.roll_no = roll_no
        self.name = name
        self.internal_marks = internal_marks
        self.exam_marks = exam_marks

    def to_dict(self):
        return {
            "Roll No": self.roll_no,
            "Name": self.name,
            "Internal Marks": self.internal_marks,
            "Exam Marks": self.exam_marks,
        }

def save_to_excel(class_name, students):
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")]
    )
    if file_path:
        df = pd.DataFrame([student.to_dict() for student in students])
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", f"Data for class {class_name} saved to {file_path}")

def load_from_excel(tree, students):
    file_name = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_name:
        try:
            df = pd.read_excel(file_name)
            required_columns = ["Roll No", "Name", "Internal Marks", "Exam Marks"]
            if not all(col in df.columns for col in required_columns):
                messagebox.showerror("Error", "The selected file lacks required columns.")
                return

            students.clear()
            tree.delete(*tree.get_children())
            for _, row in df.iterrows():
                student = Student(
                    int(row["Roll No"]),
                    row["Name"],
                    float(row["Internal Marks"]),
                    float(row["Exam Marks"]),
                )
                students.append(student)
                tree.insert("", "end", values=student.to_dict().values())
            messagebox.showinfo("Success", f"Data loaded from {file_name}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {e}")

def add_student(root, tree, students):
    # Create a new dialog window
    dialog = tb.Toplevel(root)
    dialog.title("Add Student")
    dialog.geometry("400x400")
    dialog.attributes('-alpha', 0)  # Start with transparent window
    dialog.grab_set()

    # Animate the appearance of the dialog
    def fade_in():
        alpha = dialog.attributes('-alpha')
        if alpha < 1:
            alpha += 0.1
            dialog.attributes('-alpha', alpha)
            dialog.after(30, fade_in)  # Adjust timing for smoother animation

    fade_in()

    # Field labels and entry widgets
    fields = [
        ("Roll No", "int"),
        ("Name", "str"),
        ("Internal Marks", "float"),
        ("Exam Marks", "float"),
    ]
    entries = {}

    # Add entry fields dynamically
    for label, _ in fields:
        label_widget = tb.Label(dialog, text=label, font=("Arial", 12))
        label_widget.pack(pady=5, anchor="w")  # Align labels to the left
        entry_widget = tb.Entry(dialog, font=("Arial", 12))
        entry_widget.pack(pady=5, fill="x", padx=10)  # Stretch entries to fit the dialog width
        entries[label] = entry_widget

    # Define the submit action
    def submit():
        try:
            # Collect and validate input values
            values = {
                label: float(entries[label].get())
                if dtype == "float"
                else int(entries[label].get())
                if dtype == "int"
                else entries[label].get()
                for label, dtype in fields
            }
            if not values["Name"].strip():
                raise ValueError("Name cannot be empty.")

            # Map dictionary keys to match Student class parameters
            student = Student(
                roll_no=values["Roll No"],
                name=values["Name"],
                internal_marks=values["Internal Marks"],
                exam_marks=values["Exam Marks"]
            )

            students.append(student)
            tree.insert("", "end", values=student.to_dict().values())
            dialog.destroy()
        except ValueError as e:
            messagebox.showerror("Input Error", f"Invalid input: {e}")

    # Submit button with hover effects
    submit_button = tb.Button(
        dialog,
        text="Submit",
        command=submit,
        bootstyle=(SUCCESS, "outline"),
        width=20,
    )
    submit_button.pack(pady=20)

    # Add hover effect for the button
    def on_enter(e):
        submit_button.configure(bootstyle=SUCCESS)

    def on_leave(e):
        submit_button.configure(bootstyle=(SUCCESS, "outline"))

    submit_button.bind("<Enter>", on_enter)
    submit_button.bind("<Leave>", on_leave)

    # Ensure the dialog is modal
    root.wait_window(dialog)

def delete_selected(tree, students):
    selected_items = tree.selection()
    if not selected_items:
        messagebox.showwarning("No Selection", "Please select a student to delete.")
        return

    for item in selected_items:
        values = tree.item(item, "values")  # Retrieve the tuple of values
        if values and isinstance(values, tuple):  # Ensure `values` is a tuple
            try:
                roll_no = int(values[0])  # Safely convert Roll No to an integer
                students[:] = [s for s in students if s.roll_no != roll_no]
                tree.delete(item)  # Remove the selected item from the Treeview
            except ValueError:
                messagebox.showerror("Error", f"Invalid Roll No: {values[0]}")

    messagebox.showinfo("Success", "Selected student(s) deleted.")

def search_students(tree, students, query):
    tree.delete(*tree.get_children())
    for student in students:
        if query.lower() in student.name.lower():
            tree.insert("", "end", values=student.to_dict().values())

def main_app():
    root = tb.Window(themename="flatly")
    root.title("Student Data Management")
    root.geometry("1000x700")

    students = []

    frame = tb.Frame(root, padding=10)
    frame.pack(fill="both", expand=True)

    # Title
    tb.Label(frame, text="Student Data Management", font=("Arial", 24), bootstyle=INFO).pack(pady=10)

    # Search Bar
    search_frame = tb.Frame(frame)
    search_frame.pack(fill="x", pady=10)
    search_entry = tb.Entry(search_frame, font=("Arial", 12))
    search_entry.pack(side="left", fill="x", expand=True, padx=5)
    tb.Button(
        search_frame,
        text="Search",
        bootstyle=PRIMARY,
        command=lambda: search_students(tree, students, search_entry.get()),
    ).pack(side="left", padx=5)

    # Treeview
    tree = tb.Treeview(frame, columns=("Roll No", "Name", "Internal Marks", "Exam Marks"), show="headings")
    for col in ("Roll No", "Name", "Internal Marks", "Exam Marks"):
        tree.heading(col, text=col)
        tree.column(col, width=150, anchor="center")
    tree.pack(fill="both", expand=True, pady=10)

    # Button Frame
    button_frame = tb.Frame(frame)
    button_frame.pack(fill="x", pady=10)
    tb.Button(
        button_frame,
        text="Add Student",
        bootstyle=SUCCESS,
        command=lambda: add_student(root, tree, students),
    ).pack(side="left", padx=10)
    tb.Button(
        button_frame,
        text="Save Data",
        bootstyle=INFO,
        command=lambda: save_to_excel("Class Data", students),
    ).pack(side="left", padx=10)
    tb.Button(
        button_frame,
        text="Load Data",
        bootstyle=WARNING,
        command=lambda: load_from_excel(tree, students),
    ).pack(side="left", padx=10)
    tb.Button(
        button_frame,
        text="Delete Student",
        bootstyle=DANGER,
        command=lambda: delete_selected(tree, students),
    ).pack(side="left", padx=10)

    root.mainloop()

if __name__ == "__main__":
    main_app()
