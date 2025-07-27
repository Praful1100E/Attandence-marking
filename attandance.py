import os
import json
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from ttkbootstrap import Style
from ttkbootstrap.constants import *
import pandas as pd

EXCEL_FILE = "Attendance_Records.xlsx"
JSON_FILE = "students.json"

def load_students():
    if os.path.exists(JSON_FILE):
        with open(JSON_FILE, 'r') as f:
            return json.load(f)
    return []

def save_students(student_list):
    with open(JSON_FILE, 'w') as f:
        json.dump(student_list, f, indent=2)

def load_or_create_excel():
    if os.path.exists(EXCEL_FILE):
        return pd.read_excel(EXCEL_FILE)
    else:
        df = pd.DataFrame(columns=["Roll Number", "Name"])
        df.to_excel(EXCEL_FILE, index=False)
        return df

def main():
    def update_display():
        student_listbox.delete(0, tk.END)
        for student in student_list:
            student_listbox.insert(tk.END, f"{student['roll']} - {student['name']}")
        count_label.config(text=f"Total Students: {len(student_list)}")
        update_date_dropdown()

    def add_student():
        name = name_entry.get().strip()
        roll = roll_entry.get().strip()
        if name and roll:
            if not any(s["roll"] == roll for s in student_list):
                student_list.append({"name": name, "roll": roll})
                save_students(student_list)
                name_entry.delete(0, tk.END)
                roll_entry.delete(0, tk.END)
                update_display()
            else:
                messagebox.showerror("Duplicate", "Roll number already exists!")

    def add_multiple_students():
        data = bulk_text.get("1.0", tk.END).strip()
        if not data:
            messagebox.showwarning("Input Missing", "Enter student data in 'Roll, Name' format.")
            return

        lines = data.split("\n")
        added = 0
        for line in lines:
            if ',' in line:
                parts = line.split(',')
                roll = parts[0].strip()
                name = parts[1].strip()
                if roll and name and not any(s["roll"] == roll for s in student_list):
                    student_list.append({"name": name, "roll": roll})
                    added += 1

        if added > 0:
            save_students(student_list)
            update_display()
            bulk_text.delete("1.0", tk.END)
            messagebox.showinfo("Success", f"{added} students added.")
        else:
            messagebox.showwarning("No Student Added", "No new valid students found.")

    def remove_student():
        selected = student_listbox.curselection()
        if selected:
            del student_list[selected[0]]
            save_students(student_list)
            update_display()

    def update_date_dropdown():
        df = load_or_create_excel()
        columns = df.columns.tolist()
        dates = [col for col in columns if col not in ["Roll Number", "Name"]]
        dates.sort()  # Ascending order
        date_options.set("")
        date_dropdown['menu'].delete(0, 'end')
        for d in dates:
            date_dropdown['menu'].add_command(label=d, command=tk._setit(date_options, d))

    def mark_attendance():
        selected = student_listbox.curselection()
        df = load_or_create_excel()

        # Determine selected date or use today
        selected_date = date_options.get().strip()
        date_to_use = selected_date if selected_date else datetime.now().strftime("%Y-%m-%d")

        if date_to_use not in df.columns:
            df[date_to_use] = "Absent"

        present_rolls = [student_list[i]["roll"] for i in selected]

        for student in student_list:
            roll = student["roll"]
            idx = df.index[df["Roll Number"] == roll]
            status = "Present" if roll in present_rolls else "Absent"

            if not idx.empty:
                df.at[idx[0], date_to_use] = status
            else:
                new_row = {
                    "Roll Number": roll,
                    "Name": student["name"],
                    date_to_use: status
                }
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

        df.to_excel(EXCEL_FILE, index=False)
        update_display()
        messagebox.showinfo("Success", f"Attendance updated for {date_to_use}.")

    # --- GUI Setup ---
    root = tk.Tk()
    root.title("🎓 Student Attendance System")
    root.geometry("650x800")
    Style(theme="darkly")

    global student_list
    student_list = load_students()

    tk.Label(root, text="📋 Attendance System", font=("Helvetica", 18, "bold")).pack(pady=15)

    # Entry Fields
    tk.Label(root, text="Name:", font=("Arial", 12)).pack()
    name_entry = tk.Entry(root, font=("Arial", 12), width=30)
    name_entry.pack(pady=5)

    tk.Label(root, text="Roll Number:", font=("Arial", 12)).pack()
    roll_entry = tk.Entry(root, font=("Arial", 12), width=30)
    roll_entry.pack(pady=5)

    # Buttons
    frame_btn = tk.Frame(root)
    frame_btn.pack(pady=10)
    tk.Button(frame_btn, text="➕ Add Student", command=add_student, width=15).grid(row=0, column=0, padx=5)
    tk.Button(frame_btn, text="❌ Remove Student", command=remove_student, width=15).grid(row=0, column=1, padx=5)
    tk.Button(frame_btn, text="✅ Mark Attendance", command=mark_attendance, width=20).grid(row=0, column=2, padx=5)

    # Bulk Add
    tk.Label(root, text="OR", font=("Arial", 11, "bold")).pack(pady=5)
    tk.Label(root, text="📥 Bulk Add (Roll, Name):", font=("Arial", 12)).pack()
    bulk_text = tk.Text(root, height=5, width=50, font=("Arial", 11))
    bulk_text.pack(pady=5)
    tk.Button(root, text="➕ Add Multiple Students", command=add_multiple_students, width=25).pack(pady=5)

    # Student List
    tk.Label(root, text="🎯 Select Students for Attendance", font=("Arial", 12, "bold")).pack(pady=10)
    student_listbox = tk.Listbox(root, height=15, width=50, font=("Arial", 12), selectmode=tk.MULTIPLE,
                                 bg="#2e2e2e", fg="white", selectbackground="gray")
    student_listbox.pack()

    count_label = tk.Label(root, text="", font=("Arial", 11))
    count_label.pack(pady=5)

    # Date Selection Dropdown
    tk.Label(root, text="📆 Select Date to Edit (optional)", font=("Arial", 12)).pack(pady=10)
    date_options = tk.StringVar()
    date_dropdown = tk.OptionMenu(root, date_options, "")
    date_dropdown.config(font=("Arial", 11), width=30)
    date_dropdown.pack()

    # Footer
    tk.Label(root, text="Made with ❤ using Python + Tkinter + Excel", font=("Arial", 9)).pack(pady=20)

    update_display()
    root.mainloop()

if __name__ == "__main__":
    main()