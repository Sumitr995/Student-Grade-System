import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os

class StudentGradeManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Student Grade Management System")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        # File path for Excel database
        self.file_path = "student_grades.xlsx"
        
        # Create Excel file if it doesn't exist
        if not os.path.exists(self.file_path):
            df = pd.DataFrame(columns=["Student ID", "Student Name", "Mathematics", "OS", "DBMS"])
            df.to_excel(self.file_path, index=False)
        
        # Create and place frames
        self.create_input_frame()
        self.create_display_frame()
        
        # Load initial data
        self.load_data()
    
    def create_input_frame(self):
        input_frame = ttk.LabelFrame(self.root, text="Student Information")
        input_frame.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        # Student ID
        ttk.Label(input_frame, text="Student ID:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.student_id_entry = ttk.Entry(input_frame, width=20)
        self.student_id_entry.grid(row=0, column=1, padx=5, pady=5)
        
        # Student Name
        ttk.Label(input_frame, text="Student Name:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.student_name_entry = ttk.Entry(input_frame, width=20)
        self.student_name_entry.grid(row=1, column=1, padx=5, pady=5)
        
        # Mathematics Grade
        ttk.Label(input_frame, text="Mathematics:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.math_grade_entry = ttk.Entry(input_frame, width=5)
        self.math_grade_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        # OS Grade
        ttk.Label(input_frame, text="OS:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.os_grade_entry = ttk.Entry(input_frame, width=5)
        self.os_grade_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        
        # DBMS Grade
        ttk.Label(input_frame, text="DBMS:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.dbms_grade_entry = ttk.Entry(input_frame, width=5)
        self.dbms_grade_entry.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        
        # Buttons
        button_frame = ttk.Frame(input_frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Add", command=self.add_student).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Update", command=self.update_student).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="Delete", command=self.delete_student).grid(row=0, column=2, padx=5)
        ttk.Button(button_frame, text="Clear", command=self.clear_entries).grid(row=0, column=3, padx=5)
    
    def create_display_frame(self):
        display_frame = ttk.LabelFrame(self.root, text="Student Records")
        display_frame.grid(row=0, column=1, rowspan=2, padx=10, pady=10, sticky="nsew")
        
        # Create Treeview
        self.tree = ttk.Treeview(display_frame, columns=("ID", "Name", "Math", "OS", "DBMS"), show="headings", height=15)
        
        # Configure columns
        self.tree.heading("ID", text="Student ID")
        self.tree.heading("Name", text="Student Name")
        self.tree.heading("Math", text="Mathematics")
        self.tree.heading("OS", text="OS")
        self.tree.heading("DBMS", text="DBMS")
        
        self.tree.column("ID", width=70)
        self.tree.column("Name", width=150)
        self.tree.column("Math", width=80)
        self.tree.column("OS", width=80)
        self.tree.column("DBMS", width=80)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(display_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Grid layout
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind event to select item
        self.tree.bind("<ButtonRelease-1>", self.select_item)
    
    def load_data(self):
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Load data from Excel
        try:
            df = pd.read_excel(self.file_path)
            for _, row in df.iterrows():
                self.tree.insert("", "end", values=(
                    row["Student ID"],
                    row["Student Name"],
                    row["Mathematics"],
                    row["OS"],
                    row["DBMS"]
                ))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {str(e)}")
    
    def add_student(self):
        # Get values from entries
        student_id = self.student_id_entry.get().strip()
        student_name = self.student_name_entry.get().strip()
        math_grade = self.math_grade_entry.get().strip()
        os_grade = self.os_grade_entry.get().strip()
        dbms_grade = self.dbms_grade_entry.get().strip()
        
        # Validate inputs
        if not all([student_id, student_name, math_grade, os_grade, dbms_grade]):
            messagebox.showwarning("Warning", "All fields are required!")
            return
        
        # Add to DataFrame and save
        try:
            df = pd.read_excel(self.file_path)
            
            # Check if student ID already exists
            if student_id in df["Student ID"].values:
                messagebox.showwarning("Warning", "Student ID already exists!")
                return
            
            # Add new record
            new_record = pd.DataFrame({
                "Student ID": [student_id],
                "Student Name": [student_name],
                "Mathematics": [math_grade],
                "OS": [os_grade],
                "DBMS": [dbms_grade]
            })
            
            df = pd.concat([df, new_record], ignore_index=True)
            df.to_excel(self.file_path, index=False)
            
            # Refresh display
            self.load_data()
            self.clear_entries()
            messagebox.showinfo("Success", "Student record added successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add record: {str(e)}")
    
    def update_student(self):
        # Get values from entries
        student_id = self.student_id_entry.get().strip()
        student_name = self.student_name_entry.get().strip()
        math_grade = self.math_grade_entry.get().strip()
        os_grade = self.os_grade_entry.get().strip()
        dbms_grade = self.dbms_grade_entry.get().strip()
        
        # Validate inputs
        if not all([student_id, student_name, math_grade, os_grade, dbms_grade]):
            messagebox.showwarning("Warning", "All fields are required!")
            return
        
        # Update DataFrame and save
        try:
            df = pd.read_excel(self.file_path)
            
            # Find the student by ID
            mask = df["Student ID"] == student_id
            if not mask.any():
                messagebox.showwarning("Warning", "Student ID not found!")
                return
            
            # Update record
            df.loc[mask, "Student Name"] = student_name
            df.loc[mask, "Mathematics"] = math_grade
            df.loc[mask, "OS"] = os_grade
            df.loc[mask, "DBMS"] = dbms_grade
            
            df.to_excel(self.file_path, index=False)
            
            # Refresh display
            self.load_data()
            self.clear_entries()
            messagebox.showinfo("Success", "Student record updated successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update record: {str(e)}")
    
    def delete_student(self):
        # Get student ID
        student_id = self.student_id_entry.get().strip()
        
        if not student_id:
            messagebox.showwarning("Warning", "Please enter Student ID to delete!")
            return
        
        # Confirm deletion
        confirm = messagebox.askyesno("Confirm", "Are you sure you want to delete this record?")
        if not confirm:
            return
        
        # Delete from DataFrame and save
        try:
            df = pd.read_excel(self.file_path)
            
            # Find the student by ID
            mask = df["Student ID"] == student_id
            if not mask.any():
                messagebox.showwarning("Warning", "Student ID not found!")
                return
            
            # Delete record
            df = df[~mask]
            df.to_excel(self.file_path, index=False)
            
            # Refresh display
            self.load_data()
            self.clear_entries()
            messagebox.showinfo("Success", "Student record deleted successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete record: {str(e)}")
    
    def select_item(self, event):
        try:
            # Get selected item
            selected_item = self.tree.focus()
            if not selected_item:
                return
                
            values = self.tree.item(selected_item, "values")
            
            # Clear entries
            self.clear_entries()
            
            # Set values to entries
            self.student_id_entry.insert(0, values[0])
            self.student_name_entry.insert(0, values[1])
            self.math_grade_entry.insert(0, values[2])
            self.os_grade_entry.insert(0, values[3])
            self.dbms_grade_entry.insert(0, values[4])
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to select item: {str(e)}")
    
    def clear_entries(self):
        # Clear all entry fields
        self.student_id_entry.delete(0, tk.END)
        self.student_name_entry.delete(0, tk.END)
        self.math_grade_entry.delete(0, tk.END)
        self.os_grade_entry.delete(0, tk.END)
        self.dbms_grade_entry.delete(0, tk.END)


if __name__ == "__main__":
    root = tk.Tk()
    app = StudentGradeManager(root)
    root.mainloop()