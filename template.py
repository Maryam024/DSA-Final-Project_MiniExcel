import tkinter as tk
import calendar
import csv
from tkinter import filedialog  # Import filedialog for selecting file path

class CalendarTemplate:
    def __init__(self, root, back_callback=None):
        self.root = root
        self.root.title("Calendar Template")
        self.rows = 7   
        self.columns = 7  
        self.current_month = 1  # Default to January
        self.current_year = 2024  # Default year
        self.back_callback = back_callback   
        self.create_calendar_ui()

    def create_calendar_ui(self): 
        # Frame for buttons
        button_frame = tk.Frame(self.root, bg='green')
        button_frame.pack(padx=10, pady=5, side=tk.TOP, fill=tk.X)

        # Add Back button
        back_button = tk.Button(button_frame, text="Back", command=self.back_to_previous_page)
        back_button.pack(side=tk.LEFT, padx=10, pady=5)

        # Add Save button
        save_button = tk.Button(button_frame, text="Save", command=self.save_calendar)
        save_button.pack(side=tk.RIGHT, padx=10, pady=5)

        # Formula bar
        self.formula_var = tk.StringVar()
        self.formula_bar = tk.Entry(self.root, textvariable=self.formula_var, font=("Arial", 12), relief="sunken")
        self.formula_bar.pack(fill=tk.X, padx=10, pady=5)

        # Create a grid frame for the calendar
        self.grid_frame = tk.Frame(self.root)
        self.grid_frame.pack(padx=10, pady=10)

        # Weekday headers
        weekdays = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
        for j, day in enumerate(weekdays):
            label = tk.Label(self.grid_frame, text=day, font=("Arial", 10, "bold"), width=10, relief="raised", bg="lightgray")
            label.grid(row=0, column=j + 1, sticky="nsew")

        # Initialize the calendar grid
        self.entries = []
        for i in range(1, self.rows):
            row_entries = []
            for j in range(self.columns):
                entry = tk.Entry(self.grid_frame, width=10, justify='center', font=("Arial", 12))
                entry.grid(row=i, column=j + 1, sticky="nsew")
                entry.bind("<FocusIn>", lambda e, r=i - 1, c=j: self.select_cell(r, c))  # Adjust for 0-based indexing
                entry.bind("<FocusOut>", lambda e, r=i - 1, c=j: self.update_data(r, c))
                self.bind_arrow_keys(entry, i - 1, j)  # Bind arrow keys
                row_entries.append(entry)
            self.entries.append(row_entries)

        # Populate the calendar
        self.update_calendar()

        # Add navigation buttons
        prev_button = tk.Button(self.root, text="<< Previous", command=self.previous_month, fg="white", bg="green")
        prev_button.pack(side=tk.LEFT, padx=10, pady=10)

        next_button = tk.Button(self.root, text="Next >>", command=self.next_month, fg="white", bg="green")
        next_button.pack(side=tk.RIGHT, padx=10, pady=10)

    def back_to_previous_page(self): 
        if self.back_callback:
            self.root.destroy()
            self.back_callback()

    def update_formula_bar(self): 
        self.formula_var.set(f"{self.get_month_name(self.current_month)} {self.current_year}")

    def update_calendar(self): 
        # Clear all existing entries
        for row_entries in self.entries:
            for entry in row_entries:
                entry.delete(0, tk.END)
 
        first_day, days_in_month = calendar.monthrange(self.current_year, self.current_month)
 
        day = 1
        for i in range(6):  # Max 6 rows for dates
            for j in range(7):  # 7 columns for days
                if (i == 0 and j < first_day) or day > days_in_month:
                    continue
                self.entries[i][j].insert(0, str(day))
                day += 1
  
        self.update_formula_bar()

    def get_month_name(self, month):
        """Return the full name of a month given its number."""
        return calendar.month_name[month]

    def next_month(self):
        """Switch to the next month."""
        if self.current_month == 12:
            self.current_month = 1
            self.current_year += 1
        else:
            self.current_month += 1
        self.update_calendar()

    def previous_month(self):
        """Switch to the previous month."""
        if self.current_month == 1:
            self.current_month = 12
            self.current_year -= 1
        else:
            self.current_month -= 1
        self.update_calendar()

    def save_calendar(self):
        """Save the current calendar data to a user-selected file."""
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])

        if file_path:
            with open(file_path, mode="w", newline="") as file:
                writer = csv.writer(file)
                writer.writerow(["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"])
                for row_entries in self.entries:
                    row_data = [entry.get() for entry in row_entries]
                    writer.writerow(row_data)
            print(f"Calendar saved as {file_path}")

    def bind_arrow_keys(self, entry, row, col):
        """Bind arrow keys for navigating the calendar grid."""
        entry.bind("<Up>", lambda e: self.move_focus(row - 1, col))
        entry.bind("<Down>", lambda e: self.move_focus(row + 1, col))
        entry.bind("<Left>", lambda e: self.move_focus(row, col - 1))
        entry.bind("<Right>", lambda e: self.move_focus(row, col + 1))

    def move_focus(self, row, col):
        """Move focus to the specified cell."""
        if 0 <= row < len(self.entries) and 0 <= col < len(self.entries[0]):
            self.entries[row][col].focus_set()
            self.highlight_active_cell(row, col)

    def select_cell(self, row, column):
        """Handle cell focus event."""
        print(f"Cell selected at row {row}, column {column}")
        self.highlight_active_cell(row, column)

    def update_data(self, row, column):
        """Handle cell data update."""
        value = self.entries[row][column].get()
        print(f"Updated cell at row {row}, column {column} with value: {value}")

    def highlight_active_cell(self, row, col):
        """Highlight the active cell with a green outline."""
        for row_entries in self.entries:
            for entry in row_entries:
                entry.config(highlightbackground="white", highlightcolor="white", highlightthickness=1)
        self.entries[row][col].config(highlightbackground="green", highlightcolor="green", highlightthickness=2)
