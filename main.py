import tkinter as tk
from excel_ui import MiniExcel
from template import CalendarTemplate

class StartWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Welcome to Mini Excel")
         
        self.root.geometry("1500x1200")
 
        left_frame = tk.Frame(root, width=300, height=1000, bg='green')
        left_frame.pack(side='left', fill='both', expand=True)
        
        right_frame = tk.Frame(root, width=500, height=1000, bg='white')
        right_frame.pack(side='right', fill='both', expand=True)
 
        welcome_label = tk.Label(left_frame, text="Welcome \n to \n Mini Excel", font=("Garamond", 30), bg='green', fg='white')
        welcome_label.pack(pady=200)  
 
        content_frame = tk.Frame(right_frame, bg='white')
        content_frame.place(relx=0.5, rely=0.5, anchor='center')  # Center the content_frame
 
        label = tk.Label(content_frame, text="Choose an option:", font=("Garamond", 24))
        label.pack(pady=10)
 
        blank_button = tk.Button(content_frame, text="Blank Page", font=("Garamond", 20), command=self.open_blank_page)
        blank_button.pack(pady=5)
 
        template_button = tk.Button(content_frame, text="Templates", font=("Garamond", 20), command=self.open_templates)
        template_button.pack(pady=5)

    def open_blank_page(self):
        self.root.withdraw()  # Hide the start window
        new_window = tk.Toplevel(self.root)
        app = MiniExcel(new_window, back_callback=self.show_start_window)  

    def open_templates(self):
        """Open the template window for calendar creation."""
        self.root.withdraw()   
        new_window = tk.Toplevel(self.root)  # Open a new window
        app = CalendarTemplate(new_window, back_callback=self.show_start_window)

    def show_start_window(self):
        """Show the start window again."""
        self.root.deiconify()  # Show the start window

if __name__ == "__main__":
    root = tk.Tk()
    app = StartWindow(root)
    root.mainloop()
