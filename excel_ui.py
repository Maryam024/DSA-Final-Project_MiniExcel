import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import csv
from structures import LinkedList, Stack, Queue, Graph
import tkinter.font as tkFont
from tkinter import simpledialog
import re
import networkx as nx
import matplotlib.pyplot as plt
import numpy as np
from tkinter import simpledialog, messagebox
import matplotlib.pyplot as plt

class MiniExcel:
    def __init__(self, root, back_callback=None):
        self.root = root
        self.root.title("Mini Excel")
        self.root.geometry("1500x1200")

        self.rows = 50
        self.columns = 50
        self.entries = []
        self.back_callback = back_callback 
        
        # Default font: "Arial" with size 10
        self.current_font = ("Arial", 10)
        # Data Structures
        self.data = LinkedList()
        for _ in range(self.rows):
            row = LinkedList()
            for _ in range(self.columns):
                row.append("")
            self.data.append(row)

        self.history = Stack()
        self.redo_stack = Stack()
        self.batch_queue = Queue()
        self.dependencies = Graph()

        # Current Selected Cell
        self.current_cell = (0, 0)

        self.create_ui()

    def create_ui(self):
    # Menu 
      menu = tk.Menu(self.root, bg="green")   
      self.root.config(menu=menu)

      file_menu = tk.Menu(menu, tearoff=0, bg= 'green', fg='white')
      file_menu.add_command(label="Load CSV", command=self.load_csv)
      file_menu.add_command(label="Save CSV", command=self.save_csv)
      menu.add_cascade(label="File", menu=file_menu)

      edit_menu = tk.Menu(menu, tearoff=0)
      edit_menu.add_command(label="Undo", command=self.undo)
      edit_menu.add_command(label="Redo", command=self.redo)
      edit_menu.add_command(label="Batch Process", command=self.process_batch)
      menu.add_cascade(label="Edit", menu=edit_menu)

     # Visualization Menu
      vis_menu = tk.Menu(menu, tearoff=0)
      vis_menu.add_command(label="Show Dependency Graph", command=self.show_dependency_graph)
      vis_menu.add_command(label="Show Dependency Tree", command=self.show_dependency_tree)
      vis_menu.add_command(label="Show Bar Graph", command=self.show_bar_graph)  # New option
      menu.add_cascade(label="Visualize", menu=vis_menu)
      # Toolbar for alignment and color
      toolbar = tk.Frame(self.root, bg="green")
      toolbar.pack(fill=tk.X, padx=5, pady=3)

    # Alignment buttons
      align_left = tk.Button(toolbar, text="Left        ", command=lambda: self.change_alignment("left"))
      align_left.pack(side=tk.LEFT, padx=2, pady=5)
     
      align_center = tk.Button(toolbar, text="     Center     ", command=lambda: self.change_alignment("center"))
      align_center.pack(side=tk.LEFT, padx=2, pady=5)
    
      align_right = tk.Button(toolbar, text="        Right", command=lambda: self.change_alignment("right"))
      align_right.pack(side=tk.LEFT, padx=2, pady=5) 
      cm_to_pixels = 37.795
      distance_in_cm = 0.02
      pad_pixels = int(distance_in_cm * cm_to_pixels) 
    # Bold, Italic, and Underline buttons
      bold_button = tk.Button(toolbar, text="B  ", command=self.toggle_bold)
      bold_button.pack(side=tk.LEFT, padx=pad_pixels, pady=pad_pixels)

      italic_button = tk.Button(toolbar, text="I  ", command=self.toggle_italic)
      italic_button.pack(side=tk.LEFT,  padx=pad_pixels, pady=pad_pixels)

      underline_button = tk.Button(toolbar, text="U  ", command=self.toggle_underline)
      underline_button.pack(side=tk.LEFT,  padx=pad_pixels, pady=pad_pixels)

              # Font Size combo box
      self.size_var = tk.StringVar(value="10")
      size_combo = ttk.Combobox(toolbar, textvariable=self.size_var, state="readonly", width=5)
      size_combo["values"] = [str(size) for size in range(8, 32, 2)]
      size_combo.pack(side=tk.LEFT, padx=5, pady=5)
      size_combo.bind("<<ComboboxSelected>>", self.change_font)

        # Font Style combo box
      self.font_var = tk.StringVar(value="Arial")
      font_combo = ttk.Combobox(toolbar, textvariable=self.font_var, state="readonly", width=15)
      font_combo["values"] = ["Arial", "Calibri", "Times New Roman", "Courier New", "Verdana", "Helvetica", "Georgia", "Tahoma", "Trebuchet MS", "Comic Sans MS"]
      font_combo.pack(side=tk.LEFT, padx=5, pady=5)
      font_combo.bind("<<ComboboxSelected>>", self.change_font)
      # Color palette dropdown 
      self.color_var = tk.StringVar(value="Choose Color")
      color_palette = ttk.Combobox(toolbar, textvariable=self.color_var, state="readonly")
      color_palette["values"] = [ "White", "Red", "Green", "Blue", "Yellow", "Cyan", "Magenta","Pink", "Black", "Grey", "Maroon", "Brown", "Silver", "Orange"]
      color_palette.pack(side=tk.LEFT, padx=5, pady=5)
    
      apply_color = tk.Button(toolbar, text="Apply Color", command=self.apply_color_to_cell)
      apply_color.pack(side=tk.LEFT, padx=5, pady=5)

      # Create Frame for Formula Bar and Canvas
      main_frame = tk.Frame(self.root)
      main_frame.pack(fill=tk.BOTH, expand=True)

      # Formula Bar
      self.formula_var = tk.StringVar()
      formula_bar = tk.Entry(main_frame, textvariable=self.formula_var, font=("Arial", 14), justify='left')
      formula_bar.pack(fill=tk.X)
      self.formula_var.trace_add("write", lambda *args: self.update_from_formula_bar())

      # Canvas for Spreadsheet
      canvas = tk.Canvas(main_frame)
      canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

      # Scrollbars
      h_scrollbar = tk.Scrollbar(main_frame, orient=tk.HORIZONTAL, command=canvas.xview)
      h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

      v_scrollbar = tk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
      v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
 
      canvas.configure(xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set)
 
      self.grid_frame = tk.Frame(canvas)
      canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")
 
      self.grid_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
 
      self.entries = []

      # Column Headers
      for j in range(self.columns):
        label = tk.Label(self.grid_frame, text=chr(65 + j), font=("Arial", 10, "bold"), width=10, relief="raised")
        label.grid(row=0, column=j + 1, sticky="nsew")

      # Row Headers and Entries
      for i in range(self.rows):
        # Row Header
        row_label = tk.Label(self.grid_frame, text=str(i + 1), font=("Arial", 10, "bold"), width=5, relief="raised")
        row_label.grid(row=i + 1, column=0, sticky="nsew")

        row_entries = []
        for j in range(self.columns):
            entry = tk.Entry(self.grid_frame, width=10, justify='center')
            entry.grid(row=i + 1, column=j + 1, sticky="nsew")
            entry.bind("<FocusOut>", lambda e, r=i, c=j: self.update_data(r, c))
            entry.bind("<FocusIn>", lambda e, r=i, c=j: self.select_cell(r, c))
            entry.bind("<Down>", lambda e, r=i, c=j: self.move_focus(r + 1, c))  # Move down
            entry.bind("<Up>", lambda e, r=i, c=j: self.move_focus(r - 1, c))    # Move up
            entry.bind("<Right>", lambda e, r=i, c=j: self.move_focus(r, c + 1)) # Move right
            entry.bind("<Left>", lambda e, r=i, c=j: self.move_focus(r, c - 1))  # Move left
            entry.bind("<Return>", lambda e, r=i, c=j: self.move_focus(r + 1, c))  # Press Enter to move down
            row_entries.append(entry)
        self.entries.append(row_entries)
      menu_frame = tk.Frame(self.root)
      menu_frame.pack(side=tk.TOP, fill=tk.X)
      menu_frame.config(bg='green')


      back_button = tk.Button(menu_frame, text="  Back  ", command=self.back_to_previous_page)
      back_button.pack(side=tk.LEFT, padx=10, pady=5)

    def change_alignment(self, alignment): 
      
      row, col = self.current_cell
      entry = self.entries[row][col]
      if alignment == "left":
        entry.configure(justify="left")
      elif alignment == "center":
        entry.configure(justify="center")
      elif alignment == "right":
        entry.configure(justify="right")

    def apply_color_to_cell(self):
      row, col = self.current_cell
      selected_color = self.color_var.get().lower()
      if selected_color == "choose color": 
            
            messagebox.showwarning("Invalid Selection", "Please choose a color before applying!")
            return 
      
      self.entries[row][col].configure(bg=selected_color)
    def back_to_previous_page(self): 

        if self.back_callback:
            self.root.destroy()   
            self.back_callback()  

    def show_dependency_graph(self): 
      
      G = nx.DiGraph()
    
      # Populate the graph
      for node, dependents in self.dependencies.adjacency_list.items():
        print(f"Node: {node}, Dependents: {dependents}")  
        for dependent in dependents:
            G.add_edge(node, dependent)

      if G.number_of_edges() == 0:
        print("No dependencies found.")
        messagebox.showinfo("Dependency Graph", "No dependencies to visualize.")
        return

      # Plot the graph
      plt.figure(figsize=(10, 6))
      pos = nx.spring_layout(G)
      nx.draw(G, pos, with_labels=True, node_color="lightblue", node_size=2000, font_size=10, font_weight="bold")
      plt.title("Dependency Graph")
      plt.show()


    def show_dependency_tree(self): 
      
      selected_cell = f"{self.current_cell[0]},{self.current_cell[1]}"
      tree = nx.DiGraph()

      def build_tree(node):
        
        dependents = self.dependencies.get_dependents(node)
        for dependent in dependents:
            tree.add_edge(node, dependent)
            build_tree(dependent)
 
      build_tree(selected_cell)

      if tree.nodes:
        # Plot the tree
        plt.figure(figsize=(10, 6))
        pos = nx.spring_layout(tree)  # Position the nodes
        nx.draw(tree, pos, with_labels=True, node_color="lightgreen", 
                node_size=2000, font_size=10, font_weight="bold")
        plt.title(f"Dependency Tree for Cell {selected_cell}")
        plt.show()
      else:
        # No dependencies to show
        messagebox.showinfo("Dependency Tree", "No dependencies found for the selected cell.")

    def select_cell(self, row, col): 

        self.current_cell = (row, col)
        current_row = self.data.get(row)
        if current_row:
            value = current_row.value.get(col).value
            self.formula_var.set(value)
        self.highlight_active_cell( row, col)

    def apply_formula(self): 

        row, col = self.current_cell
        formula = self.formula_var.get()
        self.entries[row][col].delete(0, tk.END)
        self.entries[row][col].insert(0, formula)
        self.update_data(row, col)
    def update_from_formula_bar(self): 
      
      row, col = self.current_cell
      formula = self.formula_var.get()
      self.entries[row][col].delete(0, tk.END)
      self.entries[row][col].insert(0, formula)
      self.update_data(row, col)
 

    def show_bar_graph(self): 
     selected_cell = f"{self.current_cell[0]},{self.current_cell[1]}"
 
     choice = simpledialog.askstring("Graph Type", "Do you want to visualize a (row/column)?")
    
     if choice not in ['row', 'column']:
        messagebox.showinfo("Invalid choice", "Please select 'row' or 'column'.")
        return
 
     values = []
     labels = []
     cell_ids = []

     if choice == 'row':
        row = self.current_cell[0]
        for col in range(self.columns):
            value = self.entries[row][col].get()
            if value.isnumeric():  
                values.append(float(value))
                labels.append(chr(65 + col))  # Column labels like A, B, C, ...
                cell_ids.append(f"{row},{col}")  # Save the cell reference
                
     elif choice == 'column':
        col = self.current_cell[1]
        for row in range(self.rows):
            value = self.entries[row][col].get()
            if value.isnumeric():   
                values.append(float(value))
                labels.append(str(row + 1))   
                cell_ids.append(f"{row},{col}")  # Save the cell reference
 
     if not values:
        messagebox.showinfo("No Data", "No numeric data found for this row/column.")
        return
 
     fig, ax = plt.subplots(figsize=(10, 6))
     bars = ax.bar(labels, values, color='lightblue')
     ax.set_xlabel('Cells')
     ax.set_ylabel('Values')
     ax.set_title(f"Bar Graph for {choice.capitalize()} {selected_cell}")
 
     plt.ion()
     plt.show()

     def on_click(event): 

        if event.inaxes != ax:
            return
 
        bar_index = int(np.floor(event.xdata))
        if bar_index < 0 or bar_index >= len(bars):
            return
 
        cell_id = cell_ids[bar_index]
        current_value = float(values[bar_index])
 
        new_value = simpledialog.askfloat("Edit Value", f"Edit value for {cell_id} (current value: {current_value}):", initialvalue=current_value)
        if new_value is None:
            return  
        
        values[bar_index] = new_value
        self.update_cell_from_graph(cell_id, new_value)
 
        bars[bar_index].set_height(new_value)
 
        fig.canvas.draw()
 
     fig.canvas.mpl_connect('button_press_event', on_click)
 
     while plt.fignum_exists(fig.number):
        plt.pause(0.1)   


    def update_cell_from_graph(self, cell_id, value): 
        row, col = map(int, cell_id.split(','))
        self.entries[row][col].delete(0, tk.END)
        self.entries[row][col].insert(0, str(value))  # Update the grid with the new value
        self.data[row][col] = value   

    def update_data(self, row, col):
     cell_id = f"{row},{col}"
     value = self.entries[row][col].get()
     current_row = self.data.get(row)

     if current_row:
        current_value = current_row.value.get(col).value
        self.history.push((row, col, current_value))
        current_row.value.update(col, value)
     self.redo_stack = Stack()  # Clear redo stack after a new change

     # Handle formulas
     if value.startswith("="):  
        references = re.findall(r'[A-Z]\d+', value)   
        for ref in references:
            ref_col = ord(ref[0].upper()) - 65
            ref_row = int(ref[1:]) - 1
            ref_id = f"{ref_row},{ref_col}"
            self.dependencies.add_edge(ref_id, cell_id)   

        # Evaluate the formula
        result = self.evaluate_formula(value[1:])  # Strip the '='
        if result is not None:
            self.entries[row][col].delete(0, tk.END)
            self.entries[row][col].insert(0, result)
            current_row.value.update(col, result)  # Save result 
     else: 
        current_row.value.update(col, value)
 
     self.recalculate_dependents(cell_id)


    def recalculate_dependents(self, cell_id): 
     
     dependents = self.dependencies.get_dependents(cell_id)
     for dependent in dependents:
        row, col = map(int, dependent.split(","))
        # Get the formula for the dependent cell
        formula = self.entries[row][col].get()
        if formula.startswith("="): 
            result = self.evaluate_formula(formula[1:])
            # Update the cell value
            self.entries[row][col].delete(0, tk.END)
            self.entries[row][col].insert(0, result)
            # Update data structure
            self.data.get(row).value.update(col, result)

 

    def evaluate_formula(self, formula): 
     try: 
        tokens = re.findall(r'[A-Z]\d+|\d+|[-+*/()]', formula)
        for token in tokens:
            if re.match(r'[A-Z]\d+', token):  # Cell reference
                col = ord(token[0].upper()) - 65
                row = int(token[1:]) - 1
                value = self.data.get(row).value.get(col).value
                 
                if value is None or value == "":
                    value = 0  
                 
                try:
                    if isinstance(value, str) and value.isdigit():
                        value = int(value)
                    else:
                        value = float(value)
                except ValueError:
                    pass
                 
                formula = formula.replace(token, str(value))
        if any(c.isalpha() for c in formula):
            return formula  
        return eval(formula)

     except Exception as e:
        print(f"Error evaluating formula: {e}")
        return None

    def undo(self):
     if not self.history.is_empty():
        row, col, previous_value = self.history.pop()
        self.redo_stack.push((row, col, self.data.get(row).value.get(col).value))
        self.data.get(row).value.update(col, previous_value)
        self.entries[row][col].delete(0, tk.END)
        self.entries[row][col].insert(0, previous_value)
        self.recalculate_dependents(f"{row},{col}")
     else:
        messagebox.showinfo("Undo", "No actions to undo.")


    def redo(self):
        if not self.redo_stack.is_empty():
            row, col, value = self.redo_stack.pop()
            self.history.push((row, col, self.data.get(row).value.get(col).value))
            self.data.get(row).value.update(col, value)
            self.entries[row][col].delete(0, tk.END)
            self.entries[row][col].insert(0, value)
        else:
            messagebox.showinfo("Redo", "No actions to redo.")

    def load_csv(self):
        """Load spreadsheet data from a CSV file."""
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return
        try:
            with open(file_path, "r") as file:
                reader = csv.reader(file)
                for i, row in enumerate(reader):
                    if i >= self.rows:
                        break
                    current_row = self.data.get(i)
                    for j, value in enumerate(row):
                        if j >= self.columns:
                            break
                        current_row.value.update(j, value)
                        self.entries[i][j].delete(0, tk.END)
                        self.entries[i][j].insert(0, value)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV: {e}")

    def save_csv(self):
        """Save spreadsheet data to a CSV file."""
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return
        try:
            with open(file_path, "w", newline="") as file:
                writer = csv.writer(file)
                for i in range(self.rows):
                    row = []
                    current_row = self.data.get(i)
                    for j in range(self.columns):
                        row.append(current_row.value.get(j).value)
                    writer.writerow(row)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save CSV: {e}")

    def add_to_batch(self, row, col, value):
        self.batch_queue.enqueue((row, col, value))

    def process_batch(self):
        while not self.batch_queue.is_empty():
            row, col, value = self.batch_queue.dequeue()
            self.entries[row][col].delete(0, tk.END)
            self.entries[row][col].insert(0, value)
            self.update_data(row, col)

    def move_focus(self, row, col):
      """Move the focus to the specified cell."""
      row = max(0, min(row, self.rows - 1))   
      col = max(0, min(col, self.columns - 1))   
      self.entries[row][col].focus_set()

    def change_font(self,event=None):
   
      font_style = self.font_var.get()  
      font_size = int(self.size_var.get())   
      self.current_font = (font_style, font_size)
      self.apply_font_to_cells()   

    def change_font_size(self):
     """Change the font size of all cells."""
     selected_size = simpledialog.askinteger("Font Size", "Enter font size (e.g., 10, 12, 14):", initialvalue=self.current_font[1])
     if selected_size and 8 <= selected_size <= 72:
        self.current_font = (self.current_font[0], selected_size)
        self.apply_font_to_cells()
     else:
        tk.messagebox.showerror("Error", "Font size must be between 8 and 72!")

    def apply_font_to_cells(self):
     """Apply the current font to all cells."""
     for row_entries in self.entries:
        for entry in row_entries:
            entry.config(font=self.current_font)

    def highlight_active_cell(self, row, col): 
      for i, row_entries in enumerate(self.entries):
        for j, entry in enumerate(row_entries):
            entry.config(highlightbackground="white", highlightcolor="white", highlightthickness=1)
     
      self.entries[row][col].config(highlightbackground="green", highlightcolor="green", highlightthickness=2)
 
    def toggle_bold(self):
     self.apply_font_style("bold")

    def toggle_italic(self):
     self.apply_font_style("italic")

    def toggle_underline(self):
     self.apply_font_style("underline")

    def apply_font_style(self, style):
     row, col = self.current_cell
     entry = self.entries[row][col]
     font = tkFont.Font(font=entry["font"])
     if style == "bold":
         font["weight"] = "bold" if font["weight"] != "bold" else "normal"
     elif style == "italic":
         font["slant"] = "italic" if font["slant"] != "italic" else "roman"
     elif style == "underline":
         font["underline"] = 1 if not font["underline"] else 0
     entry.configure(font=font)
