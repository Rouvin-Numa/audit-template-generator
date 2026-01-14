import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import zipfile
import csv
import io
import os
from collections import defaultdict


class ZipCSVReaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Audit Template Generator")
        self.root.geometry("1200x800")

        # Modern color scheme
        self.bg_color = "#f5f6fa"
        self.primary_color = "#4834df"
        self.secondary_color = "#686de0"
        self.accent_color = "#30336b"
        self.success_color = "#26de81"
        self.text_color = "#2f3542"

        # Configure root background
        self.root.configure(bg=self.bg_color)

        # Configure modern ttk style
        style = ttk.Style()
        style.theme_use('clam')

        # Configure Frame style
        style.configure("Modern.TFrame", background=self.bg_color)

        # Configure Button style
        style.configure("Modern.TButton",
                       padding=10,
                       relief="flat",
                       background=self.primary_color,
                       foreground="white",
                       borderwidth=0,
                       focuscolor="none")
        style.map("Modern.TButton",
                 background=[("active", self.secondary_color),
                            ("pressed", self.accent_color)])

        # Configure Label style
        style.configure("Modern.TLabel",
                       background=self.bg_color,
                       foreground=self.text_color,
                       font=("Segoe UI", 10))

        style.configure("Title.TLabel",
                       background=self.bg_color,
                       foreground=self.accent_color,
                       font=("Segoe UI", 14, "bold"))

        # Configure Notebook style
        style.configure("Modern.TNotebook",
                       background=self.bg_color,
                       borderwidth=0)
        style.configure("Modern.TNotebook.Tab",
                       padding=[20, 10],
                       font=("Segoe UI", 10, "bold"),
                       background="#dfe4ea",
                       foreground=self.text_color)
        style.map("Modern.TNotebook.Tab",
                 background=[("selected", "white")],
                 foreground=[("selected", self.primary_color)])

        # Create main frame
        main_frame = ttk.Frame(root, padding="20", style="Modern.TFrame")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weight
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

        # Drop zone frame with modern styling
        self.drop_frame = tk.Frame(main_frame,
                                   relief=tk.FLAT,
                                   borderwidth=0,
                                   bg="white",
                                   highlightbackground="#dfe4ea",
                                   highlightthickness=2)
        self.drop_frame.grid(row=0, column=0, pady=15, sticky=(tk.W, tk.E))

        # Label and button container
        button_container = tk.Frame(self.drop_frame, bg="white")
        button_container.pack(expand=True, pady=20, padx=30)

        # Instructions
        instructions_frame = tk.Frame(button_container, bg="white")
        instructions_frame.pack(pady=(0, 20))

        instructions_title = tk.Label(
            instructions_frame,
            text="Steps to Use",
            font=("Segoe UI", 12, "bold"),
            bg="white",
            fg=self.accent_color
        )
        instructions_title.pack(anchor=tk.W)

        instructions_text = tk.Label(
            instructions_frame,
            text=(
                "1) Navigate to the following dashboard:\n"
                "   https://numberai.cloud.looker.com/dashboards/602?Rooftop%20Name=&Inbox%20ID=&Calls%20Created%20Date=14%20day&Call%20Count=%3C5&Rooftop%20ID=\n\n"
                "2) Enter the appropriate Rooftop IDs and Inbox IDs into the dashboard filters.\n\n"
                "3) Download the dashboard data as a CSV file.\n\n"
                "4) Upload the CSV using the button below."
            ),
            font=("Segoe UI", 9),
            bg="white",
            fg=self.text_color,
            justify=tk.LEFT,
            wraplength=700
        )
        instructions_text.pack(anchor=tk.W, pady=(5, 0))

        self.drop_label = tk.Label(
            button_container,
            text="📁 Select ZIP or CSV file(s) to view contents",
            font=("Segoe UI", 13),
            bg="white",
            fg=self.text_color
        )
        self.drop_label.pack(pady=12)

        # Browse button with modern style
        self.browse_button = tk.Button(
            button_container,
            text="Browse for ZIP or CSV file",
            command=self.browse_file,
            font=("Segoe UI", 11, "bold"),
            bg=self.primary_color,
            fg="white",
            activebackground=self.secondary_color,
            activeforeground="white",
            relief=tk.FLAT,
            borderwidth=0,
            padx=25,
            pady=12,
            cursor="hand2"
        )
        self.browse_button.pack(pady=8)

        # Current file label
        self.current_file_label = tk.Label(
            button_container,
            text="",
            font=("Segoe UI", 9),
            bg="white",
            fg="#7f8fa6"
        )
        self.current_file_label.pack(pady=8)

        # Main notebook for templates only
        self.notebook = ttk.Notebook(main_frame, style="Modern.TNotebook")
        self.notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)

        # Collapsible CSV Data section
        self.csv_section_frame = tk.Frame(main_frame, bg=self.bg_color)
        self.csv_section_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 5))

        self.csv_expanded = tk.BooleanVar(value=False)

        self.csv_toggle_btn = tk.Button(
            self.csv_section_frame,
            text="▶ View Raw CSV Data",
            command=self.toggle_csv_section,
            font=("Segoe UI", 9),
            bg="#dfe4ea",
            fg=self.text_color,
            activebackground="#c8d6e5",
            activeforeground=self.text_color,
            relief=tk.FLAT,
            borderwidth=0,
            padx=10,
            pady=5,
            cursor="hand2",
            anchor=tk.W
        )
        self.csv_toggle_btn.pack(fill=tk.X)

        # CSV notebook (hidden by default)
        self.csv_notebook_frame = tk.Frame(main_frame, bg=self.bg_color)
        self.csv_notebook = ttk.Notebook(self.csv_notebook_frame, style="Modern.TNotebook")
        self.csv_notebook.pack(fill=tk.BOTH, expand=True)

        # Status bar with modern style
        status_frame = tk.Frame(main_frame, bg=self.bg_color)
        status_frame.grid(row=4, column=0, sticky=(tk.W, tk.E))

        self.status_label = tk.Label(
            status_frame,
            text="Ready",
            font=("Segoe UI", 9),
            bg="#ecf0f1",
            fg=self.text_color,
            anchor=tk.W,
            padx=10,
            pady=8,
            relief=tk.FLAT
        )
        self.status_label.pack(fill=tk.X)

        # Enable drag and drop using Windows-specific method
        self.setup_drag_drop()

    def toggle_csv_section(self):
        """Toggle the CSV data section visibility"""
        if self.csv_expanded.get():
            # Collapse
            self.csv_notebook_frame.grid_forget()
            self.csv_toggle_btn.config(text="▶ View Raw CSV Data")
            self.csv_expanded.set(False)
        else:
            # Expand
            self.csv_notebook_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
            self.csv_notebook_frame.configure(height=200)
            self.csv_toggle_btn.config(text="▼ Hide Raw CSV Data")
            self.csv_expanded.set(True)

    def setup_drag_drop(self):
        """Setup drag and drop functionality"""
        try:
            # Windows drag-and-drop support
            self.drop_frame.drop_target_register('DND_Files')
            self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
        except:
            # If drag-and-drop fails, just use the button
            pass

    def browse_file(self, event=None):
        """Open file browser dialog"""
        filename = filedialog.askopenfilename(
            title="Select ZIP or CSV file",
            filetypes=[("ZIP and CSV files", "*.zip *.csv"), ("ZIP files", "*.zip"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            if filename.lower().endswith('.zip'):
                self.process_zip_file(filename)
            elif filename.lower().endswith('.csv'):
                self.process_csv_files([filename])
            else:
                self.status_label.config(text="Please select a ZIP or CSV file")

    def on_drop(self, event):
        """Handle file drop event"""
        try:
            # Get the dropped file path
            file_path = event.data
            # Remove curly braces if present (Windows behavior)
            file_path = file_path.strip('{}')

            if file_path.lower().endswith('.zip'):
                self.process_zip_file(file_path)
            elif file_path.lower().endswith('.csv'):
                self.process_csv_files([file_path])
            else:
                self.status_label.config(text="Please drop a ZIP or CSV file")
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}")

    def process_zip_file(self, zip_path):
        """Process the dropped/selected zip file"""
        try:
            # Clear existing tabs from both notebooks
            for tab in self.notebook.tabs():
                self.notebook.forget(tab)
            for tab in self.csv_notebook.tabs():
                self.csv_notebook.forget(tab)

            self.status_label.config(text=f"Processing: {os.path.basename(zip_path)}")
            self.current_file_label.config(text=f"Current file: {os.path.basename(zip_path)}")

            # Store CSV data for template generation
            csv_data = {}

            # Open and read the zip file
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                # Get all CSV files in the zip
                csv_files = [f for f in zip_ref.namelist()
                           if f.lower().endswith('.csv') and not f.startswith('__MACOSX')]

                if not csv_files:
                    self.status_label.config(text="No CSV files found in ZIP archive")
                    return

                # Process each CSV file
                for csv_filename in csv_files:
                    rows = self.display_csv(zip_ref, csv_filename)
                    csv_data[os.path.basename(csv_filename).lower()] = rows

                # Generate templates if we have the required files
                self.generate_templates(csv_data)

                self.status_label.config(
                    text=f"✓ Loaded {len(csv_files)} CSV file(s) from {os.path.basename(zip_path)}"
                )

        except zipfile.BadZipFile:
            self.status_label.config(text="Error: Invalid ZIP file")
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}")

    def process_csv_files(self, csv_paths):
        """Process standalone CSV files (not in a ZIP)"""
        try:
            # Clear existing tabs from both notebooks
            for tab in self.notebook.tabs():
                self.notebook.forget(tab)
            for tab in self.csv_notebook.tabs():
                self.csv_notebook.forget(tab)

            self.status_label.config(text=f"Processing CSV file(s)...")

            # Store CSV data for template generation
            csv_data = {}

            for csv_path in csv_paths:
                try:
                    with open(csv_path, 'r', encoding='utf-8') as f:
                        csv_reader = csv.reader(f)
                        rows = list(csv_reader)

                        # Display the CSV
                        self.display_csv_from_rows(rows, os.path.basename(csv_path))
                        csv_data[os.path.basename(csv_path).lower()] = rows

                except Exception as e:
                    self.status_label.config(text=f"Error reading {os.path.basename(csv_path)}: {str(e)}")
                    continue

            # Generate templates if we have the required files
            self.generate_templates(csv_data)

            self.status_label.config(text=f"✓ Loaded {len(csv_data)} CSV file(s)")

        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}")

    def display_csv_from_rows(self, rows, filename):
        """Display CSV rows in a new tab"""
        try:
            # Print to terminal
            self.print_csv_to_terminal(filename, rows)

            # Create a frame for this CSV in the CSV notebook (not main notebook)
            frame = ttk.Frame(self.csv_notebook)
            self.csv_notebook.add(frame, text=os.path.basename(filename))

            # Search bar frame
            search_frame = tk.Frame(frame, bg="#f8f9fa")
            search_frame.pack(fill=tk.X, padx=5, pady=5)

            search_label = tk.Label(
                search_frame,
                text="Search:",
                font=("Segoe UI", 9),
                bg="#f8f9fa",
                fg=self.text_color
            )
            search_label.pack(side=tk.LEFT, padx=(5, 5))

            search_var = tk.StringVar()
            search_entry = tk.Entry(
                search_frame,
                textvariable=search_var,
                font=("Segoe UI", 10),
                relief=tk.FLAT,
                bg="white",
                fg=self.text_color,
                width=40
            )
            search_entry.pack(side=tk.LEFT, padx=5, ipady=4)

            search_count_label = tk.Label(
                search_frame,
                text=f"{len(rows)-1} rows",
                font=("Segoe UI", 9),
                bg="#f8f9fa",
                fg="#7f8fa6"
            )
            search_count_label.pack(side=tk.RIGHT, padx=10)

            # Create treeview for tabular display
            tree_frame = ttk.Frame(frame)
            tree_frame.pack(fill=tk.BOTH, expand=True)

            # Add scrollbars
            vsb = ttk.Scrollbar(tree_frame, orient="vertical")
            hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

            tree = ttk.Treeview(
                tree_frame,
                yscrollcommand=vsb.set,
                xscrollcommand=hsb.set
            )

            vsb.config(command=tree.yview)
            hsb.config(command=tree.xview)

            # Grid layout
            tree.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
            vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
            hsb.grid(row=1, column=0, sticky=(tk.E, tk.W))

            tree_frame.columnconfigure(0, weight=1)
            tree_frame.rowconfigure(0, weight=1)

            # Store all items for search filtering
            all_items = []

            if rows:
                # Set up columns
                headers = rows[0]
                tree["columns"] = headers
                tree["show"] = "headings"

                # Configure columns
                for header in headers:
                    tree.heading(header, text=header)
                    tree.column(header, width=150, minwidth=50)

                # Add data rows
                for row in rows[1:]:
                    # Ensure row has same length as headers
                    while len(row) < len(headers):
                        row.append('')
                    item_id = tree.insert("", tk.END, values=row[:len(headers)])
                    all_items.append((item_id, row[:len(headers)]))

                # Search functionality
                def on_search(*args):
                    search_term = search_var.get().lower().strip()
                    visible_count = 0

                    for item_id, values in all_items:
                        # Check if search term is in any cell
                        row_text = ' '.join(str(v).lower() for v in values)
                        if search_term == '' or search_term in row_text:
                            # Show item - reinsert if hidden
                            try:
                                tree.reattach(item_id, '', tk.END)
                                visible_count += 1
                            except:
                                pass
                        else:
                            # Hide item
                            tree.detach(item_id)

                    if search_term:
                        search_count_label.config(text=f"{visible_count} of {len(all_items)} rows")
                    else:
                        search_count_label.config(text=f"{len(all_items)} rows")

                search_var.trace('w', on_search)

                # Add info label
                info_label = ttk.Label(
                    frame,
                    text=f"Rows: {len(rows)-1} | Columns: {len(headers)}"
                )
                info_label.pack(side=tk.BOTTOM, pady=5)
            else:
                empty_label = ttk.Label(frame, text="CSV file is empty")
                empty_label.pack(pady=20)

        except Exception as e:
            error_frame = ttk.Frame(self.csv_notebook)
            self.csv_notebook.add(error_frame, text=os.path.basename(filename))
            error_label = ttk.Label(
                error_frame,
                text=f"Error reading file:\n{str(e)}",
                foreground="red"
            )
            error_label.pack(pady=20)

    def display_csv(self, zip_ref, csv_filename):
        """Display CSV file content in a new tab and return the rows"""
        try:
            # Read CSV content from zip
            with zip_ref.open(csv_filename) as csv_file:
                # Try different encodings
                encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
                content = None

                for encoding in encodings:
                    try:
                        csv_file.seek(0)
                        content = csv_file.read().decode(encoding)
                        break
                    except UnicodeDecodeError:
                        continue

                if content is None:
                    raise ValueError(f"Could not decode {csv_filename} with any standard encoding")

                # Parse CSV
                csv_reader = csv.reader(io.StringIO(content))
                rows = list(csv_reader)

                # Print to terminal
                self.print_csv_to_terminal(csv_filename, rows)

                # Create a frame for this CSV in the CSV notebook (not main notebook)
                frame = ttk.Frame(self.csv_notebook)
                self.csv_notebook.add(frame, text=os.path.basename(csv_filename))

                # Search bar frame
                search_frame = tk.Frame(frame, bg="#f8f9fa")
                search_frame.pack(fill=tk.X, padx=5, pady=5)

                search_label = tk.Label(
                    search_frame,
                    text="Search:",
                    font=("Segoe UI", 9),
                    bg="#f8f9fa",
                    fg=self.text_color
                )
                search_label.pack(side=tk.LEFT, padx=(5, 5))

                search_var = tk.StringVar()
                search_entry = tk.Entry(
                    search_frame,
                    textvariable=search_var,
                    font=("Segoe UI", 10),
                    relief=tk.FLAT,
                    bg="white",
                    fg=self.text_color,
                    width=40
                )
                search_entry.pack(side=tk.LEFT, padx=5, ipady=4)

                search_count_label = tk.Label(
                    search_frame,
                    text=f"{len(rows)-1} rows",
                    font=("Segoe UI", 9),
                    bg="#f8f9fa",
                    fg="#7f8fa6"
                )
                search_count_label.pack(side=tk.RIGHT, padx=10)

                # Create treeview for tabular display
                tree_frame = ttk.Frame(frame)
                tree_frame.pack(fill=tk.BOTH, expand=True)

                # Add scrollbars
                vsb = ttk.Scrollbar(tree_frame, orient="vertical")
                hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

                tree = ttk.Treeview(
                    tree_frame,
                    yscrollcommand=vsb.set,
                    xscrollcommand=hsb.set
                )

                vsb.config(command=tree.yview)
                hsb.config(command=tree.xview)

                # Grid layout
                tree.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
                vsb.grid(row=0, column=1, sticky=(tk.N, tk.S))
                hsb.grid(row=1, column=0, sticky=(tk.E, tk.W))

                tree_frame.columnconfigure(0, weight=1)
                tree_frame.rowconfigure(0, weight=1)

                # Store all items for search filtering
                all_items = []

                if rows:
                    # Set up columns
                    headers = rows[0]
                    tree["columns"] = headers
                    tree["show"] = "headings"

                    # Configure columns
                    for header in headers:
                        tree.heading(header, text=header)
                        tree.column(header, width=150, minwidth=50)

                    # Add data rows
                    for row in rows[1:]:
                        # Ensure row has same length as headers
                        while len(row) < len(headers):
                            row.append('')
                        item_id = tree.insert("", tk.END, values=row[:len(headers)])
                        all_items.append((item_id, row[:len(headers)]))

                    # Search functionality
                    def on_search(*args):
                        search_term = search_var.get().lower().strip()
                        visible_count = 0

                        for item_id, values in all_items:
                            # Check if search term is in any cell
                            row_text = ' '.join(str(v).lower() for v in values)
                            if search_term == '' or search_term in row_text:
                                # Show item - reinsert if hidden
                                try:
                                    tree.reattach(item_id, '', tk.END)
                                    visible_count += 1
                                except:
                                    pass
                            else:
                                # Hide item
                                tree.detach(item_id)

                        if search_term:
                            search_count_label.config(text=f"{visible_count} of {len(all_items)} rows")
                        else:
                            search_count_label.config(text=f"{len(all_items)} rows")

                    search_var.trace('w', on_search)

                    # Add info label
                    info_label = ttk.Label(
                        frame,
                        text=f"Rows: {len(rows)-1} | Columns: {len(headers)}"
                    )
                    info_label.pack(side=tk.BOTTOM, pady=5)
                else:
                    empty_label = ttk.Label(frame, text="CSV file is empty")
                    empty_label.pack(pady=20)

                return rows

        except Exception as e:
            error_frame = ttk.Frame(self.csv_notebook)
            self.csv_notebook.add(error_frame, text=os.path.basename(csv_filename))
            error_label = ttk.Label(
                error_frame,
                text=f"Error reading file:\n{str(e)}",
                foreground="red"
            )
            error_label.pack(pady=20)
            return []

    def format_phone_number(self, phone):
        """Format phone number as (111) 222-3333"""
        # Remove all non-digit characters
        digits = ''.join(filter(str.isdigit, str(phone)))

        # Format based on length
        if len(digits) == 10:
            return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"
        elif len(digits) == 11 and digits[0] == '1':
            # Handle numbers starting with 1
            return f"({digits[1:4]}) {digits[4:7]}-{digits[7:]}"
        else:
            # Return original if format is unexpected
            return phone

    def capitalize_name(self, name):
        """Capitalize name: each word's first character uppercase, rest lowercase"""
        if not name:
            return name
        # Use title() to capitalize each word properly
        return name.title()

    def get_first_name(self, full_name):
        """Extract first name from full name"""
        if not full_name:
            return full_name
        # Split by space and take the first part
        return full_name.strip().split()[0] if full_name.strip() else full_name

    def generate_templates(self, csv_data):
        """Generate email templates based on CSV data"""
        # Check if we have the required files
        lines_file = None
        rooftop_file = None
        desk_phones_file = None

        for filename, rows in csv_data.items():
            if 'lines_with_low' in filename and 'call_volume' in filename:
                lines_file = rows
            elif 'rooftop_information' in filename or 'rooftop_informatio' in filename:
                rooftop_file = rows
            elif 'desk_phones' in filename:
                desk_phones_file = rows

        if not lines_file or not rooftop_file:
            print("\n" + "!"*80)
            print("WARNING: Could not find required CSV files for template generation")
            print(f"Available files: {list(csv_data.keys())}")
            print("Required: 'rooftop_information.csv' and 'lines_with_low_*_call_volume.csv'")
            print("!"*80 + "\n")
            return

        # Build desk phone lookup (display name -> phone number) if desk_phones file exists
        desk_phone_lookup = {}
        if desk_phones_file and len(desk_phones_file) > 1:
            desk_headers = desk_phones_file[0]
            desk_data = desk_phones_file[1:]

            # Find column indices for desk phones file
            desk_display_name_idx = None
            desk_phone_number_idx = None

            for idx, header in enumerate(desk_headers):
                header_lower = header.lower().strip()
                # Check for display name column
                if 'display name' in header_lower or 'display_name' in header_lower:
                    desk_display_name_idx = idx
                # Check for phone number column - handle various naming conventions
                if 'phone number' in header_lower or 'phone_number' in header_lower or 'phone numbers' in header_lower:
                    desk_phone_number_idx = idx

            if desk_display_name_idx is not None and desk_phone_number_idx is not None:
                for row in desk_data:
                    if len(row) > max(desk_display_name_idx, desk_phone_number_idx):
                        display_name = row[desk_display_name_idx].strip().lower()
                        phone_number = row[desk_phone_number_idx].strip()
                        if display_name and phone_number:
                            desk_phone_lookup[display_name] = phone_number

        try:
            # Parse lines_with_low_call_volume.csv
            lines_headers = lines_file[0]
            lines_data = lines_file[1:]

            # Find column indices (case-insensitive)
            def find_col_idx(headers, possible_names):
                for name in possible_names:
                    for idx, header in enumerate(headers):
                        if name.lower() in header.lower():
                            return idx
                return None

            # Find column indices with exact matching for "Name" to avoid confusion with "Display Name"
            def find_exact_col_idx(headers, target_name):
                """Find column index with exact name match (case-insensitive)"""
                for idx, header in enumerate(headers):
                    if header.strip().lower() == target_name.lower():
                        return idx
                return None

            display_name_idx = find_col_idx(lines_headers, ['display name', 'display_name'])
            phone_number_idx = find_col_idx(lines_headers, ['phone number', 'phone_number', 'number'])
            rooftop_name_idx = find_col_idx(lines_headers, ['rooftop name', 'rooftop_name', 'rooftop'])
            inbox_name_idx = find_col_idx(lines_headers, ['inbox name', 'inbox_name', 'inbox'])
            owner_type_idx = find_col_idx(lines_headers, ['owner type', 'owner_type', 'ownertype'])
            # Use exact match for "Name" to avoid matching "Display Name"
            name_idx = find_exact_col_idx(lines_headers, 'name')

            if None in [display_name_idx, phone_number_idx, rooftop_name_idx, inbox_name_idx]:
                print("\nERROR: Could not find all required columns in lines_with_low_call_volume.csv")
                print(f"Found headers: {lines_headers}")
                return

            # Group data by rooftop
            rooftops = defaultdict(lambda: {'inbox_name': '', 'lines': []})

            for row in lines_data:
                # Find the maximum index we need to check
                max_idx = max(display_name_idx, phone_number_idx, rooftop_name_idx, inbox_name_idx)
                if owner_type_idx is not None:
                    max_idx = max(max_idx, owner_type_idx)
                if name_idx is not None:
                    max_idx = max(max_idx, name_idx)

                if len(row) > max_idx:
                    rooftop = row[rooftop_name_idx].strip()
                    if rooftop:  # Skip empty rooftops
                        rooftops[rooftop]['inbox_name'] = row[inbox_name_idx].strip()

                        # Get display name with fallback logic
                        display_name = row[display_name_idx].strip() if display_name_idx < len(row) else ''

                        # If display name is empty/null, check owner type
                        if not display_name:
                            owner_type = row[owner_type_idx].strip().upper() if owner_type_idx is not None and owner_type_idx < len(row) else ''
                            name_value = row[name_idx].strip() if name_idx is not None and name_idx < len(row) else ''

                            if owner_type == 'USER':
                                display_name = f"Unassigned line - [{self.capitalize_name(name_value)}]"
                            elif owner_type == 'DEPARTMENT':
                                display_name = f"Unassigned line - [{self.capitalize_name(name_value)}]"
                            else:
                                display_name = self.capitalize_name(name_value) if name_value else 'Unknown'
                        else:
                            # Capitalize the display name
                            display_name = self.capitalize_name(display_name)

                        # Get raw name value for desk phone table
                        raw_name = row[name_idx].strip() if name_idx is not None and name_idx < len(row) else ''
                        raw_display_name = row[display_name_idx].strip() if display_name_idx < len(row) else ''

                        # Look up desk phone number by matching display name (case-insensitive)
                        desk_phone = desk_phone_lookup.get(raw_display_name.lower(), '')

                        rooftops[rooftop]['lines'].append({
                            'display_name': display_name,
                            'phone_number': self.format_phone_number(row[phone_number_idx].strip()),
                            'raw_display_name': raw_display_name,
                            'raw_name': raw_name,
                            'desk_phone': desk_phone
                        })

            # Generate templates
            print("\n" + "="*80)
            print("GENERATED EMAIL TEMPLATES")
            print("="*80 + "\n")

            template_text = ""
            for rooftop_name, data in rooftops.items():
                inbox_name = data['inbox_name']
                lines = data['lines']

                # Separate lines into regular users and department/unassigned lines
                regular_lines = []
                department_unassigned_lines = []

                for line in lines:
                    display_name = line['display_name']
                    if display_name.startswith('Unassigned line'):
                        department_unassigned_lines.append(line)
                    else:
                        regular_lines.append(line)

                template = f"""Good morning [Dealership POC],

We've recently noticed a drop in call volume on your account ({rooftop_name} – {inbox_name})

To ensure you're getting the most out of your Numa subscription, please confirm that missed calls on the following users' direct lines are forwarding to their respective Numa IT forwarding lines after 4 rings (approximately 20 seconds), rather than going to local voicemail (including DND, busy, and after-hours scenarios):
"""
                # Add regular lines first
                for line in regular_lines:
                    template += f"• {line['display_name']} – Numa IT forwarding number: {line['phone_number']}\n"

                # Add department/unassigned lines at the bottom
                for line in department_unassigned_lines:
                    template += f"• {line['display_name']} – Numa IT forwarding number: {line['phone_number']}\n"

                template += """If you have any questions, feel free to email us at support@numa.com.
"""

                template_text += template + "\n" + "="*80 + "\n"
                print(template)
                print("="*80 + "\n")

            # Generate CSM templates first
            self.generate_csm_templates(rooftop_file, rooftops)

            # Create a tab with dealership templates
            self.create_template_tab(template_text, rooftops, "Dealership Templates")

        except Exception as e:
            print(f"\nERROR generating templates: {str(e)}")
            import traceback
            traceback.print_exc()

    def generate_csm_templates(self, rooftop_file, rooftops):
        """Generate CSM templates grouped by CSM Owner"""
        try:
            # Parse rooftop_information.csv
            rooftop_headers = rooftop_file[0]
            rooftop_data = rooftop_file[1:]

            # Find column indices (case-insensitive)
            def find_col_idx(headers, possible_names):
                for name in possible_names:
                    for idx, header in enumerate(headers):
                        if name.lower() in header.lower():
                            return idx
                return None

            rooftop_name_col_idx = find_col_idx(rooftop_headers, ['rooftop name', 'rooftop_name', 'rooftop'])
            csm_owner_idx = find_col_idx(rooftop_headers, ['csm owner', 'csm_owner', 'csmowner'])

            if rooftop_name_col_idx is None or csm_owner_idx is None:
                print("\nWARNING: Could not find CSM Owner or Rooftop Name in rooftop_information.csv")
                print(f"Found headers: {rooftop_headers}")
                return

            # Build a mapping of rooftop name to CSM owner and track all rooftops per CSM
            rooftop_to_csm = {}
            all_rooftops_by_csm = defaultdict(list)  # Track all rooftops per CSM from rooftop_information.csv
            for row in rooftop_data:
                if len(row) > max(rooftop_name_col_idx, csm_owner_idx):
                    rooftop_name = row[rooftop_name_col_idx].strip()
                    csm_owner = row[csm_owner_idx].strip()
                    if rooftop_name and csm_owner:
                        rooftop_to_csm[rooftop_name] = csm_owner
                        all_rooftops_by_csm[csm_owner].append(rooftop_name)

            # Group rooftops by CSM - only rooftops that have lines data
            csm_rooftops = defaultdict(lambda: {'included': [], 'skipped': []})
            for rooftop_name, data in rooftops.items():
                csm_owner = rooftop_to_csm.get(rooftop_name, 'Unknown CSM')
                inbox_name = data['inbox_name']
                csm_rooftops[csm_owner]['included'].append({
                    'rooftop_name': rooftop_name,
                    'inbox_name': inbox_name
                })

            # Find skipped rooftops for each CSM (in rooftop_information but not in lines file)
            for csm_owner, rooftop_names in all_rooftops_by_csm.items():
                if csm_owner not in csm_rooftops:
                    csm_rooftops[csm_owner] = {'included': [], 'skipped': []}
                for rooftop_name in rooftop_names:
                    if rooftop_name not in rooftops:
                        csm_rooftops[csm_owner]['skipped'].append(rooftop_name)

            # Generate CSM templates
            print("\n" + "="*80)
            print("GENERATED CSM TEMPLATES")
            print("="*80 + "\n")

            csm_template_text = ""
            for csm_owner, data in csm_rooftops.items():
                rooftop_list = data['included']
                skipped_list = data['skipped']

                # Skip CSMs that have no included rooftops
                if len(rooftop_list) == 0:
                    continue

                template = f"Hi {self.get_first_name(csm_owner)},\n\n"
                template += "We've identified the following dealerships with low call volume over the past two weeks. To help us follow up, could you please provide a point of contact for each location so we can reach out to them?\n"

                for rooftop_info in rooftop_list:
                    template += f"• {rooftop_info['rooftop_name']} – {rooftop_info['inbox_name']}\n"

                template += "\n"

                csm_template_text += template + "\n" + "="*80 + "\n"
                print(template)
                if skipped_list:
                    print(f"  [Skipped {len(skipped_list)} rooftop(s) - not in lines file: {', '.join(skipped_list)}]")
                print("="*80 + "\n")

            # Create a tab with CSM templates
            self.create_csm_template_tab(csm_template_text, csm_rooftops)

        except Exception as e:
            print(f"\nERROR generating CSM templates: {str(e)}")
            import traceback
            traceback.print_exc()

    def create_template_tab(self, template_text, rooftops, tab_name="Dealership Templates"):
        """Create a new tab to display generated templates"""
        frame = tk.Frame(self.notebook, bg=self.bg_color)
        self.notebook.add(frame, text=tab_name)

        # Add a canvas with scrollbar for multiple template cards
        canvas = tk.Canvas(frame, bg=self.bg_color, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.bg_color)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Create individual template cards for each rooftop
        for idx, (rooftop_name, data) in enumerate(rooftops.items(), 1):
            inbox_name = data['inbox_name']
            lines = data['lines']

            # Separate lines into regular users and department/unassigned lines
            regular_lines = []
            department_unassigned_lines = []

            for line in lines:
                display_name = line['display_name']
                if display_name.startswith('Unassigned line'):
                    department_unassigned_lines.append(line)
                else:
                    regular_lines.append(line)

            # Generate clean template for this rooftop
            template = f"Good morning [Dealership POC],\n\n"
            template += f"We've recently noticed a drop in call volume on your account ({rooftop_name} – {inbox_name})\n\n"
            template += "To ensure you're getting the most out of your Numa subscription, please confirm that missed calls on the following users' direct lines are forwarding to their respective Numa IT forwarding lines after 4 rings (approximately 20 seconds), rather than going to local voicemail (including DND, busy, and after-hours scenarios):\n"

            # Add regular lines first
            for line in regular_lines:
                template += f"• {line['display_name']} – Numa IT forwarding number: {line['phone_number']}\n"

            # Add department/unassigned lines at the bottom
            for line in department_unassigned_lines:
                template += f"• {line['display_name']} – Numa IT forwarding number: {line['phone_number']}\n"

            template += "\nAdditionally, when you have a moment, kindly update the following roster with the latest desk phone numbers for your staff members\nRoster link [insert roster link here]\n"
            template += "\nIf you have any questions, feel free to email us at support@numa.com."

            # Create card frame with modern styling
            card_frame = tk.Frame(
                scrollable_frame,
                bg="white",
                highlightbackground="#dfe4ea",
                highlightthickness=1,
                relief=tk.FLAT
            )
            card_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

            # Card header
            header_frame = tk.Frame(card_frame, bg="white")
            header_frame.pack(fill=tk.X, padx=15, pady=(15, 10))

            header_label = tk.Label(
                header_frame,
                text=f"Template {idx}: {rooftop_name}",
                font=("Segoe UI", 11, "bold"),
                bg="white",
                fg=self.accent_color,
                anchor=tk.W
            )
            header_label.pack(side=tk.LEFT)

            # Subject line display
            subject_line = f"{rooftop_name} - {inbox_name}: Phoneline forwarding"

            subject_frame = tk.Frame(card_frame, bg="white")
            subject_frame.pack(fill=tk.X, padx=15, pady=(0, 5))

            subject_label = tk.Label(
                subject_frame,
                text="Subject: ",
                font=("Segoe UI", 10, "bold"),
                bg="white",
                fg=self.text_color
            )
            subject_label.pack(side=tk.LEFT)

            subject_text = tk.Entry(
                subject_frame,
                font=("Segoe UI", 10),
                relief=tk.FLAT,
                bg="#f8f9fa",
                fg=self.text_color,
                width=60
            )
            subject_text.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
            subject_text.insert(0, subject_line)

            # Content frame - side by side layout for template and desk phones table
            content_frame = tk.Frame(card_frame, bg="white")
            content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 10))

            # Left side - Template text widget (editable)
            left_frame = tk.Frame(content_frame, bg="white")
            left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            text_widget = tk.Text(
                left_frame,
                wrap=tk.WORD,
                height=14,
                font=("Segoe UI", 10),
                relief=tk.FLAT,
                borderwidth=0,
                padx=12,
                pady=8,
                bg="#f8f9fa",
                fg=self.text_color
            )
            text_widget.pack(fill=tk.BOTH, expand=True)
            text_widget.insert(1.0, template)
            # Template is now editable - no state=DISABLED

            # Right side - Possible desk phones table
            right_frame = tk.Frame(content_frame, bg="white")
            right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(15, 0))

            desk_phones_label = tk.Label(
                right_frame,
                text="Possible Desk Phones",
                font=("Segoe UI", 10, "bold"),
                bg="white",
                fg=self.accent_color
            )
            desk_phones_label.pack(anchor=tk.W, pady=(0, 5))

            # Create treeview for desk phones table
            desk_tree_frame = tk.Frame(right_frame, bg="white")
            desk_tree_frame.pack(fill=tk.BOTH, expand=True)

            desk_tree = ttk.Treeview(
                desk_tree_frame,
                columns=("display_name", "name"),
                show="headings",
                height=10
            )
            desk_tree.heading("display_name", text="Display Name")
            desk_tree.heading("name", text="Name")
            desk_tree.column("display_name", width=150, minwidth=100)
            desk_tree.column("name", width=250, minwidth=180)

            # Add scrollbar for desk phones table
            desk_scrollbar = ttk.Scrollbar(desk_tree_frame, orient="vertical", command=desk_tree.yview)
            desk_tree.configure(yscrollcommand=desk_scrollbar.set)

            desk_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            desk_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            # Populate desk phones table with all lines (regular + unassigned)
            for line in regular_lines + department_unassigned_lines:
                raw_display = line.get('raw_display_name', '')
                raw_name = line.get('raw_name', '')
                desk_tree.insert("", tk.END, values=(raw_display, raw_name))

            # Button frame
            button_frame = tk.Frame(card_frame, bg="white")
            button_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

            # Copy button for this template - reads from text widget
            def make_copy_func(tw, subj_entry, btn, orig_text, rname):
                def copy_func():
                    current_text = tw.get("1.0", tk.END).strip()
                    current_subject = subj_entry.get().strip()
                    full_copy = f"Subject: {current_subject}\n\n{current_text}"
                    self.root.clipboard_clear()
                    self.root.clipboard_append(full_copy)
                    self.status_label.config(text=f"✓ Copied template for {rname} to clipboard", bg=self.success_color, fg="white")
                    btn.config(text="✓ Copied!", bg=self.success_color)
                    # Reset button text after 2 seconds
                    self.root.after(2000, lambda: (btn.config(text=orig_text, bg=self.primary_color),
                                                   self.status_label.config(bg="#ecf0f1", fg=self.text_color)))
                return copy_func

            button_text = f"📋 Copy Template {idx}"
            copy_btn = tk.Button(
                button_frame,
                text=button_text,
                font=("Segoe UI", 9, "bold"),
                bg=self.primary_color,
                fg="white",
                activebackground=self.secondary_color,
                activeforeground="white",
                relief=tk.FLAT,
                borderwidth=0,
                padx=15,
                pady=8,
                cursor="hand2"
            )
            copy_btn.config(command=make_copy_func(text_widget, subject_text, copy_btn, button_text, rooftop_name))
            copy_btn.pack(side=tk.LEFT, padx=(0, 5))

            # Info label
            info_label = tk.Label(
                button_frame,
                text=f"{len(lines)} phone line(s)",
                font=("Segoe UI", 9),
                bg="white",
                fg="#7f8fa6"
            )
            info_label.pack(side=tk.RIGHT, padx=5)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Add summary at bottom
        summary_frame = tk.Frame(frame, bg=self.bg_color)
        summary_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=15)

        summary_label = tk.Label(
            summary_frame,
            text=f"✓ Generated {len(rooftops)} template(s) - One per rooftop",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color,
            fg=self.accent_color
        )
        summary_label.pack(side=tk.LEFT)

        # Copy all button
        def copy_all():
            self.root.clipboard_clear()
            self.root.clipboard_append(template_text)
            self.status_label.config(text=f"✓ Copied all {len(rooftops)} template(s) to clipboard", bg=self.success_color, fg="white")
            copy_all_btn.config(text="✓ Copied!", bg=self.success_color)
            # Reset button text after 2 seconds
            self.root.after(2000, lambda: (copy_all_btn.config(text="📋 Copy All Templates", bg=self.secondary_color),
                                          self.status_label.config(bg="#ecf0f1", fg=self.text_color)))

        copy_all_btn = tk.Button(
            summary_frame,
            text="📋 Copy All Templates",
            command=copy_all,
            font=("Segoe UI", 9, "bold"),
            bg=self.secondary_color,
            fg="white",
            activebackground=self.accent_color,
            activeforeground="white",
            relief=tk.FLAT,
            borderwidth=0,
            padx=20,
            pady=10,
            cursor="hand2"
        )
        copy_all_btn.pack(side=tk.RIGHT)

    def create_csm_template_tab(self, template_text, csm_rooftops):
        """Create a new tab to display CSM templates"""
        frame = tk.Frame(self.notebook, bg=self.bg_color)
        self.notebook.add(frame, text="CSM Templates")

        # Add a canvas with scrollbar for multiple template cards
        canvas = tk.Canvas(frame, bg=self.bg_color, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.bg_color)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Create individual template cards for each CSM
        template_count = 0
        for csm_owner, data in csm_rooftops.items():
            rooftop_list = data['included']
            skipped_list = data['skipped']

            # Skip CSMs that have no included rooftops
            if len(rooftop_list) == 0:
                continue

            template_count += 1
            idx = template_count

            # Generate clean template for this CSM
            template = f"Hi {self.get_first_name(csm_owner)},\n\n"
            template += "We've identified the following dealerships with low call volume over the past two weeks. To help us follow up, could you please provide a point of contact for each location so we can reach out directly?\n"

            for rooftop_info in rooftop_list:
                template += f"• {rooftop_info['rooftop_name']} – {rooftop_info['inbox_name']}\n"

            template += "\nPlease let us know whether the lines are intentionally not forwarding, or if you'd prefer that we avoid contacting any of the dealerships mentioned above."

            # Create card frame with modern styling
            card_frame = tk.Frame(
                scrollable_frame,
                bg="white",
                highlightbackground="#dfe4ea",
                highlightthickness=1,
                relief=tk.FLAT
            )
            card_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

            # Card header
            header_frame = tk.Frame(card_frame, bg="white")
            header_frame.pack(fill=tk.X, padx=15, pady=(15, 10))

            header_label = tk.Label(
                header_frame,
                text=f"Template {idx}: {csm_owner}",
                font=("Segoe UI", 11, "bold"),
                bg="white",
                fg=self.accent_color,
                anchor=tk.W
            )
            header_label.pack(side=tk.LEFT)

            # Text widget for this template (editable)
            text_widget = tk.Text(
                card_frame,
                wrap=tk.WORD,
                height=10,
                font=("Segoe UI", 10),
                relief=tk.FLAT,
                borderwidth=0,
                padx=12,
                pady=8,
                bg="#f8f9fa",
                fg=self.text_color
            )
            text_widget.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 10))
            text_widget.insert(1.0, template)
            # Template is now editable - no state=DISABLED

            # Skipped rooftops note (if any)
            if skipped_list:
                skipped_frame = tk.Frame(card_frame, bg="#fff8e6", highlightbackground="#f5d77a", highlightthickness=1)
                skipped_frame.pack(fill=tk.X, padx=20, pady=(0, 10))

                skipped_title = tk.Label(
                    skipped_frame,
                    text=f"Skipped {len(skipped_list)} rooftop(s) - not in lines_with_low_call_volume.csv:",
                    font=("Segoe UI", 9, "bold"),
                    bg="#fff8e6",
                    fg="#856404",
                    anchor=tk.W
                )
                skipped_title.pack(fill=tk.X, padx=10, pady=(8, 2))

                for skipped_name in skipped_list:
                    skipped_item = tk.Label(
                        skipped_frame,
                        text=f"  • {skipped_name}",
                        font=("Segoe UI", 9),
                        bg="#fff8e6",
                        fg="#856404",
                        anchor=tk.W
                    )
                    skipped_item.pack(fill=tk.X, padx=10, pady=1)

                # Add bottom padding
                tk.Frame(skipped_frame, bg="#fff8e6", height=8).pack(fill=tk.X)

            # Button frame
            button_frame = tk.Frame(card_frame, bg="white")
            button_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

            # Copy button for this template - reads from text widget
            def make_copy_func(tw, btn, orig_text, csm):
                def copy_func():
                    current_text = tw.get("1.0", tk.END).strip()
                    self.root.clipboard_clear()
                    self.root.clipboard_append(current_text)
                    self.status_label.config(text=f"✓ Copied CSM template for {csm} to clipboard", bg=self.success_color, fg="white")
                    btn.config(text="✓ Copied!", bg=self.success_color)
                    # Reset button text after 2 seconds
                    self.root.after(2000, lambda: (btn.config(text=orig_text, bg=self.primary_color),
                                                   self.status_label.config(bg="#ecf0f1", fg=self.text_color)))
                return copy_func

            button_text = f"📋 Copy Template {idx}"
            copy_btn = tk.Button(
                button_frame,
                text=button_text,
                font=("Segoe UI", 9, "bold"),
                bg=self.primary_color,
                fg="white",
                activebackground=self.secondary_color,
                activeforeground="white",
                relief=tk.FLAT,
                borderwidth=0,
                padx=15,
                pady=8,
                cursor="hand2"
            )
            copy_btn.config(command=make_copy_func(text_widget, copy_btn, button_text, csm_owner))
            copy_btn.pack(side=tk.LEFT, padx=(0, 5))

            # Info label
            info_label = tk.Label(
                button_frame,
                text=f"{len(rooftop_list)} rooftop(s)",
                font=("Segoe UI", 9),
                bg="white",
                fg="#7f8fa6"
            )
            info_label.pack(side=tk.RIGHT, padx=5)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Show skipped CSMs (CSMs with rooftops in rooftop_information but none in lines file)
        skipped_csms = [(csm, data['skipped']) for csm, data in csm_rooftops.items()
                        if len(data['included']) == 0 and len(data['skipped']) > 0]
        if skipped_csms:
            skipped_csms_frame = tk.Frame(frame, bg="#fff8e6", highlightbackground="#f5d77a", highlightthickness=1)
            skipped_csms_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=(0, 10))

            skipped_csms_title = tk.Label(
                skipped_csms_frame,
                text="No templates generated for the following CSM(s) - all their rooftops are missing from lines_with_low_call_volume.csv:",
                font=("Segoe UI", 9, "bold"),
                bg="#fff8e6",
                fg="#856404",
                anchor=tk.W,
                wraplength=800
            )
            skipped_csms_title.pack(fill=tk.X, padx=10, pady=(8, 5))

            for csm_owner, skipped_list in skipped_csms:
                csm_item = tk.Label(
                    skipped_csms_frame,
                    text=f"  {csm_owner}: {', '.join(skipped_list)}",
                    font=("Segoe UI", 9),
                    bg="#fff8e6",
                    fg="#856404",
                    anchor=tk.W
                )
                csm_item.pack(fill=tk.X, padx=10, pady=2)

            tk.Frame(skipped_csms_frame, bg="#fff8e6", height=8).pack(fill=tk.X)

        # Add summary at bottom
        summary_frame = tk.Frame(frame, bg=self.bg_color)
        summary_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=15)

        summary_label = tk.Label(
            summary_frame,
            text=f"✓ Generated {template_count} CSM template(s) - One per CSM",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color,
            fg=self.accent_color
        )
        summary_label.pack(side=tk.LEFT)

        # Copy all button
        def copy_all():
            self.root.clipboard_clear()
            self.root.clipboard_append(template_text)
            self.status_label.config(text=f"✓ Copied all {template_count} CSM template(s) to clipboard", bg=self.success_color, fg="white")
            copy_all_btn.config(text="✓ Copied!", bg=self.success_color)
            # Reset button text after 2 seconds
            self.root.after(2000, lambda: (copy_all_btn.config(text="📋 Copy All Templates", bg=self.secondary_color),
                                          self.status_label.config(bg="#ecf0f1", fg=self.text_color)))

        copy_all_btn = tk.Button(
            summary_frame,
            text="📋 Copy All Templates",
            command=copy_all,
            font=("Segoe UI", 9, "bold"),
            bg=self.secondary_color,
            fg="white",
            activebackground=self.accent_color,
            activeforeground="white",
            relief=tk.FLAT,
            borderwidth=0,
            padx=20,
            pady=10,
            cursor="hand2"
        )
        copy_all_btn.pack(side=tk.RIGHT)

    def print_csv_to_terminal(self, filename, rows):
        """Print CSV content to terminal in a formatted table"""
        print("\n" + "="*80)
        print(f"CSV FILE: {os.path.basename(filename)}")
        print("="*80)

        if not rows:
            print("(Empty file)")
            return

        # Calculate column widths
        col_widths = []
        for col_idx in range(len(rows[0])):
            max_width = max(len(str(row[col_idx])) if col_idx < len(row) else 0
                          for row in rows)
            col_widths.append(min(max_width, 40))  # Cap at 40 chars

        # Print header
        header = rows[0]
        header_line = " | ".join(str(header[i])[:col_widths[i]].ljust(col_widths[i])
                                 for i in range(len(header)))
        print(header_line)
        print("-" * len(header_line))

        # Print data rows
        for row in rows[1:]:
            row_line = " | ".join(
                str(row[i])[:col_widths[i]].ljust(col_widths[i])
                if i < len(row) else " " * col_widths[i]
                for i in range(len(header))
            )
            print(row_line)

        print("\n" + f"Total rows: {len(rows)-1}")
        print("="*80 + "\n")


def main():
    root = tk.Tk()
    app = ZipCSVReaderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
