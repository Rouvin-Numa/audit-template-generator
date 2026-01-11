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
        self.root.geometry("900x700")

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
        button_container.pack(expand=True, pady=30, padx=30)

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

        # Notebook for displaying CSV files
        self.notebook = ttk.Notebook(main_frame, style="Modern.TNotebook")
        self.notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)

        # Status bar with modern style
        status_frame = tk.Frame(main_frame, bg=self.bg_color)
        status_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))

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
            # Clear existing tabs
            for tab in self.notebook.tabs():
                self.notebook.forget(tab)

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
            # Clear existing tabs
            for tab in self.notebook.tabs():
                self.notebook.forget(tab)

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

            # Create a frame for this CSV
            frame = ttk.Frame(self.notebook)
            self.notebook.add(frame, text=os.path.basename(filename))

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
                    tree.insert("", tk.END, values=row[:len(headers)])

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
            error_frame = ttk.Frame(self.notebook)
            self.notebook.add(error_frame, text=os.path.basename(filename))
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

                # Create a frame for this CSV
                frame = ttk.Frame(self.notebook)
                self.notebook.add(frame, text=os.path.basename(csv_filename))

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
                        tree.insert("", tk.END, values=row[:len(headers)])

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
            error_frame = ttk.Frame(self.notebook)
            self.notebook.add(error_frame, text=os.path.basename(csv_filename))
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
        """Capitalize name: first character uppercase, rest lowercase"""
        if not name:
            return name
        return name[0].upper() + name[1:].lower() if len(name) > 0 else name

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

        for filename, rows in csv_data.items():
            if 'lines_with_low' in filename and 'call_volume' in filename:
                lines_file = rows
            elif 'rooftop_information' in filename or 'rooftop_informatio' in filename:
                rooftop_file = rows

        if not lines_file or not rooftop_file:
            print("\n" + "!"*80)
            print("WARNING: Could not find required CSV files for template generation")
            print(f"Available files: {list(csv_data.keys())}")
            print("Required: 'rooftop_information.csv' and 'lines_with_low_*_call_volume.csv'")
            print("!"*80 + "\n")
            return

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
                                display_name = f"Staff line - User unassigned [{self.capitalize_name(name_value)}]"
                            elif owner_type == 'DEPARTMENT':
                                display_name = f"Department line - [{self.capitalize_name(name_value)}]"
                            else:
                                display_name = self.capitalize_name(name_value) if name_value else 'Unknown'
                        else:
                            # Capitalize the display name
                            display_name = self.capitalize_name(display_name)

                        rooftops[rooftop]['lines'].append({
                            'display_name': display_name,
                            'phone_number': self.format_phone_number(row[phone_number_idx].strip())
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
                    if display_name.startswith('Department line') or display_name.startswith('Staff line - User unassigned'):
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

            # Build a mapping of rooftop name to CSM owner
            rooftop_to_csm = {}
            for row in rooftop_data:
                if len(row) > max(rooftop_name_col_idx, csm_owner_idx):
                    rooftop_name = row[rooftop_name_col_idx].strip()
                    csm_owner = row[csm_owner_idx].strip()
                    if rooftop_name and csm_owner:
                        rooftop_to_csm[rooftop_name] = csm_owner

            # Group rooftops by CSM
            csm_rooftops = defaultdict(list)
            for rooftop_name, data in rooftops.items():
                csm_owner = rooftop_to_csm.get(rooftop_name, 'Unknown CSM')
                inbox_name = data['inbox_name']
                csm_rooftops[csm_owner].append({
                    'rooftop_name': rooftop_name,
                    'inbox_name': inbox_name
                })

            # Generate CSM templates
            print("\n" + "="*80)
            print("GENERATED CSM TEMPLATES")
            print("="*80 + "\n")

            csm_template_text = ""
            for csm_owner, rooftop_list in csm_rooftops.items():
                template = f"Hi {self.get_first_name(csm_owner)},\n\n"
                template += "We've identified the following dealerships with low call volume over the past two weeks. To help us follow up, could you please provide a point of contact for each location so we can reach out to them?\n"

                for rooftop_info in rooftop_list:
                    template += f"• {rooftop_info['rooftop_name']} – {rooftop_info['inbox_name']}\n"

                template += "\n"

                csm_template_text += template + "\n" + "="*80 + "\n"
                print(template)
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
                if display_name.startswith('Department line') or display_name.startswith('Staff line - User unassigned'):
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

            template += "\nIf you have any questions, feel free to email us at support@numa.com."

            # Create card frame with modern styling
            card_frame = tk.Frame(
                scrollable_frame,
                bg="white",
                highlightbackground="#dfe4ea",
                highlightthickness=1,
                relief=tk.FLAT
            )
            card_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=8)

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

            # Text widget for this template
            text_widget = tk.Text(
                card_frame,
                wrap=tk.WORD,
                height=10,
                font=("Segoe UI", 10),
                relief=tk.FLAT,
                borderwidth=0,
                padx=15,
                pady=10,
                bg="#f8f9fa",
                fg=self.text_color
            )
            text_widget.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))
            text_widget.insert(1.0, template)
            text_widget.config(state=tk.DISABLED)

            # Button frame
            button_frame = tk.Frame(card_frame, bg="white")
            button_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

            # Copy button for this template
            def make_copy_func(t, btn, orig_text):
                def copy_func():
                    self.root.clipboard_clear()
                    self.root.clipboard_append(t)
                    self.status_label.config(text=f"✓ Copied template for {rooftop_name} to clipboard", bg=self.success_color, fg="white")
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
            copy_btn.config(command=make_copy_func(template, copy_btn, button_text))
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
        for idx, (csm_owner, rooftop_list) in enumerate(csm_rooftops.items(), 1):
            # Generate clean template for this CSM
            template = f"Hi {self.get_first_name(csm_owner)},\n\n"
            template += "We've identified the following dealerships with low call volume over the past two weeks. To help us follow up, could you please provide a point of contact for each location so we can reach out directly?\n"

            for rooftop_info in rooftop_list:
                template += f"• {rooftop_info['rooftop_name']} – {rooftop_info['inbox_name']}\n"

            template += "\n"

            # Create card frame with modern styling
            card_frame = tk.Frame(
                scrollable_frame,
                bg="white",
                highlightbackground="#dfe4ea",
                highlightthickness=1,
                relief=tk.FLAT
            )
            card_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=8)

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

            # Text widget for this template
            text_widget = tk.Text(
                card_frame,
                wrap=tk.WORD,
                height=8,
                font=("Segoe UI", 10),
                relief=tk.FLAT,
                borderwidth=0,
                padx=15,
                pady=10,
                bg="#f8f9fa",
                fg=self.text_color
            )
            text_widget.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))
            text_widget.insert(1.0, template)
            text_widget.config(state=tk.DISABLED)

            # Button frame
            button_frame = tk.Frame(card_frame, bg="white")
            button_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

            # Copy button for this template
            def make_copy_func(t, btn, orig_text, csm):
                def copy_func():
                    self.root.clipboard_clear()
                    self.root.clipboard_append(t)
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
            copy_btn.config(command=make_copy_func(template, copy_btn, button_text, csm_owner))
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

        # Add summary at bottom
        summary_frame = tk.Frame(frame, bg=self.bg_color)
        summary_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=15)

        summary_label = tk.Label(
            summary_frame,
            text=f"✓ Generated {len(csm_rooftops)} CSM template(s) - One per CSM",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color,
            fg=self.accent_color
        )
        summary_label.pack(side=tk.LEFT)

        # Copy all button
        def copy_all():
            self.root.clipboard_clear()
            self.root.clipboard_append(template_text)
            self.status_label.config(text=f"✓ Copied all {len(csm_rooftops)} CSM template(s) to clipboard", bg=self.success_color, fg="white")
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
