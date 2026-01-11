# Audit Template Generator

A web-based tool for generating audit templates from CSV files. Processes low call volume data and rooftop information to create customized email templates for dealerships and CSMs.

## Features

- 📁 Drag & drop or browse for ZIP/CSV files
- 📊 Display CSV data in interactive tables
- 📧 Generate dealership-specific email templates
- 👥 Generate CSM-grouped templates
- 📋 One-click copy to clipboard
- 🎨 Modern, responsive UI
- 🚀 Works entirely in the browser (no server needed)

## How to Use

1. Visit the hosted website (see deployment instructions below)
2. Click "Browse for ZIP or CSV file" or drag and drop your file
3. Upload a ZIP file containing both required CSVs, or upload CSV files individually
4. View the generated templates in the tabs
5. Click the copy button to copy templates to clipboard

## Required CSV Files

The tool expects two CSV files:

1. **lines_with_low_inbound_call_volume.csv** (or similar name with "lines_with_low" and "call_volume")
   - Required columns: Display Name, Phone Number, Rooftop Name, Inbox Name, Owner Type, Name

2. **rooftop_information.csv**
   - Required columns: Rooftop Name, CSM Owner

## Deployment to GitHub Pages

### Option 1: Using GitHub Web Interface

1. Create a new repository on GitHub (e.g., "audit-template-generator")
2. Upload the following files to the repository:
   - `index.html`
   - `styles.css`
   - `script.js`
   - `README.md` (optional)

3. Go to repository Settings → Pages
4. Under "Source", select "Deploy from a branch"
5. Select branch: `main` (or `master`)
6. Select folder: `/ (root)`
7. Click Save
8. Wait a few minutes for deployment
9. Your site will be available at: `https://[your-username].github.io/[repo-name]/`

### Option 2: Using Git Command Line

```bash
# Initialize git repository
cd "c:\Users\Rouvin Rebello\Projects\Audit templates"
git init

# Add files
git add index.html styles.css script.js README.md

# Create initial commit
git commit -m "Initial commit: Audit Template Generator"

# Add remote repository (create repo on GitHub first)
git remote add origin https://github.com/[your-username]/[repo-name].git

# Push to GitHub
git branch -M main
git push -u origin main

# Enable GitHub Pages in repository settings
```

Then follow steps 3-9 from Option 1.

## Technologies Used

- **HTML5** - Structure
- **CSS3** - Modern styling
- **JavaScript (ES6+)** - Logic and interactivity
- **JSZip** - ZIP file processing
- **PapaParse** - CSV parsing

## Browser Compatibility

Works on all modern browsers:
- Chrome/Edge (recommended)
- Firefox
- Safari
- Opera

## Local Development

To run locally:

1. Clone the repository
2. Open `index.html` in a web browser
3. No build process or server required!

## Features Breakdown

### Template Generation
- Automatically groups phone lines by rooftop
- Formats phone numbers as (111) 222-3333
- Handles null display names with fallback logic
- Capitalizes names properly
- Groups department/unassigned lines at bottom

### CSM Templates
- Groups rooftops by CSM owner
- Uses first name only in greeting
- Lists all rooftops for each CSM

### User Interface
- Modern card-based design
- Color-coded status messages
- Smooth animations and transitions
- Responsive layout for mobile devices
- Accessible keyboard navigation

## Privacy & Security

- **100% Client-Side**: All processing happens in your browser
- **No Data Upload**: Files never leave your computer
- **No Tracking**: No analytics or tracking code
- **No Storage**: Files are processed in memory only

## Support

For issues or feature requests, please create an issue in the GitHub repository.

## License

Free to use for internal business purposes.
