// Global variables
let csvData = {};
let currentTab = null;
let currentCsvTab = null;
let csvExpanded = false;

// DOM Elements
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const browseButton = document.getElementById('browseButton');
const currentFileLabel = document.getElementById('currentFileLabel');
const statusLabel = document.getElementById('statusLabel');
const statusBar = statusLabel.parentElement;
const tabButtons = document.getElementById('tabButtons');
const tabContent = document.getElementById('tabContent');

// CSV Section Elements
const csvToggleBtn = document.getElementById('csvToggleBtn');
const csvDataContainer = document.getElementById('csvDataContainer');
const csvTabButtons = document.getElementById('csvTabButtons');
const csvTabContent = document.getElementById('csvTabContent');

// Toggle CSV Section
function toggleCsvSection() {
    csvExpanded = !csvExpanded;
    if (csvExpanded) {
        csvDataContainer.style.display = 'block';
        csvToggleBtn.textContent = '▼ Hide Raw CSV Data';
    } else {
        csvDataContainer.style.display = 'none';
        csvToggleBtn.textContent = '▶ View Raw CSV Data';
    }
}

// Event Listeners
browseButton.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', handleFileSelect);

// Drag and drop events
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const files = e.dataTransfer.files;
    handleFiles(files);
});

// File handling
function handleFileSelect(e) {
    handleFiles(e.target.files);
}

function handleFiles(files) {
    if (files.length === 0) return;

    csvData = {};
    clearTabs();

    const file = files[0];
    currentFileLabel.textContent = `Current file: ${file.name}`;
    updateStatus(`Processing: ${file.name}`);

    if (file.name.toLowerCase().endsWith('.zip')) {
        processZipFile(file);
    } else if (file.name.toLowerCase().endsWith('.csv')) {
        processCSVFile(file);
    } else {
        updateStatus('Please select a ZIP or CSV file');
    }
}

// Process ZIP file
async function processZipFile(file) {
    try {
        const zip = await JSZip.loadAsync(file);
        const csvFiles = Object.keys(zip.files).filter(name =>
            name.toLowerCase().endsWith('.csv') && !name.startsWith('__MACOSX')
        );

        if (csvFiles.length === 0) {
            updateStatus('No CSV files found in ZIP archive');
            return;
        }

        for (const filename of csvFiles) {
            const content = await zip.files[filename].async('string');
            const parsed = Papa.parse(content, { header: false });
            csvData[filename.toLowerCase()] = parsed.data;
            displayCSVTab(filename, parsed.data);
        }

        generateTemplates();
        updateStatus(`✓ Loaded ${csvFiles.length} CSV file(s) from ${file.name}`);
    } catch (error) {
        updateStatus(`Error: ${error.message}`);
        console.error(error);
    }
}

// Process CSV file
function processCSVFile(file) {
    Papa.parse(file, {
        complete: (results) => {
            csvData[file.name.toLowerCase()] = results.data;
            displayCSVTab(file.name, results.data);
            generateTemplates();
            updateStatus(`✓ Loaded ${file.name}`);
        },
        error: (error) => {
            updateStatus(`Error: ${error.message}`);
        }
    });
}

// Display CSV in a tab (in the collapsible CSV section)
function displayCSVTab(filename, rows) {
    if (!rows || rows.length === 0) return;

    const tabId = `csv-tab-${Date.now()}-${Math.random()}`;

    // Create tab button in CSV section
    const tabButton = document.createElement('button');
    tabButton.className = 'csv-tab-button';
    tabButton.textContent = filename;
    tabButton.onclick = () => switchCsvTab(tabId, tabButton);
    csvTabButtons.appendChild(tabButton);

    // Create tab panel in CSV section
    const tabPanel = document.createElement('div');
    tabPanel.id = tabId;
    tabPanel.className = 'csv-tab-panel';

    const headers = rows[0];
    const dataRows = rows.slice(1).filter(row => row.some(cell => cell && cell.trim()));

    // Create table
    const tableContainer = document.createElement('div');
    tableContainer.className = 'csv-table-container';

    const table = document.createElement('table');
    table.className = 'csv-table';

    // Table header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Table body
    const tbody = document.createElement('tbody');
    dataRows.forEach(row => {
        const tr = document.createElement('tr');
        headers.forEach((_, idx) => {
            const td = document.createElement('td');
            td.textContent = row[idx] || '';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    tableContainer.appendChild(table);
    tabPanel.appendChild(tableContainer);

    // Info label
    const info = document.createElement('div');
    info.className = 'csv-info';
    info.textContent = `Rows: ${dataRows.length} | Columns: ${headers.length}`;
    tabPanel.appendChild(info);

    csvTabContent.appendChild(tabPanel);

    // Switch to first CSV tab
    if (csvTabButtons.children.length === 1) {
        switchCsvTab(tabId, tabButton);
    }
}

// Switch CSV tabs (in collapsible section)
function switchCsvTab(tabId, buttonElement) {
    // Update buttons
    Array.from(csvTabButtons.children).forEach(btn => btn.classList.remove('active'));
    buttonElement.classList.add('active');

    // Update panels
    Array.from(csvTabContent.children).forEach(panel => panel.classList.remove('active'));
    document.getElementById(tabId).classList.add('active');

    currentCsvTab = tabId;
}

// Switch tabs
function switchTab(tabId, buttonElement) {
    // Update buttons
    Array.from(tabButtons.children).forEach(btn => btn.classList.remove('active'));
    buttonElement.classList.add('active');

    // Update panels
    Array.from(tabContent.children).forEach(panel => panel.classList.remove('active'));
    document.getElementById(tabId).classList.add('active');

    currentTab = tabId;
}

// Clear all tabs
function clearTabs() {
    // Clear main template tabs
    tabButtons.innerHTML = '';
    tabContent.innerHTML = '';
    currentTab = null;

    // Clear CSV tabs
    csvTabButtons.innerHTML = '';
    csvTabContent.innerHTML = '';
    currentCsvTab = null;
}

// Update status
function updateStatus(message, isSuccess = false) {
    statusLabel.textContent = message;
    if (isSuccess) {
        statusBar.classList.add('success');
        setTimeout(() => statusBar.classList.remove('success'), 2000);
    } else {
        statusBar.classList.remove('success');
    }
}

// Utility functions
function formatPhoneNumber(phone) {
    const digits = phone.toString().replace(/\D/g, '');
    if (digits.length === 10) {
        return `(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`;
    } else if (digits.length === 11 && digits[0] === '1') {
        return `(${digits.slice(1, 4)}) ${digits.slice(4, 7)}-${digits.slice(7)}`;
    }
    return phone;
}

function capitalizeName(name) {
    if (!name) return name;
    // Capitalize each word (title case)
    return name.toLowerCase().replace(/\b\w/g, char => char.toUpperCase());
}

function getFirstName(fullName) {
    if (!fullName) return fullName;
    return fullName.trim().split(/\s+/)[0];
}

function findColIdx(headers, possibleNames) {
    for (const name of possibleNames) {
        for (let idx = 0; idx < headers.length; idx++) {
            if (headers[idx].toLowerCase().includes(name.toLowerCase())) {
                return idx;
            }
        }
    }
    return null;
}

function findExactColIdx(headers, targetName) {
    for (let idx = 0; idx < headers.length; idx++) {
        if (headers[idx].trim().toLowerCase() === targetName.toLowerCase()) {
            return idx;
        }
    }
    return null;
}

// Generate templates
function generateTemplates() {
    let linesFile = null;
    let rooftopFile = null;
    let deskPhonesFile = null;

    for (const [filename, rows] of Object.entries(csvData)) {
        if (filename.includes('lines_with_low') && filename.includes('call_volume')) {
            linesFile = rows;
        } else if (filename.includes('rooftop_information') || filename.includes('rooftop_informatio')) {
            rooftopFile = rows;
        } else if (filename.includes('desk_phones')) {
            deskPhonesFile = rows;
        }
    }

    if (!linesFile || !rooftopFile) {
        console.log('Required CSV files not found for template generation');
        return;
    }

    // Build desk phone lookup (display name -> phone number) if desk_phones file exists
    const deskPhoneLookup = {};
    if (deskPhonesFile && deskPhonesFile.length > 1) {
        const deskHeaders = deskPhonesFile[0];
        const deskData = deskPhonesFile.slice(1).filter(row => row.some(cell => cell && cell.trim()));

        // Find column indices for desk phones file
        let deskDisplayNameIdx = null;
        let deskPhoneNumberIdx = null;

        deskHeaders.forEach((header, idx) => {
            const headerLower = header.toLowerCase().trim();
            // Check for display name column
            if (headerLower.includes('display name') || headerLower.includes('display_name')) {
                deskDisplayNameIdx = idx;
            }
            // Check for phone number column - handle various naming conventions (separate if, not elif)
            if (headerLower.includes('phone number') || headerLower.includes('phone_number') || headerLower.includes('phone numbers')) {
                deskPhoneNumberIdx = idx;
            }
        });

        if (deskDisplayNameIdx !== null && deskPhoneNumberIdx !== null) {
            deskData.forEach(row => {
                if (row.length > Math.max(deskDisplayNameIdx, deskPhoneNumberIdx)) {
                    const displayName = (row[deskDisplayNameIdx] || '').trim().toLowerCase();
                    const phoneNumber = (row[deskPhoneNumberIdx] || '').trim();
                    if (displayName && phoneNumber) {
                        deskPhoneLookup[displayName] = phoneNumber;
                    }
                }
            });
        }
    }

    try {
        const linesHeaders = linesFile[0];
        const linesData = linesFile.slice(1).filter(row => row.some(cell => cell && cell.trim()));

        const displayNameIdx = findColIdx(linesHeaders, ['display name', 'display_name']);
        const phoneNumberIdx = findColIdx(linesHeaders, ['phone number', 'phone_number', 'number']);
        const rooftopNameIdx = findColIdx(linesHeaders, ['rooftop name', 'rooftop_name', 'rooftop']);
        const inboxNameIdx = findColIdx(linesHeaders, ['inbox name', 'inbox_name', 'inbox']);
        const ownerTypeIdx = findColIdx(linesHeaders, ['owner type', 'owner_type', 'ownertype']);
        const nameIdx = findExactColIdx(linesHeaders, 'name');

        if ([displayNameIdx, phoneNumberIdx, rooftopNameIdx, inboxNameIdx].includes(null)) {
            console.error('Could not find required columns');
            return;
        }

        // Group by rooftop
        const rooftops = {};
        linesData.forEach(row => {
            if (row.length <= Math.max(displayNameIdx, phoneNumberIdx, rooftopNameIdx, inboxNameIdx)) return;

            const rooftop = row[rooftopNameIdx]?.trim();
            if (!rooftop) return;

            if (!rooftops[rooftop]) {
                rooftops[rooftop] = {
                    inbox_name: '',
                    lines: []
                };
            }

            rooftops[rooftop].inbox_name = row[inboxNameIdx]?.trim() || '';

            let displayName = row[displayNameIdx]?.trim() || '';

            if (!displayName) {
                const ownerType = row[ownerTypeIdx]?.trim().toUpperCase() || '';
                const nameValue = row[nameIdx]?.trim() || '';

                if (ownerType === 'USER') {
                    displayName = `Unassigned line - [${capitalizeName(nameValue)}]`;
                } else if (ownerType === 'DEPARTMENT') {
                    displayName = `Unassigned line - [${capitalizeName(nameValue)}]`;
                } else {
                    displayName = capitalizeName(nameValue) || 'Unknown';
                }
            } else {
                displayName = capitalizeName(displayName);
            }

            // Get raw values for desk phone table
            const rawDisplayName = row[displayNameIdx]?.trim() || '';
            const rawName = row[nameIdx]?.trim() || '';

            // Look up desk phone number by matching display name (case-insensitive)
            const deskPhone = deskPhoneLookup[rawDisplayName.toLowerCase()] || '';

            rooftops[rooftop].lines.push({
                display_name: displayName,
                phone_number: formatPhoneNumber(row[phoneNumberIdx]?.trim() || ''),
                raw_display_name: rawDisplayName,
                raw_name: rawName,
                desk_phone: deskPhone
            });
        });

        // Create CSM templates first
        generateCSMTemplates(rooftopFile, rooftops);

        // Create dealership templates
        createDealershipTemplatesTab(rooftops);

    } catch (error) {
        console.error('Error generating templates:', error);
    }
}

// Create dealership templates tab
function createDealershipTemplatesTab(rooftops) {
    const tabId = `tab-dealership-${Date.now()}`;

    // Create tab button
    const tabButton = document.createElement('button');
    tabButton.className = 'tab-button';
    tabButton.textContent = 'Dealership Templates';
    tabButton.onclick = () => switchTab(tabId, tabButton);
    tabButtons.appendChild(tabButton);

    // Create tab panel
    const tabPanel = document.createElement('div');
    tabPanel.id = tabId;
    tabPanel.className = 'tab-panel';

    const templatesContainer = document.createElement('div');
    templatesContainer.className = 'templates-container';

    let allTemplatesText = '';
    let idx = 1;

    for (const [rooftopName, data] of Object.entries(rooftops)) {
        const inboxName = data.inbox_name;
        const lines = data.lines;

        // Separate lines
        const regularLines = [];
        const departmentUnassignedLines = [];

        lines.forEach(line => {
            if (line.display_name.startsWith('Unassigned line')) {
                departmentUnassignedLines.push(line);
            } else {
                regularLines.push(line);
            }
        });

        // Generate template
        let template = `Good morning [Dealership POC],

We've recently noticed a drop in call volume on your account (${rooftopName} – ${inboxName})

To ensure you're getting the most out of your Numa subscription, please confirm that missed calls on the following users' direct lines are forwarding to their respective Numa IT forwarding lines after 4 rings (approximately 20 seconds), rather than going to local voicemail (including DND, busy, and after-hours scenarios):
`;

        regularLines.forEach(line => {
            template += `• ${line.display_name} – Numa IT forwarding number: ${line.phone_number}\n`;
        });

        departmentUnassignedLines.forEach(line => {
            template += `• ${line.display_name} – Numa IT forwarding number: ${line.phone_number}\n`;
        });

        template += `\nAdditionally, when you have a moment, kindly update the following roster with the latest desk phone numbers for your staff members
Roster link [insert roster link here]

If you have any questions, feel free to email us at support@numa.com.`;

        allTemplatesText += template + '\n\n' + '='.repeat(80) + '\n\n';

        // Create card with subject line and desk phones data
        const subjectLine = `${rooftopName} - ${inboxName}: Phoneline forwarding`;
        const deskPhonesData = [...regularLines, ...departmentUnassignedLines];
        const card = createTemplateCard(idx, rooftopName, template, lines.length, rooftopName, false, subjectLine, deskPhonesData);
        templatesContainer.appendChild(card);
        idx++;
    }

    tabPanel.appendChild(templatesContainer);

    // Summary section
    const summary = createSummarySection(Object.keys(rooftops).length, allTemplatesText, 'template');
    tabPanel.appendChild(summary);

    tabContent.appendChild(tabPanel);
}

// Generate CSM templates
function generateCSMTemplates(rooftopFile, rooftops) {
    const rooftopHeaders = rooftopFile[0];
    const rooftopData = rooftopFile.slice(1).filter(row => row.some(cell => cell && cell.trim()));

    const rooftopNameColIdx = findColIdx(rooftopHeaders, ['rooftop name', 'rooftop_name', 'rooftop']);
    const csmOwnerIdx = findColIdx(rooftopHeaders, ['csm owner', 'csm_owner', 'csmowner']);

    if (rooftopNameColIdx === null || csmOwnerIdx === null) {
        console.log('Could not find CSM Owner or Rooftop Name columns');
        return;
    }

    // Build mapping
    const rooftopToCSM = {};
    rooftopData.forEach(row => {
        if (row.length > Math.max(rooftopNameColIdx, csmOwnerIdx)) {
            const rooftopName = row[rooftopNameColIdx]?.trim();
            const csmOwner = row[csmOwnerIdx]?.trim();
            if (rooftopName && csmOwner) {
                rooftopToCSM[rooftopName] = csmOwner;
            }
        }
    });

    // Group by CSM
    const csmRooftops = {};
    for (const [rooftopName, data] of Object.entries(rooftops)) {
        const csmOwner = rooftopToCSM[rooftopName] || 'Unknown CSM';
        if (!csmRooftops[csmOwner]) {
            csmRooftops[csmOwner] = [];
        }
        csmRooftops[csmOwner].push({
            rooftop_name: rooftopName,
            inbox_name: data.inbox_name
        });
    }

    createCSMTemplatesTab(csmRooftops);
}

// Create CSM templates tab
function createCSMTemplatesTab(csmRooftops) {
    const tabId = `tab-csm-${Date.now()}`;

    // Create tab button
    const tabButton = document.createElement('button');
    tabButton.className = 'tab-button';
    tabButton.textContent = 'CSM Templates';
    tabButton.onclick = () => switchTab(tabId, tabButton);
    tabButtons.appendChild(tabButton);

    // Create tab panel
    const tabPanel = document.createElement('div');
    tabPanel.id = tabId;
    tabPanel.className = 'tab-panel';

    const templatesContainer = document.createElement('div');
    templatesContainer.className = 'templates-container';

    let allTemplatesText = '';
    let idx = 1;

    for (const [csmOwner, rooftopList] of Object.entries(csmRooftops)) {
        let template = `Hi ${getFirstName(csmOwner)},

We've identified the following dealerships with low call volume over the past two weeks. To help us follow up, could you please provide a point of contact for each location so we can reach out directly?
`;

        rooftopList.forEach(rooftop => {
            template += `• ${rooftop.rooftop_name} – ${rooftop.inbox_name}\n`;
        });

        template += '\nPlease let us know whether the lines are intentionally not forwarding, or if you\'d prefer that we avoid contacting any of the dealerships mentioned above.';

        allTemplatesText += template + '\n' + '='.repeat(80) + '\n\n';

        // Create card
        const card = createTemplateCard(idx, csmOwner, template, rooftopList.length, csmOwner, true);
        templatesContainer.appendChild(card);
        idx++;
    }

    tabPanel.appendChild(templatesContainer);

    // Summary section
    const summary = createSummarySection(Object.keys(csmRooftops).length, allTemplatesText, 'CSM template');
    tabPanel.appendChild(summary);

    tabContent.appendChild(tabPanel);
}

// Create template card
function createTemplateCard(idx, title, template, count, entityName, isCSM = false, subjectLine = null, deskPhonesData = null) {
    const card = document.createElement('div');
    card.className = 'template-card';

    // Header
    const header = document.createElement('div');
    header.className = 'card-header';
    const titleEl = document.createElement('div');
    titleEl.className = 'card-title';
    titleEl.textContent = `Template ${idx}: ${title}`;
    header.appendChild(titleEl);

    // Subject line (only for dealership templates)
    let subjectInput = null;
    if (subjectLine) {
        const subjectContainer = document.createElement('div');
        subjectContainer.className = 'subject-container';

        const subjectLabel = document.createElement('span');
        subjectLabel.className = 'subject-label';
        subjectLabel.textContent = 'Subject: ';

        subjectInput = document.createElement('input');
        subjectInput.type = 'text';
        subjectInput.className = 'subject-input';
        subjectInput.value = subjectLine;

        subjectContainer.appendChild(subjectLabel);
        subjectContainer.appendChild(subjectInput);
        card.appendChild(header);
        card.appendChild(subjectContainer);
    } else {
        card.appendChild(header);
    }

    // Content container - side by side layout for template and desk phones table
    const contentContainer = document.createElement('div');
    contentContainer.className = 'card-content-container';

    // Left side - Template text (editable textarea)
    const leftSide = document.createElement('div');
    leftSide.className = 'card-left-side';

    const textArea = document.createElement('textarea');
    textArea.className = 'template-text';
    textArea.value = template;
    textArea.rows = 12;
    leftSide.appendChild(textArea);

    contentContainer.appendChild(leftSide);

    // Right side - Possible desk phones table (only for dealership templates)
    if (deskPhonesData && deskPhonesData.length > 0) {
        const rightSide = document.createElement('div');
        rightSide.className = 'card-right-side';

        const deskPhonesLabel = document.createElement('div');
        deskPhonesLabel.className = 'desk-phones-label';
        deskPhonesLabel.textContent = 'Possible Desk Phones';
        rightSide.appendChild(deskPhonesLabel);

        const tableContainer = document.createElement('div');
        tableContainer.className = 'desk-phones-table-container';

        const table = document.createElement('table');
        table.className = 'desk-phones-table';

        // Table header
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        const th1 = document.createElement('th');
        th1.textContent = 'Display Name';
        const th2 = document.createElement('th');
        th2.textContent = 'Name';
        headerRow.appendChild(th1);
        headerRow.appendChild(th2);
        thead.appendChild(headerRow);
        table.appendChild(thead);

        // Table body
        const tbody = document.createElement('tbody');
        deskPhonesData.forEach(line => {
            const tr = document.createElement('tr');
            const td1 = document.createElement('td');
            td1.textContent = line.raw_display_name || '';
            const td2 = document.createElement('td');
            td2.textContent = line.raw_name || '';
            tr.appendChild(td1);
            tr.appendChild(td2);
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);

        tableContainer.appendChild(table);
        rightSide.appendChild(tableContainer);
        contentContainer.appendChild(rightSide);
    }

    card.appendChild(contentContainer);

    // Footer
    const footer = document.createElement('div');
    footer.className = 'card-footer';

    const copyBtn = document.createElement('button');
    copyBtn.className = 'copy-button';
    copyBtn.textContent = `📋 Copy Template ${idx}`;
    copyBtn.onclick = () => {
        // Get the current (possibly edited) text from textarea
        let textToCopy = textArea.value;
        // Include subject line if present
        if (subjectInput) {
            textToCopy = `Subject: ${subjectInput.value}\n\n${textToCopy}`;
        }
        navigator.clipboard.writeText(textToCopy).then(() => {
            const originalText = copyBtn.textContent;
            copyBtn.textContent = '✓ Copied!';
            copyBtn.classList.add('copied');
            updateStatus(isCSM ?
                `✓ Copied CSM template for ${entityName} to clipboard` :
                `✓ Copied template for ${entityName} to clipboard`, true);
            setTimeout(() => {
                copyBtn.textContent = originalText;
                copyBtn.classList.remove('copied');
            }, 2000);
        });
    };

    const info = document.createElement('span');
    info.className = 'card-info';
    info.textContent = isCSM ? `${count} rooftop(s)` : `${count} phone line(s)`;

    footer.appendChild(copyBtn);
    footer.appendChild(info);

    card.appendChild(footer);

    return card;
}

// Create summary section
function createSummarySection(count, allText, type) {
    const summary = document.createElement('div');
    summary.className = 'summary-section';

    const label = document.createElement('div');
    label.className = 'summary-label';
    label.textContent = `✓ Generated ${count} ${type}(s) - One per ${type.includes('CSM') ? 'CSM' : 'rooftop'}`;

    const copyAllBtn = document.createElement('button');
    copyAllBtn.className = 'copy-all-button';
    copyAllBtn.textContent = '📋 Copy All Templates';
    copyAllBtn.onclick = () => {
        navigator.clipboard.writeText(allText).then(() => {
            const originalText = copyAllBtn.textContent;
            copyAllBtn.textContent = '✓ Copied!';
            copyAllBtn.classList.add('copied');
            updateStatus(`✓ Copied all ${count} ${type}(s) to clipboard`, true);
            setTimeout(() => {
                copyAllBtn.textContent = originalText;
                copyAllBtn.classList.remove('copied');
            }, 2000);
        });
    };

    summary.appendChild(label);
    summary.appendChild(copyAllBtn);

    return summary;
}
