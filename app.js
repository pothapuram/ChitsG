// Store sheet data globally
let sheets = [];

// Initialize Google API
function initGAPI() {
    return new Promise((resolve, reject) => {
        gapi.load('client', () => {
            gapi.client.init({
                apiKey: 'AIzaSyDRyk27hRoCv2KLYhMx1mKAKWTnEXGk35g',
                discoveryDocs: ["https://www.googleapis.com/discovery/v1/apis/drive/v3/rest"]
            }).then(() => {
                resolve();
            }, (error) => {
                reject(error);
            });
        });
    });
}

// Function to load Excel file from Google Drive
async function loadExcelFromDrive() {
    try {
        // Initialize Google API
        await initGAPI();
        
        // Replace with your Google Drive file ID
        const fileId = '1e9NoVbSkXKYYjJ30Wxy4h13GCYVvJwSl';
        
        // Create a direct download URL using the API key
        const downloadUrl = `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media&key=AIzaSyDRyk27hRoCv2KLYhMx1mKAKWTnEXGk35g`;
        
        // Fetch the file using the download URL
        const response = await fetch(downloadUrl);
        
        if (!response.ok) {
            throw new Error(`Failed to fetch file: ${response.status} ${response.statusText}`);
        }
        
        // Handle the response as array buffer
        const arrayBuffer = await response.arrayBuffer();
        
        // Convert to Uint8Array
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });

        // Clear existing sheets
        const sheetContainer = document.getElementById('sheet-container');
        const navbar = document.querySelector('.navbar');
        if (!sheetContainer || !navbar) {
            throw new Error('Required DOM elements not found');
        }
        
        sheetContainer.innerHTML = '';
        navbar.innerHTML = '';

        // Process each sheet
        const sheetNames = workbook.SheetNames;
        if (!sheetNames || sheetNames.length === 0) {
            throw new Error('No sheets found in the Excel file');
        }

        sheetNames.forEach((sheetName, index) => {
            try {
                const worksheet = workbook.Sheets[sheetName];
                const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                
                // Store sheet data
                sheets[index] = json;
                
                // Add sheet name to navbar
                const navLink = document.createElement('a');
                navLink.href = '#';
                navLink.textContent = sheetName;
                navLink.dataset.sheet = index;
                navLink.onclick = (e) => {
                    e.preventDefault();
                    showSheet(index);
                };
                navbar.appendChild(navLink);
            } catch (sheetError) {
                console.error(`Error processing sheet ${sheetName}:`, sheetError);
            }
        });
        
        // Show first sheet by default if available
        if (sheetNames.length > 0) {
            showSheet(0);
        }
    } catch (error) {
        console.error('Error loading Excel file:', error);
        const errorMessage = error.message || 'Error loading Excel file';
        alert(`Error: ${errorMessage}. Please check the console for more details.`);
    }
}

// Function to display sheet data in a table
function showSheet(sheetIndex) {
    const container = document.getElementById('sheet-container');
    if (!container) return;
    
    const data = sheets[sheetIndex];
    if (!data || data.length === 0) {
        container.innerHTML = '<p>No data available for this sheet.</p>';
        return;
    }
    
    try {
        // Create table
        const table = document.createElement('table');
        table.className = 'excel-table';
        table.id = 'members';
        
        // Create header row if data exists
        if (data.length > 0) {
            const thead = document.createElement('thead');
            const headerRow = document.createElement('tr');
            
            data[0].forEach((header, colIndex) => {
                const th = document.createElement('th');
                th.textContent = header || `Column ${colIndex + 1}`;
                headerRow.appendChild(th);
            });
            
            thead.appendChild(headerRow);
            table.appendChild(thead);
        }
        
        // Create data rows
        if (data.length > 1) {
            debugger;
            const tbody = document.createElement('tbody');
            let isMobileNumber = false;
            for (let i = 1; i < data.length; i++) {
                const row = document.createElement('tr');
                data[i].forEach((cell, colIndex) => {
                    const td = document.createElement('td');

                    if(cell === "Total Amount Paid Per Head :")
                        {
                            td.colSpan = 2;
                            td.style.fontWeight = "bold";
                            td.style.textAlign = "right";
                            td.textContent = cell;
                            row.appendChild(td);
                            return;
                        }

                    if(cell === "சீட்டு எடுக்காதவர்கள்")
                    {
                        td.colSpan = 3;
                        td.style.fontWeight = "bold";
                        td.style.textAlign = "center";
                        td.textContent = cell;
                        row.appendChild(td);
                        return;
                    }
                    else if(cell === "S.No" || cell === "Name" || cell === "Mobile Number")
                    {
                        td.style.fontWeight = "bold";
                        td.style.textAlign = "center";
                        td.textContent = cell;
                        row.appendChild(td);
                        isMobileNumber = true;
                        return;
                    }
                    else if(isMobileNumber && colIndex === 2 && typeof cell === 'number' && !isNaN(cell))
                    {
                         //td.appendChild(document.createElement('a').textContent = `tel:${cell}`)
                         //td.appendChild(document.createElement('a').href = `tel:${cell}`)
                         td.appendChild(Object.assign(document.createElement('a'), { href: `tel:${cell}`, textContent: cell }));
                         //td.style.fontWeight = "bold";
                         td.style.textAlign = "left";
                         //td.textContent = cell;
                        row.appendChild(td);
                        return;
                    }
                    else{
                        td.textContent = cell;
                        row.appendChild(td);
                    }
                });
                tbody.appendChild(row);
            }
            table.appendChild(tbody);
        }
        
        container.innerHTML = '';
        container.appendChild(table);
    } catch (error) {
        console.error('Error displaying sheet:', error);
        container.innerHTML = '<p>Error displaying sheet data. Please check the console for details.</p>';
    }
}

// Load file when page loads
window.onload = loadExcelFromDrive;
