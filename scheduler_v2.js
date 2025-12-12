// SchedulePro V2 - Grid-based scheduler with swimlane day configuration
// Data structures
let courses = [];
let events = [];
let eventDays = [];
let instructorUnavailable = []; // Raw unavailability data from CSV
let unavailabilityMap = {}; // Pre-calculated: { 'Instructor-EventID': [blockedDayNumbers] }
let assignments = {}; // { courseId: [eventIds] }
let schedule = {}; // { eventId: { courseId: { startDay, days: [] } } }
let rooms = {}; // { eventId: { courseId: "Room Name" } }
let draggedBlock = null;
let currentTimeline = null;

// Change logging
let changeLog = [];
let uploadsLog = [];
let errorsLog = [];

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    loadLogs();
    loadRoundData(); // Load saved round data if it exists
    
    // Start with blank slate - no auto-loading of default data
    // Client must upload their own Dates.csv to begin
    
    setupFileInput();
    setupUnavailabilityFileInput();
    setupScheduleFileInput();
    setupEventsFileInput();
    setupExcelFileInput();
    // setupEventDaysFileInput(); // Commented out - element removed from HTML (using consolidated Dates.csv)
});

// Load logs from localStorage
function loadLogs() {
    const savedChangelog = localStorage.getItem('schedulepro_changelog');
    const savedUploads = localStorage.getItem('schedulepro_uploads');
    const savedErrors = localStorage.getItem('schedulepro_errors');
    
    if (savedChangelog) changeLog = JSON.parse(savedChangelog);
    if (savedUploads) uploadsLog = JSON.parse(savedUploads);
    if (savedErrors) errorsLog = JSON.parse(savedErrors);
}

// Save logs to localStorage
function saveLogs() {
    localStorage.setItem('schedulepro_changelog', JSON.stringify(changeLog));
    localStorage.setItem('schedulepro_uploads', JSON.stringify(uploadsLog));
    localStorage.setItem('schedulepro_errors', JSON.stringify(errorsLog));
}

// Clear change log with two-stage confirmation
function clearChangeLog() {
    // Create first dialog
    const dialog1 = document.createElement('div');
    dialog1.style.cssText = 'position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); display: flex; align-items: center; justify-content: center; z-index: 10000;';
    
    const box1 = document.createElement('div');
    box1.style.cssText = 'background: white; padding: 30px; border-radius: 10px; max-width: 500px; box-shadow: 0 10px 40px rgba(0,0,0,0.3);';
    
    box1.innerHTML = `
        <h3 style="margin-bottom: 20px; color: #667eea;">Export Change Log?</h3>
        <p style="margin-bottom: 25px; line-height: 1.6;">Would you like to export the current change log before clearing it?</p>
        <div style="display: flex; gap: 10px; justify-content: flex-end;">
            <button id="exportAndClear" style="padding: 10px 20px; background: #667eea; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: 600;">Export & Clear</button>
            <button id="clearWithoutExport" style="padding: 10px 20px; background: #6c757d; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: 600;">Clear Without Exporting</button>
            <button id="cancelClear1" style="padding: 10px 20px; background: #495057; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: 600;">Cancel</button>
        </div>
    `;
    
    dialog1.appendChild(box1);
    document.body.appendChild(dialog1);
    
    document.getElementById('exportAndClear').onclick = () => {
        document.body.removeChild(dialog1);
        exportChangelog();
        performChangeLogClear();
    };
    
    document.getElementById('clearWithoutExport').onclick = () => {
        document.body.removeChild(dialog1);
        showFinalWarning();
    };
    
    document.getElementById('cancelClear1').onclick = () => {
        document.body.removeChild(dialog1);
    };
    
    // Close on background click
    dialog1.onclick = (e) => {
        if (e.target === dialog1) {
            document.body.removeChild(dialog1);
        }
    };
}

// Show final warning before clearing change log
function showFinalWarning() {
    const dialog2 = document.createElement('div');
    dialog2.style.cssText = 'position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); display: flex; align-items: center; justify-content: center; z-index: 10000;';
    
    const box2 = document.createElement('div');
    box2.style.cssText = 'background: white; padding: 30px; border-radius: 10px; max-width: 550px; box-shadow: 0 10px 40px rgba(0,0,0,0.3);';
    
    box2.innerHTML = `
        <h3 style="margin-bottom: 20px; color: #dc3545;">⚠️ WARNING</h3>
        <p style="margin-bottom: 25px; line-height: 1.6; font-size: 1.05em;">
            This will <strong>permanently delete all changes logged since the beginning of time</strong>. 
            This action cannot be undone.
        </p>
        <p style="margin-bottom: 25px; line-height: 1.6; font-weight: 600;">Are you sure you wish to proceed?</p>
        <div style="display: flex; gap: 10px; justify-content: flex-end;">
            <button id="confirmDelete" style="padding: 10px 20px; background: #dc3545; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: 600;">Yes, Delete All Changes</button>
            <button id="cancelClear2" style="padding: 10px 20px; background: #667eea; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: 600;">Cancel</button>
        </div>
    `;
    
    dialog2.appendChild(box2);
    document.body.appendChild(dialog2);
    
    document.getElementById('confirmDelete').onclick = () => {
        document.body.removeChild(dialog2);
        performChangeLogClear();
    };
    
    document.getElementById('cancelClear2').onclick = () => {
        document.body.removeChild(dialog2);
    };
    
    // Close on background click
    dialog2.onclick = (e) => {
        if (e.target === dialog2) {
            document.body.removeChild(dialog2);
        }
    };
}

// Actually perform the change log clear
function performChangeLogClear() {
    changeLog = [];
    localStorage.setItem('schedulepro_changelog', JSON.stringify(changeLog));
    alert('✅ Change log has been cleared successfully.');
}

// Clear uploads and errors log with single confirmation
function clearUploadsAndErrors() {
    const dialog = document.createElement('div');
    dialog.style.cssText = 'position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); display: flex; align-items: center; justify-content: center; z-index: 10000;';
    
    const box = document.createElement('div');
    box.style.cssText = 'background: white; padding: 30px; border-radius: 10px; max-width: 500px; box-shadow: 0 10px 40px rgba(0,0,0,0.3);';
    
    box.innerHTML = `
        <h3 style="margin-bottom: 20px; color: #667eea;">Clear Uploads & Errors Log?</h3>
        <p style="margin-bottom: 25px; line-height: 1.6;">Are you sure you want to clear the Uploads & Errors log?</p>
        <div style="display: flex; gap: 10px; justify-content: flex-end;">
            <button id="confirmClearLogs" style="padding: 10px 20px; background: #dc3545; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: 600;">Yes</button>
            <button id="cancelClearLogs" style="padding: 10px 20px; background: #667eea; color: white; border: none; border-radius: 5px; cursor: pointer; font-weight: 600;">Cancel</button>
        </div>
    `;
    
    dialog.appendChild(box);
    document.body.appendChild(dialog);
    
    document.getElementById('confirmClearLogs').onclick = () => {
        document.body.removeChild(dialog);
        uploadsLog = [];
        errorsLog = [];
        localStorage.setItem('schedulepro_uploads', JSON.stringify(uploadsLog));
        localStorage.setItem('schedulepro_errors', JSON.stringify(errorsLog));
        alert('✅ Uploads & Errors log has been cleared successfully.');
    };
    
    document.getElementById('cancelClearLogs').onclick = () => {
        document.body.removeChild(dialog);
    };
    
    // Close on background click
    dialog.onclick = (e) => {
        if (e.target === dialog) {
            document.body.removeChild(dialog);
        }
    };
}

// Auto-save round data to localStorage
function autoSaveRound() {
    const roundData = {
        courses,
        events,
        eventDays,
        instructorUnavailable,
        unavailabilityMap,
        assignments,
        schedule,
        rooms,
        timestamp: new Date().toISOString()
    };
    localStorage.setItem('schedulepro_autosave', JSON.stringify(roundData));
}

// Load round data from localStorage
function loadRoundData() {
    const saved = localStorage.getItem('schedulepro_autosave');
    if (saved) {
        try {
            const data = JSON.parse(saved);
            courses = data.courses || [];
            events = data.events || [];
            eventDays = data.eventDays || [];
            instructorUnavailable = data.instructorUnavailable || [];
            unavailabilityMap = data.unavailabilityMap || {};
            assignments = data.assignments || {};
            schedule = data.schedule || {};
            rooms = data.rooms || {};
            
            // Re-render if data was loaded
            if (courses.length > 0 || events.length > 0) {
                renderAssignmentGrid();
                renderSwimlanes();
                updateStats();
                updateConfigureDaysButton();
            }
            return true; // Data was loaded
        } catch (e) {
            console.error('Failed to load saved round:', e);
        }
    }
    return false; // No data found
}

// Get current timestamp
function getTimestamp() {
    const now = new Date();
    return now.toLocaleString('en-US', { 
        year: 'numeric', 
        month: '2-digit', 
        day: '2-digit', 
        hour: '2-digit', 
        minute: '2-digit',
        hour12: true 
    });
}

// Get date range for a course placement
function getDateRange(eventId, dayNumbers) {
    if (!dayNumbers || dayNumbers.length === 0) return { firstDay: '', lastDay: '' };
    
    const days = eventDays.filter(d => d.Event_ID === eventId);
    const firstDayNum = Math.min(...dayNumbers);
    const lastDayNum = Math.max(...dayNumbers);
    
    const firstDay = days.find(d => parseInt(d.Day_Number) === firstDayNum)?.Day_Date || '';
    const lastDay = days.find(d => parseInt(d.Day_Number) === lastDayNum)?.Day_Date || '';
    
    return { firstDay, lastDay };
}

// Log a change action
function logChange(action, courseId, eventId, newDays, oldDays = null, notes = '') {
    const course = courses.find(c => c.Course_ID === courseId);
    if (!course) return;
    
    const newDates = getDateRange(eventId, newDays);
    const oldDates = oldDays ? getDateRange(eventId, oldDays) : { firstDay: '', lastDay: '' };
    
    const entry = {
        timestamp: getTimestamp(),
        action: action,
        courseId: courseId,
        courseTitle: course.Course_Name,
        instructor: course.Instructor,
        eventId: eventId || '',
        firstDay: newDates.firstDay,
        lastDay: newDates.lastDay,
        oldFirstDay: oldDates.firstDay,
        oldLastDay: oldDates.lastDay,
        notes: notes,
        csiTicket: ''
    };
    
    changeLog.push(entry);
    saveLogs();
}

// Log an upload action
function logUpload(uploadType, fileName, recordsCount, status, notes = '') {
    const entry = {
        timestamp: getTimestamp(),
        uploadType: uploadType,
        fileName: fileName,
        recordsCount: recordsCount,
        status: status,
        notes: notes
    };
    
    uploadsLog.push(entry);
    saveLogs();
}

// Log an error
function logError(errorType, courseId, courseTitle, firstDay, lastDay, errorMessage, sourceFile = '') {
    const entry = {
        timestamp: getTimestamp(),
        errorType: errorType,
        courseId: courseId || '',
        courseTitle: courseTitle || '',
        firstDay: firstDay || '',
        lastDay: lastDay || '',
        errorMessage: errorMessage,
        sourceFile: sourceFile
    };
    
    errorsLog.push(entry);
    saveLogs();
}

// Load events from consolidated Dates.csv
async function loadEvents() {
    try {
        const datesData = await fetchCSV('Data/Dates.csv');
        const parsedDates = parseCSV(datesData);
        
        // Process consolidated dates file
        processConsolidatedDates(parsedDates);
        
        renderAssignmentGrid();
        updateStats();
    } catch (error) {
        console.error('Error loading events:', error);
        // Fallback to old format if Dates.csv doesn't exist
        try {
            const eventsData = await fetchCSV('Data/events.csv');
            const daysData = await fetchCSV('Data/event_days.csv');
            events = parseCSV(eventsData);
            eventDays = parseCSV(daysData);
            renderAssignmentGrid();
            updateStats();
        } catch (fallbackError) {
            console.error('Error loading fallback events:', fallbackError);
        }
    }
}

// Helper: Process consolidated dates CSV into events[] and eventDays[] arrays
function processConsolidatedDates(datesData) {
    events = [];
    eventDays = [];
    
    datesData.forEach(row => {
        // Handle both CSV strings and Excel data (which may have Date objects, numbers, etc)
        const eventId = typeof row.Event_ID === 'string' ? row.Event_ID.trim() : String(row.Event_ID || '').trim();
        const eventName = typeof row.Event === 'string' ? row.Event.trim() : String(row.Event || '').trim();
        
        // Handle dates - Excel returns Date objects, CSV returns strings
        let firstDay = row.First_Event_Day;
        let lastDay = row.Last_Event_Day;
        
        if (firstDay instanceof Date) {
            // Already a Date object from Excel
        } else if (typeof firstDay === 'string') {
            firstDay = firstDay.trim();
        } else {
            firstDay = String(firstDay || '').trim();
        }
        
        if (lastDay instanceof Date) {
            // Already a Date object from Excel
        } else if (typeof lastDay === 'string') {
            lastDay = lastDay.trim();
        } else {
            lastDay = String(lastDay || '').trim();
        }
        
        // Only require the 4 essential fields users know when starting fresh
        if (!eventId || !eventName || !firstDay || !lastDay) {
            console.warn('Skipping row - missing required field:', { eventId, eventName, firstDay, lastDay });
            return;
        }
        
        // Auto-generate event days from First_Event_Day to Last_Event_Day
        const startDate = new Date(firstDay);
        const endDate = new Date(lastDay);
        
        // Validate dates - skip this event entirely if dates are invalid
        if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
            console.warn(`Skipping event ${eventId}: Invalid dates ${firstDay} to ${lastDay}`);
            return;
        }
        
        // Auto-calculate Total_Days from date range if not provided
        const daysDiff = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1; // +1 to include both start and end
        const totalDays = row.Total_Days && !isNaN(parseInt(row.Total_Days)) ? parseInt(row.Total_Days) : daysDiff;
        
        // Create event entry AFTER validating dates
        // Handle optional fields from both CSV and Excel
        const hotelLocation = typeof row.Hotel_Location === 'string' ? row.Hotel_Location.trim() : String(row.Hotel_Location || '').trim();
        const earlyBirdDate = typeof row.EarlyBird_End_Date === 'string' ? row.EarlyBird_End_Date.trim() : String(row.EarlyBird_End_Date || '').trim();
        const notes = typeof row.Notes === 'string' ? row.Notes.trim() : String(row.Notes || '').trim();
        
        events.push({
            Event_ID: eventId,
            Event: eventName,
            Total_Days: totalDays,
            Hotel_Location: hotelLocation,
            EarlyBird_End_Date: earlyBirdDate,
            Notes: notes
        });
        
        let currentDate = new Date(startDate);
        let dayNumber = 1;
        
        while (currentDate <= endDate && dayNumber <= totalDays) {
            eventDays.push({
                Event_ID: eventId,
                Event_Name: eventName,
                Day_Number: dayNumber,
                Day_Date: formatDate(currentDate)
            });
            
            currentDate.setDate(currentDate.getDate() + 1);
            dayNumber++;
        }
    });
    
    console.log(`Processed ${events.length} events, generated ${eventDays.length} event days`);
}

// Helper: Format date as M/D/YYYY
function formatDate(date) {
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();
    return `${month}/${day}/${year}`;
}

// Fetch CSV file
async function fetchCSV(url) {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`Failed to load ${url}`);
    return await response.text();
}

// Parse CSV data
function parseCSV(csvText) {
    const lines = csvText.trim().split('\n');
    const headers = parseCSVLine(lines[0]).map(h => h.trim());
    
    return lines.slice(1).map(line => {
        const values = parseCSVLine(line);
        const obj = {};
        headers.forEach((header, index) => {
            // Handle undefined (missing columns) and empty strings properly
            const value = values[index];
            obj[header] = (value !== undefined && value !== null) ? value.trim() : '';
        });
        return obj;
    });
}

// Parse a single CSV line (handles quoted values with commas and escaped quotes)
function parseCSVLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        const nextChar = line[i + 1];
        
        if (char === '"' && inQuotes && nextChar === '"') {
            // Escaped quote - add one quote to result
            current += '"';
            i++; // Skip next quote
        } else if (char === '"') {
            // Toggle quote mode
            inQuotes = !inQuotes;
        } else if (char === ',' && !inQuotes) {
            // Field delimiter
            result.push(current);
            current = '';
        } else {
            current += char;
        }
    }
    result.push(current);
    return result;
}

// Setup file input for courses
function setupFileInput() {
    const fileInput = document.getElementById('coursesFile');
    fileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        const text = await file.text();
        courses = parseCSV(text);
        
        // Validate courses
        if (courses.length === 0) {
            alert('No courses found in the file');
            return;
        }
        
        // Check for required columns
        const requiredColumns = ['Course_ID', 'Instructor', 'Course_Name', 'Duration_Days', 'Topic'];
        const hasAllColumns = requiredColumns.every(col => col in courses[0]);
        
        if (!hasAllColumns) {
            alert('CSV must have columns: Course_ID, Instructor, Course_Name, Duration_Days, Topic');
            return;
        }
        
        renderAssignmentGrid();
        updateStats();
        
        // Log upload
        logUpload('Courses', file.name, courses.length, 'Success');
        saveLogs();
        autoSaveRound();
        
        // Reset file input
        fileInput.value = '';
    });
}

// Setup file input for instructor unavailability
function setupUnavailabilityFileInput() {
    const fileInput = document.getElementById('unavailableFile');
    fileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        const text = await file.text();
        instructorUnavailable = parseCSV(text);
        
        // Validate
        if (instructorUnavailable.length === 0) {
            alert('No unavailability data found in the file');
            return;
        }
        
        const requiredColumns = ['Instructor', 'Unavailable_Start', 'Unavailable_End'];
        const hasAllColumns = requiredColumns.every(col => col in instructorUnavailable[0]);
        
        if (!hasAllColumns) {
            alert('CSV must have columns: Instructor, Unavailable_Start, Unavailable_End');
            return;
        }
        
        // Pre-calculate unavailability map
        calculateUnavailabilityMap();
        
        // Re-render grid to show constraints
        renderAssignmentGrid();
        
        autoSaveRound();
        
        alert(`Loaded unavailability for ${instructorUnavailable.length} entries`);
        
        // Reset file input
        fileInput.value = '';
    });
}

// Setup Excel file input (all-in-one import)
function setupExcelFileInput() {
    const fileInput = document.getElementById('excelFile');
    if (!fileInput) return;
    
    fileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        try {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data, { type: 'array' });
            
            let successCount = 0;
            const errors = [];
            
            // Process Dates tab
            if (workbook.SheetNames.includes('Dates')) {
                try {
                    const sheet = workbook.Sheets['Dates'];
                    // Use cellDates to get proper Date objects, then we'll convert them
                    const jsonData = XLSX.utils.sheet_to_json(sheet, { 
                        cellDates: true,
                        raw: false 
                    });
                    console.log('Excel Dates data:', jsonData);
                    processConsolidatedDates(jsonData);
                    logUpload('Excel: Dates', `${events.length} events loaded from Excel`);
                    successCount++;
                } catch (err) {
                    console.error('Dates tab error:', err);
                    errors.push(`Dates tab: ${err.message}`);
                }
            } else {
                errors.push('Dates tab: Tab not found in Excel file');
            }
            
            // Process Courses tab
            if (workbook.SheetNames.includes('Courses')) {
                try {
                    const sheet = workbook.Sheets['Courses'];
                    // Keep raw: true to preserve numbers like Duration_Days
                    const rawCourses = XLSX.utils.sheet_to_json(sheet, { raw: true });
                    
                    // Convert to proper format ensuring Course_ID is string
                    courses = rawCourses.map(course => ({
                        Course_ID: String(course.Course_ID || ''),
                        Course_Name: String(course.Course_Name || ''),
                        Instructor: String(course.Instructor || ''),
                        Duration_Days: course.Duration_Days // Keep as number
                    }));
                    
                    console.log('Excel Courses data:', courses.length, 'courses');
                    logUpload('Excel: Courses', `${courses.length} courses loaded from Excel`);
                    successCount++;
                } catch (err) {
                    console.error('Courses tab error:', err);
                    errors.push(`Courses tab: ${err.message}`);
                }
            } else {
                errors.push('Courses tab: Tab not found in Excel file');
            }
            
            // Process Instructor_Away tab
            if (workbook.SheetNames.includes('Instructor_Away')) {
                try {
                    const sheet = workbook.Sheets['Instructor_Away'];
                    // Use cellDates for proper date handling
                    const rawData = XLSX.utils.sheet_to_json(sheet, { 
                        cellDates: true,
                        raw: false 
                    });
                    
                    // Format dates properly
                    instructorUnavailable = rawData.map(record => ({
                        Instructor: String(record.Instructor || ''),
                        Unavailable_Start: record.Unavailable_Start,
                        Unavailable_End: record.Unavailable_End
                    }));
                    
                    console.log('Excel Instructor_Away data:', instructorUnavailable.length, 'records');
                    calculateUnavailabilityMap();
                    logUpload('Excel: Instructor_Away', `${instructorUnavailable.length} unavailability records loaded from Excel`);
                    successCount++;
                } catch (err) {
                    console.error('Instructor_Away tab error:', err);
                    errors.push(`Instructor_Away tab: ${err.message}`);
                }
            } else {
                errors.push('Instructor_Away tab: Tab not found in Excel file');
            }
            
            // Update UI
            renderAssignmentGrid();
            updateStats();
            updateConfigureDaysButton();
            autoSaveRound();
            
            // Show results
            let message = `Excel import completed:\n✓ ${successCount} tab(s) imported successfully`;
            if (errors.length > 0) {
                message += `\n\n⚠ ${errors.length} error(s):\n${errors.join('\n')}`;
            }
            alert(message);
            
        } catch (error) {
            alert('Error reading Excel file: ' + error.message);
            logError('Excel Import', '', '', '', '', error.message, file.name);
        }
        
        fileInput.value = '';
    });
}

// Pre-calculate which days instructors are unavailable for each event
function calculateUnavailabilityMap() {
    unavailabilityMap = {};
    
    // For each instructor's unavailability period
    instructorUnavailable.forEach(unavail => {
        const instructor = unavail.Instructor.trim();
        const startDate = new Date(unavail.Unavailable_Start);
        const endDate = new Date(unavail.Unavailable_End);
        
        // For each event
        events.forEach(event => {
            const eventId = event.Event_ID;
            const eventName = event.Event;
            const totalDays = parseInt(event['Total_Days']);
            
            // Get the actual dates for this event - match by Event_ID for accuracy
            const days = eventDays.filter(d => d.Event_ID === eventId);
            
            const blockedDays = [];
            
            // Check each day of the event
            days.forEach((day, index) => {
                if (!day.Day_Date) return; // Skip if no date
                const dayDate = new Date(day.Day_Date);
                
                // If this day falls within unavailability period
                if (dayDate >= startDate && dayDate <= endDate) {
                    blockedDays.push(index + 1); // Day numbers are 1-indexed
                }
            });
            
            // Store the blocked days if any
            if (blockedDays.length > 0) {
                const key = `${instructor}-${eventId}`;
                if (!unavailabilityMap[key]) {
                    unavailabilityMap[key] = [];
                }
                unavailabilityMap[key].push(...blockedDays);
                // Remove duplicates and sort
                unavailabilityMap[key] = [...new Set(unavailabilityMap[key])].sort((a, b) => a - b);
            }
        });
    });
    
    console.log('Unavailability Map:', unavailabilityMap);
}

// Get blocked days for an instructor at a specific event
function getBlockedDays(instructor, eventId) {
    const key = `${instructor}-${eventId}`;
    return unavailabilityMap[key] || [];
}

// Check if instructor has enough available days for a course
function hasEnoughAvailableDays(instructor, eventId, courseDuration) {
    const event = events.find(e => e.Event_ID === eventId);
    if (!event) return false;
    
    const totalDays = parseInt(event['Total_Days']);
    const blockedDays = getBlockedDays(instructor, eventId);
    const availableDays = totalDays - blockedDays.length;
    const daysNeeded = Math.ceil(parseFloat(courseDuration));
    
    return availableDays >= daysNeeded;
}

// Render assignment grid
function renderAssignmentGrid() {
    const table = document.getElementById('assignmentGrid');
    const thead = table.querySelector('thead tr');
    const tbody = table.querySelector('tbody');
    
    console.log('renderAssignmentGrid called - events:', events.length, 'courses:', courses.length);
    
    // Clear existing content except first header
    while (thead.children.length > 1) {
        thead.removeChild(thead.lastChild);
    }
    tbody.innerHTML = '';
    
    if (events.length === 0 || courses.length === 0) {
        console.warn('Grid render stopped - missing data');
        tbody.innerHTML = '<tr><td colspan="13" style="text-align: center; padding: 40px;">Load courses and events to begin</td></tr>';
        return;
    }
    
    // Add event headers
    events.forEach(event => {
        const th = document.createElement('th');
        th.className = 'event-header';
        
        // Get month from first event day
        const eventFirstDay = eventDays.find(d => d.Event_ID === event.Event_ID);
        let monthStr = '';
        if (eventFirstDay && eventFirstDay.Day_Date) {
            const date = new Date(eventFirstDay.Day_Date);
            const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
            monthStr = months[date.getMonth()];
        }
        
        th.innerHTML = `${event.Event}<br><small>${monthStr}</small>`;
        th.title = `${event.Event} (${event['Total_Days']} days)`;
        thead.appendChild(th);
    });
    
    // Add course rows
    courses.forEach(course => {
        const tr = document.createElement('tr');
        
        // Course info cell
        const tdCourse = document.createElement('td');
        tdCourse.className = 'course-info';
        tdCourse.innerHTML = `
            <div class="course-name-cell">
                ${course.Instructor} - 
                <span contenteditable="true" 
                      data-course-id="${course.Course_ID}"
                      style="outline: none; cursor: text; border-bottom: 1px dashed transparent;"
                      onblur="updateCourseName('${course.Course_ID}', this.textContent)"
                      onfocus="this.style.borderBottom='1px dashed #667eea'"
                      onmouseout="if(document.activeElement !== this) this.style.borderBottom='1px dashed transparent'"
                      onmouseover="this.style.borderBottom='1px dashed #ccc'"
                      title="Click to edit course name">${course.Course_Name}</span>
            </div>
            <div class="course-duration-cell">📏 ${course.Duration_Days} days</div>
        `;
        tr.appendChild(tdCourse);
        
        // Event checkboxes
        events.forEach(event => {
            const td = document.createElement('td');
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.dataset.courseId = course.Course_ID;
            checkbox.dataset.eventId = event.Event_ID;
            
            // Check unavailability
            const blockedDays = getBlockedDays(course.Instructor, event.Event_ID);
            const hasEnough = hasEnoughAvailableDays(course.Instructor, event.Event_ID, course.Duration_Days);
            
            // Check if already assigned
            if (assignments[course.Course_ID]?.includes(event.Event_ID)) {
                checkbox.checked = true;
            }
            
            // If instructor doesn't have enough available days, disable checkbox
            if (!hasEnough && blockedDays.length > 0) {
                checkbox.disabled = true;
                td.classList.add('unavailable-cell');
                const totalDays = parseInt(event['Total_Days']);
                const availableDays = totalDays - blockedDays.length;
                td.title = `${course.Instructor} unavailable ${blockedDays.length} days (${availableDays} available, needs ${Math.ceil(parseFloat(course.Duration_Days))})`;
                
                const icon = document.createElement('span');
                icon.className = 'unavailable-icon';
                icon.textContent = '⚠️';
                icon.title = td.title;
                td.appendChild(icon);
            } else if (blockedDays.length > 0) {
                // Has some blocked days but course still fits
                td.classList.add('unavailable-cell');
                const totalDays = parseInt(event['Total_Days']);
                const availableDays = totalDays - blockedDays.length;
                td.title = `${course.Instructor} unavailable days ${blockedDays.join(', ')} (${availableDays} days available)`;
                
                const icon = document.createElement('span');
                icon.className = 'unavailable-icon';
                icon.textContent = 'ℹ️';
                icon.title = td.title;
                td.appendChild(icon);
            }
            
            checkbox.addEventListener('change', (e) => {
                handleAssignmentChange(course.Course_ID, event.Event_ID, e.target.checked);
            });
            
            td.appendChild(checkbox);
            tr.appendChild(td);
        });
        
        tbody.appendChild(tr);
    });
}

// Handle assignment checkbox change
function handleAssignmentChange(courseId, eventId, isChecked) {
    if (!assignments[courseId]) {
        assignments[courseId] = [];
    }
    
    if (isChecked) {
        if (!assignments[courseId].includes(eventId)) {
            assignments[courseId].push(eventId);
            // Log ADD with no days yet
            logChange('ADD', courseId, eventId, null);
        }
    } else {
        assignments[courseId] = assignments[courseId].filter(id => id !== eventId);
        // Also remove from schedule if configured
        const hadDays = schedule[eventId] && schedule[eventId][courseId];
        const oldDays = hadDays ? schedule[eventId][courseId].days : null;
        
        if (hadDays) {
            delete schedule[eventId][courseId];
        }
        
        // Log REMOVE with old days if configured
        logChange('REMOVE', courseId, eventId, null, oldDays);
    }
    
    updateStats();
    updateConfigureDaysButton();
    saveLogs();
    autoSaveRound();
}

// Update course name globally
function updateCourseName(courseId, newName) {
    const trimmedName = newName.trim();
    
    // Find the course
    const course = courses.find(c => c.Course_ID === courseId);
    if (!course) return;
    
    const oldName = course.Course_Name;
    
    // No change
    if (trimmedName === oldName || !trimmedName) {
        // Restore original name if empty
        renderAssignmentGrid();
        return;
    }
    
    // Update course name globally
    course.Course_Name = trimmedName;
    
    // Log the change with proper format
    const entry = {
        timestamp: getTimestamp(),
        action: 'Course Name Edit',
        courseId: courseId,
        courseTitle: trimmedName,
        instructor: course.Instructor,
        eventId: '',
        firstDay: '',
        lastDay: '',
        oldFirstDay: '',
        oldLastDay: '',
        notes: `Renamed from "${oldName}" to "${trimmedName}"`,
        csiTicket: ''
    };
    
    changeLog.push(entry);
    saveLogs();
    
    // Re-render to show updated name everywhere
    renderAssignmentGrid();
    renderSwimlanes();
    updateReports();
    autoSaveRound();
}

// Update configure days button state
function updateConfigureDaysButton() {
    const btn = document.getElementById('configureDaysBtn');
    // Always enable - users can add courses via dropdown in Configure Days view
    const hasData = events.length > 0 && courses.length > 0;
    btn.disabled = !hasData;
}

// Toggle stats details panel
function toggleStatsDetails() {
    const detailsPanel = document.getElementById('statsDetails');
    const isExpanded = detailsPanel.classList.contains('expanded');
    
    if (isExpanded) {
        detailsPanel.classList.remove('expanded');
    } else {
        detailsPanel.classList.add('expanded');
        populateStatsDetails();
    }
}

// Complete reset - clear all data and start fresh
function completeReset() {
    console.log('completeReset function called');
    
    const userConfirmed = confirm('⚠️ WARNING: This will completely reset SchedulePro and clear ALL data including:\n\n• All course assignments\n• All day configurations\n• All saved rounds\n• All change logs\n• All uploads\n\nYou will need to re-upload your CSV files to start over.\n\nAre you sure you want to proceed?');
    
    console.log('User confirmation result:', userConfirmed);
    
    if (userConfirmed === true) {
        console.log('Reset confirmed - clearing data...');
        
        try {
            // Clear all data structures
            courses = [];
            events = [];
            eventDays = [];
            instructorUnavailable = [];
            unavailabilityMap = {};
            assignments = {};
            schedule = {};
            rooms = {};
            changeLog = [];
            uploadsLog = [];
            errorsLog = [];
            
            console.log('Data structures cleared');
            
            // Clear localStorage
            localStorage.removeItem('schedulepro_round');
            localStorage.removeItem('schedulepro_changelog');
            localStorage.removeItem('schedulepro_uploads');
            localStorage.removeItem('schedulepro_errors');
            
            console.log('localStorage cleared');
            console.log('Reloading page...');
            
            // Force reload from server, not cache
            window.location.href = window.location.href;
        } catch (error) {
            console.error('Error during reset:', error);
        }
    } else {
        console.log('Reset cancelled by user');
    }
}

// Populate stats details with unfilled days
function populateStatsDetails() {
    
    // Get events with unfilled days
    const eventsWithUnfilledDays = [];
    events.forEach(event => {
        const eventId = event.Event_ID;
        const totalDays = parseInt(event['Total_Days']);
        let filledDays = 0;
        
        for (let day = 1; day <= totalDays; day++) {
            let hasCourse = false;
            if (schedule[eventId]) {
                for (const courseId in schedule[eventId]) {
                    const placement = schedule[eventId][courseId];
                    if (placement && placement.days && Array.isArray(placement.days) && placement.days.includes(day)) {
                        hasCourse = true;
                        break;
                    }
                }
            }
            if (hasCourse) {
                filledDays++;
            }
        }
        
        const unfilledDays = totalDays - filledDays;
        if (unfilledDays > 0) {
            eventsWithUnfilledDays.push({
                event: event.Event,
                filled: filledDays,
                total: totalDays,
                unfilled: unfilledDays
            });
        }
    });
    
    const unfilledList = document.getElementById('unfilledDaysList');
    if (eventsWithUnfilledDays.length === 0) {
        unfilledList.innerHTML = '<li style="color: #28a745; border-left-color: #28a745;">✅ All event days are filled!</li>';
    } else {
        unfilledList.innerHTML = eventsWithUnfilledDays.map(item => {
            return `<li class="event-unfilled">
                <span><strong>${item.event}</strong></span>
                <span style="color: #6c757d;">${item.filled}/${item.total} days filled (${item.unfilled} empty)</span>
            </li>`;
        }).join('');
    }
}

// Update statistics
function updateStats() {
    const totalCourses = courses.length;
    document.getElementById('totalCourses').textContent = totalCourses;
    
    if (totalCourses === 0) {
        document.getElementById('percentAssigned').textContent = '0%';
        document.getElementById('percentDaysFilled').textContent = '0%';
        document.getElementById('avgCoursesPerEvent').textContent = '0';
        document.getElementById('avgCoursesPerDay').textContent = '0';
        return;
    }
    
    // 1. Percent of courses assigned to events
    const assignedCount = Object.keys(assignments).filter(courseId => {
        return assignments[courseId] && assignments[courseId].length > 0;
    }).length;
    const percentAssigned = totalCourses > 0 ? Math.round((assignedCount / totalCourses) * 100) : 0;
    document.getElementById('percentAssigned').textContent = percentAssigned + '%';
    
    // 2. Percent of event days with at least one course
    let totalEventDays = 0;
    let daysWithCourses = 0;
    
    events.forEach(event => {
        const eventId = event.Event_ID;
        const totalDays = parseInt(event['Total_Days']);
        totalEventDays += totalDays;
        
        for (let day = 1; day <= totalDays; day++) {
            let hasCourse = false;
            if (schedule[eventId]) {
                for (const courseId in schedule[eventId]) {
                    const placement = schedule[eventId][courseId];
                    // Check if placement exists, has days array, and includes this day number
                    if (placement && placement.days && Array.isArray(placement.days) && placement.days.includes(day)) {
                        hasCourse = true;
                        break;
                    }
                }
            }
            if (hasCourse) {
                daysWithCourses++;
            }
        }
    });
    
    const percentDaysFilled = totalEventDays > 0 ? Math.round((daysWithCourses / totalEventDays) * 100) : 0;
    document.getElementById('percentDaysFilled').textContent = `${percentDaysFilled}% (${daysWithCourses}/${totalEventDays})`;
    
    // 3. Average courses per event
    let totalAssignments = 0;
    events.forEach(event => {
        const eventId = event.Event_ID;
        if (schedule[eventId]) {
            totalAssignments += Object.keys(schedule[eventId]).length;
        }
    });
    const avgPerEvent = events.length > 0 ? (totalAssignments / events.length).toFixed(1) : 0;
    document.getElementById('avgCoursesPerEvent').textContent = avgPerEvent;
    
    // 4. Average courses per day (across all event days)
    const avgPerDay = totalEventDays > 0 ? (totalAssignments / totalEventDays).toFixed(2) : 0;
    document.getElementById('avgCoursesPerDay').textContent = avgPerDay;
}

// Go to configure days view
function goToConfigureDays() {
    document.getElementById('gridView').classList.remove('active');
    document.getElementById('configureDaysView').classList.add('active');
    renderSwimlanes();
    updateReports(); // Update reports when entering Configure Days view
}

// Back to grid view
function backToGrid() {
    document.getElementById('configureDaysView').classList.remove('active');
    document.getElementById('gridView').classList.add('active');
    
    // Ensure Step 2 section is expanded when returning from configure days
    const step2Content = document.getElementById('step2Content');
    const step2Toggle = document.getElementById('step2Toggle');
    if (step2Content && step2Toggle) {
        step2Content.classList.remove('collapsed');
        step2Toggle.textContent = '▼ Collapse Section';
    }
}

// Render swimlanes for day configuration
function renderSwimlanes() {
    const container = document.getElementById('swimlanesContainer');
    
    // Save current expanded/collapsed state
    const expandedState = {};
    document.querySelectorAll('.event-swimlane-body').forEach(body => {
        const eventId = body.id.replace('body-', '');
        expandedState[eventId] = !body.classList.contains('collapsed');
    });
    
    container.innerHTML = '';
    
    events.forEach(event => {
        const eventId = event.Event_ID;
        const eventName = event.Event;
        const totalDays = parseInt(event['Total_Days']);
        
        // Get courses assigned to this event
        const assignedCourses = courses.filter(course => 
            assignments[course.Course_ID]?.includes(eventId)
        );
        
        // Check if any courses in this event have conflicts
        let hasEventConflict = false;
        assignedCourses.forEach(course => {
            const placement = schedule[eventId]?.[course.Course_ID];
            if (placement?.startDay) {
                const blockedDays = getBlockedDays(course.Instructor, eventId);
                const daysNeeded = Math.ceil(parseFloat(course.Duration_Days));
                const courseDays = [];
                for (let i = 0; i < daysNeeded; i++) {
                    courseDays.push(placement.startDay + i);
                }
                if (courseDays.some(day => blockedDays.includes(day))) {
                    hasEventConflict = true;
                }
            }
        });
        
        // Show all events, even if no courses assigned (removed return statement)
        
        // Get days for this event
        const days = eventDays.filter(d => d.Event_ID === eventId);
        
        // Create swimlane
        const swimlane = document.createElement('div');
        swimlane.className = 'event-swimlane';
        swimlane.dataset.eventId = eventId;
        
        // Get month from first event day
        const eventFirstDay = days[0];
        let monthStr = '';
        if (eventFirstDay && eventFirstDay.Day_Date) {
            const date = new Date(eventFirstDay.Day_Date);
            const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
            monthStr = months[date.getMonth()];
        }
        
        // Determine if this event should be expanded (restore previous state or default collapsed)
        const isExpanded = expandedState[eventId] || false;
        const bodyClass = isExpanded ? 'event-swimlane-body' : 'event-swimlane-body collapsed';
        const toggleText = isExpanded ? '▼ Collapse' : '▶ Expand';
        
        // Add conflict indicator to header if conflicts exist
        const conflictIndicator = hasEventConflict ? '<span style="color: #ff9800; font-size: 1.5em; margin-left: 10px;">●●</span>' : '';
        
        swimlane.innerHTML = `
            <div class="event-swimlane-header" onclick="toggleEventSwimlane('${eventId}')">
                <span>${eventName} ${totalDays} days • ${monthStr}${conflictIndicator}</span>
                <span id="toggle-${eventId}">${toggleText}</span>
            </div>
            <div class="${bodyClass}" id="body-${eventId}">
                <div class="day-timeline" data-event-id="${eventId}">
                    ${days.map((day, index) => `
                        <div class="day-slot" data-day-num="${day.Day_Number}">
                            <div class="day-label">Day ${day.Day_Number}</div>
                            <div class="day-date">${day.Day_Date || ''}</div>
                        </div>
                    `).join('')}
                </div>
                ${assignedCourses.map(course => renderCourseSwimlane(course, eventId, totalDays)).join('')}
                <div class="add-course-section">
                    <select class="add-course-dropdown" id="add-course-${eventId}" onchange="addCourseToEvent('${eventId}', this.value, this)">
                        <option value="">+ Add Course to Event...</option>
                    </select>
                </div>
            </div>
        `;
        
        container.appendChild(swimlane);
        
        // Populate the dropdown with available courses
        populateCourseDropdown(eventId, assignedCourses);
    });
    
    // Setup drag and drop for all course blocks
    setupDragAndDrop();
}

// Populate course dropdown for an event
function populateCourseDropdown(eventId, assignedCourses) {
    const dropdown = document.getElementById(`add-course-${eventId}`);
    if (!dropdown) return;
    
    // Get IDs of courses already assigned to this event
    const assignedIds = new Set(assignedCourses.map(c => c.Course_ID));
    
    // Get event info for validation
    const event = events.find(e => e.Event_ID === eventId);
    const totalDays = event ? parseInt(event.Total_Days) : 0;
    
    // Add options for unassigned courses
    courses.forEach(course => {
        if (!assignedIds.has(course.Course_ID)) {
            const option = document.createElement('option');
            option.value = course.Course_ID;
            
            const duration = parseFloat(course.Duration_Days);
            const daysNeeded = Math.ceil(duration);
            
            // Check if course fits in event
            const warningIcon = daysNeeded > totalDays ? '⚠️ ' : '';
            
            // Check instructor availability
            const blockedDays = getBlockedDays(course.Instructor, eventId);
            const availableDays = totalDays - blockedDays.length;
            const availWarning = availableDays < daysNeeded ? '❌ ' : '';
            
            option.textContent = `${warningIcon}${availWarning}${course.Instructor} - ${course.Course_Name} (${course.Duration_Days} days)`;
            dropdown.appendChild(option);
        }
    });
}

// Add a course to an event from dropdown
function addCourseToEvent(eventId, courseId, selectElement) {
    if (!courseId) return; // User selected the placeholder option
    
    // Call existing assignment handler
    handleAssignmentChange(courseId, eventId, true);
    
    // Reset dropdown to placeholder
    selectElement.value = '';
    
    // Re-render swimlanes to show the new course
    renderSwimlanes();
    updateReports(); // Update reports when course is added
}

// Render a single course swimlane
function renderCourseSwimlane(course, eventId, totalDays) {
    const courseId = course.Course_ID;
    const duration = parseFloat(course.Duration_Days);
    const daysNeeded = Math.ceil(duration);
    
    // Get blocked days for this instructor at this event
    const blockedDays = getBlockedDays(course.Instructor, eventId);
    
    // Get current placement if exists
    const placement = schedule[eventId]?.[courseId];
    const startDay = placement?.startDay;
    
    // Calculate block width as percentage
    const blockWidth = (100 / totalDays) * daysNeeded;
    const blockLeft = startDay ? ((startDay - 1) / totalDays) * 100 : null;
    
    // Check if placed course has conflict with instructor unavailability
    let hasConflict = false;
    if (startDay && blockedDays.length > 0) {
        const courseDays = [];
        for (let i = 0; i < daysNeeded; i++) {
            courseDays.push(startDay + i);
        }
        hasConflict = courseDays.some(day => blockedDays.includes(day));
    }
    
    // Generate unavailability warning if any
    let unavailWarning = '';
    if (blockedDays.length > 0) {
        unavailWarning = `<div style="color: #dc3545; font-size: 0.85em; margin-top: 3px;">⚠️ Unavailable: Days ${blockedDays.join(', ')}</div>`;
    }
    
    // Get current room assignment
    const currentRoom = rooms[eventId]?.[courseId] || '';
    
    return `
        <div class="course-swimlane" data-course-id="${courseId}" data-event-id="${eventId}">
            <div class="course-info-sidebar">
                <div class="course-info-name">${course.Course_Name}</div>
                <div class="course-info-instructor">${course.Instructor}</div>
                <div class="course-info-duration">📏 ${course.Duration_Days} days</div>
                <div style="margin-top: 5px; display: flex; align-items: center; gap: 5px; font-size: 0.85em;">
                    <span>🏠 Room:</span>
                    <input type="text" 
                           value="${currentRoom}" 
                           placeholder="Room"
                           style="width: 80px; padding: 3px 6px; border: 1px solid #ccc; border-radius: 4px; font-size: 0.85em;"
                           onchange="updateRoomAssignment('${eventId}', '${courseId}', this.value)"
                           onclick="event.stopPropagation()">
                </div>
                ${unavailWarning}
            </div>
            <div class="course-timeline" data-course-id="${courseId}" data-event-id="${eventId}" data-total-days="${totalDays}" data-instructor="${course.Instructor}" data-blocked-days="${blockedDays.join(',')}">
                <div class="course-block ${startDay ? '' : 'unplaced'} ${hasConflict ? 'has-conflict' : ''}" 
                     data-course-id="${courseId}"
                     data-event-id="${eventId}"
                     data-days-needed="${daysNeeded}"
                     draggable="true"
                     style="${startDay ? `position: absolute; left: ${blockLeft}%; width: ${blockWidth}%; top: 5px; height: 40px; line-height: 40px;` : ''}">
                    ${startDay ? `Days ${startDay}-${startDay + daysNeeded - 1}` : 'Drag to timeline'}
                </div>
            </div>
            <div class="course-actions">
                <button class="btn btn-danger btn-small" onclick="removeCourseFromEvent('${courseId}', '${eventId}')">
                    ✗ Remove
                </button>
            </div>
        </div>
    `;
}

// Setup drag and drop
function setupDragAndDrop() {
    const blocks = document.querySelectorAll('.course-block');
    const timelines = document.querySelectorAll('.course-timeline');
    
    blocks.forEach(block => {
        block.addEventListener('dragstart', handleBlockDragStart);
        block.addEventListener('dragend', handleBlockDragEnd);
    });
    
    timelines.forEach(timeline => {
        timeline.addEventListener('dragover', handleTimelineDragOver);
        timeline.addEventListener('dragleave', handleTimelineDragLeave);
        timeline.addEventListener('drop', handleTimelineDrop);
        
        // Conflicts are now shown on course blocks (yellow) instead of day slots (red hashing)
    });
}

// Drag handlers
function handleBlockDragStart(e) {
    draggedBlock = {
        courseId: e.target.dataset.courseId,
        eventId: e.target.dataset.eventId,
        daysNeeded: parseInt(e.target.dataset.daysNeeded)
    };
    e.target.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
}

function handleBlockDragEnd(e) {
    e.target.classList.remove('dragging');
    draggedBlock = null;
    // Remove any drop indicators
    document.querySelectorAll('.drop-indicator').forEach(el => el.remove());
}

function handleTimelineDragOver(e) {
    e.preventDefault();
    
    if (!draggedBlock) return;
    
    // Only allow drop on the correct timeline
    const timelineCourseId = e.currentTarget.dataset.courseId;
    const timelineEventId = e.currentTarget.dataset.eventId;
    
    if (timelineCourseId !== draggedBlock.courseId || timelineEventId !== draggedBlock.eventId) {
        return;
    }
    
    e.dataTransfer.dropEffect = 'move';
    
    // Show drop indicator
    const timeline = e.currentTarget;
    const totalDays = parseInt(timeline.dataset.totalDays);
    const rect = timeline.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const dayWidth = rect.width / totalDays;
    
    // Calculate which day we're over
    let targetDay = Math.floor(x / dayWidth) + 1;
    
    // Calculate valid snap position
    const snapDay = calculateSnapPosition(targetDay, draggedBlock.daysNeeded, totalDays);
    
    if (snapDay !== null) {
        showDropIndicator(timeline, snapDay, draggedBlock.daysNeeded, totalDays);
    }
}

function handleTimelineDragLeave(e) {
    // Remove drop indicator when leaving timeline
    const relatedTarget = e.relatedTarget;
    if (!relatedTarget || !e.currentTarget.contains(relatedTarget)) {
        e.currentTarget.querySelectorAll('.drop-indicator').forEach(el => el.remove());
    }
}

function handleTimelineDrop(e) {
    e.preventDefault();
    
    if (!draggedBlock) return;
    
    const timeline = e.currentTarget;
    const timelineCourseId = timeline.dataset.courseId;
    const timelineEventId = timeline.dataset.eventId;
    
    if (timelineCourseId !== draggedBlock.courseId || timelineEventId !== draggedBlock.eventId) {
        return;
    }
    
    const totalDays = parseInt(timeline.dataset.totalDays);
    const rect = timeline.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const dayWidth = rect.width / totalDays;
    
    let targetDay = Math.floor(x / dayWidth) + 1;
    const snapDay = calculateSnapPosition(targetDay, draggedBlock.daysNeeded, totalDays);
    
    if (snapDay !== null) {
        // Get blocked days from timeline
        const blockedDaysStr = timeline.dataset.blockedDays || '';
        const blockedDays = blockedDaysStr.length > 0 ? blockedDaysStr.split(',').map(d => parseInt(d)) : [];
        
        // Check if any of the course days overlap with blocked days
        const courseDays = [];
        for (let i = snapDay; i < snapDay + draggedBlock.daysNeeded; i++) {
            courseDays.push(i);
        }
        
        const hasConflict = courseDays.some(day => blockedDays.includes(day));
        
        if (hasConflict) {
            alert(`Cannot place course on these days. Instructor unavailable on: ${blockedDays.filter(d => courseDays.includes(d)).join(', ')}`);
            return;
        }
        
        // Check if this is a new placement or a change
        const existingPlacement = schedule[draggedBlock.eventId]?.[draggedBlock.courseId];
        const oldDays = existingPlacement ? existingPlacement.days : null;
        const action = oldDays ? 'CHANGE' : 'ADD';
        
        // Save placement
        if (!schedule[draggedBlock.eventId]) {
            schedule[draggedBlock.eventId] = {};
        }
        
        schedule[draggedBlock.eventId][draggedBlock.courseId] = {
            startDay: snapDay,
            days: courseDays
        };
        
        // Log the change
        logChange(action, draggedBlock.courseId, draggedBlock.eventId, courseDays, oldDays);
        
        // Re-render this swimlane
        renderSwimlanes();
        updateStats();
        updateReports(); // Update reports after drag/drop
        saveLogs();
        autoSaveRound();
    }
}

// Calculate valid snap position
function calculateSnapPosition(targetDay, daysNeeded, totalDays) {
    // Ensure the course fits within the event
    if (targetDay < 1) targetDay = 1;
    
    // Calculate the last possible starting day
    const lastValidStart = totalDays - daysNeeded + 1;
    
    if (targetDay > lastValidStart) {
        targetDay = lastValidStart;
    }
    
    // If course is longer than event, return null (can't place)
    if (daysNeeded > totalDays) {
        return null;
    }
    
    return targetDay;
}

// Show drop indicator
function showDropIndicator(timeline, startDay, daysNeeded, totalDays) {
    // Remove existing indicators
    timeline.querySelectorAll('.drop-indicator').forEach(el => el.remove());
    
    const indicator = document.createElement('div');
    indicator.className = 'drop-indicator';
    
    const left = ((startDay - 1) / totalDays) * 100;
    const width = (daysNeeded / totalDays) * 100;
    
    indicator.style.left = `${left}%`;
    indicator.style.width = `${width}%`;
    
    timeline.appendChild(indicator);
}

// Update room assignment
function updateRoomAssignment(eventId, courseId, roomName) {
    // Initialize event object if needed
    if (!rooms[eventId]) {
        rooms[eventId] = {};
    }
    
    // Get old value for change log
    const oldRoom = rooms[eventId][courseId] || '';
    
    // Update room assignment
    rooms[eventId][courseId] = roomName.trim();
    
    // Log change
    const course = courses.find(c => c.Course_ID === courseId);
    const event = events.find(e => e.Event_ID === eventId);
    logChange(
        'Room Assignment',
        `${event?.Event || eventId}`,
        course?.Course_Name || courseId,
        oldRoom || '(none)',
        roomName.trim() || '(none)'
    );
    
    // Auto-save
    autoSaveRound();
}

// Remove course from event
function removeCourseFromEvent(courseId, eventId) {
    // Get old days before removing
    const oldDays = schedule[eventId]?.[courseId]?.days || null;
    
    // Remove from assignments
    if (assignments[courseId]) {
        assignments[courseId] = assignments[courseId].filter(id => id !== eventId);
    }
    
    // Remove from schedule
    if (schedule[eventId]) {
        delete schedule[eventId][courseId];
    }
    
    // Log the removal
    logChange('REMOVE', courseId, eventId, null, oldDays);
    
    // Update grid checkbox
    const checkbox = document.querySelector(`input[data-course-id="${courseId}"][data-event-id="${eventId}"]`);
    if (checkbox) {
        checkbox.checked = false;
    }
    
    // Re-render
    renderSwimlanes();
    updateStats();
    updateConfigureDaysButton();
}

// Helper function to escape CSV values
function escapeCSV(value) {
    if (value === null || value === undefined) return '';
    const str = String(value);
    // If contains comma, quote, newline, or carriage return, wrap in quotes and escape internal quotes
    if (str.includes(',') || str.includes('"') || str.includes('\n') || str.includes('\r')) {
        return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
}

// Export to Excel
function exportToExcel() {
    if (courses.length === 0) {
        alert('Please load courses first');
        return;
    }
    
    if (events.length === 0) {
        alert('Please load events first');
        return;
    }
    
    // Create CSV content
    let csv = 'Event_ID,Event,Day,Date,Course_ID,Instructor,Course_Name,Duration_Days,Room,Conflict,Configured\n';
    
    events.forEach(event => {
        const eventId = event.Event_ID;
        const eventName = event.Event;
        const totalDays = parseInt(event['Total_Days']);
        
        // Get days for this event - match by Event_ID
        const days = eventDays.filter(d => d.Event_ID === eventId)
                              .sort((a, b) => parseInt(a.Day_Number) - parseInt(b.Day_Number));
        
        days.forEach(day => {
            const dayNum = parseInt(day.Day_Number);
            
            // Find courses assigned to this day
            const coursesOnDay = [];
            if (schedule[eventId]) {
                for (const courseId in schedule[eventId]) {
                    const placement = schedule[eventId][courseId];
                    if (placement.days.includes(dayNum)) {
                        const course = courses.find(c => c.Course_ID === courseId);
                        if (course) {
                            coursesOnDay.push(course);
                        }
                    }
                }
            }
            
            if (coursesOnDay.length === 0) {
                csv += `${escapeCSV(eventId)},${escapeCSV(eventName)},${escapeCSV(dayNum)},${escapeCSV(day.Day_Date)},,,,,,No\n`;
            } else {
                coursesOnDay.forEach(course => {
                    const roomAssignment = rooms[eventId]?.[course.Course_ID] || '';
                    
                    // Check if instructor is unavailable on this day
                    const blockedDays = getBlockedDays(course.Instructor, eventId);
                    const hasConflict = blockedDays.includes(dayNum);
                    const conflictStatus = hasConflict ? 'YES - Instructor Unavailable' : '';
                    
                    csv += `${escapeCSV(eventId)},${escapeCSV(eventName)},${escapeCSV(dayNum)},${escapeCSV(day.Day_Date)},${escapeCSV(course.Course_ID)},${escapeCSV(course.Instructor)},${escapeCSV(course.Course_Name)},${escapeCSV(course.Duration_Days)},${escapeCSV(roomAssignment)},${escapeCSV(conflictStatus)},Yes\n`;
                });
            }
        });
    });
    
    // Add UTF-8 BOM for proper encoding in Excel
    const BOM = '\ufeff';
    const csvWithBOM = BOM + csv;
    
    // Download as CSV (Excel compatible)
    downloadFile(csvWithBOM, 'schedule_export.csv', 'text/csv;charset=utf-8');
    
    alert('Schedule exported successfully! Open in Excel to view.');
}

// Save schedule as JSON
function saveSchedule() {
    if (courses.length === 0) {
        alert('No data to save');
        return;
    }
    
    const data = {
        courses: courses,
        assignments: assignments,
        schedule: schedule,
        exportDate: new Date().toISOString()
    };
    
    const json = JSON.stringify(data, null, 2);
    downloadFile(json, 'schedule_backup.json', 'application/json');
}

// Load schedule from JSON
function loadSchedule() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    
    input.onchange = async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        try {
            const text = await file.text();
            const data = JSON.parse(text);
            
            courses = data.courses || [];
            assignments = data.assignments || {};
            schedule = data.schedule || {};
            
            renderAssignmentGrid();
            updateStats();
            updateConfigureDaysButton();
            
            alert('Schedule loaded successfully!');
        } catch (error) {
            alert('Error loading schedule: ' + error.message);
        }
    };
    
    input.click();
}

// Setup events file input
function setupEventsFileInput() {
    const fileInput = document.getElementById('eventsFile');
    fileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        const text = await file.text();
        const uploadedData = parseCSV(text);
        
        console.log('=== DATES CSV UPLOAD ===');
        console.log('Rows in CSV:', uploadedData.length);
        console.log('CSV data:', uploadedData);
        
        if (uploadedData.length === 0) {
            alert('No events found in the file');
            return;
        }
        
        // Check if this is the new consolidated format (Dates.csv)
        const hasConsolidatedFormat = 'First_Event_Day' in uploadedData[0] && 'Last_Event_Day' in uploadedData[0];
        
        if (hasConsolidatedFormat) {
            // New format: Process consolidated dates
            const requiredColumns = ['Event_ID', 'Event', 'First_Event_Day', 'Last_Event_Day'];
            const hasAllColumns = requiredColumns.every(col => col in uploadedData[0]);
            
            if (!hasAllColumns) {
                alert('CSV must have columns: Event_ID, Event, First_Event_Day, Last_Event_Day\n(Optional columns: Total_Days, Hotel_Location, EarlyBird_End_Date, Notes)');
                return;
            }
            
            console.log('Before processing - events:', events.length);
            processConsolidatedDates(uploadedData);
            console.log('After processing - events:', events.length, events);
            logUpload('Dates.csv', `${events.length} events with ${eventDays.length} days loaded`);
            alert(`Loaded ${events.length} events with ${eventDays.length} days auto-generated from date ranges`);
        } else {
            // Old format: Just events list
            const requiredColumns = ['Event_ID', 'Event', 'Total_Days'];
            const hasAllColumns = requiredColumns.every(col => col in uploadedData[0]);
            
            if (!hasAllColumns) {
                alert('CSV must have columns: Event_ID, Event, Total_Days');
                return;
            }
            
            events = uploadedData;
            logUpload('Events.csv', `${events.length} events loaded`);
            alert(`Loaded ${events.length} events (you may need to upload event_days.csv separately)`);
        }
        
        // Recalculate unavailability map if instructor unavailability was already loaded
        if (instructorUnavailable.length > 0) {
            calculateUnavailabilityMap();
        }
        
        renderAssignmentGrid();
        updateStats();
        updateConfigureDaysButton();
        autoSaveRound();
        
        fileInput.value = '';
    });
}

// Setup event days file input
function setupEventDaysFileInput() {
    const fileInput = document.getElementById('eventDaysFile');
    if (!fileInput) return; // Element doesn't exist (using consolidated Dates.csv)
    
    fileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        const text = await file.text();
        const newEventDays = parseCSV(text);
        
        if (newEventDays.length === 0) {
            alert('No event days found in the file');
            return;
        }
        
        const requiredColumns = ['Event_ID', 'Event_Name', 'Day_Number', 'Day_Date'];
        const hasAllColumns = requiredColumns.every(col => col in newEventDays[0]);
        
        if (!hasAllColumns) {
            alert('CSV must have columns: Event_ID, Event_Name, Day_Number, Day_Date');
            return;
        }
        
        eventDays = newEventDays;
        
        // Recalculate unavailability if it was already loaded
        if (instructorUnavailable.length > 0) {
            calculateUnavailabilityMap();
        }
        
        renderAssignmentGrid();
        updateStats();
        autoSaveRound();
        
        alert(`Loaded ${eventDays.length} event days`);
        fileInput.value = '';
    });
}

// Setup schedule file input for importing existing schedules
function setupScheduleFileInput() {
    const fileInput = document.getElementById('scheduleFile');
    fileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        const text = await file.text();
        const scheduleData = parseCSV(text);
        
        // Validate
        if (scheduleData.length === 0) {
            alert('No schedule data found in the file');
            return;
        }
        
        const requiredColumns = ['Course_ID', 'Duration_Days', 'First_Day', 'Last_Day'];
        const hasAllColumns = requiredColumns.every(col => col in scheduleData[0]);
        
        if (!hasAllColumns) {
            alert('CSV must have columns: Course_ID, Duration_Days, First_Day, Last_Day');
            return;
        }
        
        // Check if courses are loaded
        if (courses.length === 0) {
            alert('Please load courses CSV first before importing schedule');
            return;
        }
        
        // Process schedule import
        importSchedule(scheduleData, file.name);
        
        // Reset file input to allow re-uploading the same file
        fileInput.value = '';
    });
}

// Import schedule and auto-detect events based on dates
function importSchedule(scheduleData, fileName) {
    const errors = [];
    const imported = [];
    let successCount = 0;
    
    // Track unique courses to ensure they exist in courses list
    const courseIds = new Set(courses.map(c => c.Course_ID));
    
    scheduleData.forEach((row, index) => {
        const courseId = row.Course_ID.trim();
        const durationDays = parseFloat(row.Duration_Days);
        const firstDay = row.First_Day.trim();
        const lastDay = row.Last_Day.trim();
        
        // Validate course exists
        if (!courseIds.has(courseId)) {
            errors.push({
                Row: index + 2,
                Course_ID: courseId,
                First_Day: firstDay,
                Last_Day: lastDay,
                Error: 'Course_ID not found in courses list'
            });
            return;
        }
        
        // Parse dates - force local timezone to avoid date shifts
        // Split date string and create date with explicit year/month/day to avoid timezone issues
        const parseLocalDate = (dateStr) => {
            const parts = dateStr.split('-');
            if (parts.length === 3) {
                return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
            }
            // Try regular parsing if not YYYY-MM-DD format
            return new Date(dateStr);
        };
        
        const startDate = parseLocalDate(firstDay);
        const endDate = parseLocalDate(lastDay);
        
        if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
            errors.push({
                Row: index + 2,
                Course_ID: courseId,
                First_Day: firstDay,
                Last_Day: lastDay,
                Error: 'Invalid date format (use YYYY-MM-DD)'
            });
            return;
        }
        
        // Find matching event and day numbers
        const match = findEventForDateRange(startDate, endDate);
        
        if (!match) {
            // Add more detail about available event dates for debugging
            const eventDatesInfo = events.map(e => {
                const days = eventDays.filter(d => d.Event_ID === e.Event_ID);
                if (days.length > 0) {
                    return `${e.Event}: ${days[0].Day_Date} to ${days[days.length-1].Day_Date}`;
                }
                return `${e.Event}: (no days)`;
            }).join('; ');
            
            errors.push({
                Row: index + 2,
                Course_ID: courseId,
                First_Day: firstDay,
                Last_Day: lastDay,
                Error: `Dates ${firstDay} to ${lastDay} do not match any event. Available events: ${eventDatesInfo}`
            });
            return;
        }
        
        // Add to assignments and schedule
        if (!assignments[courseId]) {
            assignments[courseId] = [];
        }
        if (!assignments[courseId].includes(match.eventId)) {
            assignments[courseId].push(match.eventId);
        }
        
        if (!schedule[match.eventId]) {
            schedule[match.eventId] = {};
        }
        
        schedule[match.eventId][courseId] = {
            startDay: match.startDayNumber,
            days: match.dayNumbers
        };
        
        successCount++;
        imported.push({
            Course_ID: courseId,
            Event: match.eventName,
            Days: `${match.startDayNumber}-${match.endDayNumber}`
        });
    });
    
    // Log errors
    errors.forEach(err => {
        const course = courses.find(c => c.Course_ID === err.Course_ID);
        logError('Schedule Import', err.Course_ID, course?.Course_Name || '', err.First_Day, err.Last_Day, err.Error, fileName);
    });
    
    // Log upload summary
    const status = errors.length > 0 ? `Partial Success (${errors.length} errors)` : 'Success';
    logUpload('Schedule Import', fileName, successCount, status);
    saveLogs();
    
    // Update UI
    renderAssignmentGrid();
    updateStats();
    updateConfigureDaysButton();
    autoSaveRound();
    
    // Show results
    let message = `Successfully imported ${successCount} course assignment(s)`;
    
    if (errors.length > 0) {
        message += `\n\n${errors.length} error(s) found. Download error report?`;
        if (confirm(message)) {
            downloadErrorReport(errors);
        }
    } else {
        alert(message);
    }
}

// Find which event contains the given date range
function findEventForDateRange(startDate, endDate) {
    console.log('Looking for date range:', startDate, 'to', endDate);
    
    for (const event of events) {
        const eventId = event.Event_ID;
        const eventName = event.Event;
        
        // Get all days for this event
        const days = eventDays.filter(d => d.Event_ID === eventId)
                             .sort((a, b) => a.Day_Number - b.Day_Number);
        
        if (days.length === 0) continue;
        
        // Check if date range falls within this event
        const eventStart = new Date(days[0].Day_Date);
        const eventEnd = new Date(days[days.length - 1].Day_Date);
        
        console.log(`  Checking ${eventName}: ${days[0].Day_Date} (${eventStart}) to ${days[days.length - 1].Day_Date} (${eventEnd})`);
        
        // Normalize dates to compare by day only (ignore time)
        const normalizeDate = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate());
        const normStart = normalizeDate(startDate);
        const normEnd = normalizeDate(endDate);
        const normEventStart = normalizeDate(eventStart);
        const normEventEnd = normalizeDate(eventEnd);
        
        console.log(`    Normalized: looking for ${normStart.toISOString()} to ${normEnd.toISOString()}, event has ${normEventStart.toISOString()} to ${normEventEnd.toISOString()}`);
        
        if (normStart >= normEventStart && normEnd <= normEventEnd) {
            // Find day numbers by matching dates
            const dayNumbers = [];
            const startDayNumber = days.find(d => {
                const dayDate = normalizeDate(new Date(d.Day_Date));
                return dayDate.getTime() === normStart.getTime();
            })?.Day_Number;
            
            const endDayNumber = days.find(d => {
                const dayDate = normalizeDate(new Date(d.Day_Date));
                return dayDate.getTime() === normEnd.getTime();
            })?.Day_Number;
            
            if (startDayNumber && endDayNumber) {
                for (let i = parseInt(startDayNumber); i <= parseInt(endDayNumber); i++) {
                    dayNumbers.push(i);
                }
                
                return {
                    eventId,
                    eventName,
                    startDayNumber: parseInt(startDayNumber),
                    endDayNumber: parseInt(endDayNumber),
                    dayNumbers
                };
            }
        }
    }
    
    return null;
}

// Download error report CSV
function downloadErrorReport(errors) {
    const headers = 'Row,Course_ID,First_Day,Last_Day,Error\n';
    const rows = errors.map(e => 
        `${e.Row},${e.Course_ID},${e.First_Day},${e.Last_Day},"${e.Error}"`
    ).join('\n');
    
    // Use setTimeout to ensure download doesn't interfere with page state
    setTimeout(() => {
        downloadFile(headers + rows, 'schedule_import_errors.csv', 'text/csv');
    }, 100);
}

// Download schedule template
function downloadScheduleTemplate() {
    const template = `Course_ID,Duration_Days,First_Day,Last_Day
C001,3,2026-01-26,2026-01-28
C001,3,2026-02-23,2026-02-25
C002,2,2026-03-16,2026-03-17
C003,4,2026-04-13,2026-04-16`;
    
    downloadFile(template, 'schedule_template.csv', 'text/csv');
}

// Download template
function downloadTemplate() {
    const template = `Course_ID,Instructor,Course_Name,Duration_Days,Topic
C001,Alfred,Mapmaking,3,Geography
C002,Betty,Cooking Basics,2,Culinary Arts
C003,Charlie,Advanced Photography,4,Visual Arts
C004,Diana,Web Design,3.5,Technology
C005,Edward,Public Speaking,1,Communication`;
    
    downloadFile(template, 'courses_template.csv', 'text/csv');
}

// Download unavailability template
function downloadUnavailabilityTemplate() {
    const template = `Instructor,Unavailable_Start,Unavailable_End
Alfred,2026-06-15,2026-06-18
Betty,2026-02-23,2026-02-24
Charlie,2026-03-16,2026-03-19
Diana,2026-05-11,2026-05-14
Edward,2026-07-20,2026-07-23`;
    
    downloadFile(template, 'instructor_unavailable_template.csv', 'text/csv');
}

// Toggle collapsible section
function toggleSection(contentId) {
    const content = document.getElementById(contentId);
    const toggle = document.getElementById(contentId.replace('Content', 'Toggle'));
    
    const isCollapsed = content.classList.toggle('collapsed');
    toggle.classList.toggle('collapsed');
    
    // Update text
    toggle.textContent = isCollapsed ? '▶ Expand Section' : '▼ Collapse Section';
}

// Toggle event swimlane
function toggleEventSwimlane(eventId) {
    const body = document.getElementById(`body-${eventId}`);
    const toggle = document.getElementById(`toggle-${eventId}`);
    
    const isCollapsed = body.classList.toggle('collapsed');
    toggle.textContent = isCollapsed ? '▶ Expand' : '▼ Collapse';
}

// Toggle Step 1 help section
function toggleStep1Help() {
    const help = document.getElementById('step1Help');
    const toggle = document.getElementById('step1HelpToggle');
    
    const isCollapsed = help.classList.contains('collapsed');
    
    if (isCollapsed) {
        help.classList.remove('collapsed');
        help.style.maxHeight = '200px';
        toggle.textContent = '❌ Hide Help';
    } else {
        help.classList.add('collapsed');
        help.style.maxHeight = '0';
        toggle.textContent = '💡 Help';
    }
}

// Export all scheduling conflicts to CSV
function exportConflicts() {
    const conflicts = [];
    
    // Loop through all events and courses to find conflicts
    events.forEach(event => {
        const eventId = event.Event_ID;
        const eventName = event.Event;
        
        courses.forEach(course => {
            // Only check if course is assigned to this event
            if (!assignments[course.Course_ID]?.includes(eventId)) return;
            
            const placement = schedule[eventId]?.[course.Course_ID];
            if (!placement?.startDay) return; // Skip unplaced courses
            
            // Get blocked days for this instructor at this event
            const blockedDays = getBlockedDays(course.Instructor, eventId);
            if (blockedDays.length === 0) return; // No unavailability
            
            // Calculate course days
            const daysNeeded = Math.ceil(parseFloat(course.Duration_Days));
            const courseDays = [];
            for (let i = 0; i < daysNeeded; i++) {
                courseDays.push(placement.startDay + i);
            }
            
            // Check for conflicts
            const conflictDays = courseDays.filter(day => blockedDays.includes(day));
            if (conflictDays.length > 0) {
                // Get actual dates for conflict days
                const days = eventDays.filter(d => d.Event_ID === eventId);
                const conflictDates = conflictDays.map(dayNum => {
                    const day = days.find(d => parseInt(d.Day_Number) === dayNum);
                    return day ? day.Day_Date : `Day ${dayNum}`;
                }).join(', ');
                
                conflicts.push({
                    Event: eventName,
                    Event_ID: eventId,
                    Course_ID: course.Course_ID,
                    Course_Name: course.Course_Name,
                    Instructor: course.Instructor,
                    Scheduled_Days: `${placement.startDay}-${placement.startDay + daysNeeded - 1}`,
                    Conflict_Days: conflictDays.join(', '),
                    Conflict_Dates: conflictDates,
                    Issue: 'Instructor Unavailable'
                });
            }
        });
    });
    
    if (conflicts.length === 0) {
        alert('No conflicts found! All courses are scheduled without instructor unavailability issues.');
        return;
    }
    
    // Generate CSV
    const headers = ['Event', 'Event_ID', 'Course_ID', 'Course_Name', 'Instructor', 'Scheduled_Days', 'Conflict_Days', 'Conflict_Dates', 'Issue'];
    const csvRows = [headers.join(',')];
    
    conflicts.forEach(conflict => {
        const row = headers.map(header => {
            const value = conflict[header] || '';
            // Escape commas and quotes
            const escaped = String(value).replace(/"/g, '""');
            return escaped.includes(',') || escaped.includes('"') || escaped.includes('\n') ? `"${escaped}"` : escaped;
        });
        csvRows.push(row.join(','));
    });
    
    const csvContent = csvRows.join('\n');
    const timestamp = new Date().toISOString().split('T')[0];
    downloadFile(csvContent, `SchedulePro_Conflicts_${timestamp}.csv`, 'text/csv');
    
    alert(`Exported ${conflicts.length} conflict(s) to CSV`);
}

// Helper function to download a file
function downloadFile(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// Export change log to CSV
function exportChangelog() {
    if (changeLog.length === 0) {
        alert('No changes to export yet.');
        return;
    }
    
    const headers = ['Timestamp', 'Action', 'Course_ID', 'Course_Title', 'Instructor', 'Event_ID', 
                     'First_Day', 'Last_Day', 'Old_First_Day', 'Old_Last_Day', 'Notes', 'CSI_Ticket'];
    
    const rows = changeLog.map(entry => [
        entry.timestamp || '',
        entry.action || '',
        entry.courseId || '',
        entry.courseTitle || '',
        entry.instructor || '',
        entry.eventId || '',
        entry.firstDay || '',
        entry.lastDay || '',
        entry.oldFirstDay || '',
        entry.oldLastDay || '',
        entry.notes || '',
        entry.csiTicket || ''
    ]);
    
    const csv = [headers, ...rows].map(row => 
        row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
    ).join('\n');
    
    // Add UTF-8 BOM for proper encoding in Excel
    const BOM = '\ufeff';
    const csvWithBOM = BOM + csv;
    
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
    downloadFile(csvWithBOM, `SchedulePro_ChangeLog_${timestamp}.csv`, 'text/csv;charset=utf-8');
}

// Export uploads and errors to separate CSVs
function exportUploadsAndErrors() {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
    
    // Export uploads log
    if (uploadsLog.length > 0) {
        const uploadHeaders = ['Timestamp', 'Upload_Type', 'File_Name', 'Records_Count', 'Status', 'Notes'];
        const uploadRows = uploadsLog.map(entry => [
            entry.timestamp || '',
            entry.uploadType || '',
            entry.fileName || '',
            entry.recordsCount !== undefined ? entry.recordsCount : '',
            entry.status || '',
            entry.notes || ''
        ]);
        
        const uploadCsv = [uploadHeaders, ...uploadRows].map(row => 
            row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
        ).join('\n');
        
        // Add UTF-8 BOM for proper encoding in Excel
        const BOM = '\ufeff';
        downloadFile(BOM + uploadCsv, `SchedulePro_Uploads_${timestamp}.csv`, 'text/csv;charset=utf-8');
    } else {
        alert('No upload records to export.');
    }
    
    // Export errors log with slight delay to prevent download collision
    setTimeout(() => {
        if (errorsLog.length > 0) {
            const errorHeaders = ['Timestamp', 'Error_Type', 'Course_ID', 'Course_Title', 
                                 'First_Day', 'Last_Day', 'Error_Message', 'Source_File'];
            const errorRows = errorsLog.map(entry => [
                entry.timestamp || '',
                entry.errorType || '',
                entry.courseId || '',
                entry.courseTitle || '',
                entry.firstDay || '',
                entry.lastDay || '',
                entry.errorMessage || '',
                entry.sourceFile || ''
            ]);
            
            const errorCsv = [errorHeaders, ...errorRows].map(row => 
                row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
            ).join('\n');
            
            // Add UTF-8 BOM for proper encoding in Excel
            const BOM = '\ufeff';
            downloadFile(BOM + errorCsv, `SchedulePro_Errors_${timestamp}.csv`, 'text/csv;charset=utf-8');
        } else if (uploadsLog.length === 0) {
            // Only show alert if no uploads were exported either
            alert('No records to export.');
        }
    }, 100);
}

// Save Round - Manual save with timestamp
function saveRound() {
    if (courses.length === 0 && Object.keys(assignments).length === 0) {
        alert('Nothing to save yet. Upload courses and make assignments first.');
        return;
    }
    
    const roundData = {
        courses,
        events,
        eventDays,
        instructorUnavailable,
        unavailabilityMap,
        assignments,
        schedule,
        timestamp: new Date().toISOString(),
        savedAt: getTimestamp()
    };
    
    const json = JSON.stringify(roundData, null, 2);
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
    downloadFile(json, `SchedulePro_Round_${timestamp}.json`, 'application/json');
    
    alert('Round saved successfully! You can load this file later to continue working.');
}

// Load Round - Import saved round from JSON file
function loadRound() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    
    input.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        try {
            const text = await file.text();
            const data = JSON.parse(text);
            
            // Validate data structure
            if (!data.courses || !data.events) {
                alert('Invalid round file format.');
                return;
            }
            
            // Load all data
            courses = data.courses || [];
            events = data.events || [];
            eventDays = data.eventDays || [];
            instructorUnavailable = data.instructorUnavailable || [];
            unavailabilityMap = data.unavailabilityMap || {};
            assignments = data.assignments || {};
            schedule = data.schedule || {};
            
            // Re-render everything
            renderAssignmentGrid();
            renderSwimlanes();
            updateStats();
            updateConfigureDaysButton();
            
            // Auto-save to localStorage
            autoSaveRound();
            
            const savedAt = data.savedAt || 'Unknown date';
            alert(`Round loaded successfully!\n\nSaved at: ${savedAt}\nCourses: ${courses.length}\nEvents: ${events.length}`);
            
        } catch (error) {
            console.error('Load error:', error);
            alert('Failed to load round file. Please check the file format.');
        }
    });
    
    input.click();
}

// Toggle Reports section
function toggleReports() {
    const content = document.getElementById('reportsContent');
    const toggle = document.getElementById('reportsToggle');
    
    if (content.classList.contains('expanded')) {
        content.classList.remove('expanded');
        toggle.textContent = '▶ Expand';
    } else {
        content.classList.add('expanded');
        toggle.textContent = '▼ Collapse';
        updateReports();
    }
}

// Toggle Finances section
function toggleFinances() {
    const content = document.getElementById('financesContent');
    const toggle = document.getElementById('financesToggle');
    
    if (content.classList.contains('expanded')) {
        content.classList.remove('expanded');
        toggle.textContent = '▶ Expand';
    } else {
        content.classList.add('expanded');
        toggle.textContent = '▼ Collapse';
        updateFinances();
    }
}

// Update all reports
function updateReports() {
    updateInstructorWorkload();
    updateTopicCoverage();
    updateEventUtilization();
    updateTopicsPerEvent();
    updateFinances();
}

// Report 1: Instructor Workload (Revised)
function updateInstructorWorkload() {
    const container = document.getElementById('instructorWorkloadReport');
    
    // Collect data per instructor
    const instructorMap = {};
    
    courses.forEach(course => {
        const instructor = course.Instructor;
        const courseId = course.Course_ID;
        const duration = parseFloat(course.Duration_Days) || 0;
        const isAssigned = assignments[courseId] && assignments[courseId].length > 0;
        const eventCount = isAssigned ? assignments[courseId].length : 0;
        
        if (!instructorMap[instructor]) {
            instructorMap[instructor] = {
                uniqueCourses: 0,
                uniqueAssigned: 0,
                scheduledInstances: 0,
                totalDays: 0
            };
        }
        
        instructorMap[instructor].uniqueCourses++;
        
        if (isAssigned) {
            instructorMap[instructor].uniqueAssigned++;
            instructorMap[instructor].scheduledInstances += eventCount;
            instructorMap[instructor].totalDays += (duration * eventCount);
        }
    });
    
    // Sort by scheduled instances (descending)
    const sorted = Object.entries(instructorMap).sort((a, b) => b[1].scheduledInstances - a[1].scheduledInstances);
    
    if (sorted.length === 0) {
        container.innerHTML = '<p style="color: #6c757d;">No courses loaded yet.</p>';
        return;
    }
    
    let html = '<table class="report-table">';
    html += '<thead><tr>';
    html += '<th>Instructor</th>';
    html += '<th>Courses Scheduled (#)</th>';
    html += '<th>Unique Courses</th>';
    html += '<th>Total Days</th>';
    html += '<th>% Assigned</th>';
    html += '</tr></thead>';
    html += '<tbody>';
    
    sorted.forEach(([instructor, data]) => {
        const percentage = data.uniqueCourses > 0 ? Math.round((data.uniqueAssigned / data.uniqueCourses) * 100) : 0;
        
        html += `<tr>
            <td><strong>${instructor}</strong></td>
            <td>${data.scheduledInstances}</td>
            <td>${data.uniqueCourses}</td>
            <td>${Math.round(data.totalDays)}</td>
            <td>${percentage}%</td>
        </tr>`;
    });
    
    html += '</tbody></table>';
    container.innerHTML = html;
}

// Report 2: Topic Coverage (Revised)
function updateTopicCoverage() {
    const container = document.getElementById('topicCoverageReport');
    
    // Count courses per topic
    const topicMap = {};
    courses.forEach(course => {
        const topic = course.Topic || 'Uncategorized';
        const courseId = course.Course_ID;
        const isAssigned = assignments[courseId] && assignments[courseId].length > 0;
        
        if (!topicMap[topic]) {
            topicMap[topic] = { total: 0, assigned: 0 };
        }
        topicMap[topic].total++;
        if (isAssigned) {
            topicMap[topic].assigned++;
        }
    });
    
    // Sort by total courses (descending)
    const sorted = Object.entries(topicMap).sort((a, b) => b[1].total - a[1].total);
    
    if (sorted.length === 0) {
        container.innerHTML = '<p style="color: #6c757d;">No courses loaded yet.</p>';
        return;
    }
    
    let html = '<table class="report-table">';
    html += '<thead><tr><th>Topic</th><th>Total Courses in Topic</th><th>Offered Across Events</th></tr></thead>';
    html += '<tbody>';
    
    sorted.forEach(([topic, data]) => {
        html += `<tr>
            <td><strong>${topic}</strong></td>
            <td>${data.total}</td>
            <td>${data.assigned}</td>
        </tr>`;
    });
    
    html += '</tbody></table>';
    container.innerHTML = html;
}

// Report 3: Event Utilization (Revised with Instructor Count)
function updateEventUtilization() {
    const container = document.getElementById('eventUtilizationReport');
    
    if (events.length === 0) {
        container.innerHTML = '<p style="color: #6c757d;">No events loaded yet.</p>';
        return;
    }
    
    let html = '<table class="report-table">';
    html += '<thead><tr><th>Event</th><th>Courses</th><th>Days Filled</th><th>Days Empty</th><th># Instructors</th></tr></thead>';
    html += '<tbody>';
    
    events.forEach(event => {
        const eventId = event.Event_ID;
        const totalDays = parseInt(event['Total_Days']);
        
        // Get courses assigned to this event
        const assignedCourseIds = [];
        courses.forEach(course => {
            if (assignments[course.Course_ID] && assignments[course.Course_ID].includes(eventId)) {
                assignedCourseIds.push(course.Course_ID);
            }
        });
        
        const courseCount = assignedCourseIds.length;
        
        // Count unique instructors at this event
        const instructorsSet = new Set();
        assignedCourseIds.forEach(courseId => {
            const course = courses.find(c => c.Course_ID === courseId);
            if (course) {
                instructorsSet.add(course.Instructor);
            }
        });
        const instructorCount = instructorsSet.size;
        
        // Count filled days
        let filledDays = 0;
        for (let day = 1; day <= totalDays; day++) {
            let hasCourse = false;
            if (schedule[eventId]) {
                for (const courseId in schedule[eventId]) {
                    const placement = schedule[eventId][courseId];
                    if (placement && placement.days && placement.days.includes(day)) {
                        hasCourse = true;
                        break;
                    }
                }
            }
            if (hasCourse) filledDays++;
        }
        
        const emptyDays = totalDays - filledDays;
        
        html += `<tr>
            <td><strong>${event.Event}</strong></td>
            <td>${courseCount}</td>
            <td>${filledDays}</td>
            <td>${emptyDays}</td>
            <td>${instructorCount}</td>
        </tr>`;
    });
    
    html += '</tbody></table>';
    container.innerHTML = html;
}

// Update Finances Report
function updateFinances() {
    const container = document.getElementById('financesReport');
    
    if (!container) return; // Container not in DOM yet
    
    if (events.length === 0 || courses.length === 0) {
        container.innerHTML = '<p style="color: #6c757d;">No events or courses loaded yet.</p>';
        return;
    }
    
    // Pricing structure based on course duration
    function getCoursePrice(durationDays) {
        const duration = parseFloat(durationDays);
        if (duration <= 0.5) return 518;  // 4 hours
        if (duration <= 1) return 881;    // 1 day
        if (duration <= 2) return 1735;   // 2 days
        if (duration <= 3) return 2427;   // 3 days
        return 3228;                      // 4+ days
    }
    
    // Seat scenarios
    const scenarios = [
        { name: 'Low', seats: 10 },
        { name: 'Mid', seats: 20 },
        { name: 'High', seats: 30 }
    ];
    
    let html = '<table class="report-table"><thead><tr>';
    html += '<th>Event</th>';
    scenarios.forEach(scenario => {
        html += `<th>${scenario.name} (${scenario.seats} seats)</th>`;
    });
    html += '</tr></thead><tbody>';
    
    // Calculate revenue for each event
    events.forEach(event => {
        const eventId = event.Event_ID;
        
        // Find all courses assigned to this event
        const eventCourses = courses.filter(course => 
            assignments[course.Course_ID]?.includes(eventId)
        );
        
        // Calculate total revenue per scenario
        const scenarioRevenues = scenarios.map(scenario => {
            let totalRevenue = 0;
            eventCourses.forEach(course => {
                const price = getCoursePrice(course.Duration_Days);
                totalRevenue += price * scenario.seats;
            });
            return totalRevenue;
        });
        
        html += `<tr>
            <td><strong>${event.Event}</strong></td>`;
        
        scenarioRevenues.forEach(revenue => {
            html += `<td>$${revenue.toLocaleString()}</td>`;
        });
        
        html += '</tr>';
    });
    
    html += '</tbody></table>';
    container.innerHTML = html;
}

// Report 4: Topics per Event (Cross-Tab)
function updateTopicsPerEvent() {
    const container = document.getElementById('topicsPerEventReport');
    
    if (events.length === 0 || courses.length === 0) {
        container.innerHTML = '<p style="color: #6c757d;">No events or courses loaded yet.</p>';
        return;
    }
    
    // Get all unique topics
    const allTopics = [...new Set(courses.map(c => c.Topic || 'Uncategorized'))].sort();
    
    // Build cross-tab data
    const eventTopicData = [];
    
    events.forEach(event => {
        const eventId = event.Event_ID;
        const eventName = event.Event;
        
        // Count courses per topic for this event
        const topicCounts = {};
        allTopics.forEach(topic => {
            topicCounts[topic] = 0;
        });
        
        courses.forEach(course => {
            const courseId = course.Course_ID;
            const courseTopic = course.Topic || 'Uncategorized';
            
            // Check if this course is assigned to this event
            if (assignments[courseId] && assignments[courseId].includes(eventId)) {
                topicCounts[courseTopic]++;
            }
        });
        
        eventTopicData.push({
            eventName,
            topicCounts
        });
    });
    
    // Build table with horizontal scroll for many topics
    let html = '<div style="overflow-x: auto;">';
    html += '<table class="report-table" style="min-width: 600px;">';
    html += '<thead><tr><th style="position: sticky; left: 0; background: #f8f9fa; z-index: 1;">Event</th>';
    
    allTopics.forEach(topic => {
        html += `<th>${topic}</th>`;
    });
    
    html += '</tr></thead><tbody>';
    
    eventTopicData.forEach(row => {
        html += '<tr>';
        html += `<td style="position: sticky; left: 0; background: white; font-weight: 600;">${row.eventName}</td>`;
        
        allTopics.forEach(topic => {
            const count = row.topicCounts[topic];
            const cellStyle = count > 0 ? 'background: #e7f3ff;' : '';
            html += `<td style="text-align: center; ${cellStyle}">${count}</td>`;
        });
        
        html += '</tr>';
    });
    
    html += '</tbody></table></div>';
    container.innerHTML = html;
}
