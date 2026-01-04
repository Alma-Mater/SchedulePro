// SchedulePro V2 - Grid-based scheduler with swimlane day configuration
// Supabase client (initialized in HTML)
let supabaseDb = null;

// Data structures
let courses = [];
let events = [];
let eventDays = [];
let eventRooms = {}; // { eventId: numberOfRooms } - Number of rooms available per event
let lockedEvents = new Set(); // Set of eventIds that are locked from editing
let instructorUnavailable = []; // Raw unavailability data from CSV
let unavailabilityMap = {}; // Pre-calculated: { 'Instructor-EventID': [blockedDayNumbers] }
let assignments = {}; // { courseId: [eventIds] }
let schedule = {}; // { eventId: { courseId: { startDay, days: [], roomNumber: 1, isDraft: false } } }
let rooms = {}; // DEPRECATED - now using roomNumber in schedule
let draggedBlock = null;
let currentTimeline = null;

// Initialize Supabase reference from global
function initSupabase() {
    if (typeof window.supabaseClient !== 'undefined') {
        supabaseDb = window.supabaseClient;
    } else if (typeof supabaseClient !== 'undefined') {
        supabaseDb = supabaseClient;
    }
}

// Debounced auto-save (prevents too many saves)
let saveTimeout = null;
function triggerAutoSave() {
    clearTimeout(saveTimeout);
    saveTimeout = setTimeout(async () => {
        if (supabaseDb) {
            await saveRoundData();
        }
    }, 2000); // Save 2 seconds after last change
}

// Save all data to Supabase
async function saveRoundData() {
    if (!supabaseDb) {
        console.warn('⚠ Supabase not initialized - data will NOT persist on refresh!');
        return;
    }
    
    try {
        console.log('💾 Saving to Supabase...');
        
        // Save courses (delete all and re-insert)
        await supabaseDb.from('courses').delete().neq('id', 0);
        if (courses.length > 0) {
            const coursesData = courses.map(c => ({
                course_id: c.Course_ID,
                course_name: c.Course_Name,
                instructor: c.Instructor,
                duration_days: c.Duration_Days,
                topic: c.Topic || null
            }));
            const { error: coursesError } = await supabaseDb.from('courses').insert(coursesData);
            if (coursesError) {
                console.error('❌ Error saving courses:', coursesError);
                return;
            }
            console.log(`✓ Saved ${courses.length} courses`);
        }
        
        // Save schedule (delete all and re-insert)
        await supabaseDb.from('schedule').delete().neq('id', 0);
        const scheduleData = [];
        for (const eventId in schedule) {
            for (const courseId in schedule[eventId]) {
                const placement = schedule[eventId][courseId];
                scheduleData.push({
                    event_id: eventId,
                    course_id: courseId,
                    start_day: placement.startDay || null,
                    days: placement.days ? placement.days.join(',') : null,
                    room_number: placement.roomNumber || null,
                    is_draft: placement.isDraft || false,
                    updated_at: new Date().toISOString()
                });
            }
        }
        if (scheduleData.length > 0) {
            const { error: scheduleError } = await supabaseDb.from('schedule').insert(scheduleData);
            if (scheduleError) {
                console.error('❌ Error saving schedule:', scheduleError);
                return;
            }
            console.log(`✓ Saved ${scheduleData.length} schedule entries`);
        }
        
        // Update room counts in events table (update each event)
        for (const eventId in eventRooms) {
            const { error: roomError } = await supabaseDb
                .from('events')
                .update({ room_count: eventRooms[eventId] })
                .eq('event_id', eventId);
            
            if (roomError) console.error(`❌ Error updating room count for ${eventId}:`, roomError);
        }
        
        console.log('✅ All data saved to Supabase successfully!');
    } catch (error) {
        console.error('❌ Critical save error:', error);
        alert('Failed to save to Supabase. Your data may not persist on refresh. Check console for details.');
    }
}

// Three-button dialog helper
function showThreeButtonDialog(message, button1Text, button2Text, button3Text) {
    const overlay = document.createElement('div');
    overlay.style.cssText = 'position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); z-index: 10000; display: flex; align-items: center; justify-content: center;';
    
    const dialog = document.createElement('div');
    dialog.style.cssText = 'background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.3); max-width: 500px; width: 90%;';
    
    const messageDiv = document.createElement('div');
    messageDiv.style.cssText = 'white-space: pre-line; margin-bottom: 25px; line-height: 1.6; color: #333;';
    messageDiv.textContent = message;
    
    const buttonContainer = document.createElement('div');
    buttonContainer.style.cssText = 'display: flex; gap: 10px; justify-content: flex-end;';
    
    let result = 'cancel';
    
    const btn1 = document.createElement('button');
    btn1.textContent = button1Text;
    btn1.className = 'btn btn-primary';
    btn1.onclick = () => { result = 'replace'; overlay.remove(); };
    
    const btn2 = document.createElement('button');
    btn2.textContent = button2Text;
    btn2.className = 'btn btn-secondary';
    btn2.onclick = () => { result = 'cancel'; overlay.remove(); };
    
    const btn3 = document.createElement('button');
    btn3.textContent = button3Text;
    btn3.className = 'btn btn-warning';
    btn3.onclick = () => { result = 'draft'; overlay.remove(); };
    
    buttonContainer.appendChild(btn1);
    buttonContainer.appendChild(btn2);
    buttonContainer.appendChild(btn3);
    
    dialog.appendChild(messageDiv);
    dialog.appendChild(buttonContainer);
    overlay.appendChild(dialog);
    document.body.appendChild(overlay);
    
    // Wait for user interaction
    return new Promise(resolve => {
        const checkRemoved = setInterval(() => {
            if (!document.body.contains(overlay)) {
                clearInterval(checkRemoved);
                resolve(result);
            }
        }, 100);
    });
}

// Change logging
let changeLog = [];
let uploadsLog = [];
let errorsLog = [];

// Initialize
document.addEventListener('DOMContentLoaded', async () => {
    // Initialize Supabase connection
    initSupabase();
    
    loadLogs();
    
    // Setup file inputs first
    setupFileInput();
    setupRoomsFileInput();
    setupUnavailabilityFileInput();
    setupScheduleFileInput();
    setupEventsFileInput();
    setupExcelFileInput();
    // setupEventDaysFileInput(); // Commented out - element removed from HTML (using consolidated Dates.csv)
    
    // Load data from Supabase (async)
    await loadRoundData();
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

// Auto-save round data to Supabase (replaces localStorage)
async function autoSaveRound() {
    if (!supabaseDb) {
        console.warn('Supabase not initialized, skipping auto-save');
        return;
    }

    try {
        // Save events
        await supabaseDb.from('events').delete().neq('id', 0); // Clear existing
        if (events.length > 0) {
            const eventsData = events.map(e => ({
                event_id: e.Event_ID,
                event_name: e.Event,
                location: e.Location || null,
                room_count: e.Room_Count || null
            }));
            await supabaseDb.from('events').insert(eventsData);
        }

        // Save courses
        await supabaseDb.from('courses').delete().neq('id', 0);
        if (courses.length > 0) {
            const coursesData = courses.map(c => ({
                course_id: c.Course_ID,
                course_name: c.Course_Name,
                instructor: c.Instructor,
                duration_days: c.Duration_Days,
                topic: c.Topic || null
            }));
            await supabaseDb.from('courses').insert(coursesData);
        }

        // Save event days
        await supabaseDb.from('event_days').delete().neq('id', 0);
        if (eventDays.length > 0) {
            const daysData = eventDays.map(d => ({
                event_id: d.Event_ID,
                day_number: d.Day_Number,
                day_date: d.Day_Date
            }));
            await supabaseDb.from('event_days').insert(daysData);
        }

        // Save instructor unavailability
        await supabaseDb.from('instructor_unavailability').delete().neq('id', 0);
        if (instructorUnavailable.length > 0) {
            const unavailData = instructorUnavailable.map(u => ({
                instructor: u.Instructor,
                event_id: u.Event_ID,
                unavailable_days: u.Unavailable_Days || null
            }));
            await supabaseDb.from('instructor_unavailability').insert(unavailData);
        }

        // Save schedule
        await supabaseDb.from('schedule').delete().neq('id', 0);
        const scheduleData = [];
        for (const eventId in schedule) {
            for (const courseId in schedule[eventId]) {
                const placement = schedule[eventId][courseId];
                scheduleData.push({
                    event_id: eventId,
                    course_id: courseId,
                    start_day: placement.startDay || null,
                    days: placement.days ? placement.days.join(',') : null,
                    room_number: placement.roomNumber || null,
                    is_draft: placement.isDraft || false,
                    updated_at: new Date().toISOString()
                });
            }
        }
        if (scheduleData.length > 0) {
            await supabaseDb.from('schedule').insert(scheduleData);
        }

        console.log('Auto-save to Supabase completed');
    } catch (error) {
        console.error('Error auto-saving to Supabase:', error);
        alert('Failed to save data to cloud. Please check your connection.');
    }
}

// Load round data from Supabase
async function loadRoundData() {
    if (!supabaseDb) {
        console.warn('Supabase not initialized');
        return false;
    }

    try {
        // Load events
        const { data: eventsData, error: eventsError } = await supabaseDb.from('events').select('*');
        if (eventsError) throw eventsError;
        events = eventsData.map(e => ({
            Event_ID: e.event_id,
            Event: e.event_name,
            Location: e.location,
            Room_Count: e.room_count
        }));

        // Load courses
        const { data: coursesData, error: coursesError } = await supabaseDb.from('courses').select('*');
        if (coursesError) throw coursesError;
        courses = coursesData.map(c => ({
            Course_ID: c.course_id,
            Course_Name: c.course_name,
            Instructor: c.instructor,
            Duration_Days: c.duration_days,
            Topic: c.topic
        }));

        // Load event days
        const { data: daysData, error: daysError } = await supabaseDb.from('event_days').select('*');
        if (daysError) throw daysError;
        eventDays = daysData.map(d => ({
            Event_ID: d.event_id,
            Day_Number: d.day_number,
            Day_Date: d.day_date
        }));

        // Load instructor unavailability
        const { data: unavailData, error: unavailError } = await supabaseDb.from('instructor_unavailability').select('*');
        if (unavailError) throw unavailError;
        instructorUnavailable = unavailData.map(u => ({
            Instructor: u.instructor,
            Event_ID: u.event_id,
            Unavailable_Days: u.unavailable_days
        }));

        // Load schedule
        const { data: scheduleData, error: scheduleError } = await supabaseDb.from('schedule').select('*');
        if (scheduleError) throw scheduleError;
        
        schedule = {};
        scheduleData.forEach(s => {
            if (!schedule[s.event_id]) {
                schedule[s.event_id] = {};
            }
            schedule[s.event_id][s.course_id] = {
                startDay: s.start_day,
                days: s.days ? s.days.split(',').map(Number) : [],
                roomNumber: s.room_number,
                isDraft: s.is_draft
            };
        });

        // Rebuild derived data
        rebuildAssignments();
        rebuildUnavailabilityMap();
        calculateEventRooms();

        // Re-render if data was loaded
        if (courses.length > 0 && events.length > 0) {
            renderAssignmentGrid();
            renderSwimlanes();
            renderSwimlanesGrid(); // Render Room Grid view
            updateReportsGrid(); // Update Room Grid reports
            renderCoursesTableGrid(); // Update courses table
            updateStats();
            updateConfigureDaysButton();
        }

        console.log('Loaded data from Supabase');
        return true;
    } catch (error) {
        console.error('Error loading from Supabase:', error);
        return false;
    }
}

// Rebuild assignments from schedule
function rebuildAssignments() {
    assignments = {};
    for (const eventId in schedule) {
        for (const courseId in schedule[eventId]) {
            if (!assignments[courseId]) {
                assignments[courseId] = [];
            }
            if (!assignments[courseId].includes(eventId)) {
                assignments[courseId].push(eventId);
            }
        }
    }
}

// Rebuild unavailability map
function rebuildUnavailabilityMap() {
    unavailabilityMap = {};
    instructorUnavailable.forEach(entry => {
        const key = `${entry.Instructor}-${entry.Event_ID}`;
        const days = entry.Unavailable_Days ? entry.Unavailable_Days.split(',').map(d => parseInt(d.trim())) : [];
        unavailabilityMap[key] = days;
    });
}

// Calculate event rooms from events data
function calculateEventRooms() {
    eventRooms = {};
    events.forEach(event => {
        if (event.Room_Count) {
            eventRooms[event.Event_ID] = event.Room_Count;
        }
    });
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
    eventRooms = {}; // Reset room counts
    
    datesData.forEach(row => {
        // Handle both CSV strings and Excel data (which may have Date objects, numbers, etc)
        const eventId = typeof row.Event_ID === 'string' ? row.Event_ID.trim() : String(row.Event_ID || '').trim();
        const eventName = typeof row.Event === 'string' ? row.Event.trim() : String(row.Event || '').trim();
        
        // Extract Room_Count if provided (optional, defaults to 1)
        let roomCount = 1; // Default
        if (row.Room_Count !== undefined && row.Room_Count !== null && row.Room_Count !== '') {
            const parsed = parseInt(row.Room_Count);
            if (!isNaN(parsed) && parsed > 0) {
                roomCount = parsed;
            }
        }
        
        // Set room count for this event (first occurrence wins)
        if (!eventRooms[eventId]) {
            eventRooms[eventId] = roomCount;
        }
        
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
        const location = typeof row.Location === 'string' ? row.Location.trim() : String(row.Location || '').trim();
        
        events.push({
            Event_ID: eventId,
            Event: eventName,
            Total_Days: totalDays,
            Hotel_Location: hotelLocation,
            EarlyBird_End_Date: earlyBirdDate,
            Notes: notes,
            Location: location
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
        triggerAutoSave();
        
        // Reset file input
        fileInput.value = '';
    });
}

// Setup file input for rooms
function setupRoomsFileInput() {
    const fileInput = document.getElementById('roomsFile');
    fileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        const text = await file.text();
        const roomsData = parseCSV(text);
        
        // Validate
        if (roomsData.length === 0) {
            alert('No rooms data found in the file');
            return;
        }
        
        const requiredColumns = ['Event_ID', 'Number_of_Rooms'];
        const hasAllColumns = requiredColumns.every(col => col in roomsData[0]);
        
        if (!hasAllColumns) {
            alert('Rooms CSV must have columns: Event_ID, Number_of_Rooms');
            return;
        }
        
        // Convert to eventRooms object
        eventRooms = {};
        roomsData.forEach(row => {
            const eventId = row.Event_ID?.trim();
            const numRooms = parseInt(row.Number_of_Rooms);
            if (eventId && numRooms > 0) {
                eventRooms[eventId] = numRooms;
            }
        });
        
        console.log('Rooms loaded:', eventRooms);
        
        // Re-render if we already have data
        if (events.length > 0) {
            renderSwimlanes();
        }
        
        // Log upload
        logUpload('Rooms', file.name, Object.keys(eventRooms).length, 'Success');
        saveLogs();
        triggerAutoSave();
        
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
        
        triggerAutoSave();
        
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
            renderSwimlanesGrid(); // Update Room Grid view if active
            updateReportsGrid(); // Update Room Grid reports
            renderCoursesTableGrid(); // Update courses table
            
            // Save immediately to Supabase (don't wait for auto-save timeout)
            if (supabaseDb) {
                await saveRoundData();
            }
            
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
    triggerAutoSave();
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
    triggerAutoSave();
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
            // Remember current view
            const isRoomGridActive = document.getElementById('roomGridView')?.classList.contains('active');
            
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
            
            // Clear Supabase data if connected
            if (supabaseDb) {
                supabaseDb.from('events').delete().neq('id', 0);
                supabaseDb.from('courses').delete().neq('id', 0);
                supabaseDb.from('event_days').delete().neq('id', 0);
                supabaseDb.from('instructor_unavailability').delete().neq('id', 0);
                supabaseDb.from('schedule').delete().neq('id', 0);
            }
            
            // Update all views
            renderAssignmentGrid();
            updateStats();
            
            // Restore room grid view if it was active
            if (isRoomGridActive) {
                // Keep room grid view active
                document.getElementById('gridView').classList.remove('active');
                document.getElementById('configureDaysView').classList.remove('active');
                document.getElementById('roomCapacityView').classList.remove('active');
                document.getElementById('roomGridView').classList.add('active');
                renderSwimlanesGrid();
                updateReportsGrid();
                console.log('Staying on room grid view');
            } else {
                // Just refresh the current view
                renderSwimlanesGrid();
                updateReportsGrid();
                console.log('Reset complete, staying on current view');
            }
            
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
    
    // 1. Percent of courses assigned to events - COUNT ONLY courses in the courses array
    let assignedCount = 0;
    courses.forEach(course => {
        if (assignments[course.Course_ID] && assignments[course.Course_ID].length > 0) {
            assignedCount++;
        }
    });
    
    const percentAssigned = totalCourses > 0 ? Math.round((assignedCount / totalCourses) * 100) : 0;
    
    // Enhanced Debug: Show the mismatch
    console.log(`📊 ASSIGNMENT STATS:`);
    console.log(`   Total courses in array: ${totalCourses}`);
    console.log(`   Courses with assignments: ${assignedCount}`);
    console.log(`   Percentage: ${percentAssigned}%`);
    console.log(`   Keys in assignments object: ${Object.keys(assignments).length}`);
    
    // Check for ALL unassigned courses (both missing from object and empty arrays)
    const unassignedCourses = courses.filter(course => {
        return !assignments[course.Course_ID] || assignments[course.Course_ID].length === 0;
    });
    
    // Check for duplicate Course_IDs
    const courseIdCounts = {};
    courses.forEach(course => {
        courseIdCounts[course.Course_ID] = (courseIdCounts[course.Course_ID] || 0) + 1;
    });
    const duplicateIds = Object.keys(courseIdCounts).filter(id => courseIdCounts[id] > 1);
    
    if (duplicateIds.length > 0) {
        console.log(`⚠️ DUPLICATE COURSE IDs FOUND (${duplicateIds.length} IDs have duplicates):`);
        duplicateIds.forEach(id => {
            const duplicates = courses.filter(c => c.Course_ID === id);
            console.log(`   Course_ID "${id}" appears ${courseIdCounts[id]} times:`);
            console.table(duplicates.map(c => ({
                Course_ID: c.Course_ID,
                Instructor: c.Instructor,
                Course_Name: c.Course_Name,
                Duration: c.Duration_Days,
                Topic: c.Topic || ''
            })));
        });
    }
    
    if (unassignedCourses.length > 0) {
        console.log(`🚨 ${unassignedCourses.length} UNASSIGNED COURSES FOUND:`);
        console.table(unassignedCourses.map(c => ({ 
            Course_ID: c.Course_ID, 
            Instructor: c.Instructor,
            Course_Name: c.Course_Name,
            Duration: c.Duration_Days,
            Topic: c.Topic || '',
            Has_Assignment: !!assignments[c.Course_ID],
            Assignment_Length: assignments[c.Course_ID]?.length || 0
        })));
    } else {
        console.log('✅ All courses are assigned to at least one event');
    }
    
    // Check for assignments that don't have a matching course
    const assignmentsWithoutCourse = Object.keys(assignments).filter(courseId => {
        return !courses.find(c => c.Course_ID === courseId);
    });
    if (assignmentsWithoutCourse.length > 0) {
        console.log(`⚠️ ${assignmentsWithoutCourse.length} Course IDs in assignments but NOT in courses array:`, assignmentsWithoutCourse);
    }
    
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
    document.getElementById('roomCapacityView').classList.remove('active');
    document.getElementById('roomGridView').classList.remove('active');
    renderSwimlanes();
    updateReports(); // Update reports when entering Configure Days view
}

// Go to room grid view
function goToRoomGridView() {
    console.log('Switching to Room Grid View...');
    document.getElementById('gridView').classList.remove('active');
    document.getElementById('configureDaysView').classList.remove('active');
    document.getElementById('roomCapacityView').classList.remove('active');
    document.getElementById('roomGridView').classList.add('active');
    
    // Render with small delay to ensure DOM is ready
    setTimeout(() => {
        try {
            renderSwimlanesGrid();
            updateReportsGrid();
            renderCoursesTableGrid();
            console.log('Room Grid View rendered successfully');
        } catch (error) {
            console.error('Error rendering Room Grid View:', error);
        }
    }, 50);
}

// Go to room capacity view
function goToRoomCapacity() {
    document.getElementById('gridView').classList.remove('active');
    document.getElementById('configureDaysView').classList.remove('active');
    document.getElementById('roomCapacityView').classList.add('active');
    document.getElementById('roomGridView').classList.remove('active');
    renderRoomCapacity();
}

// Back to grid view
function backToGrid() {
    document.getElementById('configureDaysView').classList.remove('active');
    document.getElementById('roomCapacityView').classList.remove('active');
    document.getElementById('roomGridView').classList.remove('active');
    document.getElementById('gridView').classList.add('active');
    
    // Ensure Step 2 section is expanded when returning from configure days
    const step2Content = document.getElementById('step2Content');
    const step2Toggle = document.getElementById('step2Toggle');
    if (step2Content && step2Toggle) {
        step2Content.classList.remove('collapsed');
        step2Toggle.textContent = '▼ Collapse Section';
    }
}

// Render room capacity table
function renderRoomCapacity() {
    const container = document.getElementById('roomCapacityContainer');
    
    if (events.length === 0) {
        container.innerHTML = '<p style="color: #6c757d;">No events loaded yet.</p>';
        return;
    }
    
    // First, collect all available rooms
    const availableRooms = [];
    
    events.forEach(event => {
        const eventId = event.Event_ID;
        const eventName = event.Event;
        const totalDays = parseInt(event['Total_Days']);
        const numRooms = eventRooms[eventId] || 1;
        
        // Get days for this event
        const days = eventDays.filter(d => d.Event_ID === eventId).sort((a, b) => parseInt(a.Day_Number) - parseInt(b.Day_Number));
        
        // Analyze each room
        for (let roomNum = 1; roomNum <= numRooms; roomNum++) {
            // Track which days are occupied in this room
            const occupiedDays = new Set();
            
            if (schedule[eventId]) {
                for (const courseId in schedule[eventId]) {
                    const placement = schedule[eventId][courseId];
                    if (placement.roomNumber === roomNum && placement.days && placement.days.length > 0) {
                        placement.days.forEach(day => occupiedDays.add(day));
                    }
                }
            }
            
            const daysOccupied = occupiedDays.size;
            const daysAvailable = totalDays - daysOccupied;
            
            // Only include rooms with availability
            if (daysAvailable > 0) {
                // Build list of available dates
                const availableDatesList = [];
                for (let day = 1; day <= totalDays; day++) {
                    if (!occupiedDays.has(day)) {
                        const dayInfo = days.find(d => parseInt(d.Day_Number) === day);
                        if (dayInfo && dayInfo.Day_Date) {
                            const date = new Date(dayInfo.Day_Date);
                            const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
                            availableDatesList.push(`${months[date.getMonth()]} ${date.getDate()}`);
                        }
                    }
                }
                
                availableRooms.push({
                    eventName,
                    roomNum,
                    daysAvailable,
                    availableDates: availableDatesList.join(', ')
                });
            }
        }
    });
    
    // Build HTML starting with available rooms table
    let html = '';
    
    if (availableRooms.length > 0) {
        html += `
            <div style="margin-bottom: 30px;">
                <h3 style="margin-top: 0; color: #667eea; cursor: pointer; display: flex; align-items: center; gap: 8px;" 
                    onclick="document.getElementById('availableRoomsTable').style.display = document.getElementById('availableRoomsTable').style.display === 'none' ? 'block' : 'none'; this.querySelector('.toggle-icon').textContent = this.querySelector('.toggle-icon').textContent === '▼' ? '▶' : '▼';">
                    <span class="toggle-icon">▼</span>
                    🟢 Available Rooms (${availableRooms.length})
                </h3>
                <div id="availableRoomsTable">
                    <table class="report-table" style="width: 100%;">
                        <thead><tr><th>Event</th><th>Room</th><th>Days Available</th><th>Available Dates</th></tr></thead>
                        <tbody>`;
        
        availableRooms.forEach(room => {
            html += `<tr style="background: rgba(255, 152, 0, 0.15);">
                <td style="color: black; font-weight: 700;"><strong>${room.eventName}</strong></td>
                <td style="color: black; font-weight: 700;">Room ${room.roomNum}</td>
                <td style="text-align: center; font-weight: 700; color: black;">${room.daysAvailable}</td>
                <td style="color: black; font-weight: 700;">${room.availableDates}</td>
            </tr>`;
        });
        
        html += `
                        </tbody>
                    </table>
                </div>
            </div>`;
    } else {
        html += '<p style="color: #28a745; font-weight: bold; margin-bottom: 20px;">✅ All rooms are fully booked!</p>';
    }
    
    // Now add the full capacity overview
    html += '<h3 style="color: #667eea;">📊 Full Capacity Overview</h3>';
    html += '<table class="report-table" style="width: 100%;">';
    html += '<thead><tr><th>Event</th><th>Room</th><th>Total Days</th><th>Days Occupied</th><th>Days Available</th><th>Available Dates</th></tr></thead>';
    html += '<tbody>';
    
    events.forEach(event => {
        const eventId = event.Event_ID;
        const eventName = event.Event;
        const totalDays = parseInt(event['Total_Days']);
        const numRooms = eventRooms[eventId] || 1;
        
        // Get days for this event
        const days = eventDays.filter(d => d.Event_ID === eventId).sort((a, b) => parseInt(a.Day_Number) - parseInt(b.Day_Number));
        
        // Analyze each room
        for (let roomNum = 1; roomNum <= numRooms; roomNum++) {
            // Track which days are occupied in this room
            const occupiedDays = new Set();
            
            if (schedule[eventId]) {
                for (const courseId in schedule[eventId]) {
                    const placement = schedule[eventId][courseId];
                    if (placement.roomNumber === roomNum && placement.days && placement.days.length > 0) {
                        placement.days.forEach(day => occupiedDays.add(day));
                    }
                }
            }
            
            const daysOccupied = occupiedDays.size;
            const daysAvailable = totalDays - daysOccupied;
            
            // Build list of available dates
            const availableDatesList = [];
            for (let day = 1; day <= totalDays; day++) {
                if (!occupiedDays.has(day)) {
                    const dayInfo = days.find(d => parseInt(d.Day_Number) === day);
                    if (dayInfo && dayInfo.Day_Date) {
                        const date = new Date(dayInfo.Day_Date);
                        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
                        availableDatesList.push(`${months[date.getMonth()]} ${date.getDate()}`);
                    }
                }
            }
            
            const availableDatesStr = availableDatesList.length > 0 ? availableDatesList.join(', ') : 'Full';
            // Green = fully booked, faded orange = empty or partially booked
            const rowColor = daysAvailable === 0 ? 'background: #d4edda;' : 'background: rgba(255, 152, 0, 0.15);';
            
            html += `<tr style="${rowColor}">
                <td style="color: black; font-weight: 700;"><strong>${eventName}</strong></td>
                <td style="color: black; font-weight: 700;">Room ${roomNum}</td>
                <td style="text-align: center; color: black; font-weight: 700;">${totalDays}</td>
                <td style="text-align: center; color: black; font-weight: 700;">${daysOccupied}</td>
                <td style="text-align: center; font-weight: 700; color: black;">${daysAvailable}</td>
                <td style="color: black; font-weight: 700;">${availableDatesStr}</td>
            </tr>`;
        }
    });
    
    html += '</tbody></table>';
    container.innerHTML = html;
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
        
        // Initialize schedule entries for assigned courses that don't have placement yet
        assignedCourses.forEach(course => {
            if (!schedule[eventId]) {
                schedule[eventId] = {};
            }
            if (!schedule[eventId][course.Course_ID]) {
                schedule[eventId][course.Course_ID] = {
                    startDay: null,
                    days: [],
                    roomNumber: 1  // Default to Room 1
                };
            }
        });
        
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
        
        // Get date range from first and last event days
        const eventFirstDay = days[0];
        const eventLastDay = days[days.length - 1];
        let dateRangeStr = '';
        if (eventFirstDay && eventFirstDay.Day_Date) {
            const startDate = new Date(eventFirstDay.Day_Date);
            const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
            const monthName = months[startDate.getMonth()];
            const startDay = startDate.getDate();
            
            if (eventLastDay && eventLastDay.Day_Date) {
                const endDate = new Date(eventLastDay.Day_Date);
                const endDay = endDate.getDate();
                dateRangeStr = `${monthName} ${startDay}-${endDay}`;
            } else {
                dateRangeStr = `${monthName} ${startDay}`;
            }
        }
        
        // Determine if this event should be expanded (restore previous state or default collapsed)
        const isExpanded = expandedState[eventId] || false;
        const bodyClass = isExpanded ? 'event-swimlane-body' : 'event-swimlane-body collapsed';
        const toggleText = isExpanded ? '▼ Collapse' : '▶ Expand';
        
        // Add conflict indicator to header if conflicts exist
        const conflictIndicator = hasEventConflict ? '<span style="color: #ff9800; font-size: 1.5em; margin-left: 10px;">●●</span>' : '';
        
        // Get number of rooms for this event (default to 1 if not specified)
        const numRooms = eventRooms[eventId] || 1;
        
        // Calculate room utilization
        const occupiedRooms = new Set();
        if (schedule[eventId]) {
            for (const courseId in schedule[eventId]) {
                const placement = schedule[eventId][courseId];
                if (placement.roomNumber) {
                    occupiedRooms.add(placement.roomNumber);
                }
            }
        }
        const roomsOccupied = occupiedRooms.size;
        const roomsAvailable = numRooms - roomsOccupied;
        const roomUtilization = `<span style="color: #ff9800; font-weight: 700;">${roomsOccupied}/${numRooms}</span>`;
        
        // Build room lanes
        let roomLanesHTML = '';
        for (let roomNum = 1; roomNum <= numRooms; roomNum++) {
            // Get courses assigned to this room
            const roomCourses = assignedCourses.filter(course => {
                const placement = schedule[eventId]?.[course.Course_ID];
                return placement && (placement.roomNumber === roomNum || (!placement.roomNumber && roomNum === 1));
            });
            
            roomLanesHTML += `
                <div class="room-lane" data-event-id="${eventId}" data-room-number="${roomNum}">
                    <div class="room-lane-header">
                        <span>🏠 Room ${roomNum}</span>
                        <div class="room-actions">
                            ${roomNum === numRooms && numRooms > 1 ? `<button onclick="removeRoom('${eventId}', ${roomNum}); event.stopPropagation();" title="Remove this room">✗</button>` : ''}
                        </div>
                    </div>
                    <div class="room-timeline-container">
                        <div class="day-timeline" data-event-id="${eventId}" data-room-number="${roomNum}">
                            ${days.map((day, index) => `
                                <div class="day-slot" data-day-num="${day.Day_Number}">
                                    <div class="day-label">Day ${day.Day_Number}</div>
                                    <div class="day-date">${day.Day_Date || ''}</div>
                                </div>
                            `).join('')}
                        </div>
                        ${roomCourses.map(course => renderCourseSwimlane(course, eventId, totalDays, roomNum)).join('')}
                    </div>
                </div>
            `;
        }
        
        swimlane.innerHTML = `
            <div class="event-swimlane-header" onclick="toggleEventSwimlane('${eventId}')">
                <span>${eventName} ${totalDays} days • ${dateRangeStr} • Rooms: ${roomUtilization} 
                    <span onclick="event.stopPropagation(); editRoomCount('${eventId}', ${numRooms})" 
                          style="cursor: pointer; opacity: 0.8; padding: 0 3px;" 
                          title="Click to edit room count">✎</span>${conflictIndicator}
                </span>
                <span id="toggle-${eventId}">${toggleText}</span>
            </div>
            <div class="${bodyClass}" id="body-${eventId}">
                ${roomLanesHTML}
                <div class="add-course-section" style="display: flex; gap: 10px; justify-content: space-between; align-items: center;">
                    <button class="btn btn-secondary btn-small" onclick="addRoom('${eventId}')">➕ Add Room</button>
                    <select class="add-course-dropdown" id="add-course-${eventId}" onchange="addCourseToEvent('${eventId}', this.value, this)" style="flex: 1;">
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
    
    // Clear existing options except the placeholder
    while (dropdown.options.length > 1) {
        dropdown.remove(1);
    }
    
    // Get IDs of courses already assigned to this event
    const assignedIds = new Set(assignedCourses.map(c => c.Course_ID));
    
    // Debug logging
    console.log(`📋 Dropdown for ${eventId}: ${courses.length} total courses, ${assignedIds.size} already assigned`);
    
    // Get event info for validation
    const event = events.find(e => e.Event_ID === eventId);
    const totalDays = event ? parseInt(event.Total_Days) : 0;
    
    let addedCount = 0;
    const skippedCourses = [];
    
    // Add options for unassigned courses
    courses.forEach(course => {
        if (!assignedIds.has(course.Course_ID)) {
            addedCount++;
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
        } else {
            skippedCourses.push({
                ID: course.Course_ID,
                Name: course.Course_Name,
                Instructor: course.Instructor
            });
        }
    });
    
    console.log(`   ✅ Added ${addedCount} courses to dropdown`);
    
    if (skippedCourses.length > 0) {
        console.log(`   ⏭️ Skipped ${skippedCourses.length} already-assigned courses:`, skippedCourses);
    }
}

// Add a course to an event from dropdown
function addCourseToEvent(eventId, courseId, selectElement) {
    if (!courseId) return; // User selected the placeholder option
    
    // Call existing assignment handler
    handleAssignmentChange(courseId, eventId, true);
    
    // Initialize schedule entry with Room 1 (unplaced on timeline)
    if (!schedule[eventId]) {
        schedule[eventId] = {};
    }
    if (!schedule[eventId][courseId]) {
        schedule[eventId][courseId] = {
            startDay: null,
            days: [],
            roomNumber: 1,  // Default to Room 1
            isDraft: false
        };
    }
    
    // Reset dropdown to placeholder
    selectElement.value = '';
    
    // Re-render swimlanes to show the new course
    renderSwimlanes();
    updateReports(); // Update reports when course is added
    triggerAutoSave();
}

// Render a single course swimlane
function renderCourseSwimlane(course, eventId, totalDays, roomNumber = null) {
    const courseId = course.Course_ID;
    const duration = parseFloat(course.Duration_Days);
    const daysNeeded = Math.ceil(duration);
    
    // Get blocked days for this instructor at this event
    const blockedDays = getBlockedDays(course.Instructor, eventId);
    
    // Get current placement if exists
    const placement = schedule[eventId]?.[courseId];
    const startDay = placement?.startDay;
    const assignedRoom = placement?.roomNumber || roomNumber || 1;
    
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
    
    // Get number of rooms for this event to populate dropdown
    const numRooms = eventRooms[eventId] || 1;
    let roomOptions = '';
    for (let r = 1; r <= numRooms; r++) {
        roomOptions += `<option value="${r}" ${r === assignedRoom ? 'selected' : ''}>Room ${r}</option>`;
    }
    
    return `
        <div class="course-swimlane" data-course-id="${courseId}" data-event-id="${eventId}" data-room-number="${assignedRoom}">
            <div class="course-info-sidebar">
                <div class="course-info-name">${course.Course_Name}</div>
                <div class="course-info-instructor">${course.Instructor}</div>
                <div class="course-info-duration">📏 ${course.Duration_Days} days</div>
                <div style="margin-top: 8px; display: flex; align-items: center; gap: 5px;">
                    <span style="font-size: 0.9em; color: #667eea;">🏠</span>
                    <select onchange="changeCourseRoom('${eventId}', '${courseId}', parseInt(this.value))" 
                            style="flex: 1; padding: 5px; border: 2px solid #667eea; border-radius: 4px; font-size: 0.85em; cursor: pointer; background: white;"
                            onclick="event.stopPropagation()">
                        ${roomOptions}
                    </select>
                </div>
                ${unavailWarning}
            </div>
            <div class="course-timeline" data-course-id="${courseId}" data-event-id="${eventId}" data-room-number="${assignedRoom}" data-total-days="${totalDays}" data-instructor="${course.Instructor}" data-blocked-days="${blockedDays.join(',')}">
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
        
        // Get room number from the timeline being dropped onto
        const roomNumber = parseInt(timeline.dataset.roomNumber) || 1;
        
        // Check if this is a new placement or a change
        const existingPlacement = schedule[draggedBlock.eventId]?.[draggedBlock.courseId];
        const oldDays = existingPlacement ? existingPlacement.days : null;
        const oldRoom = existingPlacement ? existingPlacement.roomNumber : null;
        const action = oldDays ? 'CHANGE' : 'ADD';
        
        // Check for room conflicts - is another course already in this room on these days?
        const roomConflicts = [];
        if (schedule[draggedBlock.eventId]) {
            for (const otherCourseId in schedule[draggedBlock.eventId]) {
                if (otherCourseId === draggedBlock.courseId) continue; // Skip self
                const otherPlacement = schedule[draggedBlock.eventId][otherCourseId];
                if (otherPlacement.roomNumber === roomNumber) {
                    // Check for day overlap
                    const overlap = courseDays.some(day => otherPlacement.days.includes(day));
                    if (overlap) {
                        const otherCourse = courses.find(c => c.Course_ID === otherCourseId);
                        roomConflicts.push(otherCourse?.Course_Name || otherCourseId);
                    }
                }
            }
        }
        
        if (roomConflicts.length > 0) {
            alert(`⚠️ Room ${roomNumber} Conflict!\n\nThis room is already occupied on some of these days by:\n${roomConflicts.join('\n')}\n\nPlease choose a different room or days.`);
            return;
        }
        
        // Save placement
        if (!schedule[draggedBlock.eventId]) {
            schedule[draggedBlock.eventId] = {};
        }
        
        schedule[draggedBlock.eventId][draggedBlock.courseId] = {
            startDay: snapDay,
            days: courseDays,
            roomNumber: roomNumber
        };
        
        // Log the change
        logChange(action, draggedBlock.courseId, draggedBlock.eventId, courseDays, oldDays);
        
        // Re-render this swimlane
        renderSwimlanes();
        updateStats();
        updateReports(); // Update reports after drag/drop
        saveLogs();
        triggerAutoSave();
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
    triggerAutoSave();
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
    saveLogs();
    triggerAutoSave();
}

// Edit room count for an event
function editRoomCount(eventId, currentCount) {
    // Check if event is locked
    if (lockedEvents.has(eventId)) {
        alert('🔒 This event is locked. Unlock it to make changes.');
        return;
    }
    
    const newCount = prompt(`Set number of rooms for this event:\n(Current: ${currentCount})`, currentCount);
    if (newCount === null) return; // Cancelled
    
    const numRooms = parseInt(newCount);
    if (isNaN(numRooms) || numRooms < 1) {
        alert('Please enter a valid number (minimum 1)');
        return;
    }
    
    if (numRooms === currentCount) return; // No change
    
    // If reducing rooms, check for courses in rooms that will be removed
    if (numRooms < currentCount) {
        const affectedCourses = [];
        if (schedule[eventId]) {
            for (const courseId in schedule[eventId]) {
                const placement = schedule[eventId][courseId];
                if (placement.roomNumber > numRooms) {
                    const course = courses.find(c => c.Course_ID === courseId);
                    affectedCourses.push({
                        name: course?.Course_Name || courseId,
                        room: placement.roomNumber
                    });
                }
            }
        }
        
        if (affectedCourses.length > 0) {
            const details = affectedCourses.map(c => `  • ${c.name} (Room ${c.room})`).join('\n');
            const proceed = confirm(`⚠️ Warning: ${affectedCourses.length} course(s) are in rooms ${numRooms + 1}-${currentCount}:\n\n${details}\n\nThese courses will be unassigned. Continue?`);
            if (!proceed) return;
            
            // Remove courses from removed rooms
            for (const courseId in schedule[eventId]) {
                const placement = schedule[eventId][courseId];
                if (placement.roomNumber > numRooms) {
                    delete schedule[eventId][courseId];
                    // Also remove from assignments
                    if (assignments[courseId]) {
                        assignments[courseId] = assignments[courseId].filter(id => id !== eventId);
                    }
                }
            }
        }
    }
    
    // Update room count
    eventRooms[eventId] = numRooms;
    
    // Log the change
    const event = events.find(e => e.Event_ID === eventId);
    const eventName = event ? event.Event : eventId;
    logUpload('ROOM_COUNT', eventName, numRooms, 'Changed', `Room count changed from ${currentCount} to ${numRooms} for event "${eventName}"`);
    
    // Re-render the appropriate view
    const roomGridView = document.getElementById('roomGridView');
    const configureDaysView = document.getElementById('configureDaysView');
    
    if (roomGridView && roomGridView.classList.contains('active')) {
        renderSwimlanesGrid();
        updateReportsGrid();
    } else if (configureDaysView && configureDaysView.classList.contains('active')) {
        renderSwimlanes();
        updateReports();
    }
    
    updateStats();
    updateConfigureDaysButton();
    saveLogs();
    triggerAutoSave();
}

// Add a room to an event
function addRoom(eventId) {
    const currentRooms = eventRooms[eventId] || 1;
    eventRooms[eventId] = currentRooms + 1;
    
    // Log the change
    const event = events.find(e => e.Event_ID === eventId);
    const eventName = event ? event.Event : eventId;
    logUpload('ROOM_ADD', eventName, currentRooms + 1, 'Added', `Room ${currentRooms + 1} added to event "${eventName}"`);
    
    // Re-render
    renderSwimlanes();
    saveLogs();
    triggerAutoSave();
}

// Change course to a different room
function changeCourseRoom(eventId, courseId, newRoomNumber) {
    const placement = schedule[eventId]?.[courseId];
    if (!placement) {
        alert('Course not yet placed on timeline. Please drag to days first.');
        return;
    }
    
    const oldRoom = placement.roomNumber || 1;
    
    // Check for room conflicts - is another course already in this room on these days?
    const roomConflicts = [];
    if (schedule[eventId]) {
        for (const otherCourseId in schedule[eventId]) {
            if (otherCourseId === courseId) continue; // Skip self
            const otherPlacement = schedule[eventId][otherCourseId];
            if (otherPlacement.roomNumber === newRoomNumber) {
                // Check for day overlap
                const overlap = placement.days.some(day => otherPlacement.days.includes(day));
                if (overlap) {
                    const otherCourse = courses.find(c => c.Course_ID === otherCourseId);
                    roomConflicts.push(otherCourse?.Course_Name || otherCourseId);
                }
            }
        }
    }
    
    if (roomConflicts.length > 0) {
        alert(`⚠️ Room ${newRoomNumber} Conflict!\n\nThis room is already occupied on some of these days by:\n${roomConflicts.join('\n')}\n\nPlease choose a different room.`);
        // Re-render to reset dropdown
        renderSwimlanes();
        return;
    }
    
    // Update room assignment
    placement.roomNumber = newRoomNumber;
    
    // Log the change
    const course = courses.find(c => c.Course_ID === courseId);
    const courseName = course ? course.Course_Name : courseId;
    logUpload('COURSE_ROOM_MOVE', courseId, 0, 'Moved', `"${courseName}" moved from Room ${oldRoom} to Room ${newRoomNumber} in event`);
    
    // Re-render to show in new room lane
    renderSwimlanes();
    saveLogs();
    triggerAutoSave();
}

// Remove a room from an event
function removeRoom(eventId, roomNumber) {
    // Check if any courses are assigned to this room
    const coursesInRoom = [];
    if (schedule[eventId]) {
        for (const courseId in schedule[eventId]) {
            const placement = schedule[eventId][courseId];
            if (placement.roomNumber === roomNumber) {
                const course = courses.find(c => c.Course_ID === courseId);
                coursesInRoom.push(course?.Course_Name || courseId);
            }
        }
    }
    
    if (coursesInRoom.length > 0) {
        const proceed = confirm(`⚠️ Room ${roomNumber} has ${coursesInRoom.length} course(s) assigned:\n\n${coursesInRoom.join('\n')}\n\nRemoving this room will unassign these courses. Continue?`);
        if (!proceed) return;
        
        // Remove courses from this room
        if (schedule[eventId]) {
            for (const courseId in schedule[eventId]) {
                const placement = schedule[eventId][courseId];
                if (placement.roomNumber === roomNumber) {
                    delete schedule[eventId][courseId];
                    // Also remove from assignments
                    if (assignments[courseId]) {
                        assignments[courseId] = assignments[courseId].filter(id => id !== eventId);
                    }
                }
            }
        }
    }
    
    // Decrease room count
    const currentRooms = eventRooms[eventId] || 1;
    if (currentRooms > 1) {
        eventRooms[eventId] = currentRooms - 1;
        
        // Log the change
        const event = events.find(e => e.Event_ID === eventId);
        const eventName = event ? event.Event : eventId;
        logUpload('ROOM_REMOVE', eventName, currentRooms - 1, 'Removed', `Room ${roomNumber} removed from event "${eventName}"`);
        
        // Re-render
        renderSwimlanes();
        updateStats();
        updateConfigureDaysButton();
        saveLogs();
        triggerAutoSave();
    }
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
                    // Get room assignment from schedule (new room-based system)
                    const placement = schedule[eventId]?.[course.Course_ID];
                    const roomNumber = placement?.roomNumber || '';
                    
                    // Check if instructor is unavailable on this day
                    const blockedDays = getBlockedDays(course.Instructor, eventId);
                    const hasConflict = blockedDays.includes(dayNum);
                    const conflictStatus = hasConflict ? 'YES - Instructor Unavailable' : '';
                    
                    csv += `${escapeCSV(eventId)},${escapeCSV(eventName)},${escapeCSV(dayNum)},${escapeCSV(day.Day_Date)},${escapeCSV(course.Course_ID)},${escapeCSV(course.Instructor)},${escapeCSV(course.Course_Name)},${escapeCSV(course.Duration_Days)},${escapeCSV(roomNumber)},${escapeCSV(conflictStatus)},Yes\n`;
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

// Export schedule in import-ready format with Event_ID and Room_Number
function exportScheduleWithRooms() {
    if (courses.length === 0) {
        alert('Please load courses first');
        return;
    }
    
    if (events.length === 0) {
        alert('Please load events first');
        return;
    }
    
    // Create CSV content with format: Course_ID,Duration_Days,First_Day,Last_Day,Event_ID,Room_Number
    let csv = 'Course_ID,Duration_Days,First_Day,Last_Day,Event_ID,Room_Number\n';
    
    // Collect all scheduled courses
    const scheduledEntries = [];
    
    events.forEach(event => {
        const eventId = event.Event_ID;
        
        if (schedule[eventId]) {
            for (const courseId in schedule[eventId]) {
                const placement = schedule[eventId][courseId];
                
                // Only export if course has been scheduled (has days)
                if (placement.days && placement.days.length > 0) {
                    const course = courses.find(c => c.Course_ID === courseId);
                    if (!course) continue;
                    
                    // Get first and last days
                    const sortedDays = [...placement.days].sort((a, b) => a - b);
                    const firstDayNum = sortedDays[0];
                    const lastDayNum = sortedDays[sortedDays.length - 1];
                    
                    // Get actual dates for these days
                    const days = eventDays.filter(d => d.Event_ID === eventId);
                    const firstDayInfo = days.find(d => parseInt(d.Day_Number) === firstDayNum);
                    const lastDayInfo = days.find(d => parseInt(d.Day_Number) === lastDayNum);
                    
                    const firstDate = firstDayInfo?.Day_Date || '';
                    const lastDate = lastDayInfo?.Day_Date || '';
                    
                    // Get room number (blank if not assigned)
                    const roomNumber = placement.roomNumber || '';
                    
                    scheduledEntries.push({
                        courseId: courseId,
                        duration: course.Duration_Days,
                        firstDate: firstDate,
                        lastDate: lastDate,
                        eventId: eventId,
                        roomNumber: roomNumber
                    });
                }
            }
        }
    });
    
    // Sort by event and course for easier reading
    scheduledEntries.sort((a, b) => {
        if (a.eventId !== b.eventId) {
            return a.eventId.localeCompare(b.eventId);
        }
        return a.courseId.localeCompare(b.courseId);
    });
    
    // Build CSV rows
    scheduledEntries.forEach(entry => {
        csv += `${escapeCSV(entry.courseId)},${escapeCSV(entry.duration)},${escapeCSV(entry.firstDate)},${escapeCSV(entry.lastDate)},${escapeCSV(entry.eventId)},${escapeCSV(entry.roomNumber)}\n`;
    });
    
    // Add UTF-8 BOM for proper encoding in Excel
    const BOM = '\ufeff';
    const csvWithBOM = BOM + csv;
    
    // Download as CSV
    downloadFile(csvWithBOM, 'schedule_import_format.csv', 'text/csv;charset=utf-8');
    
    alert('Schedule exported in import format! This CSV includes Event_ID and Room_Number columns and can be re-imported.');
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
        triggerAutoSave();
        
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
        
        // Extract room counts from event days data if Room_Count column exists
        const roomCountsByEvent = {};
        newEventDays.forEach(row => {
            const eventId = row.Event_ID;
            if (row.Room_Count !== undefined && row.Room_Count !== null && row.Room_Count !== '') {
                const parsed = parseInt(row.Room_Count);
                if (!isNaN(parsed) && parsed > 0 && !roomCountsByEvent[eventId]) {
                    roomCountsByEvent[eventId] = parsed;
                }
            }
        });
        
        // Update eventRooms with discovered room counts
        Object.entries(roomCountsByEvent).forEach(([eventId, count]) => {
            eventRooms[eventId] = count;
        });
        
        // Recalculate unavailability if it was already loaded
        if (instructorUnavailable.length > 0) {
            calculateUnavailabilityMap();
        }
        
        renderAssignmentGrid();
        updateStats();
        triggerAutoSave();
        
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
        
        // Check for required columns - be flexible with Duration column name
        const firstRow = scheduleData[0];
        const hasBasicColumns = 'Course_ID' in firstRow && 'First_Day' in firstRow && 'Last_Day' in firstRow;
        const hasDuration = 'Duration_Days' in firstRow || 'Duration' in firstRow;
        
        if (!hasBasicColumns || !hasDuration) {
            const foundColumns = Object.keys(firstRow).join(', ');
            alert(`CSV must have columns: Course_ID, Duration_Days (or Duration), First_Day, Last_Day (optional: Event_ID, Room_Number)\n\nFound columns: ${foundColumns}`);
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
    
    // Warn if room assignments will be lost due to missing room configuration
    const roomConfigWarnings = new Set();
    scheduleData.forEach(row => {
        if (row.Event_ID && row.Room_Number) {
            const eventId = row.Event_ID.trim();
            const roomNum = parseInt(row.Room_Number);
            const currentRoomCount = eventRooms[eventId] || 1;
            if (!isNaN(roomNum) && roomNum > currentRoomCount) {
                roomConfigWarnings.add(`Event "${eventId}" has Room ${roomNum} in CSV but only ${currentRoomCount} room(s) configured`);
            }
        }
    });
    
    if (roomConfigWarnings.size > 0) {
        const warningMessage = 'Room configuration mismatch detected:\n\n' + 
            Array.from(roomConfigWarnings).join('\n') + 
            '\n\nPlease configure room counts for each event before importing room assignments.\n\nContinue anyway?';
        if (!confirm(warningMessage)) {
            return;
        }
    }
    
    scheduleData.forEach((row, index) => {
        const courseId = row.Course_ID.trim();
        const durationDays = parseFloat(row.Duration_Days);
        const firstDay = row.First_Day.trim();
        const lastDay = row.Last_Day.trim();
        
        // Check if Event_ID is provided (new format)
        const hasEventId = row.Event_ID && row.Event_ID.trim();
        
        // Check if Room_Number is provided (new format)
        const roomNumberStr = row.Room_Number ? row.Room_Number.trim() : '';
        const roomNumber = roomNumberStr ? parseInt(roomNumberStr) : null;
        
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
        
        // Find matching event - use Event_ID if provided, otherwise match by date range
        let match;
        
        if (hasEventId) {
            const eventId = row.Event_ID.trim();
            console.log('Processing with Event_ID:', eventId, 'for course:', courseId);
            const event = events.find(e => e.Event_ID === eventId);
            
            if (!event) {
                console.error('Event not found:', eventId);
                errors.push({
                    Row: index + 2,
                    Course_ID: courseId,
                    First_Day: firstDay,
                    Last_Day: lastDay,
                    Error: `Event_ID "${eventId}" not found`
                });
                return;
            }
            
            // Find day numbers for this event that match the dates
            console.log('Finding event days for:', eventId, startDate, endDate);
            match = findEventDaysForDates(eventId, startDate, endDate);
            console.log('Match result:', match);
            
            if (!match) {
                errors.push({
                    Row: index + 2,
                    Course_ID: courseId,
                    First_Day: firstDay,
                    Last_Day: lastDay,
                    Error: `Dates ${firstDay} to ${lastDay} do not match days in event "${eventId}"`
                });
                return;
            }
        } else {
            // Legacy mode: auto-detect event from date range
            match = findEventForDateRange(startDate, endDate);
            
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
        }
        
        // Validate room number if provided
        if (roomNumber !== null) {
            const eventRoomCount = eventRooms[match.eventId] || 1;
            if (roomNumber < 1 || roomNumber > eventRoomCount) {
                errors.push({
                    Row: index + 2,
                    Course_ID: courseId,
                    First_Day: firstDay,
                    Last_Day: lastDay,
                    Error: `Room_Number ${roomNumber} is invalid for event "${match.eventName}" (valid range: 1-${eventRoomCount})`
                });
                return;
            }
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
            days: match.dayNumbers,
            roomNumber: roomNumber,
            isDraft: false
        };
        
        successCount++;
        imported.push({
            Course_ID: courseId,
            Event: match.eventName,
            Days: `${match.startDayNumber}-${match.endDayNumber}`,
            Room: roomNumber !== null ? roomNumber : 'Unassigned'
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
    renderSwimlanesGrid(); // Update Room Grid view if active
    updateReportsGrid(); // Update Room Grid reports
    renderCoursesTableGrid(); // Update courses table
    triggerAutoSave();
    
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

// Find day numbers for a specific event given date range
function findEventDaysForDates(eventId, startDate, endDate) {
    console.log('findEventDaysForDates called with:', { eventId, startDate, endDate });
    const event = events.find(e => e.Event_ID === eventId);
    if (!event) {
        console.error('Event not found in events array:', eventId);
        return null;
    }
    
    const eventName = event.Event;
    console.log('Found event:', eventName);
    
    // Get all days for this event
    const days = eventDays.filter(d => d.Event_ID === eventId)
                         .sort((a, b) => a.Day_Number - b.Day_Number);
    
    console.log('Event days found:', days.length, days);
    
    if (days.length === 0) return null;
    
    // Normalize dates to compare by day only (ignore time)
    const normalizeDate = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate());
    const normStart = normalizeDate(startDate);
    const normEnd = normalizeDate(endDate);
    
    // Find day numbers by matching dates
    const startDayNumber = days.find(d => {
        const dayDate = normalizeDate(new Date(d.Day_Date));
        return dayDate.getTime() === normStart.getTime();
    })?.Day_Number;
    
    const endDayNumber = days.find(d => {
        const dayDate = normalizeDate(new Date(d.Day_Date));
        return dayDate.getTime() === normEnd.getTime();
    })?.Day_Number;
    
    if (!startDayNumber || !endDayNumber) return null;
    
    const dayNumbers = [];
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

// Download rooms template
function downloadRoomsTemplate() {
    const template = `Event_ID,Number_of_Rooms
ATL,10
DFW,8
DTW,5
CHI,12
PHX,10
STL,8
PHI,10
SLC,8
NYC,15
CLE,6
HOU,10
IND,7`;
    
    downloadFile(template, 'rooms_template.csv', 'text/csv');
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

// Export Trainer Contracts - Excel with one tab per trainer
function exportTrainerContracts() {
    // Helper: Extract last names from instructor field
    function getLastNames(instructorField) {
        // Split by comma for multiple instructors: "Lisa Harding, Thomas Jeffrey"
        const instructors = instructorField.split(',').map(i => i.trim());
        const lastNames = instructors.map(fullName => {
            const parts = fullName.trim().split(' ');
            return parts[parts.length - 1]; // Get last word as last name
        });
        return lastNames.join(' and ');
    }
    
    // Helper: Format date as "Mar 17 2026"
    function formatDate(dateStr) {
        if (!dateStr) return '';
        const date = new Date(dateStr);
        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        return `${months[date.getMonth()]} ${date.getDate()} ${date.getFullYear()}`;
    }
    
    // Helper: Get month name
    function getMonthName(dateStr) {
        if (!dateStr) return '';
        const date = new Date(dateStr);
        const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
        return months[date.getMonth()];
    }
    
    // Collect all scheduled courses with their details
    const scheduledCourses = [];
    
    events.forEach(event => {
        const eventId = event.Event_ID;
        const eventName = event.Event;
        const eventLocation = event.Location || eventName; // Use Location field if available, fallback to event name
        const isVirtual = isVirtualEvent(eventName);
        const days = eventDays.filter(d => d.Event_ID === eventId);
        
        courses.forEach(course => {
            const courseId = course.Course_ID;
            const placement = schedule[eventId]?.[courseId];
            
            // Only export courses that are scheduled (have days assigned)
            // For virtual events, roomNumber might be 0 or null, so check days instead
            if (!placement || !placement.days || placement.days.length === 0) {
                return;
            }
            
            // Get start and end dates
            const startDayNum = Math.min(...placement.days);
            const endDayNum = Math.max(...placement.days);
            
            const startDayObj = days.find(d => parseInt(d.Day_Number) === startDayNum);
            const endDayObj = days.find(d => parseInt(d.Day_Number) === endDayNum);
            
            if (!startDayObj || !endDayObj) return;
            
            const startDate = startDayObj.Day_Date;
            const endDate = endDayObj.Day_Date;
            
            scheduledCourses.push({
                instructor: course.Instructor,
                startDate: startDate,
                endDate: endDate,
                eventLocation: eventLocation,
                isVirtual: isVirtual,
                courseTitle: course.Course_Name,
                sortDate: new Date(startDate)
            });
        });
    });
    
    if (scheduledCourses.length === 0) {
        alert('No scheduled courses found. Please assign courses to rooms and days before exporting.');
        return;
    }
    
    // Group by instructor
    const byInstructor = {};
    scheduledCourses.forEach(sc => {
        // Split instructors if multiple (comma-separated)
        const instructors = sc.instructor.split(',').map(i => i.trim());
        
        instructors.forEach(instructor => {
            if (!byInstructor[instructor]) {
                byInstructor[instructor] = [];
            }
            byInstructor[instructor].push(sc);
        });
    });
    
    // Create workbook
    const wb = XLSX.utils.book_new();
    
    // Create a sheet for each instructor
    Object.keys(byInstructor).sort().forEach(instructor => {
        const courses = byInstructor[instructor];
        
        // Sort chronologically
        courses.sort((a, b) => a.sortDate - b.sortDate);
        
        // Build data rows
        const rows = courses.map(course => {
            const monthName = getMonthName(course.startDate);
            const location = course.eventLocation; // Using Location field from event data
            const monthLocation = `${monthName} - ${location}`;
            
            return {
                'Start Date': formatDate(course.startDate),
                'End Date': formatDate(course.endDate),
                'Month/Location': monthLocation,
                'Modality': course.isVirtual ? 'Virtual' : 'In Person',
                'Training Title': course.courseTitle,
                'Trainer(s)': getLastNames(course.instructor)
            };
        });
        
        // Create worksheet
        const ws = XLSX.utils.json_to_sheet(rows);
        
        // Set column widths
        ws['!cols'] = [
            { wch: 15 }, // Start Date
            { wch: 15 }, // End Date
            { wch: 30 }, // Month/Location
            { wch: 12 }, // Modality
            { wch: 40 }, // Training Title
            { wch: 20 }  // Trainer(s)
        ];
        
        // Use first name + last initial as sheet name (Excel limit: 31 chars)
        const parts = instructor.trim().split(' ');
        let sheetName = instructor;
        if (parts.length >= 2) {
            sheetName = `${parts[0]} ${parts[parts.length - 1].charAt(0)}`;
        }
        sheetName = sheetName.substring(0, 31); // Excel sheet name limit
        
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
    
    // Generate and download
    const timestamp = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Trainer_Contracts_${timestamp}.xlsx`);
    
    alert(`Exported contracts for ${Object.keys(byInstructor).length} trainer(s) to Excel`);
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
            triggerAutoSave();
            
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
    // Check which view is active
    const isRoomGrid = document.getElementById('roomGridView').classList.contains('active');
    const contentId = isRoomGrid ? 'reportsContentGrid' : 'reportsContent';
    const toggleId = isRoomGrid ? 'reportsToggleGrid' : 'reportsToggle';
    
    const content = document.getElementById(contentId);
    const toggle = document.getElementById(toggleId);
    
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
// Toggle conflicts section
function toggleConflictsGrid() {
    const content = document.getElementById('conflictsContentGrid');
    const toggle = document.getElementById('conflictsToggleGrid');
    
    if (content.classList.contains('expanded')) {
        content.classList.remove('expanded');
        toggle.textContent = '▶ Expand';
    } else {
        content.classList.add('expanded');
        toggle.textContent = '▼ Collapse';
        updateConflictsGrid();
    }
}

// Toggle courses table section
function toggleCoursesTable() {
    const content = document.getElementById('coursesTableContent');
    const toggle = document.getElementById('coursesTableToggle');
    
    if (content.classList.contains('expanded')) {
        content.classList.remove('expanded');
        toggle.textContent = '▶ Expand';
    } else {
        content.classList.add('expanded');
        toggle.textContent = '▼ Collapse';
        renderCoursesTable();
    }
}

// Toggle courses table section for Room Grid view
function toggleCoursesTableGrid() {
    const content = document.getElementById('coursesTableContentGrid');
    const toggle = document.getElementById('coursesTableToggleGrid');
    
    if (content.classList.contains('expanded')) {
        content.classList.remove('expanded');
        toggle.textContent = '▶ Expand';
    } else {
        content.classList.add('expanded');
        toggle.textContent = '▼ Collapse';
        renderCoursesTableGrid();
    }
}

// Toggle change log section for Room Grid view
function toggleChangeLogGrid() {
    const content = document.getElementById('changeLogContentGrid');
    const toggle = document.getElementById('changeLogToggleGrid');
    
    if (content.classList.contains('expanded')) {
        content.classList.remove('expanded');
        toggle.textContent = '▶ Expand';
    } else {
        content.classList.add('expanded');
        toggle.textContent = '▼ Collapse';
        renderChangeLogGrid();
    }
}

// Render change log table for Room Grid view
function renderChangeLogGrid() {
    const container = document.getElementById('changeLogTableContainerGrid');
    if (!container) return;
    
    // Combine all log types
    const allLogs = [
        ...changeLog.map(entry => ({ ...entry, logType: 'change' })),
        ...uploadsLog.map(entry => ({ ...entry, logType: 'upload' }))
    ];
    
    // Sort by timestamp (most recent first)
    allLogs.sort((a, b) => {
        const timeA = new Date(a.timestamp || 0);
        const timeB = new Date(b.timestamp || 0);
        return timeB - timeA;
    });
    
    if (allLogs.length === 0) {
        container.innerHTML = '<p style="color: #6c757d;">No changes logged yet.</p>';
        return;
    }
    
    let html = '<table class="report-table"><thead><tr>';
    html += '<th>Timestamp</th><th>Type</th><th>Action</th><th>Details</th><th>Notes</th>';
    html += '</tr></thead><tbody>';
    
    // Show most recent 50 entries
    allLogs.slice(0, 50).forEach(entry => {
        if (entry.logType === 'change') {
            // Schedule change log entry
            const action = entry.action || '';
            const courseInfo = entry.courseTitle ? `${entry.courseTitle} (${entry.courseId})` : entry.courseId;
            const dateInfo = entry.firstDay && entry.lastDay ? `${entry.firstDay} to ${entry.lastDay}` : '';
            const oldDateInfo = entry.oldFirstDay && entry.oldLastDay ? `Was: ${entry.oldFirstDay} to ${entry.oldLastDay}` : '';
            
            html += `<tr>
                <td style="white-space: nowrap;">${entry.timestamp || ''}</td>
                <td><span style="background: #e3f2fd; color: #1976d2; padding: 2px 8px; border-radius: 3px; font-size: 0.85em;">Schedule</span></td>
                <td>${action}</td>
                <td>${courseInfo}<br><small style="color: #666;">${dateInfo}</small></td>
                <td><small style="color: #666;">${oldDateInfo || entry.notes || ''}</small></td>
            </tr>`;
        } else {
            // Upload/system log entry
            const uploadType = entry.uploadType || '';
            const fileName = entry.fileName || '';
            const status = entry.status || '';
            const notes = entry.notes || '';
            const recordCount = entry.recordsCount !== undefined ? `(${entry.recordsCount} records)` : '';
            
            const typeColor = uploadType.includes('COURSE') ? '#fff3e0' : 
                            uploadType.includes('ROOM') ? '#f3e5f5' :
                            uploadType.includes('EVENT') ? '#e8f5e9' : '#f5f5f5';
            const textColor = uploadType.includes('COURSE') ? '#e65100' : 
                            uploadType.includes('ROOM') ? '#6a1b9a' :
                            uploadType.includes('EVENT') ? '#2e7d32' : '#424242';
            
            html += `<tr>
                <td style="white-space: nowrap;">${entry.timestamp || ''}</td>
                <td><span style="background: ${typeColor}; color: ${textColor}; padding: 2px 8px; border-radius: 3px; font-size: 0.85em;">${uploadType}</span></td>
                <td>${status}</td>
                <td>${fileName} ${recordCount}</td>
                <td><small style="color: #666;">${notes}</small></td>
            </tr>`;
        }
    });
    
    html += '</tbody></table>';
    
    if (allLogs.length > 50) {
        html += `<p style="color: #6c757d; text-align: center; margin-top: 10px;">Showing most recent 50 of ${allLogs.length} entries</p>`;
    }
    
    container.innerHTML = html;
}

// Render editable courses table
function renderCoursesTable() {
    const container = document.getElementById('coursesTableContainer');
    if (!container) return;
    
    if (courses.length === 0) {
        container.innerHTML = '<p style="color: #6c757d;">No courses loaded yet. Upload an Excel file or CSV to get started.</p>';
        return;
    }
    
    // Find duplicates (trim IDs for comparison)
    const courseIdCounts = {};
    courses.forEach(c => {
        const id = String(c.Course_ID || '').trim();
        courseIdCounts[id] = (courseIdCounts[id] || 0) + 1;
    });
    
    // Separate duplicates from non-duplicates
    const duplicates = [];
    const nonDuplicates = [];
    
    courses.forEach((course, index) => {
        const courseId = String(course.Course_ID || '').trim();
        const item = { course, index };
        if (courseIdCounts[courseId] > 1) {
            duplicates.push(item);
        } else {
            nonDuplicates.push(item);
        }
    });
    
    // Sort duplicates by Course_ID (so same IDs are together)
    duplicates.sort((a, b) => {
        const idA = String(a.course.Course_ID || '').trim().toLowerCase();
        const idB = String(b.course.Course_ID || '').trim().toLowerCase();
        return idA.localeCompare(idB);
    });
    
    // Sort non-duplicates by Course_ID
    nonDuplicates.sort((a, b) => {
        const idA = String(a.course.Course_ID || '').trim().toLowerCase();
        const idB = String(b.course.Course_ID || '').trim().toLowerCase();
        return idA.localeCompare(idB);
    });
    
    let html = '<table class="report-table"><thead><tr>';
    html += '<th>Course ID</th><th>Course Name</th><th>Instructor</th><th>Duration (Days)</th><th>Topic</th><th>Actions</th>';
    html += '</tr></thead><tbody>';
    
    // Render duplicates first
    if (duplicates.length > 0) {
        html += '<tr style="background: #fff4e6;"><td colspan="6" style="padding: 8px 10px; color: #856404;">Duplicate Course IDs detected</td></tr>';
        
        duplicates.forEach(item => {
            const course = item.course;
            const originalIndex = item.index;
            
            // Find where this course is scheduled
            const courseId = String(course.Course_ID || '').trim();
            const scheduledLocations = [];
            
            // Check if this course is assigned to any events (assignments is object: { courseId: [eventIds] })
            const assignedEventIds = assignments[courseId] || [];
            assignedEventIds.forEach(eventId => {
                if (schedule[eventId] && schedule[eventId][courseId]) {
                    const courseSchedule = schedule[eventId][courseId];
                    const event = events.find(e => e.Event_ID === eventId);
                    const eventName = event ? event.Event : eventId;
                    scheduledLocations.push(`${eventName} - Day ${courseSchedule.startDay}, Room ${courseSchedule.roomNumber}`);
                }
            });
            
            const scheduleInfo = scheduledLocations.length > 0 
                ? 'Scheduled in: ' + scheduledLocations.join('; ') 
                : 'Not yet scheduled';
            
            html += `<tr class="duplicate-warning" title="${scheduleInfo}">
                <td class="editable-cell" data-index="${originalIndex}" data-field="Course_ID" onclick="editCell(this)" title="${scheduleInfo}">${course.Course_ID || ''}</td>
                <td class="editable-cell" data-index="${originalIndex}" data-field="Course_Name" onclick="editCell(this)" title="${scheduleInfo}">${course.Course_Name || ''}</td>
                <td class="editable-cell" data-index="${originalIndex}" data-field="Instructor" onclick="editCell(this)">${course.Instructor || ''}</td>
                <td class="editable-cell" data-index="${originalIndex}" data-field="Duration_Days" onclick="editCell(this)">${course.Duration_Days || ''}</td>
                <td class="editable-cell" data-index="${originalIndex}" data-field="Topic" onclick="editCell(this)">${course.Topic || ''}</td>
                <td><button class="btn btn-small btn-warning" onclick="deleteCourse(${originalIndex})" title="Delete this course">🗑️</button></td>
            </tr>`;
        });
        
        // Add separator row
        html += '<tr style="background: #f8f9fa; height: 2px;"><td colspan="6" style="padding: 0; border-top: 2px solid #dee2e6;"></td></tr>';
        html += '<tr style="background: #f8f9fa;"><td colspan="6" style="padding: 8px 10px; color: #6c757d; font-size: 0.9em;">All courses</td></tr>';
    }
    
    // Render non-duplicates
    nonDuplicates.forEach(item => {
        const course = item.course;
        const originalIndex = item.index;
        html += `<tr>
            <td class="editable-cell" data-index="${originalIndex}" data-field="Course_ID" onclick="editCell(this)">${course.Course_ID || ''}</td>
            <td class="editable-cell" data-index="${originalIndex}" data-field="Course_Name" onclick="editCell(this)">${course.Course_Name || ''}</td>
            <td class="editable-cell" data-index="${originalIndex}" data-field="Instructor" onclick="editCell(this)">${course.Instructor || ''}</td>
            <td class="editable-cell" data-index="${originalIndex}" data-field="Duration_Days" onclick="editCell(this)">${course.Duration_Days || ''}</td>
            <td class="editable-cell" data-index="${originalIndex}" data-field="Topic" onclick="editCell(this)">${course.Topic || ''}</td>
            <td><button class="btn btn-small btn-warning" onclick="deleteCourse(${originalIndex})" title="Delete this course">🗑️</button></td>
        </tr>`;
    });
    
    html += '</tbody></table>';
    
    // Show duplicate note if any found
    if (duplicates.length > 0) {
        const duplicateCount = Object.values(courseIdCounts).filter(count => count > 1).length;
        html = `<div style="background: #fff4e6; padding: 10px; border-left: 3px solid #ffc107; margin-bottom: 15px;">
            <span style="color: #856404;">Note:</span> ${duplicateCount} duplicate Course ID(s) detected at the top of the table. Hover over rows to see where they're scheduled.
        </div>` + html;
    }
    
    container.innerHTML = html;
}

// Render editable courses table for Room Grid view
function renderCoursesTableGrid() {
    const container = document.getElementById('coursesTableContainerGrid');
    if (!container) return;
    
    if (courses.length === 0) {
        container.innerHTML = '<p style="color: #6c757d;">No courses loaded yet. Upload an Excel file or CSV to get started.</p>';
        return;
    }
    
    // Find duplicates (trim IDs for comparison)
    const courseIdCounts = {};
    courses.forEach(c => {
        const id = String(c.Course_ID || '').trim();
        courseIdCounts[id] = (courseIdCounts[id] || 0) + 1;
    });
    
    // Separate duplicates from non-duplicates
    const duplicates = [];
    const nonDuplicates = [];
    
    courses.forEach((course, index) => {
        const courseId = String(course.Course_ID || '').trim();
        const item = { course, index };
        if (courseIdCounts[courseId] > 1) {
            duplicates.push(item);
        } else {
            nonDuplicates.push(item);
        }
    });
    
    // Sort duplicates by Course_ID (so same IDs are together)
    duplicates.sort((a, b) => {
        const idA = String(a.course.Course_ID || '').trim().toLowerCase();
        const idB = String(b.course.Course_ID || '').trim().toLowerCase();
        return idA.localeCompare(idB);
    });
    
    // Sort non-duplicates by Course_ID
    nonDuplicates.sort((a, b) => {
        const idA = String(a.course.Course_ID || '').trim().toLowerCase();
        const idB = String(b.course.Course_ID || '').trim().toLowerCase();
        return idA.localeCompare(idB);
    });
    
    let html = '<table class="report-table"><thead><tr>';
    html += '<th>Course ID</th><th>Course Name</th><th>Instructor</th><th>Duration (Days)</th><th>Topic</th><th>Actions</th>';
    html += '</tr></thead><tbody>';
    
    // Render duplicates first
    if (duplicates.length > 0) {
        html += '<tr style="background: #fff4e6;"><td colspan="6" style="padding: 8px 10px; color: #856404;">Duplicate Course IDs detected</td></tr>';
        
        duplicates.forEach(item => {
            const course = item.course;
            const originalIndex = item.index;
            
            // Find where this course is scheduled
            const courseId = String(course.Course_ID || '').trim();
            const scheduledLocations = [];
            
            // Check if this course is assigned to any events (assignments is object: { courseId: [eventIds] })
            const assignedEventIds = assignments[courseId] || [];
            assignedEventIds.forEach(eventId => {
                if (schedule[eventId] && schedule[eventId][courseId]) {
                    const courseSchedule = schedule[eventId][courseId];
                    const event = events.find(e => e.Event_ID === eventId);
                    const eventName = event ? event.Event : eventId;
                    scheduledLocations.push(`${eventName} - Day ${courseSchedule.startDay}, Room ${courseSchedule.roomNumber}`);
                }
            });
            
            const scheduleInfo = scheduledLocations.length > 0 
                ? 'Scheduled in: ' + scheduledLocations.join('; ') 
                : 'Not yet scheduled';
            
            html += `<tr class="duplicate-warning" title="${scheduleInfo}">
                <td class="editable-cell" data-index="${originalIndex}" data-field="Course_ID" onclick="editCell(this)" title="${scheduleInfo}">${course.Course_ID || ''}</td>
                <td class="editable-cell" data-index="${originalIndex}" data-field="Course_Name" onclick="editCell(this)" title="${scheduleInfo}">${course.Course_Name || ''}</td>
                <td class="editable-cell" data-index="${originalIndex}" data-field="Instructor" onclick="editCell(this)">${course.Instructor || ''}</td>
                <td class="editable-cell" data-index="${originalIndex}" data-field="Duration_Days" onclick="editCell(this)">${course.Duration_Days || ''}</td>
                <td class="editable-cell" data-index="${originalIndex}" data-field="Topic" onclick="editCell(this)">${course.Topic || ''}</td>
                <td>
                    <button class="btn btn-small" style="background: #856404; color: white; margin-right: 4px;" onclick="mergeDuplicates('${courseId}')" title="Merge duplicates with this Course ID">Merge</button>
                    <button class="btn btn-small btn-warning" onclick="deleteCourse(${originalIndex})" title="Delete this course">🗑️</button>
                </td>
            </tr>`;
        });
        
        // Add separator row
        html += '<tr style="background: #f8f9fa; height: 2px;"><td colspan="6" style="padding: 0; border-top: 2px solid #dee2e6;"></td></tr>';
        html += '<tr style="background: #f8f9fa;"><td colspan="6" style="padding: 8px 10px; color: #6c757d; font-size: 0.9em;">All courses</td></tr>';
    }
    
    // Render non-duplicates
    nonDuplicates.forEach(item => {
        const course = item.course;
        const originalIndex = item.index;
        html += `<tr>
            <td class="editable-cell" data-index="${originalIndex}" data-field="Course_ID" onclick="editCell(this)">${course.Course_ID || ''}</td>
            <td class="editable-cell" data-index="${originalIndex}" data-field="Course_Name" onclick="editCell(this)">${course.Course_Name || ''}</td>
            <td class="editable-cell" data-index="${originalIndex}" data-field="Instructor" onclick="editCell(this)">${course.Instructor || ''}</td>
            <td class="editable-cell" data-index="${originalIndex}" data-field="Duration_Days" onclick="editCell(this)">${course.Duration_Days || ''}</td>
            <td class="editable-cell" data-index="${originalIndex}" data-field="Topic" onclick="editCell(this)">${course.Topic || ''}</td>
            <td><button class="btn btn-small btn-warning" onclick="deleteCourse(${originalIndex})" title="Delete this course">🗑️</button></td>
        </tr>`;
    });
    
    html += '</tbody></table>';
    
    // Show duplicate note if any found
    if (duplicates.length > 0) {
        const duplicateCount = Object.values(courseIdCounts).filter(count => count > 1).length;
        html = `<div style="background: #fff4e6; padding: 10px; border-left: 3px solid #ffc107; margin-bottom: 15px;">
            <span style="color: #856404;">Note:</span> ${duplicateCount} duplicate Course ID(s) detected at the top of the table. Hover over rows to see where they're scheduled.
        </div>` + html;
    }
    
    container.innerHTML = html;
}

// Edit cell inline
function editCell(cell) {
    if (cell.querySelector('input')) return; // Already editing
    
    const originalValue = cell.textContent;
    const index = parseInt(cell.dataset.index);
    const field = cell.dataset.field;
    
    const input = document.createElement('input');
    input.type = field === 'Duration_Days' ? 'number' : 'text';
    input.value = originalValue;
    input.step = field === 'Duration_Days' ? '0.5' : undefined;
    
    input.onblur = function() {
        const newValue = this.value;
        cell.textContent = newValue;
        
        const course = courses[index];
        const oldValue = originalValue;
        
        // Update course data
        courses[index][field] = field === 'Duration_Days' ? parseFloat(newValue) : newValue;
        
        // Log the change
        if (oldValue !== newValue) {
            const fieldName = field.replace(/_/g, ' ');
            logUpload('COURSE_EDIT', course.Course_ID, 0, 'Modified', `${fieldName}: "${oldValue}" → "${newValue}"`);
        }
        
        // Re-render to check for duplicates
        renderCoursesTable();
        
        // Update other views
        renderAssignmentGrid();
        renderSwimlanesGrid();
        
        // Save changes
        triggerAutoSave();
    };
    
    input.onkeydown = function(e) {
        if (e.key === 'Enter') {
            this.blur();
        } else if (e.key === 'Escape') {
            cell.textContent = originalValue;
        }
    };
    
    cell.textContent = '';
    cell.appendChild(input);
    input.focus();
    input.select();
}

// Merge duplicate courses with the same Course_ID
function mergeDuplicates(courseId) {
    const trimmedId = String(courseId).trim();
    
    // Find all courses with this Course_ID
    const duplicateCourses = [];
    courses.forEach((course, index) => {
        if (String(course.Course_ID || '').trim() === trimmedId) {
            duplicateCourses.push({ course, index });
        }
    });
    
    if (duplicateCourses.length <= 1) {
        alert('No duplicates found for this Course ID.');
        return;
    }
    
    // Build selection dialog
    let message = `Found ${duplicateCourses.length} courses with ID "${trimmedId}".\n\nSelect which one to KEEP (others will be deleted):\n\n`;
    duplicateCourses.forEach((item, i) => {
        const c = item.course;
        message += `${i + 1}. ${c.Course_Name || 'Unnamed'} - Instructor: ${c.Instructor || 'N/A'} - Duration: ${c.Duration_Days || 'N/A'} days\n`;
    });
    message += `\nEnter the number (1-${duplicateCourses.length}) of the course to keep:`;
    
    const choice = prompt(message);
    if (!choice) return; // User cancelled
    
    const choiceNum = parseInt(choice);
    if (isNaN(choiceNum) || choiceNum < 1 || choiceNum > duplicateCourses.length) {
        alert('Invalid selection. Please enter a number between 1 and ' + duplicateCourses.length);
        return;
    }
    
    // Keep the selected course, delete the others
    const keepIndex = duplicateCourses[choiceNum - 1].index;
    
    // Sort indices in reverse order to delete from end to beginning (to maintain correct indices)
    const indicesToDelete = duplicateCourses
        .map(item => item.index)
        .filter(idx => idx !== keepIndex)
        .sort((a, b) => b - a);
    
    // Delete the courses we're not keeping
    indicesToDelete.forEach(idx => {
        courses.splice(idx, 1);
    });
    
    // Log the merge
    const keptCourse = courses.find(c => String(c.Course_ID || '').trim() === trimmedId);
    logUpload('COURSE_MERGE', trimmedId, indicesToDelete.length, 'Merged', `Kept "${keptCourse?.Course_Name || 'Unknown'}" and removed ${indicesToDelete.length} duplicate(s) with Course ID "${trimmedId}"`);
    
    // Update all views
    renderCoursesTable();
    renderCoursesTableGrid();
    renderSwimlanesGrid();
    updateReportsGrid();
    
    alert(`Merged successfully! Kept 1 course and removed ${indicesToDelete.length} duplicate(s).`);
}

// Delete course
function deleteCourse(index) {
    const course = courses[index];
    if (!confirm(`Delete course "${course.Course_Name}" (${course.Course_ID})?\n\nThis will remove this course entry from the list.`)) {
        return;
    }
    
    const courseId = String(course.Course_ID || '').trim();
    
    // Remove from courses array
    courses.splice(index, 1);
    
    // Check if there are any other courses with the same Course_ID (trimmed comparison)
    const otherCoursesWithSameId = courses.some(c => String(c.Course_ID || '').trim() === courseId);
    
    // Only remove assignments and schedule if no other courses have this ID
    if (!otherCoursesWithSameId) {
        // Remove from assignments
        delete assignments[courseId];
        
        // Remove from schedule
        for (const eventId in schedule) {
            if (schedule[eventId][courseId]) {
                delete schedule[eventId][courseId];
            }
        }
    }
    
    // Log the deletion
    logUpload('COURSE_DELETE', courseId, 0, 'Deleted', `Course "${course.Course_Name}" (${courseId}) removed from courses list`);
    
    // Update all views
    renderCoursesTable();
    renderCoursesTableGrid();
    renderSwimlanesGrid();
    updateReportsGrid();
    renderAssignmentGrid();
    renderSwimlanesGrid();
    updateStats();
    
    // Save changes
    triggerAutoSave();
}

function toggleFinances() {
    // Check which view is active
    const isRoomGrid = document.getElementById('roomGridView').classList.contains('active');
    const contentId = isRoomGrid ? 'financesContentGrid' : 'financesContent';
    const toggleId = isRoomGrid ? 'financesToggleGrid' : 'financesToggle';
    
    const content = document.getElementById(contentId);
    const toggle = document.getElementById(toggleId);
    
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
    updateEventUtilization();
    updateTopicsPerEvent();
    updateFinances();
    
    // Also update grid view reports if they exist
    if (document.getElementById('instructorWorkloadReportGrid')) {
        updateInstructorWorkloadGrid();
        updateEventUtilizationGrid();
        updateTopicsPerEventGrid();
        updateFinancesGrid();
        updateConflictsGrid();
    }
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
    html += '<th>Event</th><th>Courses</th>';
    scenarios.forEach(scenario => {
        html += `<th>${scenario.name} (${scenario.seats} seats)</th>`;
    });
    html += '</tr></thead><tbody>';
    
    let scenarioTotals = [0, 0, 0];
    let totalCourses = 0;
    
    // Calculate revenue for each event
    events.forEach(event => {
        const eventId = event.Event_ID;
        
        // Find all courses assigned to this event (exclude drafts)
        const eventCourses = courses.filter(course => {
            const isAssigned = assignments[course.Course_ID]?.includes(eventId);
            const isDraft = schedule[eventId]?.[course.Course_ID]?.isDraft;
            return isAssigned && !isDraft;
        });
        
        // Calculate total revenue per scenario
        const scenarioRevenues = scenarios.map((scenario, idx) => {
            let totalRevenue = 0;
            eventCourses.forEach(course => {
                const price = getCoursePrice(course.Duration_Days);
                totalRevenue += price * scenario.seats;
            });
            scenarioTotals[idx] += totalRevenue;
            return totalRevenue;
        });
        
        html += `<tr>
            <td><strong>${event.Event}</strong></td>
            <td>${eventCourses.length}</td>`;
        
        scenarioRevenues.forEach(revenue => {
            html += `<td>$${revenue.toLocaleString()}</td>`;
        });
        
        html += '</tr>';
        totalCourses += eventCourses.length;
    });
    
    html += `<tr style="background: #e7f3ff; font-weight: 700;">
        <td><strong>Total</strong></td>
        <td>${totalCourses}</td>`;
    
    scenarioTotals.forEach(total => {
        html += `<td>$${total.toLocaleString()}</td>`;
    });
    
    html += '</tr></tbody></table>';
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

// ========================================
// ROOM GRID VIEW FUNCTIONS
// ========================================

// Update reports for Room Grid view
function updateReportsGrid() {
    updateInstructorWorkloadGrid();
    updateTopicCoverageGrid();
    updateEventUtilizationGrid();
    updateTopicsPerEventGrid();
    updateFinancesGrid();
}

// Copy all report functions for grid view
function updateInstructorWorkloadGrid() {
    const container = document.getElementById('instructorWorkloadReportGrid');
    if (!container) return;
    
    // Reuse the same logic but target different container
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

function updateTopicCoverageGrid() {
    const container = document.getElementById('topicCoverageReportGrid');
    if (!container) return;
    
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

function updateEventUtilizationGrid() {
    const container = document.getElementById('eventUtilizationReportGrid');
    if (!container) return;
    
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
        
        const assignedCourseIds = [];
        courses.forEach(course => {
            if (assignments[course.Course_ID] && assignments[course.Course_ID].includes(eventId)) {
                assignedCourseIds.push(course.Course_ID);
            }
        });
        
        const courseCount = assignedCourseIds.length;
        
        const instructorsSet = new Set();
        assignedCourseIds.forEach(courseId => {
            const course = courses.find(c => c.Course_ID === courseId);
            if (course) {
                instructorsSet.add(course.Instructor);
            }
        });
        const instructorCount = instructorsSet.size;
        
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

function updateConflictsGrid() {
    const container = document.getElementById('conflictsReportGrid');
    if (!container) return;
    
    if (events.length === 0 || courses.length === 0) {
        container.innerHTML = '<p style="color: #6c757d;">No events or courses loaded yet.</p>';
        return;
    }
    
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
                    Course_Name: course.Course_Name,
                    Instructor: course.Instructor,
                    Conflict_Dates: conflictDates
                });
            }
        });
    });
    
    if (conflicts.length === 0) {
        container.innerHTML = '<p style="color: #28a745; font-weight: bold;">✅ No conflicts found! All courses are scheduled without instructor unavailability issues.</p>';
        return;
    }
    
    let html = `<p style="color: #dc3545; font-weight: bold; margin-bottom: 15px;">⚠️ Found ${conflicts.length} scheduling conflict${conflicts.length !== 1 ? 's' : ''}:</p>`;
    html += '<table class="report-table"><thead><tr>';
    html += '<th>Event</th><th>Course</th><th>Instructor</th><th>Conflict Dates</th>';
    html += '</tr></thead><tbody>';
    
    conflicts.forEach(conflict => {
        html += `<tr>
            <td>${conflict.Event}</td>
            <td>${conflict.Course_Name}</td>
            <td>${conflict.Instructor}</td>
            <td style="color: #dc3545;">${conflict.Conflict_Dates}</td>
        </tr>`;
    });
    
    html += '</tbody></table>';
    container.innerHTML = html;
}

function updateFinancesGrid() {
    const container = document.getElementById('financesReportGrid');
    if (!container) return;
    
    if (events.length === 0 || courses.length === 0) {
        container.innerHTML = '<p style="color: #6c757d;">No events or courses loaded yet.</p>';
        return;
    }
    
    function getCoursePrice(durationDays) {
        const duration = parseFloat(durationDays);
        if (duration <= 0.5) return 518;
        if (duration <= 1) return 881;
        if (duration <= 2) return 1735;
        if (duration <= 3) return 2427;
        return 3228;
    }
    
    // Seat scenarios
    const scenarios = [
        { name: 'Low', seats: 10 },
        { name: 'Mid', seats: 20 },
        { name: 'High', seats: 30 }
    ];
    
    let html = '<table class="report-table"><thead><tr>';
    html += '<th>Event</th><th>Courses</th>';
    scenarios.forEach(scenario => {
        html += `<th>${scenario.name} (${scenario.seats} seats)</th>`;
    });
    html += '</tr></thead><tbody>';
    
    let scenarioTotals = [0, 0, 0];
    let totalCourses = 0;
    
    events.forEach(event => {
        const eventId = event.Event_ID;
        
        // Find all courses assigned to this event (exclude drafts)
        const eventCourses = courses.filter(course => {
            const isAssigned = assignments[course.Course_ID]?.includes(eventId);
            const isDraft = schedule[eventId]?.[course.Course_ID]?.isDraft;
            return isAssigned && !isDraft;
        });
        
        // Calculate total revenue per scenario
        const scenarioRevenues = scenarios.map((scenario, idx) => {
            let totalRevenue = 0;
            eventCourses.forEach(course => {
                const price = getCoursePrice(course.Duration_Days);
                totalRevenue += price * scenario.seats;
            });
            scenarioTotals[idx] += totalRevenue;
            return totalRevenue;
        });
        
        html += `<tr>
            <td><strong>${event.Event}</strong></td>
            <td>${eventCourses.length}</td>`;
        
        scenarioRevenues.forEach(revenue => {
            html += `<td>$${revenue.toLocaleString()}</td>`;
        });
        
        html += '</tr>';
        totalCourses += eventCourses.length;
    });
    
    html += `<tr style="background: #e7f3ff; font-weight: 700;">
        <td><strong>Total</strong></td>
        <td>${totalCourses}</td>`;
    
    scenarioTotals.forEach(total => {
        html += `<td>$${total.toLocaleString()}</td>`;
    });
    
    html += '</tr></tbody></table>';
    container.innerHTML = html;
}

function updateTopicsPerEventGrid() {
    const container = document.getElementById('topicsPerEventReportGrid');
    if (!container) return;
    
    if (events.length === 0 || courses.length === 0) {
        container.innerHTML = '<p style="color: #6c757d;">No events or courses loaded yet.</p>';
        return;
    }
    
    const allTopics = [...new Set(courses.map(c => c.Topic || 'Uncategorized'))].sort();
    const eventTopicData = [];
    
    events.forEach(event => {
        const eventId = event.Event_ID;
        const eventName = event.Event;
        const topicCounts = {};
        
        allTopics.forEach(topic => {
            topicCounts[topic] = 0;
        });
        
        courses.forEach(course => {
            const courseTopic = course.Topic || 'Uncategorized';
            const courseId = course.Course_ID;
            
            if (assignments[courseId] && assignments[courseId].includes(eventId)) {
                topicCounts[courseTopic]++;
            }
        });
        
        eventTopicData.push({
            eventName,
            topicCounts
        });
    });
    
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

// Render swimlanes for room grid view
// Check if event is virtual (by name)
function isVirtualEvent(eventName) {
    return eventName.toLowerCase().includes('virtual');
}

function renderSwimlanesGrid() {
    const container = document.getElementById('swimlanesContainerGrid');
    
    // Save current expanded/collapsed state
    const expandedState = {};
    document.querySelectorAll('.event-swimlane-body').forEach(body => {
        const eventId = body.id.replace('body-grid-', '');
        expandedState[eventId] = !body.classList.contains('collapsed');
    });
    
    container.innerHTML = '';
    
    // Build consolidated "All Open Bookings" section at the top
    const allOpenBookingsSection = document.createElement('div');
    allOpenBookingsSection.style.cssText = 'margin-bottom: 30px; background: white; border: 2px solid #dee2e6; border-radius: 10px; overflow: hidden;';
    allOpenBookingsSection.innerHTML = `
        <div style="background: linear-gradient(135deg, #28a745 0%, #20c997 100%); color: white; padding: 15px 20px; font-size: 1.3em; font-weight: 700; cursor: pointer; user-select: none; display: flex; justify-content: space-between; align-items: center;" onclick="toggleAllOpenBookings()">
            <span>📊 All Open Bookings</span>
            <span id="toggle-all-bookings">▶ Expand</span>
        </div>
        <div id="all-open-bookings-content" style="padding: 20px; display: none; max-height: 600px; overflow-y: auto;"></div>
    `;
    container.appendChild(allOpenBookingsSection);
    
    // Populate the All Open Bookings content
    buildAllOpenBookingsContent();
    
    // Sort events by first day's date (earliest first)
    const sortedEvents = [...events].sort((a, b) => {
        const daysA = eventDays.filter(d => d.Event_ID === a.Event_ID);
        const daysB = eventDays.filter(d => d.Event_ID === b.Event_ID);
        
        if (daysA.length === 0) return 1;
        if (daysB.length === 0) return -1;
        
        const dateA = new Date(daysA[0].Day_Date);
        const dateB = new Date(daysB[0].Day_Date);
        
        return dateA - dateB;
    });
    
    sortedEvents.forEach(event => {
        const eventId = event.Event_ID;
        const eventName = event.Event;
        const totalDays = parseInt(event['Total_Days']);
        const isVirtual = isVirtualEvent(eventName);
        
        // Get courses assigned to this event
        const assignedCourses = courses.filter(course => 
            assignments[course.Course_ID]?.includes(eventId)
        );
        
        // Initialize schedule entries for assigned courses that don't have placement yet
        assignedCourses.forEach(course => {
            if (!schedule[eventId]) {
                schedule[eventId] = {};
            }
            if (!schedule[eventId][course.Course_ID]) {
                schedule[eventId][course.Course_ID] = {
                    startDay: null,
                    days: [],
                    roomNumber: null
                };
            }
        });
        
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
        
        // Get days for this event
        const days = eventDays.filter(d => d.Event_ID === eventId);
        
        // Create swimlane
        const swimlane = document.createElement('div');
        const isLocked = lockedEvents.has(eventId);
        swimlane.className = isLocked ? 'event-swimlane locked' : 'event-swimlane';
        swimlane.dataset.eventId = eventId;
        
        // Get date range from first and last event days
        const eventFirstDay = days[0];
        const eventLastDay = days[days.length - 1];
        let dateRangeStr = '';
        if (eventFirstDay && eventFirstDay.Day_Date) {
            const startDate = new Date(eventFirstDay.Day_Date);
            const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
            const monthName = months[startDate.getMonth()];
            const startDay = startDate.getDate();
            
            if (eventLastDay && eventLastDay.Day_Date) {
                const endDate = new Date(eventLastDay.Day_Date);
                const endDay = endDate.getDate();
                dateRangeStr = `${monthName} ${startDay}-${endDay}`;
            } else {
                dateRangeStr = `${monthName} ${startDay}`;
            }
        }
        
        // Determine if this event should be expanded
        const isExpanded = expandedState[eventId] || false;
        const bodyClass = isExpanded ? 'event-swimlane-body' : 'event-swimlane-body collapsed';
        const toggleText = isExpanded ? '▼ Collapse' : '▶ Expand';
        
        // Add conflict indicator to header if conflicts exist
        const conflictIndicator = hasEventConflict ? '<span style="color: #ff9800; font-size: 1.5em; margin-left: 10px;">●●</span>' : '';
        
        // Get number of rooms for this event
        const numRooms = eventRooms[eventId] || 1;
        
        // Build day timeline (single timeline for all rooms in this view)
        let dayTimelineHTML = `
            <div class="day-timeline" data-event-id="${eventId}">
                ${days.map((day, index) => `
                    <div class="day-slot" data-day-num="${day.Day_Number}">
                        <div class="day-label">Day ${day.Day_Number}</div>
                        <div class="day-date">${day.Day_Date || ''}</div>
                    </div>
                `).join('')}
            </div>
        `;
        
        // Calculate room availability for each room (skip for virtual events)
        const roomAvailability = {};
        let bookingStatus = '';
        
        if (!isVirtual) {
            for (let roomNum = 1; roomNum <= numRooms; roomNum++) {
                roomAvailability[roomNum] = new Set();
                // Initially all days are available
                for (let day = 1; day <= totalDays; day++) {
                    roomAvailability[roomNum].add(day);
                }
            }
            
            // Track unique room-day combinations to avoid double-counting when drafts overlap
            const occupiedRoomDays = new Set();
            if (schedule[eventId]) {
                for (const courseId in schedule[eventId]) {
                    const placement = schedule[eventId][courseId];
                    if (placement.roomNumber && placement.days && placement.days.length > 0) {
                        placement.days.forEach(day => {
                            // Add unique room-day combination to set
                            occupiedRoomDays.add(`${placement.roomNumber}-${day}`);
                            roomAvailability[placement.roomNumber]?.delete(day);
                        });
                    }
                }
            }
            
            // Calculate booking status using unique room-day count
            const totalRoomDaysAvailable = numRooms * totalDays;
            const totalRoomDaysUsed = occupiedRoomDays.size;
            const unbookedRoomDays = totalRoomDaysAvailable - totalRoomDaysUsed;
            bookingStatus = unbookedRoomDays === 0 
                ? 'Fully Booked'
                : `Unbooked: ${unbookedRoomDays} room-days`;
        }
        
        // Lock status (isLocked already declared above)
        const lockIcon = isLocked ? '🔒' : '🔓';
        const lockTitle = isLocked ? 'Click to unlock event' : 'Click to lock event';
        
        // Build room availability bars (skip for virtual events)
        let roomAvailabilityHTML = '';
        let hasAnyAvailability = false;
        
        if (!isVirtual) {
            for (let roomNum = 1; roomNum <= numRooms; roomNum++) {
            const availableDays = Array.from(roomAvailability[roomNum]).sort((a, b) => a - b);
            if (availableDays.length > 0) {
                hasAnyAvailability = true;
                
                // Group consecutive days into ranges
                const ranges = [];
                let rangeStart = availableDays[0];
                let rangeEnd = availableDays[0];
                
                for (let i = 1; i < availableDays.length; i++) {
                    if (availableDays[i] === rangeEnd + 1) {
                        rangeEnd = availableDays[i];
                    } else {
                        ranges.push({ start: rangeStart, end: rangeEnd });
                        rangeStart = availableDays[i];
                        rangeEnd = availableDays[i];
                    }
                }
                ranges.push({ start: rangeStart, end: rangeEnd });
                
                // Create bars for each range
                ranges.forEach((range, rangeIndex) => {
                    const daysInRange = range.end - range.start + 1;
                    const blockWidth = (100 / totalDays) * daysInRange;
                    const blockLeft = ((range.start - 1) / totalDays) * 100;
                    
                    // Find courses that could fit in this gap
                    const suitableCourses = courses.filter(course => {
                        const courseId = course.Course_ID;
                        const courseDuration = Math.ceil(parseFloat(course.Duration_Days));
                        
                        // Course must fit in the available days
                        if (courseDuration > daysInRange) return false;
                        
                        // Check if course is already scheduled in this event (any room, any days)
                        const placement = schedule[eventId]?.[courseId];
                        if (placement && placement.days && placement.days.length > 0) {
                            // Already scheduled in this event, not available
                            return false;
                        }
                        
                        // Check instructor availability during this range
                        const blockedDays = getBlockedDays(course.Instructor, eventId);
                        const rangeDays = [];
                        for (let d = range.start; d <= range.end; d++) {
                            rangeDays.push(d);
                        }
                        const hasConflict = rangeDays.some(day => blockedDays.includes(day));
                        if (hasConflict) return false;
                        
                        return true;
                    });
                    
                    // Create dropdown options
                    let dropdownOptions = '<option value="">+ Add Draft Option...</option>';
                    suitableCourses.forEach(course => {
                        const courseDuration = Math.ceil(parseFloat(course.Duration_Days));
                        dropdownOptions += `<option value="${course.Course_ID}">${course.Instructor} - ${course.Course_Name} (${courseDuration} ${courseDuration === 1 ? 'day' : 'days'})</option>`;
                    });
                    
                    const uniqueId = `room-${roomNum}-range-${rangeIndex}-event-${eventId}`;
                    
                    roomAvailabilityHTML += `
                        <div style="background: #f8f9fa; border-radius: 8px; padding: 12px; margin-bottom: 10px; border: 2px solid #28a745;">
                            <div style="display: flex; align-items: center; gap: 15px; margin-bottom: 10px;">
                                <div style="min-width: 150px; font-weight: 700; color: #28a745;">
                                    🟢 Room ${roomNum}
                                </div>
                                <div style="flex: 1; position: relative; min-height: 40px; background: white; border-radius: 8px; padding: 5px;">
                                    <div style="position: absolute; left: ${blockLeft}%; width: ${blockWidth}%; top: 5px; height: 30px; background: linear-gradient(135deg, #28a745 0%, #20c997 100%); color: white; border-radius: 6px; display: flex; align-items: center; justify-content: center; font-weight: 600; font-size: 0.85em; box-shadow: 0 2px 8px rgba(40, 167, 69, 0.3);">
                                        Days ${range.start}${range.start !== range.end ? `-${range.end}` : ''} (${daysInRange} ${daysInRange === 1 ? 'day' : 'days'} available)
                                    </div>
                                </div>
                            </div>
                            ${suitableCourses.length > 0 ? `
                                <div style="display: flex; gap: 10px; align-items: stretch;">
                                    <select class="add-course-dropdown" id="draft-select-${uniqueId}" onchange="addToDraftList('${eventId}', ${roomNum}, ${range.start}, ${range.end}, '${uniqueId}', this.value); this.value='';" style="flex: 1;">
                                        ${dropdownOptions}
                                    </select>
                                </div>
                                <div id="draft-list-${uniqueId}" style="margin-top: 10px; display: none;">
                                    <div style="font-weight: 600; color: #667eea; margin-bottom: 5px; font-size: 0.9em;">📋 Draft Selections:</div>
                                    <div id="draft-items-${uniqueId}"></div>
                                </div>
                            ` : '<div style="color: #6c757d; font-size: 0.9em; padding: 8px; text-align: center; background: white; border-radius: 6px;">No courses fit in this time slot</div>'}
                        </div>
                    `;
                });
            }
        }
        
        if (hasAnyAvailability) {
            roomAvailabilityHTML = `
                <div style="margin: 15px 0; padding: 15px; background: #e7f3e7; border-radius: 8px; border: 2px solid #28a745;">
                    <div style="font-weight: 700; color: #28a745; margin-bottom: 10px; font-size: 1.1em; cursor: pointer; user-select: none; display: flex; align-items: center; gap: 8px;" 
                         onclick="toggleRoomAvailability('${eventId}')">
                        <span id="room-avail-icon-${eventId}">▶</span>
                        <span>📊 Room Availability</span>
                    </div>
                    <div id="room-avail-content-${eventId}" style="display: none;">
                        ${roomAvailabilityHTML}
                    </div>
                </div>
            `;
            }
        } // End if (!isVirtual) for room availability
        
        // Sort courses: blank/unassigned rooms at top, then by room number (1, 2, 3, etc.)
        const sortedCourses = [...assignedCourses].sort((a, b) => {
            const roomA = schedule[eventId]?.[a.Course_ID]?.roomNumber;
            const roomB = schedule[eventId]?.[b.Course_ID]?.roomNumber;
            
            // Treat null, undefined, and 0 as "no room assigned"
            const aIsBlank = roomA === null || roomA === undefined || roomA === 0;
            const bIsBlank = roomB === null || roomB === undefined || roomB === 0;
            
            // Both blank - maintain order
            if (aIsBlank && bIsBlank) return 0;
            // A is blank - A comes first (top)
            if (aIsBlank) return -1;
            // B is blank - B comes first (top)
            if (bIsBlank) return 1;
            
            // Both have room assignments - sort by room number ascending (1, 2, 3...)
            return roomA - roomB;
        });
        
        // Count drafts for this event
        let eventDraftCount = 0;
        if (schedule[eventId]) {
            for (const courseId in schedule[eventId]) {
                if (schedule[eventId][courseId].isDraft) {
                    eventDraftCount++;
                }
            }
        }
        const draftIndicator = eventDraftCount > 0 ? ` • <span style="color: #ff9800; font-weight: 700;">Drafts: ${eventDraftCount}</span>` : '';
        
        // Build course swimlanes
        let courseSwimlanesHTML = sortedCourses.map(course => 
            renderCourseSwimlaneGrid(course, eventId, totalDays, numRooms, isVirtual)
        ).join('');
        
        swimlane.innerHTML = `
            <div class="event-swimlane-header" onclick="toggleEventSwimlaneGrid('${eventId}')" style="display: flex; justify-content: space-between; align-items: center;">
                <div style="display: flex; gap: 15px; align-items: center; flex: 1;">
                    <span style="min-width: 320px; max-width: 320px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${eventName}">${eventName}</span>
                    <span style="min-width: 80px;">${totalDays} days</span>
                    <span style="min-width: 120px;">${dateRangeStr}</span>
                    <span style="min-width: 100px;">${isVirtual ? 'V Rooms' : 'Rooms'}: <span style="color: #ff9800; font-weight: 700;">${numRooms}</span>
                        <span onclick="event.stopPropagation(); editRoomCount('${eventId}', ${numRooms})" 
                              style="cursor: pointer; opacity: 0.8; padding: 0 3px;" 
                              title="Click to edit room count">✎</span>
                    </span>
                    ${!isVirtual ? `<span style="min-width: 150px;">${bookingStatus}</span>` : ''}
                    ${eventDraftCount > 0 ? `<span style="color: #ff9800; font-weight: 700; min-width: 80px;">Drafts: ${eventDraftCount}</span>` : ''}
                    <span onclick="event.stopPropagation(); toggleEventLock('${eventId}')" 
                          style="cursor: pointer; opacity: 0.9; padding: 0 5px; font-size: 1.1em;" 
                          title="${lockTitle}">${lockIcon}</span>
                    ${conflictIndicator}
                </div>
                <span id="toggle-grid-${eventId}">${toggleText}</span>
            </div>
            <div class="${bodyClass}" id="body-grid-${eventId}">
                ${dayTimelineHTML}
                ${roomAvailabilityHTML}
                ${courseSwimlanesHTML}
                <div class="add-course-section">
                    <select class="add-course-dropdown" id="add-course-grid-${eventId}" onchange="addCourseToEventGrid('${eventId}', this.value, this)" style="flex: 1;">
                        <option value="">+ Add Course to Event...</option>
                    </select>
                </div>
            </div>
        `;
        
        container.appendChild(swimlane);
        
        // Populate the dropdown with available courses
        populateCourseDropdownGrid(eventId, assignedCourses);
    });
    
    // Setup drag and drop for all course blocks
    setupDragAndDropGrid();
}

// Render a single course swimlane for room grid view
function renderCourseSwimlaneGrid(course, eventId, totalDays, numRooms, isVirtual = false) {
    const courseId = course.Course_ID;
    const duration = parseFloat(course.Duration_Days);
    const daysNeeded = Math.ceil(duration);
    
    // Get blocked days for this instructor at this event
    const blockedDays = getBlockedDays(course.Instructor, eventId);
    
    // Get current placement if exists
    const placement = schedule[eventId]?.[courseId];
    const startDay = placement?.startDay;
    const assignedRoom = placement?.roomNumber || null;
    const isDraft = placement?.isDraft || false;
    
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
    
    // Add draft class and label
    const draftClass = isDraft ? 'draft' : '';
    const draftLabel = isDraft ? ' DRAFT' : '';
    
    // Generate unavailability warning if any
    let unavailWarning = '';
    if (blockedDays.length > 0) {
        unavailWarning = `<div style="color: #dc3545; font-size: 0.85em; margin-top: 3px;">⚠️ Unavailable: Days ${blockedDays.join(', ')}</div>`;
    }
    
    // Build room grid (skip for virtual events)
    let roomGridHTML = '';
    if (!isVirtual) {
        roomGridHTML = '<div class="room-grid"><span style="font-size: 0.9em; color: #667eea; font-weight: 600; margin-right: 8px;">🏠 Room:</span>';
        for (let r = 1; r <= numRooms; r++) {
            const isSelected = (assignedRoom !== null && r === assignedRoom) ? 'selected' : '';
            roomGridHTML += `
                <div class="room-grid-cell ${isSelected}" 
                     data-room="${r}" 
                     data-course-id="${courseId}"
                     data-event-id="${eventId}"
                     onclick="selectRoomGrid('${eventId}', '${courseId}', ${r})">
                    ${r}
                </div>
            `;
        }
        roomGridHTML += '</div>';
    }
    
    return `
        <div class="course-swimlane" data-course-id="${courseId}" data-event-id="${eventId}" data-room-number="${assignedRoom}">
            <div class="course-info-sidebar">
                <div class="course-info-name">${course.Course_Name}</div>
                <div class="course-info-instructor">${course.Instructor}</div>
                <div class="course-info-duration">📏 ${course.Duration_Days} days</div>
                ${unavailWarning}
            </div>
            <div class="timeline-container">
                <div class="course-timeline" data-course-id="${courseId}" data-event-id="${eventId}" data-room-number="${assignedRoom}" data-total-days="${totalDays}" data-instructor="${course.Instructor}" data-blocked-days="${blockedDays.join(',')}">
                    <div class="course-block ${startDay ? '' : 'unplaced'} ${hasConflict ? 'has-conflict' : ''} ${draftClass}" 
                         data-course-id="${courseId}"
                         data-event-id="${eventId}"
                         data-days-needed="${daysNeeded}"
                         draggable="true"
                         style="${startDay ? `position: absolute; left: ${blockLeft}%; width: ${blockWidth}%; top: 5px; height: 40px; line-height: 40px;` : ''}">
                        ${startDay ? `Days ${startDay}-${startDay + daysNeeded - 1}${draftLabel}` : 'Drag to timeline'}
                    </div>
                </div>
                ${roomGridHTML}
            </div>
            <div class="course-actions">
                <button class="btn btn-danger btn-small" onclick="removeCourseFromEventGrid('${courseId}', '${eventId}')">
                    ✗ Remove
                </button>
                <button class="btn btn-secondary btn-small" style="background: #e0e0e0; color: #666;" onclick="toggleDraftStatus('${eventId}', '${courseId}')">
                    ${isDraft ? '✓ Finalize' : '📝 Draft'}
                </button>
            </div>
        </div>
    `;
}

// Room selection handler for grid view
async function selectRoomGrid(eventId, courseId, roomNumber) {
    // Check if event is locked
    if (lockedEvents.has(eventId)) {
        alert('🔒 This event is locked. Unlock it to make changes.');
        return;
    }
    
    if (!schedule[eventId]) {
        schedule[eventId] = {};
    }
    
    if (!schedule[eventId][courseId]) {
        schedule[eventId][courseId] = {
            startDay: null,
            days: [],
            roomNumber: roomNumber
        };
    } else {
        const placement = schedule[eventId][courseId];
        
        // Toggle: if clicking the same room, deselect it
        if (placement.roomNumber === roomNumber) {
            schedule[eventId][courseId].roomNumber = null;
            const course = courses.find(c => c.Course_ID === courseId);
            const courseName = course ? course.Course_Name : courseId;
            logUpload('ROOM_SELECT', courseName, 0, 'Deselected', `Course "${courseName}" deselected from any room`);
            renderSwimlanesGrid();
            saveLogs();
            triggerAutoSave();
            return;
        }
        
        // Check if changing room - validate no conflicts
        if (placement.startDay && placement.days && placement.days.length > 0) {
            // Check for room conflicts
            const roomConflicts = [];
            if (schedule[eventId]) {
                for (const otherCourseId in schedule[eventId]) {
                    if (otherCourseId === courseId) continue;
                    const otherPlacement = schedule[eventId][otherCourseId];
                    if (otherPlacement.roomNumber === roomNumber) {
                        const overlap = placement.days.some(day => otherPlacement.days.includes(day));
                        if (overlap) {
                            const otherCourse = courses.find(c => c.Course_ID === otherCourseId);
                            roomConflicts.push(otherCourse?.Course_Name || otherCourseId);
                        }
                    }
                }
            }
            
            if (roomConflicts.length > 0) {
                const currentCourse = courses.find(c => c.Course_ID === courseId);
                const conflictMessage = `⚠️ Room ${roomNumber} Conflict!\n\nThis room is already occupied on some of these days by:\n${roomConflicts.join('\n')}\n\nWhat would you like to do?`;
                
                // Create custom dialog with 3 buttons
                const result = await showThreeButtonDialog(
                    conflictMessage,
                    'Replace (clear conflicting rooms)',
                    'Cancel (keep current)',
                    'Draft Both (mark as drafts)'
                );
                
                if (result === 'cancel') {
                    return;
                }
                
                if (result === 'draft') {
                    // Mark all conflicting courses as draft
                    if (schedule[eventId]) {
                        for (const otherCourseId in schedule[eventId]) {
                            const otherPlacement = schedule[eventId][otherCourseId];
                            if (otherPlacement.roomNumber === roomNumber) {
                                const overlap = placement.days.some(day => otherPlacement.days.includes(day));
                                if (overlap || otherCourseId === courseId) {
                                    schedule[eventId][otherCourseId].isDraft = true;
                                    const otherCourse = courses.find(c => c.Course_ID === otherCourseId);
                                    const otherCourseName = otherCourse ? otherCourse.Course_Name : otherCourseId;
                                    logUpload('DRAFT_STATUS', otherCourseName, 0, 'Draft', `Marked as draft due to room conflict`);
                                }
                            }
                        }
                    }
                    // Assign room and mark as draft
                    schedule[eventId][courseId].roomNumber = roomNumber;
                    schedule[eventId][courseId].isDraft = true;
                    logChange('Room Selection', courseId, eventId, `Room ${roomNumber} (Draft)`, null);
                    renderSwimlanesGrid();
                    saveLogs();
                    triggerAutoSave();
                    return;
                }
                
                // result === 'replace'
                // Clear room assignments for conflicting courses
                if (schedule[eventId]) {
                    for (const otherCourseId in schedule[eventId]) {
                        if (otherCourseId === courseId) continue;
                        const otherPlacement = schedule[eventId][otherCourseId];
                        if (otherPlacement.roomNumber === roomNumber) {
                            const overlap = placement.days.some(day => otherPlacement.days.includes(day));
                            if (overlap) {
                                // Clear the room assignment but keep the day schedule
                                schedule[eventId][otherCourseId].roomNumber = null;
                                const otherCourse = courses.find(c => c.Course_ID === otherCourseId);
                                logChange('Room Selection', otherCourseId, eventId, 'None (cleared due to conflict)', `Room ${roomNumber}`);
                            }
                        }
                    }
                }
            }
        }
        
        schedule[eventId][courseId].roomNumber = roomNumber;
    }
    
    // Log the change
    const course = courses.find(c => c.Course_ID === courseId);
    logChange('Room Selection', courseId, eventId, `Room ${roomNumber}`, null);
    
    // Re-render to show updated selection
    renderSwimlanesGrid();
    saveLogs();
    triggerAutoSave();
}

// Toggle draft status for a course
function toggleDraftStatus(eventId, courseId) {
    if (!schedule[eventId] || !schedule[eventId][courseId]) return;
    
    // Check if event is locked
    if (lockedEvents.has(eventId)) {
        alert('🔒 This event is locked. Unlock it to make changes.');
        return;
    }
    
    const placement = schedule[eventId][courseId];
    placement.isDraft = !placement.isDraft;
    
    const course = courses.find(c => c.Course_ID === courseId);
    const status = placement.isDraft ? 'Draft' : 'Finalized';
    logChange('Draft Status', courseId, eventId, status, null);
    
    // Check if draft was cleared and auto-clear others if conflict resolved
    if (!placement.isDraft && placement.roomNumber && placement.startDay) {
        checkAndClearResolvedConflicts(eventId, courseId);
    }
    
    renderSwimlanesGrid();
    saveLogs();
    triggerAutoSave();
}

// Auto-clear draft status when conflicts are resolved
function checkAndClearResolvedConflicts(eventId, courseId) {
    const placement = schedule[eventId][courseId];
    if (!placement) return;
    
    const roomNumber = placement.roomNumber;
    if (!roomNumber) return;
    
    // Find other courses in same room
    const conflictingCourses = [];
    for (const otherCourseId in schedule[eventId]) {
        if (otherCourseId === courseId) continue;
        const otherPlacement = schedule[eventId][otherCourseId];
        if (otherPlacement.roomNumber === roomNumber && otherPlacement.isDraft) {
            // Check for day overlap
            const overlap = placement.days.some(day => otherPlacement.days.includes(day));
            if (!overlap) {
                // No overlap means conflict is resolved, clear draft
                schedule[eventId][otherCourseId].isDraft = false;
                logChange('Draft Status', otherCourseId, eventId, 'Auto-finalized (conflict resolved)', null);
            }
        }
    }
}

// Setup drag and drop for room grid view
function setupDragAndDropGrid() {
    const blocks = document.querySelectorAll('.course-block');
    const timelines = document.querySelectorAll('.course-timeline');
    
    blocks.forEach(block => {
        block.addEventListener('dragstart', handleBlockDragStartGrid);
        block.addEventListener('dragend', handleBlockDragEndGrid);
    });
    
    timelines.forEach(timeline => {
        timeline.addEventListener('dragover', handleTimelineDragOverGrid);
        timeline.addEventListener('dragleave', handleTimelineDragLeaveGrid);
        timeline.addEventListener('drop', handleTimelineDropGrid);
    });
}

// Drag handlers for grid view - reuse most logic from original
function handleBlockDragStartGrid(e) {
    draggedBlock = {
        courseId: e.target.dataset.courseId,
        eventId: e.target.dataset.eventId,
        daysNeeded: parseInt(e.target.dataset.daysNeeded)
    };
    e.target.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
}

function handleBlockDragEndGrid(e) {
    e.target.classList.remove('dragging');
    draggedBlock = null;
    document.querySelectorAll('.drop-indicator').forEach(el => el.remove());
}

function handleTimelineDragOverGrid(e) {
    e.preventDefault();
    
    if (!draggedBlock) return;
    
    const timelineCourseId = e.currentTarget.dataset.courseId;
    const timelineEventId = e.currentTarget.dataset.eventId;
    
    if (timelineCourseId !== draggedBlock.courseId || timelineEventId !== draggedBlock.eventId) {
        return;
    }
    
    e.dataTransfer.dropEffect = 'move';
    
    const timeline = e.currentTarget;
    const totalDays = parseInt(timeline.dataset.totalDays);
    const rect = timeline.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const dayWidth = rect.width / totalDays;
    
    let targetDay = Math.floor(x / dayWidth) + 1;
    const snapDay = calculateSnapPosition(targetDay, draggedBlock.daysNeeded, totalDays);
    
    if (snapDay !== null) {
        showDropIndicator(timeline, snapDay, draggedBlock.daysNeeded, totalDays);
    }
}

function handleTimelineDragLeaveGrid(e) {
    const relatedTarget = e.relatedTarget;
    if (!relatedTarget || !e.currentTarget.contains(relatedTarget)) {
        e.currentTarget.querySelectorAll('.drop-indicator').forEach(el => el.remove());
    }
}

function handleTimelineDropGrid(e) {
    e.preventDefault();
    
    if (!draggedBlock) return;
    
    const timeline = e.currentTarget;
    const timelineCourseId = timeline.dataset.courseId;
    const timelineEventId = timeline.dataset.eventId;
    
    // Check if event is locked
    if (lockedEvents.has(timelineEventId)) {
        alert('🔒 This event is locked. Unlock it to make changes.');
        return;
    }
    
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
        const blockedDaysStr = timeline.dataset.blockedDays || '';
        const blockedDays = blockedDaysStr.length > 0 ? blockedDaysStr.split(',').map(d => parseInt(d)) : [];
        
        const courseDays = [];
        for (let i = snapDay; i < snapDay + draggedBlock.daysNeeded; i++) {
            courseDays.push(i);
        }
        
        const hasConflict = courseDays.some(day => blockedDays.includes(day));
        
        if (hasConflict) {
            alert(`Cannot place course on these days. Instructor unavailable on: ${blockedDays.filter(d => courseDays.includes(d)).join(', ')}`);
            return;
        }
        
        // Get room number from the current assignment (use room 1 if not selected)
        const currentPlacement = schedule[draggedBlock.eventId]?.[draggedBlock.courseId];
        const roomNumber = currentPlacement?.roomNumber ?? 1;
        
        // Check if room is selected (not null)
        if (roomNumber === null || currentPlacement?.roomNumber === null) {
            alert('⚠️ Please select a room first!\n\nClick a room number below the course to assign it to a room before scheduling days.');
            return;
        }
        
        const existingPlacement = schedule[draggedBlock.eventId]?.[draggedBlock.courseId];
        const oldDays = existingPlacement ? existingPlacement.days : null;
        const action = oldDays ? 'CHANGE' : 'ADD';
        
        // Check for room conflicts
        const roomConflicts = [];
        if (schedule[draggedBlock.eventId]) {
            for (const otherCourseId in schedule[draggedBlock.eventId]) {
                if (otherCourseId === draggedBlock.courseId) continue;
                const otherPlacement = schedule[draggedBlock.eventId][otherCourseId];
                if (otherPlacement.roomNumber === roomNumber) {
                    const overlap = courseDays.some(day => otherPlacement.days.includes(day));
                    if (overlap) {
                        const otherCourse = courses.find(c => c.Course_ID === otherCourseId);
                        roomConflicts.push(otherCourse?.Course_Name || otherCourseId);
                    }
                }
            }
        }
        
        if (roomConflicts.length > 0) {
            alert(`⚠️ Room ${roomNumber} Conflict!\n\nThis room is already occupied on some of these days by:\n${roomConflicts.join('\n')}\n\nPlease select a different room using the room grid.`);
            return;
        }
        
        // Save placement
        if (!schedule[draggedBlock.eventId]) {
            schedule[draggedBlock.eventId] = {};
        }
        
        schedule[draggedBlock.eventId][draggedBlock.courseId] = {
            startDay: snapDay,
            days: courseDays,
            roomNumber: roomNumber
        };
        
        // Log the change
        logChange(action, draggedBlock.courseId, draggedBlock.eventId, courseDays, oldDays);
        
        // Re-render
        renderSwimlanesGrid();
        updateStats();
        updateReportsGrid();
        saveLogs();
        triggerAutoSave();
    }
}

// Populate course dropdown for room grid view
function populateCourseDropdownGrid(eventId, assignedCourses) {
    const dropdown = document.getElementById(`add-course-grid-${eventId}`);
    if (!dropdown) return;
    
    while (dropdown.options.length > 1) {
        dropdown.remove(1);
    }
    
    const assignedIds = new Set(assignedCourses.map(c => c.Course_ID));
    const event = events.find(e => e.Event_ID === eventId);
    const totalDays = event ? parseInt(event.Total_Days) : 0;
    
    courses.forEach(course => {
        if (!assignedIds.has(course.Course_ID)) {
            const option = document.createElement('option');
            option.value = course.Course_ID;
            
            const duration = parseFloat(course.Duration_Days);
            const daysNeeded = Math.ceil(duration);
            
            const warningIcon = daysNeeded > totalDays ? '⚠️ ' : '';
            const blockedDays = getBlockedDays(course.Instructor, eventId);
            const availableDays = totalDays - blockedDays.length;
            const availWarning = availableDays < daysNeeded ? '❌ ' : '';
            
            option.textContent = `${warningIcon}${availWarning}${course.Instructor} - ${course.Course_Name} (${course.Duration_Days} days)`;
            dropdown.appendChild(option);
        }
    });
}

// Add course to event from dropdown in grid view
function addCourseToEventGrid(eventId, courseId, selectElement) {
    if (!courseId) return;
    
    // Check if event is locked
    if (lockedEvents.has(eventId)) {
        alert('🔒 This event is locked. Unlock it to make changes.');
        selectElement.value = '';
        return;
    }
    
    handleAssignmentChange(courseId, eventId, true);
    
    if (!schedule[eventId]) {
        schedule[eventId] = {};
    }
    if (!schedule[eventId][courseId]) {
        schedule[eventId][courseId] = {
            startDay: null,
            days: [],
            roomNumber: null,
            isDraft: false
        };
    }
    
    selectElement.value = '';
    
    renderSwimlanesGrid();
    updateReportsGrid();
    triggerAutoSave();
}

// Remove course from event in grid view
function removeCourseFromEventGrid(courseId, eventId) {
    // Check if event is locked
    if (lockedEvents.has(eventId)) {
        alert('🔒 This event is locked. Unlock it to make changes.');
        return;
    }
    
    const oldDays = schedule[eventId]?.[courseId]?.days || null;
    
    if (assignments[courseId]) {
        assignments[courseId] = assignments[courseId].filter(id => id !== eventId);
    }
    
    if (schedule[eventId]) {
        delete schedule[eventId][courseId];
    }
    
    logChange('REMOVE', courseId, eventId, null, oldDays);
    
    const checkbox = document.querySelector(`input[data-course-id="${courseId}"][data-event-id="${eventId}"]`);
    if (checkbox) {
        checkbox.checked = false;
    }
    
    renderSwimlanesGrid();
    updateStats();
    updateConfigureDaysButton();
    saveLogs();
    triggerAutoSave();
}

// Toggle event swimlane in grid view
function toggleEventSwimlaneGrid(eventId) {
    const body = document.getElementById(`body-grid-${eventId}`);
    const toggle = document.getElementById(`toggle-grid-${eventId}`);
    
    if (body.classList.contains('collapsed')) {
        body.classList.remove('collapsed');
        toggle.textContent = '▼ Collapse';
    } else {
        body.classList.add('collapsed');
        toggle.textContent = '▶ Expand';
    }
}

// Toggle All Open Bookings section
function toggleAllOpenBookings() {
    const content = document.getElementById('all-open-bookings-content');
    const toggle = document.getElementById('toggle-all-bookings');
    
    if (content.style.display === 'none') {
        content.style.display = 'block';
        toggle.textContent = '▼ Collapse';
    } else {
        content.style.display = 'none';
        toggle.textContent = '▶ Expand';
    }
}

// Build content for All Open Bookings section
function buildAllOpenBookingsContent() {
    const contentDiv = document.getElementById('all-open-bookings-content');
    if (!contentDiv) return;
    
    let allBookingsHTML = '';
    
    // Loop through all non-virtual events
    events.forEach(event => {
        const eventId = event.Event_ID;
        const eventName = event.Event;
        const totalDays = parseInt(event['Total_Days']);
        const isVirtual = isVirtualEvent(eventName);
        
        // Skip virtual events
        if (isVirtual) return;
        
        const numRooms = eventRooms[eventId] || 1;
        const days = eventDays.filter(d => d.Event_ID === eventId);
        
        // Calculate room availability
        const roomAvailability = {};
        for (let roomNum = 1; roomNum <= numRooms; roomNum++) {
            roomAvailability[roomNum] = new Set();
            for (let day = 1; day <= totalDays; day++) {
                roomAvailability[roomNum].add(day);
            }
        }
        
        // Remove occupied days
        if (schedule[eventId]) {
            for (const courseId in schedule[eventId]) {
                const placement = schedule[eventId][courseId];
                if (placement.roomNumber && placement.days && placement.days.length > 0) {
                    placement.days.forEach(day => {
                        roomAvailability[placement.roomNumber]?.delete(day);
                    });
                }
            }
        }
        
        // Check if there's any availability
        let hasAvailability = false;
        for (let roomNum = 1; roomNum <= numRooms; roomNum++) {
            if (roomAvailability[roomNum].size > 0) {
                hasAvailability = true;
                break;
            }
        }
        
        if (!hasAvailability) return;
        
        // Build availability HTML for this event
        let eventAvailabilityHTML = `<div style="margin-bottom: 25px; padding: 15px; background: #f8f9fa; border-radius: 10px; border: 2px solid #28a745;"><div style="font-weight: 700; color: #667eea; font-size: 1.1em; margin-bottom: 15px;">${eventName}</div>`;
        
        for (let roomNum = 1; roomNum <= numRooms; roomNum++) {
            const availableDays = Array.from(roomAvailability[roomNum]).sort((a, b) => a - b);
            if (availableDays.length === 0) continue;
            
            // Group consecutive days into ranges
            const ranges = [];
            let rangeStart = availableDays[0];
            let rangeEnd = availableDays[0];
            
            for (let i = 1; i < availableDays.length; i++) {
                if (availableDays[i] === rangeEnd + 1) {
                    rangeEnd = availableDays[i];
                } else {
                    ranges.push({ start: rangeStart, end: rangeEnd });
                    rangeStart = availableDays[i];
                    rangeEnd = availableDays[i];
                }
            }
            ranges.push({ start: rangeStart, end: rangeEnd });
            
            // Create bars for each range
            ranges.forEach((range, rangeIndex) => {
                const daysInRange = range.end - range.start + 1;
                const blockWidth = (100 / totalDays) * daysInRange;
                const blockLeft = ((range.start - 1) / totalDays) * 100;
                
                // Find suitable courses
                const suitableCourses = courses.filter(course => {
                    const courseId = course.Course_ID;
                    const courseDuration = Math.ceil(parseFloat(course.Duration_Days));
                    
                    if (courseDuration > daysInRange) return false;
                    
                    const placement = schedule[eventId]?.[courseId];
                    if (placement && placement.days && placement.days.length > 0) {
                        return false;
                    }
                    
                    const blockedDays = getBlockedDays(course.Instructor, eventId);
                    const rangeDays = [];
                    for (let d = range.start; d <= range.end; d++) {
                        rangeDays.push(d);
                    }
                    const hasConflict = rangeDays.some(day => blockedDays.includes(day));
                    if (hasConflict) return false;
                    
                    return true;
                });
                
                let dropdownOptions = '<option value="">+ Add Draft Option...</option>';
                suitableCourses.forEach(course => {
                    const courseDuration = Math.ceil(parseFloat(course.Duration_Days));
                    dropdownOptions += `<option value="${course.Course_ID}">${course.Instructor} - ${course.Course_Name} (${courseDuration} ${courseDuration === 1 ? 'day' : 'days'})</option>`;
                });
                
                const uniqueId = `all-room-${roomNum}-range-${rangeIndex}-event-${eventId}`;
                
                eventAvailabilityHTML += `
                    <div style="background: white; border-radius: 8px; padding: 12px; margin-bottom: 10px; border: 2px solid #28a745;">
                        <div style="display: flex; align-items: center; gap: 15px; margin-bottom: 10px;">
                            <div style="min-width: 150px; font-weight: 700; color: #28a745;">
                                🟢 Room ${roomNum}
                            </div>
                            <div style="flex: 1; position: relative; min-height: 40px; background: #f8f9fa; border-radius: 8px; padding: 5px;">
                                <div style="position: absolute; left: ${blockLeft}%; width: ${blockWidth}%; top: 5px; height: 30px; background: linear-gradient(135deg, #28a745 0%, #20c997 100%); color: white; border-radius: 6px; display: flex; align-items: center; justify-content: center; font-weight: 600; font-size: 0.85em; box-shadow: 0 2px 8px rgba(40, 167, 69, 0.3);">
                                    Days ${range.start}${range.start !== range.end ? `-${range.end}` : ''} (${daysInRange} ${daysInRange === 1 ? 'day' : 'days'} available)
                                </div>
                            </div>
                        </div>
                        ${suitableCourses.length > 0 ? `
                            <div style="display: flex; gap: 10px; align-items: stretch;">
                                <select class="add-course-dropdown" id="draft-select-${uniqueId}" onchange="addToDraftList('${eventId}', ${roomNum}, ${range.start}, ${range.end}, '${uniqueId}', this.value); this.value='';" style="flex: 1;">
                                    ${dropdownOptions}
                                </select>
                            </div>
                            <div id="draft-list-${uniqueId}" style="margin-top: 10px; display: none;">
                                <div style="font-weight: 600; color: #667eea; margin-bottom: 5px; font-size: 0.9em;">📋 Draft Selections:</div>
                                <div id="draft-items-${uniqueId}"></div>
                            </div>
                        ` : '<div style="color: #6c757d; font-size: 0.9em; padding: 8px; text-align: center; background: #f8f9fa; border-radius: 6px;">No courses fit in this time slot</div>'}
                    </div>
                `;
            });
        }
        
        eventAvailabilityHTML += '</div>';
        allBookingsHTML += eventAvailabilityHTML;
    });
    
    if (allBookingsHTML === '') {
        contentDiv.innerHTML = '<div style="text-align: center; color: #6c757d; padding: 30px; font-size: 1.1em;">🎉 No open bookings - all events are fully scheduled!</div>';
    } else {
        contentDiv.innerHTML = allBookingsHTML;
    }
}

// Toggle event lock status
function toggleEventLock(eventId) {
    const event = events.find(e => e.Event_ID === eventId);
    const eventName = event ? event.Event : eventId;
    
    if (lockedEvents.has(eventId)) {
        lockedEvents.delete(eventId);
        logUpload('EVENT', eventName, 0, 'Unlocked', `Event "${eventName}" unlocked for editing`);
    } else {
        lockedEvents.add(eventId);
        logUpload('EVENT', eventName, 0, 'Locked', `Event "${eventName}" locked to prevent changes`);
    }
    
    triggerAutoSave();
    renderSwimlanesGrid();
}

// Toggle room availability section
function toggleRoomAvailability(eventId) {
    const content = document.getElementById(`room-avail-content-${eventId}`);
    const icon = document.getElementById(`room-avail-icon-${eventId}`);
    
    if (content.style.display === 'none') {
        content.style.display = 'block';
        icon.textContent = '▼';
    } else {
        content.style.display = 'none';
        icon.textContent = '▶';
    }
}

// Track draft lists per availability slot
const draftLists = {};

// Add course to draft list
function addToDraftList(eventId, roomNumber, startDay, endDay, uniqueId, courseId) {
    if (!courseId) return;
    
    // Check if event is locked
    if (lockedEvents.has(eventId)) {
        alert('🔒 This event is locked. Unlock it to make changes.');
        return;
    }
    
    const course = courses.find(c => c.Course_ID === courseId);
    if (!course) return;
    
    // Initialize draft list for this slot
    if (!draftLists[uniqueId]) {
        draftLists[uniqueId] = [];
    }
    
    // Check if already in draft list
    if (draftLists[uniqueId].some(item => item.courseId === courseId)) {
        return; // Already added
    }
    
    // Add to draft list
    draftLists[uniqueId].push({
        courseId,
        courseName: course.Course_Name,
        instructor: course.Instructor,
        duration: Math.ceil(parseFloat(course.Duration_Days)),
        eventId,
        roomNumber,
        startDay,
        endDay
    });
    
    // Update UI
    updateDraftListUI(uniqueId);
}

// Update draft list UI
function updateDraftListUI(uniqueId) {
    const listContainer = document.getElementById(`draft-list-${uniqueId}`);
    const itemsContainer = document.getElementById(`draft-items-${uniqueId}`);
    
    if (!listContainer || !itemsContainer) return;
    
    const items = draftLists[uniqueId] || [];
    
    if (items.length === 0) {
        listContainer.style.display = 'none';
        return;
    }
    
    listContainer.style.display = 'block';
    
    itemsContainer.innerHTML = items.map((item, index) => `
        <div style="display: flex; align-items: center; gap: 10px; padding: 8px; background: white; border-radius: 6px; margin-bottom: 5px; border: 2px solid #667eea;">
            <div style="flex: 1;">
                <div style="font-weight: 600; color: #667eea;">${item.instructor} - ${item.courseName}</div>
                <div style="font-size: 0.85em; color: #6c757d;">${item.duration} ${item.duration === 1 ? 'day' : 'days'}</div>
            </div>
            <button class="btn btn-success btn-small" onclick="applyDraftCourse('${uniqueId}', ${index})" style="white-space: nowrap;">
                ✓ Apply
            </button>
            <button class="btn btn-danger btn-small" onclick="removeDraftCourse('${uniqueId}', ${index})" style="white-space: nowrap;">
                ✗ Remove
            </button>
        </div>
    `).join('');
}

// Remove course from draft list
function removeDraftCourse(uniqueId, index) {
    if (!draftLists[uniqueId]) return;
    
    // Check if event is locked (get eventId from the draft item)
    const item = draftLists[uniqueId][index];
    if (item && lockedEvents.has(item.eventId)) {
        alert('🔒 This event is locked. Unlock it to make changes.');
        return;
    }
    
    draftLists[uniqueId].splice(index, 1);
    updateDraftListUI(uniqueId);
}

// Apply single draft course
function applyDraftCourse(uniqueId, index) {
    if (!draftLists[uniqueId] || !draftLists[uniqueId][index]) return;
    
    const item = draftLists[uniqueId][index];
    const { courseId, eventId, roomNumber, startDay, endDay, duration } = item;
    
    // Check if event is locked
    if (lockedEvents.has(eventId)) {
        alert('🔒 This event is locked. Unlock it to make changes.');
        return;
    }
    
    const daysAvailable = endDay - startDay + 1;
    
    // Check if course still fits
    if (duration > daysAvailable) {
        alert('This course no longer fits in the available space.');
        return;
    }
    
    // Check if still unscheduled
    const placement = schedule[eventId]?.[courseId];
    if (placement && placement.days && placement.days.length > 0) {
        alert('This course has already been scheduled.');
        removeDraftCourse(uniqueId, index);
        return;
    }
    
    // Calculate days this course will occupy
    const courseDays = [];
    for (let i = 0; i < duration; i++) {
        courseDays.push(startDay + i);
    }
    
    // Check for instructor conflicts
    const course = courses.find(c => c.Course_ID === courseId);
    const blockedDays = getBlockedDays(course.Instructor, eventId);
    const hasConflict = courseDays.some(day => blockedDays.includes(day));
    
    if (hasConflict) {
        alert(`Cannot schedule this course. Instructor ${course.Instructor} is unavailable on some of these days.`);
        return;
    }
    
    // Apply the course
    if (!schedule[eventId]) {
        schedule[eventId] = {};
    }
    
    // Add to assignments if not already assigned to this event
    if (!assignments[courseId]) {
        assignments[courseId] = [];
    }
    if (!assignments[courseId].includes(eventId)) {
        assignments[courseId].push(eventId);
    }
    
    schedule[eventId][courseId] = {
        startDay: startDay,
        days: courseDays,
        roomNumber: roomNumber,
        isDraft: false
    };
    
    logChange('ADD', courseId, eventId, courseDays, null);
    
    // Remove from draft list
    removeDraftCourse(uniqueId, index);
    
    // Re-render and save
    renderSwimlanesGrid();
    updateStats();
    updateReportsGrid();
    saveLogs();
    triggerAutoSave();
}
