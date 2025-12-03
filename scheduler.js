// SchedulePro - Main JavaScript
// Data structures
let courses = [];
let events = [];
let eventDays = [];
let schedule = {}; // { eventId: { dayNum: [courseAssignments] } }
let draggedCourse = null;
let currentDropTarget = null;

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    loadEvents();
    setupFileInput();
});

// Load events and event days from CSV
async function loadEvents() {
    try {
        const eventsData = await fetchCSV('Data/events.csv');
        const daysData = await fetchCSV('Data/event_days.csv');
        
        events = parseCSV(eventsData);
        eventDays = parseCSV(daysData);
        
        renderScheduleGrid();
        updateStats();
    } catch (error) {
        console.error('Error loading events:', error);
        document.getElementById('scheduleGrid').innerHTML = 
            '<div class="alert alert-warning">Error loading event data. Make sure Data/events.csv and Data/event_days.csv exist.</div>';
    }
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
    const headers = lines[0].split(',').map(h => h.trim());
    
    return lines.slice(1).map(line => {
        const values = parseCSVLine(line);
        const obj = {};
        headers.forEach((header, index) => {
            obj[header] = values[index] ? values[index].trim() : '';
        });
        return obj;
    });
}

// Parse a single CSV line (handles quoted values with commas)
function parseCSVLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        if (char === '"') {
            inQuotes = !inQuotes;
        } else if (char === ',' && !inQuotes) {
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
        const requiredColumns = ['Course_ID', 'Instructor', 'Course_Name', 'Duration_Days'];
        const hasAllColumns = requiredColumns.every(col => col in courses[0]);
        
        if (!hasAllColumns) {
            alert('CSV must have columns: Course_ID, Instructor, Course_Name, Duration_Days');
            return;
        }
        
        renderCoursesList();
        updateStats();
    });
}

// Render courses list
function renderCoursesList() {
    const container = document.getElementById('coursesList');
    
    if (courses.length === 0) {
        container.innerHTML = '<div class="no-courses">Load a courses CSV file to get started</div>';
        return;
    }
    
    container.innerHTML = courses.map(course => {
        const isAssigned = isCourseAssigned(course.Course_ID);
        return `
            <div class="course-card ${isAssigned ? 'assigned' : ''}" 
                 draggable="true"
                 data-course-id="${course.Course_ID}"
                 ondragstart="handleDragStart(event)"
                 ondragend="handleDragEnd(event)">
                <div class="course-instructor">${course.Instructor}</div>
                <div class="course-name">${course.Course_Name}</div>
                <div class="course-duration">📏 ${course.Duration_Days} days</div>
                ${isAssigned ? '<div style="color: #28a745; font-size: 0.8em; margin-top: 5px;">✓ Assigned</div>' : ''}
            </div>
        `;
    }).join('');
}

// Check if course is assigned anywhere
function isCourseAssigned(courseId) {
    for (const eventId in schedule) {
        for (const dayNum in schedule[eventId]) {
            if (schedule[eventId][dayNum].some(a => a.courseId === courseId)) {
                return true;
            }
        }
    }
    return false;
}

// Render schedule grid
function renderScheduleGrid() {
    const container = document.getElementById('scheduleGrid');
    
    if (events.length === 0) {
        container.innerHTML = '<div class="loading">No events loaded</div>';
        return;
    }
    
    container.innerHTML = events.map(event => {
        const eventId = event.Event_ID;
        const totalDays = parseInt(event['Total Days']);
        const eventName = event.Event;
        
        // Get days for this event
        const days = eventDays.filter(d => 
            d.Event.toLowerCase().includes(eventName.toLowerCase().split('-')[0])
        ).slice(0, totalDays);
        
        return `
            <div class="event-card" data-event-id="${eventId}">
                <div class="event-header">
                    <div>
                        <div class="event-name">${eventName}</div>
                        <div class="event-info">${eventId}</div>
                    </div>
                    <div class="event-info">${totalDays} days</div>
                </div>
                <div class="event-days">
                    ${days.map((day, index) => {
                        const dayNum = index + 1;
                        const assignments = getAssignmentsForDay(eventId, dayNum);
                        
                        return `
                            <div class="day-slot" 
                                 data-event-id="${eventId}"
                                 data-day-num="${dayNum}"
                                 ondragover="handleDragOver(event)"
                                 ondragleave="handleDragLeave(event)"
                                 ondrop="handleDrop(event)">
                                <div class="day-label">Day ${dayNum} - ${day['Date text'] || ''}</div>
                                ${assignments.map(assignment => renderAssignment(assignment, eventId, dayNum)).join('')}
                            </div>
                        `;
                    }).join('')}
                </div>
            </div>
        `;
    }).join('');
}

// Get assignments for a specific day
function getAssignmentsForDay(eventId, dayNum) {
    if (!schedule[eventId] || !schedule[eventId][dayNum]) {
        return [];
    }
    return schedule[eventId][dayNum];
}

// Render a course assignment
function renderAssignment(assignment, eventId, dayNum) {
    const course = courses.find(c => c.Course_ID === assignment.courseId);
    if (!course) return '';
    
    const dayRange = assignment.days.length > 1 
        ? `Days ${Math.min(...assignment.days)}-${Math.max(...assignment.days)}`
        : `Day ${assignment.days[0]}`;
    
    return `
        <div class="assigned-course">
            <button class="remove-btn" onclick="removeAssignment('${eventId}', ${dayNum}, '${assignment.courseId}')">×</button>
            <div class="assigned-course-name">${course.Course_Name}</div>
            <div class="assigned-course-instructor">${course.Instructor} (${course.Duration_Days} days)</div>
            <div style="font-size: 0.85em; margin-top: 3px;">${dayRange}</div>
        </div>
    `;
}

// Drag and drop handlers
function handleDragStart(event) {
    const courseId = event.target.dataset.courseId;
    draggedCourse = courses.find(c => c.Course_ID === courseId);
    event.target.classList.add('dragging');
    event.dataTransfer.effectAllowed = 'move';
}

function handleDragEnd(event) {
    event.target.classList.remove('dragging');
    draggedCourse = null;
}

function handleDragOver(event) {
    event.preventDefault();
    event.dataTransfer.dropEffect = 'move';
    
    const daySlot = event.currentTarget;
    if (!daySlot.classList.contains('drag-over')) {
        daySlot.classList.add('drag-over');
    }
}

function handleDragLeave(event) {
    const daySlot = event.currentTarget;
    daySlot.classList.remove('drag-over');
}

function handleDrop(event) {
    event.preventDefault();
    const daySlot = event.currentTarget;
    daySlot.classList.remove('drag-over');
    
    if (!draggedCourse) return;
    
    const eventId = daySlot.dataset.eventId;
    const dayNum = parseInt(daySlot.dataset.dayNum);
    
    // Open modal to select which days
    openDaySelectionModal(draggedCourse, eventId, dayNum);
}

// Day selection modal
function openDaySelectionModal(course, eventId, startDay) {
    const modal = document.getElementById('dayModal');
    const event = events.find(e => e.Event_ID === eventId);
    const totalDays = parseInt(event['Total Days']);
    const courseDuration = parseFloat(course.Duration_Days);
    
    // Calculate how many days are needed
    const daysNeeded = Math.ceil(courseDuration);
    
    // Show course info
    document.getElementById('modalCourseInfo').innerHTML = `
        <div class="alert alert-info">
            <strong>${course.Instructor} - ${course.Course_Name}</strong><br>
            Duration: ${course.Duration_Days} days<br>
            Event: ${event.Event} (${totalDays} days)<br>
            Select which consecutive days this course will occupy:
        </div>
    `;
    
    // Generate day options
    const dayOptions = [];
    for (let i = 1; i <= totalDays - daysNeeded + 1; i++) {
        dayOptions.push(i);
    }
    
    document.getElementById('daySelector').innerHTML = dayOptions.map(startDayOption => {
        const endDay = startDayOption + daysNeeded - 1;
        return `
            <div class="day-option" 
                 data-start-day="${startDayOption}"
                 data-end-day="${endDay}"
                 onclick="selectDayOption(this)">
                <div style="font-weight: 700;">Days ${startDayOption}-${endDay}</div>
                <div style="font-size: 0.9em; margin-top: 5px;">${daysNeeded} days</div>
            </div>
        `;
    }).join('');
    
    // Store current selection context
    modal.dataset.courseId = course.Course_ID;
    modal.dataset.eventId = eventId;
    modal.dataset.suggestedDay = startDay;
    
    modal.classList.add('active');
}

function selectDayOption(element) {
    // Remove previous selection
    document.querySelectorAll('.day-option').forEach(opt => opt.classList.remove('selected'));
    // Add new selection
    element.classList.add('selected');
}

function confirmDaySelection() {
    const modal = document.getElementById('dayModal');
    const selectedOption = document.querySelector('.day-option.selected');
    
    if (!selectedOption) {
        alert('Please select which days for this course');
        return;
    }
    
    const courseId = modal.dataset.courseId;
    const eventId = modal.dataset.eventId;
    const startDay = parseInt(selectedOption.dataset.startDay);
    const endDay = parseInt(selectedOption.dataset.endDay);
    
    // Create days array
    const days = [];
    for (let i = startDay; i <= endDay; i++) {
        days.push(i);
    }
    
    // Add assignment to schedule
    if (!schedule[eventId]) {
        schedule[eventId] = {};
    }
    
    // Add to all days this course spans
    days.forEach(day => {
        if (!schedule[eventId][day]) {
            schedule[eventId][day] = [];
        }
        // Check if this course is already assigned to this day
        const existingIndex = schedule[eventId][day].findIndex(a => a.courseId === courseId);
        if (existingIndex === -1) {
            schedule[eventId][day].push({
                courseId: courseId,
                days: days
            });
        }
    });
    
    closeModal();
    renderScheduleGrid();
    renderCoursesList();
    updateStats();
}

function closeModal() {
    document.getElementById('dayModal').classList.remove('active');
}

// Remove assignment
function removeAssignment(eventId, dayNum, courseId) {
    if (!schedule[eventId] || !schedule[eventId][dayNum]) return;
    
    // Find the assignment to get all days it spans
    const assignment = schedule[eventId][dayNum].find(a => a.courseId === courseId);
    if (!assignment) return;
    
    // Remove from all days it spans
    assignment.days.forEach(day => {
        if (schedule[eventId][day]) {
            schedule[eventId][day] = schedule[eventId][day].filter(a => a.courseId !== courseId);
            if (schedule[eventId][day].length === 0) {
                delete schedule[eventId][day];
            }
        }
    });
    
    renderScheduleGrid();
    renderCoursesList();
    updateStats();
}

// Update statistics
function updateStats() {
    document.getElementById('totalCourses').textContent = courses.length;
    
    const assignedCourseIds = new Set();
    for (const eventId in schedule) {
        for (const dayNum in schedule[eventId]) {
            schedule[eventId][dayNum].forEach(a => assignedCourseIds.add(a.courseId));
        }
    }
    document.getElementById('assignedCourses').textContent = assignedCourseIds.size;
    document.getElementById('totalEvents').textContent = events.length;
}

// Export to Excel
async function exportToExcel() {
    if (courses.length === 0) {
        alert('Please load courses first');
        return;
    }
    
    // Create CSV content
    let csv = 'Event_ID,Event,Day,Date,Course_ID,Instructor,Course_Name,Duration_Days\n';
    
    events.forEach(event => {
        const eventId = event.Event_ID;
        const eventName = event.Event;
        const totalDays = parseInt(event['Total Days']);
        
        // Get days for this event
        const days = eventDays.filter(d => 
            d.Event.toLowerCase().includes(eventName.toLowerCase().split('-')[0])
        ).slice(0, totalDays);
        
        days.forEach((day, index) => {
            const dayNum = index + 1;
            const assignments = getAssignmentsForDay(eventId, dayNum);
            
            if (assignments.length === 0) {
                csv += `${eventId},${eventName},${dayNum},${day['Date text'] || ''},,,,\n`;
            } else {
                assignments.forEach(assignment => {
                    const course = courses.find(c => c.Course_ID === assignment.courseId);
                    if (course) {
                        csv += `${eventId},${eventName},${dayNum},${day['Date text'] || ''},${course.Course_ID},${course.Instructor},${course.Course_Name},${course.Duration_Days}\n`;
                    }
                });
            }
        });
    });
    
    // Download as CSV (Excel compatible)
    downloadFile(csv, 'schedule_export.csv', 'text/csv');
    
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
            schedule = data.schedule || {};
            
            renderCoursesList();
            renderScheduleGrid();
            updateStats();
            
            alert('Schedule loaded successfully!');
        } catch (error) {
            alert('Error loading schedule: ' + error.message);
        }
    };
    
    input.click();
}

// Download template
function downloadTemplate() {
    const template = `Course_ID,Instructor,Course_Name,Duration_Days
C001,Alfred,Mapmaking,3
C002,Betty,Cooking Basics,2
C003,Charlie,Advanced Photography,4
C004,Diana,Web Design,3.5`;
    
    downloadFile(template, 'courses_template.csv', 'text/csv');
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

// Close modal when clicking outside
document.getElementById('dayModal').addEventListener('click', (e) => {
    if (e.target.id === 'dayModal') {
        closeModal();
    }
});
