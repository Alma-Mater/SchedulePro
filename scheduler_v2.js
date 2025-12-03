// SchedulePro V2 - Grid-based scheduler with swimlane day configuration
// Data structures
let courses = [];
let events = [];
let eventDays = [];
let assignments = {}; // { courseId: [eventIds] }
let schedule = {}; // { eventId: { courseId: { startDay, days: [] } } }
let draggedBlock = null;
let currentTimeline = null;

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
        
        renderAssignmentGrid();
        updateStats();
    } catch (error) {
        console.error('Error loading events:', error);
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
        
        renderAssignmentGrid();
        updateStats();
    });
}

// Render assignment grid
function renderAssignmentGrid() {
    const table = document.getElementById('assignmentGrid');
    const thead = table.querySelector('thead tr');
    const tbody = table.querySelector('tbody');
    
    // Clear existing content except first header
    while (thead.children.length > 1) {
        thead.removeChild(thead.lastChild);
    }
    tbody.innerHTML = '';
    
    if (events.length === 0 || courses.length === 0) {
        tbody.innerHTML = '<tr><td colspan="13" style="text-align: center; padding: 40px;">Load courses and events to begin</td></tr>';
        return;
    }
    
    // Add event headers
    events.forEach(event => {
        const th = document.createElement('th');
        th.className = 'event-header';
        th.textContent = event.Event;
        th.title = `${event.Event} (${event['Total Days']} days)`;
        thead.appendChild(th);
    });
    
    // Add course rows
    courses.forEach(course => {
        const tr = document.createElement('tr');
        
        // Course info cell
        const tdCourse = document.createElement('td');
        tdCourse.className = 'course-info';
        tdCourse.innerHTML = `
            <div class="course-name-cell">${course.Instructor} - ${course.Course_Name}</div>
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
            
            // Check if already assigned
            if (assignments[course.Course_ID]?.includes(event.Event_ID)) {
                checkbox.checked = true;
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
        }
    } else {
        assignments[courseId] = assignments[courseId].filter(id => id !== eventId);
        // Also remove from schedule if configured
        if (schedule[eventId] && schedule[eventId][courseId]) {
            delete schedule[eventId][courseId];
        }
    }
    
    updateStats();
    updateConfigureDaysButton();
}

// Update configure days button state
function updateConfigureDaysButton() {
    const btn = document.getElementById('configureDaysBtn');
    const hasAssignments = Object.values(assignments).some(arr => arr.length > 0);
    btn.disabled = !hasAssignments;
}

// Update statistics
function updateStats() {
    document.getElementById('totalCourses').textContent = courses.length;
    
    const assignedCount = Object.values(assignments).filter(arr => arr.length > 0).length;
    document.getElementById('assignedCourses').textContent = assignedCount;
    
    let configuredCount = 0;
    for (const eventId in schedule) {
        for (const courseId in schedule[eventId]) {
            if (schedule[eventId][courseId].startDay !== null) {
                configuredCount++;
            }
        }
    }
    document.getElementById('configuredCourses').textContent = configuredCount;
}

// Go to configure days view
function goToConfigureDays() {
    document.getElementById('gridView').classList.remove('active');
    document.getElementById('configureDaysView').classList.add('active');
    renderSwimlanes();
}

// Back to grid view
function backToGrid() {
    document.getElementById('configureDaysView').classList.remove('active');
    document.getElementById('gridView').classList.add('active');
}

// Render swimlanes for day configuration
function renderSwimlanes() {
    const container = document.getElementById('swimlanesContainer');
    container.innerHTML = '';
    
    events.forEach(event => {
        const eventId = event.Event_ID;
        const eventName = event.Event;
        const totalDays = parseInt(event['Total Days']);
        
        // Get courses assigned to this event
        const assignedCourses = courses.filter(course => 
            assignments[course.Course_ID]?.includes(eventId)
        );
        
        if (assignedCourses.length === 0) return;
        
        // Get days for this event
        const days = eventDays.filter(d => 
            d.Event.toLowerCase().includes(eventName.toLowerCase().split('-')[0])
        ).slice(0, totalDays);
        
        // Create swimlane
        const swimlane = document.createElement('div');
        swimlane.className = 'event-swimlane';
        swimlane.dataset.eventId = eventId;
        
        swimlane.innerHTML = `
            <div class="event-swimlane-header">
                ${eventName} (${totalDays} days)
            </div>
            <div class="event-swimlane-body">
                <div class="day-timeline" data-event-id="${eventId}">
                    ${days.map((day, index) => `
                        <div class="day-slot" data-day-num="${index + 1}">
                            <div class="day-label">Day ${index + 1}</div>
                            <div class="day-date">${day['Date text'] || ''}</div>
                        </div>
                    `).join('')}
                </div>
                ${assignedCourses.map(course => renderCourseSwimlane(course, eventId, totalDays)).join('')}
            </div>
        `;
        
        container.appendChild(swimlane);
    });
    
    // Setup drag and drop for all course blocks
    setupDragAndDrop();
}

// Render a single course swimlane
function renderCourseSwimlane(course, eventId, totalDays) {
    const courseId = course.Course_ID;
    const duration = parseFloat(course.Duration_Days);
    const daysNeeded = Math.ceil(duration);
    
    // Get current placement if exists
    const placement = schedule[eventId]?.[courseId];
    const startDay = placement?.startDay;
    
    // Calculate block width as percentage
    const blockWidth = (100 / totalDays) * daysNeeded;
    const blockLeft = startDay ? ((startDay - 1) / totalDays) * 100 : null;
    
    return `
        <div class="course-swimlane" data-course-id="${courseId}" data-event-id="${eventId}">
            <div class="course-info-sidebar">
                <div class="course-info-name">${course.Course_Name}</div>
                <div class="course-info-instructor">${course.Instructor}</div>
                <div class="course-info-duration">📏 ${course.Duration_Days} days</div>
            </div>
            <div class="course-timeline" data-course-id="${courseId}" data-event-id="${eventId}" data-total-days="${totalDays}">
                <div class="course-block ${startDay ? '' : 'unplaced'}" 
                     data-course-id="${courseId}"
                     data-event-id="${eventId}"
                     data-days-needed="${daysNeeded}"
                     draggable="true"
                     style="${startDay ? `position: absolute; left: ${blockLeft}%; width: ${blockWidth}%;` : ''}">
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
        // Save placement
        if (!schedule[draggedBlock.eventId]) {
            schedule[draggedBlock.eventId] = {};
        }
        
        const days = [];
        for (let i = snapDay; i < snapDay + draggedBlock.daysNeeded; i++) {
            days.push(i);
        }
        
        schedule[draggedBlock.eventId][draggedBlock.courseId] = {
            startDay: snapDay,
            days: days
        };
        
        // Re-render this swimlane
        renderSwimlanes();
        updateStats();
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

// Remove course from event
function removeCourseFromEvent(courseId, eventId) {
    // Remove from assignments
    if (assignments[courseId]) {
        assignments[courseId] = assignments[courseId].filter(id => id !== eventId);
    }
    
    // Remove from schedule
    if (schedule[eventId]) {
        delete schedule[eventId][courseId];
    }
    
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

// Export to Excel
function exportToExcel() {
    if (courses.length === 0) {
        alert('Please load courses first');
        return;
    }
    
    // Create CSV content
    let csv = 'Event_ID,Event,Day,Date,Course_ID,Instructor,Course_Name,Duration_Days,Configured\n';
    
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
                csv += `${eventId},${eventName},${dayNum},${day['Date text'] || ''},,,,, No\n`;
            } else {
                coursesOnDay.forEach(course => {
                    csv += `${eventId},${eventName},${dayNum},${day['Date text'] || ''},${course.Course_ID},${course.Instructor},${course.Course_Name},${course.Duration_Days},Yes\n`;
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

// Download template
function downloadTemplate() {
    const template = `Course_ID,Instructor,Course_Name,Duration_Days
C001,Alfred,Mapmaking,3
C002,Betty,Cooking Basics,2
C003,Charlie,Advanced Photography,4
C004,Diana,Web Design,3.5
C005,Edward,Public Speaking,1`;
    
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
