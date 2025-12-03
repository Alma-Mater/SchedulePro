# 📅 SchedulePro - Event Course Scheduler

**Easy-to-use scheduling software for event planners**

SchedulePro helps you schedule courses with instructors across multiple events, manage course durations, and export beautiful schedules to Excel.

## ✨ Features

- 📁 **Import courses from CSV** - Load your courses, instructors, and durations
- 🎯 **Drag & Drop Interface** - Intuitive course assignment to events
- 📆 **Smart Day Selection** - Automatically suggests day ranges based on course duration
- 📊 **Visual Schedule Grid** - See all your events and courses at a glance
- 📈 **Real-time Statistics** - Track total courses, assignments, and events
- 💾 **Save & Load** - Backup your work as JSON files
- 📤 **Excel Export** - Export schedules to CSV (Excel-compatible)
- 🎨 **Beautiful UI** - Professional gradient design

## 🚀 Quick Start

### 1. Open the Application

Simply open `index.html` in your web browser. No installation or server required!

```powershell
# Navigate to the project folder
cd "C:\Users\Dell\Documents\SchedulePro"

# Open in default browser
start index.html
```

### 2. Prepare Your Data

The application needs:
- **Events Data** (already set up in `Data/events.csv`)
- **Event Days** (already set up in `Data/event_days.csv`)
- **Courses CSV** (you need to create this)

### 3. Create Your Courses CSV

Use the provided template (`Data/courses_template.csv`) or create your own with these columns:

```csv
Course_ID,Instructor,Course_Name,Duration_Days
C001,Alfred,Mapmaking,3
C002,Betty,Cooking Basics,2
C003,Charlie,Advanced Photography,4
C004,Diana,Web Design,3.5
C005,Edward,Public Speaking,1
```

**Requirements:**
- `Course_ID`: Unique identifier (e.g., C001, C002)
- `Instructor`: Instructor name
- `Course_Name`: Name of the course
- `Duration_Days`: Length in days (0.5 to 4 days, can use decimals)

## 📖 How to Use

### Step 1: Load Your Courses
1. Click **"📁 Load Courses CSV"**
2. Select your courses CSV file
3. Your courses will appear in the "Available Courses" section

### Step 2: Assign Courses to Events
1. **Drag** a course card from the "Available Courses" section
2. **Drop** it onto any day slot in an event below
3. A modal will appear asking you to select which days

### Step 3: Select Specific Days
1. The modal shows available day ranges based on course duration
2. Click on the day range you want (e.g., "Days 1-3")
3. Click **"✓ Confirm"** to assign the course

### Step 4: Fine-Tune Your Schedule
- **Remove assignments**: Click the **×** button on any assigned course
- **Reassign courses**: Just drag and drop again
- **Multiple courses per event**: Drop multiple courses onto different days

### Step 5: Export Your Schedule
- Click **"📊 Export to Excel"** to download a CSV file
- Open in Microsoft Excel or Google Sheets
- Click **"💾 Save Schedule"** to backup your work as JSON
- Click **"📂 Load Schedule"** to restore a previous session

## 📁 File Structure

```
SchedulePro/
├── index.html              # Main application
├── scheduler.js            # JavaScript logic
├── convert_to_csv.py       # Convert Excel to CSV (helper)
├── README.md              # This file
└── Data/
    ├── events.csv         # 12 events with durations
    ├── event_days.csv     # All event days with dates
    ├── Events.xlsx        # Original Excel file
    ├── Event_Days.xlsx    # Original Excel file
    └── courses_template.csv  # Sample courses template
```

## 🔧 Technical Details

### Events
- **Total Events**: 12 pre-planned events
- **Duration**: Most events are 4 days (Detroit is 2 days)
- **Format**: Each event has specific dates in 2026

### Courses
- **Duration Range**: 0.5 to 4 days
- **Assignment**: One instructor per course
- **Flexibility**: Same course can be at multiple events
- **Concurrent**: Multiple courses can run simultaneously

### Browser Compatibility
- ✅ Chrome/Edge (Recommended)
- ✅ Firefox
- ✅ Safari
- No internet connection required!

## 💡 Tips & Best Practices

1. **Start with a template**: Download the course template and fill it with your data
2. **Save often**: Use "Save Schedule" to backup your work regularly
3. **Export early**: Test exports to make sure everything looks right
4. **Multiple assignments**: The same course-instructor can be at different events
5. **Visual feedback**: Assigned courses show a green checkmark in the courses list

## 🎯 Example Workflow

Let's schedule Alfred's Mapmaking course (3 days) to Orlando (4 days):

1. **Load courses.csv** with Alfred-Mapmaking listed as 3 days
2. **Drag** the "Alfred - Mapmaking" card
3. **Drop** it onto Orlando's Day 1 slot
4. **Select** "Days 1-3" in the modal (since it's 3 days long)
5. **Confirm** - The course now appears on Days 1, 2, and 3 of Orlando
6. **Export** to see the schedule in Excel

## 📊 Export Format

The Excel export includes:
- Event ID and Name
- Day number and date
- Course ID, Instructor, Course Name
- Duration in days

Perfect for:
- Sharing with stakeholders
- Printing schedules
- Further analysis in Excel

## 🐛 Troubleshooting

**Courses won't load?**
- Check CSV format matches template
- Ensure all required columns exist
- No special characters in CSV

**Can't see events?**
- Make sure `Data/events.csv` and `Data/event_days.csv` exist
- Run `convert_to_csv.py` if needed

**Drag and drop not working?**
- Use a modern browser (Chrome, Edge, Firefox)
- Make sure JavaScript is enabled

## 🔄 Converting Excel Files to CSV

If you need to update the events data:

```powershell
# Make sure Python environment is activated
C:/Users/Dell/Documents/SchedulePro/.venv/Scripts/python.exe convert_to_csv.py
```

This will regenerate:
- `Data/events.csv`
- `Data/event_days.csv`

## 📝 Notes

- All events are in 2026
- Detroit is the only 2-day event (all others are 4 days)
- Course durations can be decimals (e.g., 0.5, 1.5, 2.5)
- No events overlap, so instructors can teach at multiple events

## 🎨 Features Coming Soon

- Visual calendar view with color coding
- Conflict detection (overlapping courses)
- Print-friendly layouts
- Bulk course assignment
- Search and filter courses

## 📧 Support

For questions or issues, refer to this README or check the code comments in `scheduler.js`.

---

**Happy Scheduling! 📅✨**
