# 📚 SchedulePro User Guide

## Table of Contents
1. [Introduction](#introduction)
2. [Getting Started](#getting-started)
3. [Understanding Your Data](#understanding-your-data)
4. [Step-by-Step Tutorial](#step-by-step-tutorial)
5. [Advanced Features](#advanced-features)
6. [Common Scenarios](#common-scenarios)
7. [Troubleshooting](#troubleshooting)

---

## Introduction

SchedulePro is a visual scheduling tool designed specifically for event planners who need to:
- Assign instructor-led courses to multiple events
- Manage courses of varying lengths (0.5 to 4 days)
- Handle the complexity of fitting courses into event days
- Export professional schedules to Excel

### Key Concepts

**Event**: A scheduled gathering (e.g., Orlando, Chicago) with a set number of days
**Course**: A class with an instructor and specific duration
**Assignment**: Placing a course into specific days within an event

---

## Getting Started

### System Requirements
- Windows, Mac, or Linux
- Modern web browser (Chrome, Edge, Firefox, Safari)
- No internet connection required
- No installation needed

### Launch Methods

**Method 1**: Double-click `index.html`

**Method 2**: Use the batch file
```
Double-click START_APP.bat
```

**Method 3**: PowerShell
```powershell
start index.html
```

### First Launch

When you first open SchedulePro, you'll see:
1. **Header**: Purple gradient with app title
2. **Load Course Data**: Section to upload your courses
3. **Available Courses**: Empty until you load data
4. **Statistics Bar**: Shows 0 courses, 0 assigned, 12 events
5. **Schedule Grid**: Displays all 12 events with their days

---

## Understanding Your Data

### Events (Pre-configured)

Your events are already set up in `Data/events.csv`:

| Event ID | Event Name | Days |
|----------|------------|------|
| EV001 | Virtual-Feb | 4 |
| EV002 | Scottsdale | 4 |
| EV003 | Virtual-Apr | 4 |
| EV004 | Chicago | 4 |
| EV005 | Orlando | 4 |
| EV006 | Virtual-Jul | 4 |
| EV007 | San Diego | 4 |
| EV008 | Denver | 4 |
| EV009 | Detroit | 2 ⚠️ |
| EV010 | Washington | 4 |
| EV011 | Virtual-Nov | 4 |
| EV012 | Las Vegas | 4 |

⚠️ **Note**: Detroit is the only 2-day event!

### Event Days (Pre-configured)

Each event has specific dates in 2026:
- Virtual-Feb: Feb 23-26, 2026
- Scottsdale: Mar 16-19, 2026
- Virtual-Apr: Apr 20-23, 2026
- And so on...

### Courses (You Create This!)

You need to create a CSV file with your courses:

```csv
Course_ID,Instructor,Course_Name,Duration_Days
C001,Alfred,Mapmaking,3
C002,Betty,Cooking Basics,2
C003,Charlie,Advanced Photography,4
```

**Column Definitions**:
- `Course_ID`: Unique identifier (required, no spaces)
- `Instructor`: Instructor's name (required)
- `Course_Name`: Name of the course (required)
- `Duration_Days`: Length in days (required, 0.5 to 4)

**Duration Examples**:
- Half day: `0.5`
- One and half days: `1.5`
- Two days: `2`
- Three and half days: `3.5`
- Four days: `4`

---

## Step-by-Step Tutorial

### Tutorial 1: Your First Schedule

**Goal**: Schedule Alfred's 3-day Mapmaking course at Orlando

#### Step 1: Prepare Your Courses File

1. Open `Data/courses_template.csv` in Excel or Notepad
2. You'll see Alfred's Mapmaking course already listed (3 days)
3. Save the file (or create your own)

#### Step 2: Load Courses

1. In SchedulePro, find **"Step 1: Load Course Data"**
2. Click **"📁 Load Courses CSV"**
3. Browse to `Data/courses_template.csv`
4. Click Open

**Result**: You'll see course cards appear in the "Available Courses" section!

#### Step 3: Assign to Event

1. Find the **"Alfred - Mapmaking"** card (purple border)
2. Click and hold your mouse on the card
3. Drag it down to the **Orlando** event section
4. Drop it on any day slot (let's try Day 1)

**Result**: A modal pops up asking you to select days!

#### Step 4: Select Days

The modal shows:
- Course info: Alfred - Mapmaking (3 days)
- Event info: Orlando (4 days)
- Day options: "Days 1-3" or "Days 2-4"

1. Click **"Days 1-3"** (to start on day 1)
2. Click **"✓ Confirm"**

**Result**: The course now appears on Days 1, 2, and 3 in Orlando!

#### Step 5: Verify

Look at the Orlando event card:
- Day 1: Shows "Mapmaking" with Alfred's name
- Day 2: Shows "Mapmaking" with Alfred's name
- Day 3: Shows "Mapmaking" with Alfred's name
- Day 4: Empty (as expected, since it's a 3-day course)

The course card in "Available Courses" now has a green checkmark (✓ Assigned)!

#### Step 6: Export

1. Scroll to **"Step 4: Export Schedule"**
2. Click **"📊 Export to Excel"**
3. The file `schedule_export.csv` downloads automatically
4. Open it in Excel to see your schedule!

**Congratulations!** 🎉 You've created your first schedule!

---

### Tutorial 2: Multiple Courses, Multiple Events

**Goal**: Schedule multiple courses across different events

#### Scenario
- Alfred's Mapmaking (3 days) → Orlando
- Betty's Cooking (2 days) → Chicago  
- Charlie's Photography (4 days) → San Diego

#### Steps

1. **Load the template** (already has all three courses)

2. **Assign Alfred to Orlando**:
   - Drag Alfred's card → Orlando Day 1
   - Select "Days 1-3"
   - Confirm

3. **Assign Betty to Chicago**:
   - Drag Betty's card → Chicago Day 1
   - Select "Days 1-2" (or "Days 2-3", or "Days 3-4")
   - Confirm

4. **Assign Charlie to San Diego**:
   - Drag Charlie's card → San Diego Day 1
   - Select "Days 1-4" (the only option since it's 4 days)
   - Confirm

**Result**: Three courses scheduled across three events!

---

### Tutorial 3: Same Course, Multiple Events

**Goal**: Schedule Edward's Public Speaking (1 day) at multiple events

#### Scenario
Edward is available to teach at Virtual-Feb, Chicago, and Washington

#### Steps

1. **First Assignment**:
   - Drag Edward's card → Virtual-Feb Day 1
   - Select "Day 1" (it's only 1 day)
   - Confirm

2. **Second Assignment** (same course!):
   - Drag Edward's card again → Chicago Day 3
   - Select "Day 3"
   - Confirm

3. **Third Assignment**:
   - Drag Edward's card again → Washington Day 4
   - Select "Day 4"
   - Confirm

**Result**: Edward teaches at three different events! Notice the course card still shows "✓ Assigned" even though it's at multiple events.

---

## Advanced Features

### Multiple Courses Per Event

You can schedule multiple courses to run simultaneously at the same event:

**Example**: Orlando (4 days)
- Days 1-3: Alfred's Mapmaking
- Days 2-4: Betty's Cooking

Both courses overlap on Days 2 and 3, which is perfectly fine!

### Removing Assignments

Made a mistake?

1. Find the assigned course in the schedule grid
2. Click the **×** button in the top-right corner
3. The course is removed from ALL days it was assigned to

### Saving Your Work

**Save as JSON**:
1. Click **"💾 Save Schedule (JSON)"**
2. Downloads `schedule_backup.json`
3. Keeps: courses, assignments, everything!

**Load from JSON**:
1. Click **"📂 Load Schedule (JSON)"**
2. Select your `.json` backup file
3. Everything restores instantly!

**When to Save**:
- Before making major changes
- At the end of each work session
- Before exporting to share

### Working with Decimal Durations

Courses can be fractional days:

**0.5 days** (half day):
- The system rounds up to 1 day for scheduling
- So a 0.5-day course occupies 1 day slot

**1.5 days**:
- Rounds up to 2 days
- Select a 2-day range

**3.5 days**:
- Rounds up to 4 days
- Select a 4-day range (fills the entire event if it's 4 days)

---

## Common Scenarios

### Scenario 1: 4-Day Course in 4-Day Event

**Course**: Helen's Project Management (4 days)
**Event**: Scottsdale (4 days)

**Solution**: Only one option!
- Drag to any day
- Only "Days 1-4" will be available
- Confirm

Result: Course fills the entire event

---

### Scenario 2: 1-Day Course with Multiple Options

**Course**: Edward's Public Speaking (1 day)
**Event**: Orlando (4 days)

**Solution**: You have 4 options!
- Days 1, 2, 3, or 4

Select based on:
- Other courses already scheduled
- Preferred timing
- Instructor availability

---

### Scenario 3: 3-Day Course in 2-Day Event (Detroit)

**Course**: Alfred's Mapmaking (3 days)
**Event**: Detroit (2 days)

**Problem**: Course is longer than event!

**Solution**: 
- The modal will show NO day options
- You cannot assign this course to Detroit
- Choose a different event (most are 4 days)

**Tip**: Save Detroit for shorter courses (1-2 days)

---

### Scenario 4: Reorganizing After Initial Assignment

**Situation**: You scheduled too many courses on the same days

**Solution**:
1. Click × to remove courses
2. Reassign them to different day ranges
3. Or assign to different events entirely

**Example**:
- Remove Alfred from Orlando Days 1-3
- Reassign to Orlando Days 2-4 instead
- Now Day 1 is free for another course!

---

## Troubleshooting

### ❌ "CSV must have columns: Course_ID, Instructor, Course_Name, Duration_Days"

**Problem**: Your CSV file is missing required columns

**Solution**:
1. Open your CSV in Excel
2. Make sure the first row has exactly these headers:
   - Course_ID
   - Instructor
   - Course_Name
   - Duration_Days
3. Spelling and capitalization must match exactly!
4. Save and try loading again

---

### ❌ Courses Won't Appear

**Problem**: Loaded CSV but no courses show

**Check**:
1. Is the CSV file empty (only headers)?
2. Are there special characters breaking the parsing?
3. Try loading `courses_template.csv` to verify the app works

---

### ❌ "Error loading event data"

**Problem**: Can't find event files

**Solution**:
1. Make sure these files exist:
   - `Data/events.csv`
   - `Data/event_days.csv`
2. If missing, run:
   ```powershell
   C:/Users/Dell/Documents/SchedulePro/.venv/Scripts/python.exe convert_to_csv.py
   ```

---

### ❌ Drag and Drop Doesn't Work

**Problem**: Can't drag course cards

**Possible Causes**:
1. Using an old browser → Update to latest version
2. JavaScript disabled → Enable in browser settings
3. File opened from email attachment → Save to your computer first

---

### ❌ Modal Shows No Day Options

**Problem**: Drop course on event but modal shows no selectable days

**Cause**: Course duration is longer than event duration

**Example**: 3-day course + Detroit (2 days) = No valid options

**Solution**: Choose a different event with more days

---

### ❌ Export Downloads But Excel Looks Wrong

**Problem**: CSV opens with weird formatting

**Solution**:
1. In Excel, go to Data → Text to Columns
2. Select "Delimited"
3. Check "Comma"
4. Finish

Or:
1. Import the CSV using Excel's Import Wizard
2. Specify comma as delimiter

---

### ❌ Lost All My Work!

**Problem**: Closed browser and schedule is gone

**Prevention**:
- Click **"💾 Save Schedule"** regularly
- Downloads a JSON file you can reload

**Recovery**:
- If you saved before, click **"📂 Load Schedule"**
- Select your backup JSON file

**Note**: The app runs in your browser's memory only. Closing the browser clears everything unless you save!

---

## Tips for Success

### ✅ Best Practices

1. **Save Early, Save Often**: Use JSON backups
2. **Start Simple**: Schedule one course at a time initially
3. **Check Statistics**: Watch the "Assigned" counter to track progress
4. **Export to Review**: Export to Excel frequently to review in a familiar format
5. **Use Templates**: Start with `courses_template.csv` and modify it

### ✅ Workflow Recommendations

**Phase 1: Preparation**
1. Create your courses CSV with all instructors and durations
2. Review event dates in `Data/event_days.csv`
3. Plan which courses should go to which events

**Phase 2: Bulk Assignment**
1. Load courses
2. Assign all courses to events (don't worry about exact days yet)
3. Save progress as JSON

**Phase 3: Fine-Tuning**
1. Review the schedule
2. Remove and reassign courses to optimize day usage
3. Ensure no conflicts or issues

**Phase 4: Export**
1. Export to Excel
2. Share with stakeholders
3. Make final adjustments if needed

### ✅ Keyboard Shortcuts

While the app doesn't have built-in shortcuts, you can use:
- **Ctrl+F**: Search on the page (find specific courses/events)
- **Ctrl+S**: (In browser) Save the page (not the schedule though!)
- **F11**: Full screen mode for better viewing

---

## FAQ

**Q: Can I edit a course after assigning it?**
A: Remove it, then reassign with different day selection.

**Q: Can one instructor teach multiple courses?**
A: Yes! Just list them as separate rows in your CSV with different Course_IDs.

**Q: What if two courses need the same days?**
A: That's fine! Multiple courses can run simultaneously at the same event.

**Q: Can I change event dates?**
A: Yes, edit `Data/Event_Days.xlsx` then run `convert_to_csv.py`.

**Q: Does this work offline?**
A: Yes! Completely offline after opening.

**Q: Can I print the schedule?**
A: Export to Excel, then print from there for best results.

**Q: What's the maximum number of courses?**
A: No hard limit, but performance may slow with 1000+ courses.

---

## Summary

You now know how to:
- ✅ Load course data from CSV
- ✅ Drag and drop courses to events  
- ✅ Select specific day ranges
- ✅ Handle multiple assignments
- ✅ Export to Excel
- ✅ Save and restore your work
- ✅ Troubleshoot common issues

**Ready to schedule?** Open the app and start planning! 🚀

---

*For technical details, see `README.md`*  
*For quick reference, see `QUICK_START.md`*
