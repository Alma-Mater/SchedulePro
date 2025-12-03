# 🎨 SchedulePro Visual Reference

## Interface Layout

```
┌────────────────────────────────────────────────────────┐
│                   📅 SchedulePro                       │
│     Easy-to-Use Course Scheduling for Event Planners   │
└────────────────────────────────────────────────────────┘

┌─ Step 1: Load Course Data ──────────────────────────┐
│  [📁 Load Courses CSV]  [📥 Download Template]      │
│  💡 Tip: Load your courses CSV with columns...      │
└─────────────────────────────────────────────────────┘

┌─ Step 2: Available Courses ─────────────────────────┐
│  ┌──────────┐  ┌──────────┐  ┌──────────┐         │
│  │  Alfred  │  │  Betty   │  │ Charlie  │         │
│  │Mapmaking │  │ Cooking  │  │  Photo   │         │
│  │ 📏 3 days│  │ 📏 2 days│  │ 📏 4 days│         │
│  └──────────┘  └──────────┘  └──────────┘         │
└─────────────────────────────────────────────────────┘

┌─ Statistics ─────────────────────────────────────────┐
│   Total: 10    Assigned: 5    Events: 12           │
└─────────────────────────────────────────────────────┘

┌─ Step 3: Drag & Drop Courses to Events ─────────────┐
│                                                      │
│  ┌──── Event: Orlando (4 days) ────────────────┐   │
│  │ Day 1     │ Day 2     │ Day 3     │ Day 4   │   │
│  │ Jun-15    │ Jun-16    │ Jun-17    │ Jun-18  │   │
│  │           │           │           │         │   │
│  │ Mapmaking │ Mapmaking │ Mapmaking │         │   │
│  │ Alfred    │ Alfred    │ Alfred    │         │   │
│  └───────────┴───────────┴───────────┴─────────┘   │
│                                                      │
│  ┌──── Event: Chicago (4 days) ─────────────────┐   │
│  │ Day 1     │ Day 2     │ Day 3     │ Day 4   │   │
│  │ May-11    │ May-12    │ May-13    │ May-14  │   │
│  └───────────┴───────────┴───────────┴─────────┘   │
└─────────────────────────────────────────────────────┘

┌─ Step 4: Export Schedule ────────────────────────────┐
│    [📊 Export to Excel]  [💾 Save]  [📂 Load]       │
└─────────────────────────────────────────────────────┘
```

## Color Guide

### Course Cards
- **Purple Border**: Unassigned course
- **Green Border + Checkmark**: Assigned course
- **Lighter Background**: While dragging

### Event Cards
- **Purple Header**: Event name and info
- **White Day Slots**: Available days
- **Dashed Border**: Empty day slot
- **Solid Purple Border**: During drag-over
- **Green Boxes**: Assigned courses

### Buttons
- **Purple Gradient**: Primary actions (Load, Confirm)
- **Gray**: Secondary actions (Cancel, Download Template)
- **Green**: Success actions (Export)
- **White ×**: Remove course

## Interaction Patterns

### Drag and Drop Flow
```
1. Click & Hold Course Card
   ↓
2. Drag Over Event Day
   (Day slot highlights blue)
   ↓
3. Release Mouse
   (Modal pops up)
   ↓
4. Select Day Range
   (Option highlights purple)
   ↓
5. Click Confirm
   (Course appears on days)
```

### Day Selection Modal
```
┌────────────────────────────────────┐
│  Select Days for Course            │
├────────────────────────────────────┤
│  Alfred - Mapmaking                │
│  Duration: 3 days                  │
│  Event: Orlando (4 days)           │
│                                    │
│  ┌──────────┐  ┌──────────┐      │
│  │ Days 1-3 │  │ Days 2-4 │      │
│  │  3 days  │  │  3 days  │      │
│  └──────────┘  └──────────┘      │
│                                    │
│    [✓ Confirm]  [✗ Cancel]        │
└────────────────────────────────────┘
```

## Status Indicators

### Course Card States
```
UNASSIGNED:
┌──────────────┐
│   Alfred     │  ← Purple border
│  Mapmaking   │
│  📏 3 days   │
└──────────────┘

ASSIGNED:
┌──────────────┐
│   Alfred     │  ← Green border
│  Mapmaking   │
│  📏 3 days   │
│ ✓ Assigned   │  ← Green checkmark
└──────────────┘

DRAGGING:
┌──────────────┐
│   Alfred     │  ← Semi-transparent
│  Mapmaking   │
│  📏 3 days   │
└──────────────┘
```

### Day Slot States
```
EMPTY:
┌──────────────┐
│ Day 1        │  ← Dashed border
│ Jun-15       │     Gray background
│              │
└──────────────┘

DRAG OVER:
┌──────────────┐
│ Day 1        │  ← Solid purple border
│ Jun-15       │     Light blue background
│              │
└──────────────┘

WITH COURSE:
┌──────────────┐
│ Day 1        │
│ Jun-15       │
│ ┌──────────┐ │
│ │Mapmaking │ │  ← Green box
│ │Alfred    │ │     with × button
│ └──────────┘ │
└──────────────┘
```

## Icons Reference

| Icon | Meaning |
|------|---------|
| 📅 | Schedule/Calendar |
| 📁 | Load/Open File |
| 📥 | Download |
| 📏 | Duration/Length |
| ✓ | Assigned/Confirmed |
| × | Remove/Delete |
| 📊 | Export to Excel |
| 💾 | Save |
| 📂 | Load |
| 💡 | Tip/Information |
| ⚠️ | Warning/Attention |
| ✅ | Success |
| ❌ | Error |

## Responsive Behavior

### Desktop (>768px)
- Course cards: 3-4 per row
- Event days: 4 columns (side by side)
- Full header with large title

### Mobile/Tablet (<768px)
- Course cards: 1 per row
- Event days: 1 column (stacked)
- Smaller header text
- Touch-friendly drag and drop

## Export Format

### Excel Export Structure
```
┌──────────┬───────────┬─────┬────────┬──────────┬────────────┬───────────┬──────────────┐
│ Event_ID │   Event   │ Day │  Date  │Course_ID │ Instructor │   Course  │ Duration_Days│
├──────────┼───────────┼─────┼────────┼──────────┼────────────┼───────────┼──────────────┤
│  EV005   │  Orlando  │  1  │ Jun-15 │   C001   │   Alfred   │Mapmaking  │      3       │
│  EV005   │  Orlando  │  2  │ Jun-16 │   C001   │   Alfred   │Mapmaking  │      3       │
│  EV005   │  Orlando  │  3  │ Jun-17 │   C001   │   Alfred   │Mapmaking  │      3       │
│  EV005   │  Orlando  │  4  │ Jun-18 │          │            │           │              │
└──────────┴───────────┴─────┴────────┴──────────┴────────────┴───────────┴──────────────┘
```

## Quick Actions

### Keyboard
- **Drag**: Click and hold, move mouse
- **Drop**: Release mouse button
- **Remove**: Click × on assigned course
- **Close Modal**: Click outside or Cancel button

### Mouse
- **Hover over course**: Card lifts up (shadow effect)
- **Hover over button**: Button lifts and glows
- **Hover over assigned course**: × button becomes more visible

## Animation Effects

### Smooth Transitions
- Course cards: 0.2s lift on hover
- Buttons: 0.2s transform and glow
- Day slots: 0.2s background change on drag-over
- Modal: Fade in/out effect

### Visual Feedback
- **Success**: Green color, checkmark appears
- **Warning**: Yellow/orange alert box
- **Error**: Red text in alert
- **Info**: Blue alert box with tips

---

## Example Workflows

### Workflow 1: Simple Assignment
```
[Course: Alfred-Mapmaking]
       ↓ DRAG
[Event: Orlando, Day 1]
       ↓ DROP
[Modal: Select Days]
       ↓ CLICK "Days 1-3"
[Confirm]
       ↓
[Result: Course on Days 1, 2, 3]
```

### Workflow 2: Multi-Event Assignment
```
[Course: Edward-Speaking (1 day)]
       ↓ DRAG & DROP
[Event: Orlando, Day 1] → Assigned
       ↓ DRAG AGAIN
[Event: Chicago, Day 3] → Assigned
       ↓ DRAG AGAIN
[Event: Denver, Day 2] → Assigned
       ↓
[Result: Same course at 3 events]
```

### Workflow 3: Correction
```
[Assigned Course on Days 1-3]
       ↓ CLICK ×
[Removed from all days]
       ↓ DRAG AGAIN
[Drop on Day 2]
       ↓ SELECT "Days 2-4"
[Result: Course moved to Days 2-4]
```

---

**Tip**: Keep this reference open in a separate window while working in SchedulePro!
