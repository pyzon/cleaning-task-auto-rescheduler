# Cleaning Task Auto-Rescheduler

An adaptive recurring task scheduler built with Google Apps Script, Google Tasks, and Google Sheets.

Instead of using fixed intervals (e.g. "every 7 days"), this system adjusts each task's next due date based on historical completion behavior using a moving average of past intervals.

## Idea

I noticed that different areas of my home require cleaning at very different frequencies. For example, the kitchen floor typically needs mopping about twice a week, while the driveway only needs sweeping roughly once a month. Instead of assigning arbitrary fixed intervals, I became curious about how often I actually clean each place and whether that data could inform better scheduling decisions.

Recording completion dates of tasks can be automated on many platforms. Since I am already familiar with Google Apps' ecosystem and I am using it on a daily basis, I chose it to accomplish this project.

But then I realized I could take the idea one step further: once completion dates are recorded over time, the intervals between cleanings can be calculated. From those intervals, some form of middle value, like a moving average can be computed - providing a data-informed estimate of when the next cleaning should be scheduled.

## What it does

Each time a task is completed:

1. The completion date is logged into a spreadsheet.
2. The spreadsheet calculates the moving average of the intervals between completions.
3. The script sets the next due date to: `today + moving_average_interval`

Over time:
- If a task is done more frequently, the interval gradually decreases.
- If it is done less frequently, the interval gradually increases.

This creates a self-adjusting rhythm based on real-world usage patterns. For example, when a new pet dog arrives at the house, it usually means the dirt accumulates more rapidly in areas where the dog is allowed in. Or, in vacation season, people spend less time at home, creating less dirt, resulting in less frequent cleaning.

As you can see, the computed due date is intended as guidance rather than a strict rule - the task should always be done when it is actually needed, and those real-world decisions are exactly what drive the moving average to adjust over time.

---

## Architecture

- **Google Tasks**: user interface
- **Google Sheets**: data store for completion history and intervals
- **Apps Script**: automation and synchronization logic

---

## Required Spreadsheet Structure

### Dates Sheet
- Row 1: task names (no gaps between columns)
- Rows 2+: completion dates (`yyyy-MM-dd`)
  - Dates must fill downward without empty cells between entries
  - They are filled automatically by the script, but you can also enter dates manually

### Intervals Sheet
- Row 1: identical to Dates sheet row 1
  - Insert into A1: `=ARRAYFORMULA(Dates!A1:1)`, no autofill needed
- Row 2: moving average - or whatever formula you choose - of the intervals (in days) for each task
  - I used a moving average of the middle 3 of the last 5 intervals, leaving out the highest and lowest. For this, I used this formula in A2 and autofill all the way to the right: `=ROUND(TRIMMEAN(A$3:A$7,0.4))`
  - These are the final values that will be read by the script
- Additional rows - customizable, but here's how I did it:
  - Rows 3 to 7:
    - The last 5 intervals copied over: `=INDEX(A$8:A,COUNT(A$8:A)+ROW()-7)` (A3, autofill down to A7, and right all the way) - this is so that the average formula is working from the same range, making it simpler
  - Rows 8 to 12:
    - Padding, filled with whatever intervals are reasonable for the task - this is for the formulas to not break when you don't have enough data yet
  - Rows 13+:
    - The actual intervals: `=IF(AND(NOT(ISBLANK(Dates!A3)),NOT(ISBLANK(Dates!A2))),Dates!A3-Dates!A2,"")` (A13, autofill down and right all the way)

---

## Setup

1. Create a Google Spreadsheet following the required structure.
2. Create a dedicated task list in Google Tasks.
3. Create a Google Apps Script project.
4. Enable Google Sheets and Google Tasks API services. They will appear on the left under Services if you've done it right.
5. Add the following Script Properties - there is a helper function for this in `helpers.gs`:
- `CLEANING_DATA_SPREADSHEET_ID` - find this in the URL of the spreadsheet between `d/` and `/edit`
- `DATES_SHEET_ID` - find this after activating the tab in the URL after `gid=`, if this is the default tab, it will be `0`
- `INTERVALS_SHEET_ID` - find this similarly
- `CLEANING_TASKLIST_ID` - find this by listing and logging all of your task lists - there is a helper function for this too
5. Deploy a daily time-driven trigger (e.g., 23:00) - there is a helper function for this too.

## Authorization

When running the script for the first time, Google will request authorization.

The script requires permission to:

- Run automatically when you are not present (time-driven trigger)
- Read and write Google Tasks
- Read and write Google Sheets

Because this is a personal Apps Script project and not a published Google-verified application, you will see a warning stating that the app is not verified by Google.

After reviewing the source code and confirming that it performs only the intended actions, click **Advanced** and proceed to grant the required permissions.

No external services are used, and all data remains within your Google account.