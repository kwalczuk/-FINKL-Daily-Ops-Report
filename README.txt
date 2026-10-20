CHICAGO DAILY OPS REPORT
========================

SETUP (once)
  1. Before unzipping: right-click the .zip > Properties > tick "Unblock" > OK.
  2. Unzip the folder anywhere (e.g. J:\Departments\Machine\4M\KACPER\DailyOps).
  Needs Excel. Close the report workbook in Excel before running (else a copy is saved instead).

EVERY DAY
  1. Download today's Raw Machine Data and the Shift Hours and OT Report(s)
     (one file per day, or a .zip with several - if a day is in there twice, the newest copy wins).
  2. Double-click "Daily Ops.bat" and press Enter.
  3. It goes step by step and asks for each file in turn:
       workbook -> report date -> raw data -> start-time trackers -> schedules
       -> absent per shift -> hours -> save
     At each step:
       drag the file(s) from Explorer into the window   (select several and drag them together,
                                                          or drop one, then another - Enter when done)
       B            opens the normal Windows file picker in that step's folder (Ctrl+click for several)
       1 / 1,3 / A  take what it found in the folder      Enter  the [suggestion]
       undo / clear take back the last / all             S      skip the step
     Dropped a file at the wrong step? It's recognised by its name and kept for its own step.
     (You can also drag everything onto the .bat at once - each file lands at its step.)
  4. Saved: the workbook in place (backup in Backups\) and the 2 pictures in
     Output\2026-10-19\ next to the script. The Summary picture is also on the clipboard.

  Where each step looks first (change with: set <name> <path>):
    Report workbook   J:\Departments\Machine\4M\KACPER                      (ReportFolder)
    Raw machine data  Downloads                                             (RawFolder)
    Start trackers    J:\Departments\Machine\4M                             (TrackerFolder)
    Schedules         J:\Departments\Machine\SOUTH\23 Attendance\Schedules  (ScheduleFolder)
    Hours / OT .zip   Downloads                                             (HoursFolder)
  Files on J: are copied to this PC first; the workbook is copied back when it's saved.

START TIMES (minutes late, the "Average time to start tool" row)
  Read from the trackers (1st/2nd/3rd shift from the file name). Type over any day:
    hms tue 12        12 minutes
    hms tue 6:25      clock time - counted from that shift's start (6:00 / 14:00 / 22:00)
    fel wed na        N/A
    hms tue -         back to the tracker value        undo hms   all typed for HMS
  Typed days show a * and are kept for 60 days (re-running the same day keeps them).

ABSENT (per shop, per shift -> "# of employee present" Actual = Scheduled - absent)
    hms 2 0 1         HMS, latest day: 2 absent 1st, 0 2nd, 1 3rd
    hms wed 2 0 1     a given day
    mms mon 3         1st shift only
    hms mon -         clear that day                    clear hms   all days
  The "present" line shows present/scheduled as you type.

WHAT IT DOES TO THE WORKBOOK (same as the buttons/macros did)
  Summary C5 = report date, C4 = week ending, day columns after the report date hidden.
  Raw Machine Data -> YEARLY DATA A:U (V:AD formulas filled down if the data got longer).
  Start times, scheduled headcount, present, Regular/OT/DT hours -> the Core Machining sheets.
  Cubic-inch pivots refreshed and filtered to the report week.
  Yellow cells right of the Performance % rows (T13, T23, T33, T43) use the latest report day.
  Formulas that were typed over / #REF! in the Summary rows (and the Core rows they use)
  are put back - listed on screen (turn off with: set FixFormulas No).

COMMANDS (type help in the window)
    run [files]              the daily update (Enter on an empty line does the same)
    pictures                 just the 2 pictures from the workbook as it is
    shops | settings | set <name> <value>
    typed | forget           typed start times / absences kept for re-runs
    open config | output | report | backups
    cancel (at any prompt)   back out

FILES
  DailyOps.ps1          the script
  Daily Ops.bat         launcher
  DailyOpsConfig.xlsx   shops (Core sheet, dept #, tracker tab, schedule depts) + settings
  Saved\typed.json      typed start times / absences
  Backups\              the last 5 versions of the report workbook (KeepBackups)
  Output\               pictures (and copies when the workbook was open elsewhere)
