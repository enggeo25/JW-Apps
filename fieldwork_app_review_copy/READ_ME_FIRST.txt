GEOTECH TRACKER - START FIX

This package fixes the startup problem.

VERY IMPORTANT:
1. Right-click the ZIP file and choose Extract All
2. Extract it into a normal folder, for example:
   Desktop\Geotech Tracker
3. Open that extracted folder
4. Double-click:
   START_GEOTECH_TRACKER.bat

DO NOT:
- run the BAT file from inside the ZIP preview
- run it from a temporary extraction window
- move or delete files while it is starting

Why the start file seemed to disappear:
If you run a BAT file from inside a ZIP preview, Windows can open a temporary copy.
That temporary view can close, which makes it look like the start file vanished.

Files to use:
- START_GEOTECH_TRACKER.bat
- DEBUG_START.bat

If the normal start does not work, run DEBUG_START.bat and send the full error text.


V2.8.2 STARTUP FIX
- Fixed the IndentationError in app.py that stopped the newest build from starting.
- This is a fix to the newest interactive-calendar build, not a rollback.

Still pending and not dropped:
- true drag-left/right calendar bar movement
- resize bar ends to change duration
- verify custom test methods flow all the way through every tab with full polish
- verify applicable-method hiding all the way through every tab with full polish
- dashboard tab
- final GIS full-screen and map-height polish


V2.8.3 CALENDAR FIX AND METERS
- Fixed the newest build again without rolling back.
- Restored planned calendar bars logic from the current interactive calendar build and tightened the rendering path.
- Added budget metre fields for Borehole, CPTU, and Geophysics in project setup and edit.
- Added overview metre summary cards.
- Added applicable-method hiding on setup and edit screens.
- Kept custom test methods in setup and edit.
- Added initial calendar drag / resize support on active bars:
  * drag the middle of a bar to move dates
  * drag near the left edge to move the start
  * drag near the right edge to move the end
- Improved map full-screen behavior and map area sizing.
- Custom test methods and applicable-method hiding remain tracked and are not dropped from notes.


V2.8.5 TEMPLATE FIX
- Fixed the Jinja template syntax error that caused project pages to crash.
- Made the 'What's new' section smaller.
- Default Data Management groups now start closed.
- Minor layout polish for setup cards.


V2.8.6 JINJA FIX
- Fixed the remaining Jinja template syntax error on the project page.
- Cleaned the broken escaped quotes in the planning-bar template.
- Made the current-version / what's-new card smaller and moved it lower on the opening screen.


V2.8.7 LAYOUT POLISH
- Moved the current-version / what's-new card to the bottom of the opening screen.
- Expanded the project setup card to use the full width more cleanly.
- Added colour styling to the planning method boxes to match the app colour scheme.


V2.8.8 FINAL LAYOUT + LOGO
- Added the Jones & Wagener logo to the top left of the app.
- Moved the current-version / what's-new card fully below the Projects section.
- Kept the create-project area full width and spaced the planning boxes evenly.


V2.9.1 COLOR + RATE + CALENDAR CLARITY
- Made actual calendar bars use the same colours as their matching methods.
- Centered planning and actual bar labels.
- Added clearer true-start and true-end markers to planned bars.
- Added projected days needed to the Overview tab.
- Reworked individual method rate maths to use average completed items per completion day.
- Expanded the individual method status spectrum to deep red / red / orange / yellow / green based on rate performance.


V2.9.3 OVERVIEW POLISH
- Softened the overview card colours with more transparency and gentler gradients.
- Reordered the overview comparisons for clearer side-by-side reading.
- Moved total items into a cleaner top-right summary block.


V2.9.4 RATE + ACTIVE COLOR FIX
- Fixed current-rate maths to average each completed item's own per-day rate.
- This handles slow multi-day items like boreholes properly.
- Fixed active calendar bars so they use their correct method colours instead of defaulting to blue.
