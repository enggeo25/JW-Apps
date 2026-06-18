GEOTECH TRACKER GIS V2.5 CALENDAR FOCUS

This version focuses on the calendar view:
- full month calendar squares
- planned bars shown by work type date windows
- actual item bars shown on their actual dates
- summary still shown below the calendar
- Data Management colors restored
- inline editing keeps the page position better
- Include Saturday checkbox alignment cleaned up

Notes:
- Calendar changes are linked because everything still works from the same records.
- This is a stronger visual calendar, but it is still the first structured version of that calendar concept.


V2.5.1 FIX
- Fixed project page internal server error caused by an unclosed Jinja block in project.html.


V2.6 CALENDAR VIEWS
- Added Planning / Active / Combined calendar subtabs
- Added continuous weekly bars instead of separate day blocks
- Planned bars are dashed/hatched, actual bars are solid
- Type colours separated for Borehole / CPTU / Test Pit / Geophysics


V2.6.1 CALENDAR FIX
- Fixed weekly calendar bars so they start on the exact day and end on the exact day
- Added bulk map date update fields for selected items
- Small visual improvements to weekly calendar view


V2.7 STABLE CALENDAR EDITOR
- Rebuilt from the last stable working package
- Fixed planned calendar bars to break over Sundays and optional Saturdays
- Added editing directly from the Calendar tab
- Kept map bulk date editing fields


V2.8 INTERACTIVE CALENDAR
- Calendar bars are clickable and edit in a side editor
- Legend buttons toggle types and planned/active on or off
- Added applicable-method selection at project setup
- Added custom test methods with name, symbol, and color
- Added map full screen toggle and tighter map layout


V2.8.4 REFINEMENTS
- Calendar bar text improved for readability
- Planned bars are now clickable and editable from the calendar side editor
- Data Management sections remember open/closed state and default closed
- Added depth/metres per item and overview metre summaries now use item depths
- Improved drag edge sensitivity
- Improved GIS map height and full-screen behavior
