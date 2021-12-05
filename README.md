# ScoutSlideMangler
AdHoc PowerPoint Slide XML mangling tool to automate creation of merit badge slides for Court of Honor
_____

With a troop of almost 100 Boy Scouts, updating PowerPoint slides for every Court of Honor can be a real time-sucker.  This sloppy hack automates the biggest manual task of creating earned badges slides.

**Current Data Formatting:**
1. Conform a list of Scouts in an Excel spreadsheet with four columns (LastName, FirstName, [Empty], Merit Badge Name)
2. Name the Excel tab "Merit Badges", and save file as "COH.xlsx"
3. You need a PowerPoint file with at least as many slides to overwrite as there are scouts in Excel spreadsheet (the hack doesn't insert)
4. PowerPoint file should be named "CourtOfHonor.pptx"
4. Run ScoutSlideMangler

**Example .XLSX and .PPTX files are included to demonstrate exact formatting necessary to generate slides**

**TODO: Expand system to create a general purpose PowerPoint data merge template system**
