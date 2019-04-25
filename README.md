# ScoutSlideMangler
AdHoc PowerPoint Slide XML mangling tool to automate creation of merit badge slides for Court of Honor
_____

With a troop of almost 100 Boy Scouts, updating PowerPoint slides for every Court of Honor can be a real time-sucker.  This sloppy hack automates the biggest manual task of creating earned badges slides.

**Hacky Manual Steps:**
1. Conform list of Scouts to an Excel spreadsheet with four columns (LastName, FirstName, [Empty], Merit Badge Name)
2. Name Excel tab "Merit Badges", save file as "C:\Temp\COH.xlsx"
3. Run ScoutSlideMangler
4. Rename a copy of your PowerPoint with a .ZIP extention
5. Take generated slides from "C:\Temp\" and replace slides inside the ZIP (from .PPTX) file
6. Rename file from .ZIP extention back to .PPTX extention


**TODO: automate .PPTX to .ZIP to .PPTX and other dumb steps**
