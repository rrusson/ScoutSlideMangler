# ScoutSlideMangler
AdHoc PowerPoint Slide XML mangling tool to automate creation of merit badge slides for Court of Honor

With a troop of almost 100 Boy Scouts, updating slides for every Court of Honor can be a real time-sucker.  This sloppy hack automates the biggest manual task of creating badges earned slides from troop data.

Hacky Manual/Automated Steps:
1 - Conform list of Scouts to an Excel spreadsheet with four columns (LastName, FirstName, [Anything], Merit Badge Name)
2 - Name Excel tab "Merit Badges", save file as "C:\Temp\COH.xlsx"
3 - Rename a copy of your PowerPoint to give it a .ZIP extention
4 - Run ScoutSlideMangler
5 - Take generated Slides from "C:\Temp\" and replace slides inside the ZIP (from .PPTX) file
6 - Rename file from .ZIP extention back to .PPTX extention

