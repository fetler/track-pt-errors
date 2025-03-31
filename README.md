# track-pt-errors
Tracks where students have been unallocated from a personal tutor group in CMIS.

Must be used with the XLSX files created by the [Missing Tutor Groups](https://github.com/fetler/missing-tutor-groups) script.

## What does this script do?
If you find any students have been mysteriously removed from a personal tutoring module, or have had their group allocation removed, this script searches for that student's ID in every XLSX file exported by the Missing Tutor Groups script. 

It it finds a match, it returns the filename of the XLSX file, the date the file was created, and the worksheet (tab) it was found in.

It then exports an XLSX file to your computer with the results, allowing you to check what personal tutoring module that student was previously enrolled onto.

<img width="504" alt="Screenshot 2025-03-31 at 11 46 08" src="https://github.com/user-attachments/assets/1f92a490-9bae-4786-ab42-aa2045ed2dd9" />


## To-do
- [ ] Return the CMIS group the student was previously allocated to
- [ ] Change inputting the student ID into the console to inputting the student ID into a text box using TKinter
- [ ] Add a save file dialog so the user can choose where to save it and what to name it, to replace the automatic saving to a default location
