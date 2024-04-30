# NewRevisionApp
Simple app which allows to copy folders and creating new ones with updated revisions, applies changes to the excel files and changes strings inside.
App made on demand, as improvement in current job. All internal data, names, id's had to be changed.
This project is done for the specific files and folders. 
Main goal is to save time when it comes to change revisions of docuemntation and eliminate manual changes as much as possible.
Unfortunately, i musn't show example of documentation.

How it works:
  User provides necessary data from checkboxes and text.
  Python def Creating_doc, based on input data creates new folders, based on original files.
  Copies the files from origin.
  Changes the names of new files.
  Based on excel input data, which is controlled by user, def Generate_change changes strings inside of excel files.
