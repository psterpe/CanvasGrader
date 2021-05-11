# CanvasGrader

**WORK IN PROGRESS**

This is the Google Apps Script code behind a Google Sheets file that acts as a (simple) frontend to Canvas (the LMS from Instructure). The Google Sheet fetches a course
from Canvas (including the roster). You can then select a student, and the sheet will fetch all assignment grades for the given student and course. Weighted averages are
calculated for the assignment groups, and an overall grade for the student is shown. For assignments that have not been graded, you can opt to omit them from calculations,
or include them as zeroes. Once assignment grades have been fetched, you can also edit the sheet to supply a what-if grade for any/all assignments; the sheet recalculates
the student's overall grade as you make the changes.
