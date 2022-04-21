# info-intro-auto-mated
written for the tutors of introduction to informatics @ TUM


## auto-corrector
Helps streamlining the process correction large quantities of homework submissions (from Moodle). Iterates through the exported homework submissions folders from Moodle and sequentially opens supported files (in this version .pdf and java projects) respectively in the standard pdf viewer and VS-code (required) of the users operating system (tested only on Windows and macOS but should work on Linux as well).
The user will be asked for the grades while name and matriculation number are extracted from the folder structure (if possible). The details are temporarily saved and can then easily be transferred to an excel file.


## auto-grader (selenium based)
Transfers the collected grades from a (specifically) formatted excel file to Moodle via a selenium automated browser instance. Due to the way of the automation it is not necessarily time efficient but still **significantly quicker** than manually transferring a big number of grades (and less cumbersome).
Not matching entries are collected and saved for manual correction later.


## contact
For more information/tutorial videos, custom edits and bug reports use GitHub or write me an email at jcosta.studies@gmail.com
