INI-file maitenance
-------------------
This application serves as a template for writing new applications
in which you can save all the userconfigurable settings into an INI-file.
Take a look at the module "modINI" and the class "clsINI"
clsINI was taken from a PSC-contribution by Eric Dalquist (I have included 
his zipfile for your reference).
I liked the concept so I decided to add some more functionalities.

Sample application description
------------------------------
This application shows how to save/restore a sections in the inifile
using "Save settings" command buttons.  Also there is an automatic save-restore
of the manual entry section.
If you play around with the add/rename/delete buttons and section/key/value fields
the contents of the inifile is displayed immediately.  Then at program restart you 
will find only the settings that are valid / maintained by the application!
All the other un-used settings are deleted automatically.
Even if you delete the entire inifile a new one will be created at startup using
hardcoded defaultvalues.

Class "clsINI" description
--------------------------
  - init takes care of a default sectionname, which is the inifilename
    but you can override that.
  - saves the currentSectionName and iniFileName so you don't have to  
    specify them (but you can) for every subroutine call.
  - add/rename/delete a section or key.
  - delete all sections.
  - returns all sectionnames, or keys within a section.
  - returns the number of sections, or keys within a section.
  - returns the version of the inifile.
  - getKeyValue: if a key is not found the supplied defaultvalue will 
    be written to the inifile and returned.

Module "modINI" description
---------------------------
  - this is the part you would change for every new application
  - here you declare all your global variables which you want to save/restore
  - there is a readDefaultSection and writeDefaultSection sub and
    for each section there is a readSection and a writeSection sub
  - if you want to stick to just one section just use the default section
  - to keep the inifile up to date during development the version of the
    inifile is checked and the number of keys in each section.  If the version
    is not equal to the hard coded value the inifile is rewritten and if
    the number of keys in the section is not correct the section is rewritten.
  - at application startup all you have to do is to call readAllSettings and 
    do the setup of your application, see the formLoad routine of frmMain.


That's all folks.
Have fun!

Eric Kok

