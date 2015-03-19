# libMAIL
libMAIL is a small LGPL library for sending emails from MS-Access
# Installation
* Create a new folder on your drive ;
* Extract files from the archive (https://github.com/Dejieres/libMAIL/archive/master.zip) ;
* Create a new database. Name it libMAIL.mdb (or .accdb) ;
* Open the database ;
* Create a new standard module and import the file Outils.bas. Save the module under the name 'Outils'. This module contains a function that **_DELETES ALL MODULES AND FORMS FROM THE DATABASE_** , excepting the one named 'Outils'. It imports then all the *.frm and *.bas files it finds in the current database's directory ;
* Ensure that DAO (Microsoft DAO Object Library) and MSOffice X.y Object Library (msoxx.dll) references are set ;
* Open the debug window (Ctrl-G) ;
* Just type `ChargeVB` followed by Enter. This will import all the objects and compile the project.

Your library is ready to be referenced by your projects.

# Quick start
## Create a mail
    Call CreeMail (Recipients, Subject, Message, From, [Username], [CC], [BCC], [Attachments()], [EditMail])
## Start the server
    Call SMTPLance(SrvName, [HELOdomain], [LogData], [LogComm], [LogFile], [Quit], [PollInterval], [Timeout])
