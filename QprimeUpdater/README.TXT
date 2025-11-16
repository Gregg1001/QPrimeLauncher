Launcher 
========================================================================
LAUNCHER is a simple utility designed to be used when a program provides 
only a command line (or even just a module file name) for launching 
another program; but either Windows registry associations should be 
used, or a DDE command, additional parameters and/or a working directory 
need to be specified.

In particular, LAUNCHER was first written to allow more flexibility in 
the use of Agent's MIME Types table.

Version 2 of LAUNCHER adds the ability to use environment variables
(such as %UserName%) to control launching.
======================================================================== 
BETA DISCLAIMER:

PRIOR TO RELEASE, THIS PROGRAM VERSION HAS BEEN TESTED ON EXACTLY ONE 
SYSTEM: MINE, A 1.6GHz PENTIUM 4 RUNNING WINDOWS 2000 PROFESSIONAL SP 2. 
WHILE I DO NOT EXPECT THAT IT WILL CAUSE ANY SYSTEM DAMAGE, THIS CANNOT 
BE GUARANTEED! USE SUITABLE PRECAUTIONS WHEN TESTING ANY BETA SOFTWARE.

THIS PROGRAM VERSION HAS NOT YET BEEN TESTED UNDER WINDOWS 95/98/ME, 
WINDOWS NT, OR WINDOWS XP. INFORMATION REGARDING EXPERIENCES IN RUNNING 
IT UNDER THOSE SYSTEMS WILL BE GREATLY APPRECIATED.
========================================================================
Release Format:   Zip file (lnch200a.zip)
Platform:         Windows NT/95 and up
Author/contact:   Software@Coises.com (Randy Fellmy)

Distribution:     FREEWARE: This program (with all original files 
                  included and unmodified, and without charge) may be 
                  freely distributed and used.
FileVersion:      2.0.0
FileDescription:  Launcher: program and document launch interface
InternalName:     Launcher
LegalCopyright:   © Copyright 1999-2002 by Randy Fellmy
OriginalFilename: launcher.exe
ProductName:      Launcher
ProductVersion:   2.0.0

Compiler:         Microsoft Visual C++ 6.0 (Win32 API)
========================================================================

Changes in Launcher version 2.0.0
---------------------------------

1. Launcher is now a native 32-bit application.

   A. This appears to have fixed a problem of unknown origin which caused 
   the "directory" parameter to be unreliable when running in Windows NT. 
   (Whether the problem occurs in either 1.1b or 2.0.0 on other 32-bit 
   versions of Windows is unknown.) Of course, this also means it will no 
   longer run under Windows 3.x (with or without Win32S).
   
   B. This also appears to have fixed another problem of unknown origin, 
   in which Microsoft Word 6 would hang on exit after Launcher had 
   launched a document into it by Windows associations. I haven't 
   observed that behavior with this version of Launcher, running under 
   Windows 2000.

2. The interpretation of the "ifexec" parameter was changed so that
if a null value is specified, no DDE conversation is initiated if the 
application must be started.

3. The processing of the "delay" parameter has changed significantly. 
Formerly, "delay" specified the minimum length of time, in milliseconds, 
that should elapse between launching the application and initiating a 
DDE conversation. Launcher now waits for the application to be ready for 
input, for a maximum of "delay" milliseconds, before attempting to 
initiate DDE. The default value for "delay" is now 32767 instead of 0.

4. Environment variables are now recognized in the "directory" parameter 
and in the first token in the "command" parameter (the file name of the 
application or document to be launched).

5. The "override" parameter has been added to allow substitution of an 
alternate initialization file based on environment variables (such as 
%UserName%).



Known problems in Launcher version 2.0.0
----------------------------------------

Launcher recognizes the appearance of a leading double-quote mark in its 
calling parameters as indicative that a file name has been specified 
without a preceding section name; also, a double-quote at the beginning 
of the "command" parameter is recognized as the beginning of a "long 
file name" which extends to the next double-quote mark. Aside from this, 
LAUNCHER does not interpret double-quote marks. It is thus necessary in 
some cases to know whether double-quote marks will be present in the 
LAUNCHER command line when configuring the launch parameters, as they 
are not automatically stripped in %1 substitution. If there are 
applications which sometimes include double-quote marks around file 
names when calling an external program and other times do not, it may be 
impossible to code parameters that will launch both cases correctly.

========================================================================

Changes in Launcher version 1.1b
--------------------------------

1. Corrected (hopefully) mishandling of long file names containing 
blanks when launched by Windows associations or entered (enclosed in 
double quotes) as the first token in the command parameter.

2. Added note in this Read Me file regarding double quotes in Windows 
initialization files, and revised known problems statement.

========================================================================

Changes in Launcher version 1.1a
--------------------------------

1. Launcher now supports request as well as execute method DDE.
     
2. Launcher now supports the "ifexec" key.

3. Launcher can now send a second DDE message to activate (give focus 
to) an application that was already running.

4. A [*] section can be defined in the initialization file to control 
launching of files from programs that do not permit the definition of 
command parameters (but do allow the specification of an application).

========================================================================

Installation
------------

The Zip file contains both source and executable files. The only files 
needed to run the program are LAUNCHER.EXE and LAUNCHER.INI; both files 
MUST be placed in the same directory. The LAUNCHER.INI file must be 
customized using a text editor: the program does not alter this file in 
any way. Then command lines in application settings or in the registry 
can be changed to use LAUNCHER when the additional flexibility of this 
program is desired.


INI file
--------

For each different set of special launching instructions, create a 
section in the INI file. You can name the sections whatever you like, so 
long as each name contains letters and numbers only (NO BLANKS!) and is 
less than 64 characters in length. (These names are not case sensitive.) 
It is also possible to define a [*] section, which is used when LAUNCHER 
is passed a file name without a section name.

(Note: Launcher uses standard Windows routines to process its 
initialization file. If the value for a parameter both begins and ends 
with a double quote ["], these routines remove the beginning and ending 
quotes. Other double quote characters within the value are not affected. 
For example, to specify that the "command" parameter is to be:
     "%1"
it is necessary to write:
     command=""%1""
in the initialization file.)

In each section, these parameters can be used:

command=
     specifies the command line used to start the application.
        This parameter is scanned for the string %1: if present, it is 
     replaced by the trailing part of the LAUNCHER command. (Only the 
     first occurrence, if any, of %1 is replaced.)
        After %1 replacement, the first token of the command line is 
     identified as the file name to be launched. Environment variables 
     (e.g., %UserName%) are recognized in the file name (but not in any 
     argument string that follows).
        In a [*] section, if this parameter is omitted (or null) the 
     trailing part of the command line is taken to be a file name to be 
     launched according to Windows file type associations; if this 
     parameter is omitted in any other section, the launch will fail 
     unless DDE parameters are specified and the target application is 
     already running.
        If DDE parameters are specified, the command is issued only if 
     the application is not already running.

directory=
     specifies a working directory (or "Start In" property) for the 
     command. Environment variables are recognized in this parameter.
     
ddeexec=             (required if application or topic is specified)
     specifies the data to be passed to the application via DDE. This 
     parameter is scanned for the string %1: if present, it is replaced 
     by the trailing part of the LAUNCHER command. (Only the first 
     occurrence, if any, of %1 is replaced.)

application=         (required if ddeexec or topic is specified)
topic=               (required if ddeexec or application is specified)
     specify the application and topic names under which the application 
     will process the command specified in the ddeexec parameter.	

ifexec=              (allowed only if ddeexec is specified)
     specifies the data to be passed via DDE to the application if it 
     was not already running, after the application has been launched. 
     This parameter is scanned for the string %1: if present, it is 
     replaced by the trailing part of the LAUNCHER command. (Only the 
     first occurrence, if any, of %1 is replaced.)
        If this parameter is omitted, the data specified in "ddeexec" is 
     used whether the application was already running or not. If this 
     parameter is specified but is null (after %1 substitution), a DDE 
     message is not sent if the application was not already running.

request=             (allowed only if ddeexec is specified)
     any non-zero numeric value specifies that "request method" DDE, 
     rather than "execute method DDE," will be used. Request method is 
     typically used by web browsers.

timeout=             (allowed only if ddeexec is specified)
     specifies the length of time, in milliseconds, the system will 
     allow for the application to accept the DDE transaction. (This 
     parameter defaults to 2500.)

delay=               (allowed only if ddeexec is specified)
     specifies the maximum length of time, in milliseconds, that should 
     elapse between launching the application and initiating a DDE 
     conversation. After launching an application, LAUNCHER waits for it 
     to be ready for user input before attempting to start a DDE 
     conversation. The waiting time is limited by this parameter, which 
     defaults to 32767.

IgnoreNotProcessed=  (allowed only if ddeexec is specified)
     any non-zero numeric value specifies that the DDE Execute error 
     code which indicates the server did not process the request should 
     be treated as indicating success. (That error code occurs when 
     launching a file into Opera using execute method DDE, even though 
     the launch is successful; this parameter suppresses the resulting 
     message. Though LAUNCHER now supports request mode DDE, which is 
     recommended for Opera and does not result in this error, I saw no 
     reason to remove the parameter.)

ActivateCommand=     (allowed only if ddeexec is specified)
ActivateTopic=       (allowed only if ddeexec is specified)
     specify a topic and command used to send a second DDE message to 
     make the application active. If only ActivateCommand is specified, 
     the topic used will be the same as for the first DDE message; if 
     only ActivateTopic is specified, "ActivateCommand=-1,0" is assumed. 
     This message is sent only when the application is not launched. The 
     same application and method (request or execute) are used as for 
     the first DDE message.
     
force=              (effective only in an [*] section)
     any non-zero numeric value specifies that the entire command line 
     following LAUNCHER is used as the "%1" substitution string: no 
     attempt is made to parse a section name.

override=
     specifies the full path to an alternate initialization file to be 
     used instead of the default initialization file. When this 
     parameter is specified in a named section, LAUNCHER searches for a 
     "command=" setting in the same section of the indicated file. If 
     one is found, all the settings are taken from that file; otherwise, 
     the settings in the original file are used. Environment variables 
     are recognized in this parameter.
        When this parameter is specified in a [*] section, it applies 
     both to the [*] section and to any named sections in which this 
     parameter is not specified.

The easiest way to determine the correct command, ddeexec, application, 
topic and ifexec parameters is probably to find a key in the registry 
for a file type that launches the application: the parameters work the 
same way in LAUNCHER (except that if an ifexec key exists in the 
registry but is null, you should omit the ifexec parameter in LAUNCHER 
rather than specify a null value). The other parameters can usually be 
omitted at first: it will be fairly obvious if they're needed. If 
ddeexec, application and topic are all omitted, DDE is not used.

You can use environment variables to specify parameters that must be 
different for different users. You can use an environment variable in 
the override parameter to allow individual users to have entirely 
different settings for some or all sections.

The LAUNCHER.INI file provided shows sample parameters for launching a 
file into Opera (e.g., from the Agent MIME Types table); for launching a 
mailto: link into Agent; for launching into a typical DDE-aware 
application; and for launching a non-DDE-aware application with a 
working directory. The [*] and [WebBrowser] sections demonstrate how to 
establish Internet Explorer as a default browser for all users, while 
allowing individual users to override this (for example, to use Opera) 
by placing a LAUNCHER.INI file in an appropriate sub-folder of their 
profile. To complete this setup, the command line:
     "C:\Program Files\Launcher\Launcher.exe" WebBrowser %1
would have to be set as the default value in the registry for the keys:
     HKEY_CLASSES_ROOT\htmlfile\shell\open\command
     HKEY_CLASSES_ROOT\http\shell\open\command
     HKEY_CLASSES_ROOT\https\shell\open\command
(with no accompanying ddeexec key).

Command lines
-------------

LAUNCHER is useful when a program that can launch other programs only 
provides for a simple command line, but something like the way
File Manager opens documents (using DDE, if defined in the registry, to 
launch multiple files into a single instance) is desired.

Normally, the command line takes either of these forms to launch a file 
according to the registry association for the file type (that is, as if 
the file were double-clicked in File Manager):

     path\launcher.exe file\path\and\name
     path\launcher.exe * file\path\and\name

however, it is possible to override this behavior by specifying a [*] 
section in the initialization file. The parameters in the [*] section, 
if it exists, are used when launching by either of the above formats.

The command line to launch according to the parameters specified in a 
named section of the INI file is:

     path\launcher.exe section file\path\and\name

where "section" is one of the section names defined in LAUNCHER.INI 
(without the enclosing square brackets). The trailing part of the 
command does not actually have to be a file path and name: the entire 
string of characters beginning with the first non-blank character 
following the section identifier and extending to the end of the command 
line is substituted for any %1 arguments in the initialization file.

If launcher.exe is copied or renamed, the file name constructed by 
replacing ".exe" with ".ini" is used to initialize that copy of 
LAUNCHER. This may be useful in connection with the [*] section (and 
possibly "force=1") to create multiple "instances" of LAUNCHER which can 
be used to control opening files from applications that permit only the 
entry of the name of an executable module to open a document, rather 
than providing for specification of a complete command line.