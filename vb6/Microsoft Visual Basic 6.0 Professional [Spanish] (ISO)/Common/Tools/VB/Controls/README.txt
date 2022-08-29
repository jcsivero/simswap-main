\Tools\Controls

This directory contains all of the ActiveX Controls that shipped with Visual 
Basic 4.0/5.0 Professional and Enterprise Editions, which are no longer shipping 
with Visual Basic 6.0.

AniBtn32.ocx
Gauge32.ocx
Graph32.ocx
Gsw32.EXE
Gswdll32.DLL
Grid32.ocx
KeySta32.ocx
MSOutl32.ocx
Spin32.ocx
Threed32.ocx
MSChart.ocx

The \Tools\Controls\BiDi directory contains a Bi-directional version of 
Grid32.Ocx.

If you have Visual Basic 5.0 Professional or Enterprise Editions installed on 
your machine, you should already have these ActiveX controls available to you in 
Visual Basic 6.0.  

Graph32.ocx has been updated to work properly in Visual Basic 6.0 and it 
requires two additional support files: gsw32.exe and gswdll32.dll.  You must 
place the three files together in the \Windows\System directory or the control 
will not function properly.

If you do not have these controls and wish to use these in Visual Basic 6.0, you 
can install them by:

1. Copy all of the files in this directory to your \WINDOWS\SYSTEM directory.

2. Register the controls by either Browsing to them in Visual Basic itself, or 
manually register them using RegSvr32.Exe.  RegSvr32.EXE can be found in the 
\Tools\RegistrationUtilities directory.  The command line is:

regsvr32.exe grid32.ocx

3. Register the design time licenses for the controls.  To do this, merge the 
vbctrls.reg file found in this directory into your registry.  You can merge this 
file into your registry using RegEdit.Exe (Win95 or WinNT4) or RegEd32.Exe 
(WinNT3.51):

regedit vbctrls.reg (or other reg files associated with the controls)


