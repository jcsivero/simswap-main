
Microsoft ® Script Control

To Install:
===========

1)	Copy all Microsoft Script Control files into your system directory
		For Windows95 this is typically C:\Windows\System
		For WindowsNT this is typically C:\WinNT\System32

	Example:
		copy msscr*.* C:\Windows\System


2)	Register the control using RegSvr32

	Example:
		cd /d C:\Windows\System
		regsvr32 msscript.ocx

File List:
==========

msscript.ocx	
	Microsoft Script Control

msscript.hlp	
msscript.cnt	
	Microsoft Script Control Help Files

msscrXXX.dll
	International Resources
