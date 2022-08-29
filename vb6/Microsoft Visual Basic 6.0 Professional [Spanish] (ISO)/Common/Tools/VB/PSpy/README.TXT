This document contains the installation instructions for
the Process Viewer - PSpy.EXE.

Table of Contents:
  I. Installation for users of Windows NT 3.51/4
 II. Installation for users of Windows 95
III. Running PSpy
 IV. What PSPY is
  V. What PSPY is Not
 VI. Working Sets


I. Installation for users of Windows NT 3.51/4:
---------------------------------------------
1. PSpy requires that Visual Basic is installed.
2. PSpy will now run from the CD.

3. If you want to install PSpy to your machine, copy the 
following files to your \System32 directory:
	PERFINFO.DLL
	WORKSET.DLL
4. Copy PSPY.EXE to your hard drive, and create a Program
Manager icon for it.

II. Installation for users of Windows 95:
-----------------------------------------
1. PSpy requires that Visual Basic is installed.
2. Copy the following file to your \System directory:
	MEMMON.VXD
3. Copy the following file to the \System\VMM32 directory:
	VMM.VXD
4. Add the following line to the [386Enh] section of your
SYSTEM.INI file:
	Device=memmon.vxd
5. Restart Windows 95.
6. PSpy will now run from the CD.

7. If you want to install PSpy to your machine, copy the 
following files to your \System32 directory:
	PERFINFO.DLL
	WORKSET.DLL
8. Copy PSPY.EXE to your hard drive, and create a Program
Manager icon for it.

III. Running PSpy:
------------------
1. To spy on a process, choose the 'Examine' option from
the Process menu.
2. Click and drag the magnifying glass on the Process you
wish to view information.
3. Click OK.
4. Click on a file name on the left half of the window to show
details for that file.
5. Click the "Flush" button to flush the active working set
for the program you are spying on (see below).
6.  Click the "Refresh" button to refresh the working set
display for the program.
7. Right clicking on either pane will bring up a context
menu.  From here you can change the view and copy the
display to the clipboard.

IV. What PSPY is:
-----------------
1. A tool to help you locate where DLLs in memory are
loaded from - much like WPS.EXE in Windows 3.1.
2. A tool to identify the versions of all DLL's that
your program is using.
3. A tool to provide information on the Working Set of a 
process.  See below for a description of what the Working
Set for a process is.

V. What PSPY is Not:
--------------------
1. PSpy is not supported by PSS, though they may have you
use it to help debug a problem.
2. PSpy does not trap window messages like the Spy and 
Spy++ utilities shipped in Microsoft Visual C++.

VI. Working Sets:
-----------------
32 bit operating systems like Windows 95 and Windows NT
support memory paging.  This means that blocks of memory
that haven't been used for a while can be "paged" to a
temporary file on disk, which frees up memory for other
programs to use.  This mechanism allows you to run more
programs than you would otherwise have memory for, and
makes the operating system run smoother because it doesn't
need to keep so much information in physical RAM.  This
memory paging happens automatically; programs are oblivious
to any paging activity.

The "Working Set" of a program is the amount of 
physical RAM that the operating system is currently giving
your program.  As a simple example, let's say you have
a program that creates a byte array that contains 5 million
elements but you only access the first ten.  If other programs
need more memory, the operating system will page out all but
the ten elements that you are using.  Your program doesn't
know the difference and doesn't pay any speed penalty.  If,
however, you later try to access element 1 million, there 
will be a slight delay while the operating system fetches
the memory block that contains element 1 million from disk.

A programs data is not the only component of the working
set.  The actual program code is another component.  PSpy
can break out the program code from the data and show which
modules are using how much memory.  It can also "Flush" the 
working set.  This is a special command to the operating 
system that says, "make this program take NO physical RAM".
As long as the program is idle, it won't take any memory.
As soon as the program needs to execute code, however, the
OS will page memory back in.  Flushing is a good way to find
out how much memory a particular operation takes.  For example, 
to find out how much memory is taken by loading a file in 
an application, you would perform the following steps in PSpy:

    1.  Start PSpy and the application you want to check.
        Begin spying on the app.

    2.  Click the flush button of PSpy to flush the app's
        working set.

    3.  Open a file with the application.

    4.  Click the refresh button on PSpy.  This will refresh
        the working set numbers for the application.

The working set numbers that are now displayed on PSpy reflect
the amount of RAM required by your program to load the file.

Advanced topic:  Working Set on Windows 95 vs Windows NT

Windows 95 treats shared DLL's differently than Windows NT. Under
Windows NT, the working set for a shared DLL accurately reflects
the working set of the DLL for the process being monitored.  Windows
95, however, uses a shared address space for shared DLL's.  Working
set numbers for shared DLL's under Windows 95 reflect the combined
working set for all processes that use the DLL.  This causes working
set numbers under Windows 95 to appear higher than on Windows NT.  
For example, a minimal application will show a subtantial working
set contribution from kernel32 because this DLL is also in use by 
the system.

If you are curious about the working set or how paging works, see
"Examples of Memory Activity and Paging" in the Windows NT Resource
Kit under the section, "Optimizing Windows NT".
