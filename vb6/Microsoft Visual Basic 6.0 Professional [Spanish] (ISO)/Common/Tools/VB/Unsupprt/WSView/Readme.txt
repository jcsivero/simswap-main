\Tools\Unsupprt\WSView

Working Set Viewer
==================

To install the Working Set Viewer:

Windows95:

- Copy MEMMON.VXD and PSAPI.DLL from \Tools\Unsupprt\WSView\Win95\ to you Windows 95 
  system directory.  For example (assuming the VB CD is in drive D:):

  cd c:\win95\system
  copy d:\Tools\Unsupprt\WSView\Win95\memmon.vxd
  copy d:\Tools\Unsupprt\WSView\Win95\psapi.dll

- Copy \Tools\Unsupprt\WSView\Win95\VMM.VXD to the \VMM32 directory under your Windows
  95 system directory.  For example:

  cd c:\win95\system\vmm32
  copy d:\Tools\Unsupprt\WSView\Win95\vmm.vxd

- Add the following line to your system.ini file in the [386Enh] section:

  Device = memmon.vxd

- Reboot Windows 95 and then run WSVIEW.EXE

Windows NT:

- Copy \Tools\Unsupprt\WSView\WinNT\PSAPI.DLL to your \System32 directory and then run
  WSVIEW.EXE



