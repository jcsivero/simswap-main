\DlgObj

Microsoft Dialog Automation Objects
===================================

The Microsoft Dialog Automation Objects component provides Common Dialog 
functionality without requiring that you place a control on a form.  

To use install the Microsoft Dialog Automation Objects:

1.  Copy DLGOBJS.DLL from VB98\Wizards\PDWizard to your \Windows\System
    directory (or \System32 directory on Windows NT).

2.  Register the design time license by merging the Registry file DLGOBJS.REG
    into your registry.

    - Under Windows NT 4.0 and Windows 95/98, right click on
      \Tools\Unsupprt\DlgObj\DLGOBJS.REG and choose 'Merge'

    - Under Windows NT 3.51, copy \Tools\Unsupprt\DlgObj\DLGOBJS.REG to your
      hard drive and merge it into the registry using RegEdt32.Exe

3.  Register DLGOBJS.DLL by either using RegSvr32.Exe found in \Tools\RegUtils
    or by Browsing for the DLGOBJS.DLL in the Project|References Dialog in 
    Visual Basic and clicking 'Open'.

4.  Once the DLL is registered, select "Microsoft Dialog Automation Objects"
    in the Project|References Dialog in Visual Basic.

