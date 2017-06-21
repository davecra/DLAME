# DLAME
Outlook Delayed Load Add-in Manager for the Enterprise

# Executive Summary
This add-in helps an enterprise delay load Outlook add-in to avoid potential issues those add-ins may cause when Outlook launches.

# Summary
The Outlook Delayed Loading of Add-ins Manager for the Enterprise (or DLAME Add-in) is provided as an option for enterprises to select specific add-ins that they wish to have load after Outlook has loaded all other add-ins and from 1 to (n) seconds after Outlook has fully loaded.

# Installation
The DLAME Add-in is supported on Outlook 2013 and Outlook 2016. To install the adding, please follow these steps:
1)	Make sure Outlook is closed
2)	Locate the “DLAME.vsto” and double-click on it.
3)	This will install the add-in into Outlook. When completed, you will get a notification that it is installed.

To install for all users a script will be required. You can install the add-in silently using the following command line:

      %commonprogramfiles%\microsoft shared\VSTO\10.0\VSTOInstaller.exe \i \ s <path>\DLAME.vsto 

The switches are:
 - /s – silent
 - /i – install

NOTE: You can uninstall with the /u switch. 

# Configuration
To configure the DLAME add-in you will create the following registry key:
  - Per User: HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\DelayedAddins
  - Per Machine: HKEY_LOCAL_MACHINE\Software\Microsoft\Office\Outlook\DelayedAddins
To specify the delay beyond 1 second, you will place a number in the DEFAULT string value already in the key. To be certain the Outlook User Interface is completely loaded, and all connection are established a setting of 5 is suggested. Next, you will create a string value under the DelayedAddins key for each add-in you wish to delay load. To do this:
1)	Look in the following Registry Key: HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins.
2)	Find the Key name (on the left side) for the add-in you wish to delay load and:
    a.	Right-click and select Rename
    b.	Select and copy the name.
    c.	Click away. Do NOT change or CUT the name.
3)	Next, click on that add-in key and change the LoadBehavior value to 0 (zero).
4)	In the DelayedAddins key, create a new string value.
5)	Paste the name of the add-in copied from step 2.
For example, if you wish to delay load the OneNote add-in, you will set the load behavior for the “OneNote.OutlookAddin” to 0 (zero) and then create a new string value in the DelayedAddins key with the same name: “OneNote.OutlookAddin.”

NOTE: If the add-in you wish to delay load is currently loaded from HKEY_LOCAL_MACHINE, you will need to remove the key from HKLM and move it to HKCU. If the concern is about loading the add-in on every start, then you should both load DLAME from HKLM and add it to the resiliency list, per the following section.

## Registry Keys Detailed
This section includes the exact REG import files you can create for your add-in and deploy via policy. Here is how you would configure a REG key to Current User:

> Windows Registry Editor Version 5.00
> [HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\DelayedAddins]
> "<add-in name>"=""
> @="5"

Here is how you would configure a REG file to Local Machine:

> Windows Registry Editor Version 5.00 
> [HKEY_LOCAL_MACHINE\Software\Microsoft\Office\Outlook\DelayedAddins]
> "<add-in name>"=""
> @="5"

## Add DLAME to Resiliency List
DLAME is a .NET/VSTO add-in and as such requires the .NET framework to load in order for it to operate. This can cause a delay in loading the add-in and Outlook might disable it. To prevent this follow the steps outlined in the following article:
https://msdn.microsoft.com/VBA/Outlook-VBA/articles/support-for-keeping-add-ins-enabled
You will create the following Policy Registry key:

> HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList
> DLAME: 0x00000001

# Operation
The DLAME add-in will load at the startup of Outlook. It will wait one second after ALL add-ins have been loaded and then read the DelayedAddins registry key as well as the default value to get additional delay load time. It will then sleep/delay for the remained of the specified time. Once completed with the delay, it will cycle through all the registered add-ins and load those specified in the DelayedAddins registry key.

NOTE: It must ALWAYS load at the startup of Outlook. You cannot delay load this add-in or no delayed loaded add-ins will ever be loaded.

# Troubleshooting
If the add-in is not behaving properly, please see the following table to issues and their possible resolutions. Also, please note that if there are any EXCEPTIONS in the add-in, they will be logged in the temp folder with a name like this: DLAMEYYYYMMDDHHnnss.log

## Known Issues and Resolutions
 - Nothing happens – the delayed add-ins are not loaded.	Either the add-ins is set to be disabled in the registry, or you attempted to delay load it as well, or there was some serious failure on startup. In the last case, please see if a temp file exists with more information about the likely cause. 
 - Not all the add-ins are delay loaded as desired	This might happen if you have specified the add-in name incorrectly in the DelayedAddins list, or the add-in you wanted to delay load failed to load when it was tasked to by the D-LAME add-in. To verify the later, you can remove it from the delay load and verify that it is functioning correctly. If it is unable to be delay loaded, then please contact the vendor.
 - The delay loaded add-in is not functioning correctly	The add-in might require to be loaded with the start of Outlook. If you remove it from the delay load and verify that it is functioning correctly when launched with Outlook, it might be unable to be delay loaded. Please contact the vendor.
 - The add-in I want to delay, delays the first time, but then loads with Outlook on the next start	The add-in might have code to make sure its load behavior is always set to 3 (load at startup). Or, there might be a “companion add-in” that performs the same operation. In either case you may not be able to delay load the add-in. Please contact the vendor.

## Known Exception Points
Throughout the DLAME add-in there are specific exception points that you might see and some very generic possible causes. Here is a list of them:
- ThisAddin_Startup() failed: This is the main entry point for the add-in. If you see an Exception here it is likely caused by the inability to attach to the Outlook Application_Startup event. This might happen if there is a conflict with another add-in or a problem with your Outlook installation.
- Thread failed: This will occur if something interrupts the background thread that is created.
-  Application_Startup(): Unable to start thread: This will occur if the system is unable to generate a background thread, which is used to delay load the add-ins without causing the entire Outlook Application to be affected.
- Unable to load: <addin>: This will occur if there is a problem enabling a specific add-in you placed in the DleayedAddins key. This can happen of that add-in has a problem (see its documentation), or if there is something interfering with the enabling of the add-in, such as another add-in or possibly anti-virus or execution prevention policies.
- loadDelayedAddins() failed: More generic problem occurred in the loading of add-ins, likely cause by the inability to read the Registry, memory access issues or a problem with Outlooks ability to read its add-in list.
- Registry cannot be accessed or DelayedAddins key is missing: This likely occurred because the Registry cannot be read or the DLAME DelayedAddins key is missing.
- readDelayedAddins() failed: Generic error reading the Registry. Likely caused by something interfering with the ability of the add-in to access the Registry or read specific values. It might be permissions or interference from another application.

# Support
This tool is provided As-is.
