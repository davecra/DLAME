using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Threading;
using Microsoft.Win32;
using System.Windows.Forms;

namespace DLAME
{
    public partial class ThisAddIn
    {
        private const string MCstrREGKEY = "Software\\Microsoft\\Office\\Outlook\\DelayedAddins";
        private List<string> MobjAddinNames = new List<string>();
        private int MintDelay = 0;

        /// <summary>
        /// STARTUP
        /// On startup we connect to the startup event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // hook to the startup event. It will only fire after all
                // the add-ins have been loaded. 
                Application.Startup += Application_Startup;
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log("ThisAddin_Startup() failed.");
            }
        }

        /// <summary>
        /// EVENT
        /// This event fires when Outlook starts, but after ALL other
        /// add-ins have been loaded.
        /// </summary>
        private void Application_Startup()
        {
            try
            {
                // PRIMARY PURPOSE:
                // start a thread, wait a second and then load the
                // add-in. We do this because we want to give Outlook
                // a chance to fully connect to the server before
                // specific add-ins load and interfere with the
                // connection to the Exchange Server
                new Thread(() =>
                   {
                       try
                       {
                           Thread.Sleep(1000); // wait one second
                           loadDelayedAddins();
                       }
                       catch (Exception PobjEx)
                       {
                           PobjEx.Log("Thread failed.");
                       }
                   }).Start();
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log("Application_Startup(): Unable to start thread.");
            }
        }

        /// <summary>
        /// HELPER
        /// When the new explorer event first, we load the add-ins listed in
        /// Delayed load registry key
        /// </summary>
        /// <param name="Explorer"></param>
        private void loadDelayedAddins()
        {
            try
            {
                if(!readDelayedAddins()) // load names
                    return; // failed - exception in load

                // next verify we have values
                if (MobjAddinNames.Count == 0)
                    return; // no add-ins were in the delayed add-in list
                            // so we exit and do nothing more

                // sleep more if specified
                if (MintDelay > 0) Thread.Sleep(MintDelay * 1000);

                // loop through each add-in from the COMAddins list
                // and if we find it from the delayed list, enable it
                foreach (Office.COMAddIn LobjAddin in Application.COMAddIns)
                {
                    if (MobjAddinNames.Contains(LobjAddin.ProgId))
                    {
                        try
                        {
                            LobjAddin.Connect = true; // load it
                        }
                        catch (Exception PobjEx)
                        {
                            PobjEx.Log("Unable to load: " + LobjAddin.ProgId);
                        }
                    }
                }
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log("loadDelayedAddins() failed.");
            }
        }

        /// <summary>
        /// HELPER
        /// Reads the list of add-ins from the delayed registry key
        /// </summary>
        /// <returns>True is successful</returns>
        private bool readDelayedAddins()
        {
            try
            {
                RegistryKey LobjKey = null;
                try
                {
                    LobjKey = Registry.CurrentUser.OpenSubKey(MCstrREGKEY);
                    if(LobjKey == null)
                    {
                        LobjKey = Registry.LocalMachine.OpenSubKey(MCstrREGKEY);
                    }
                }
                catch { } // ignore a failure here
                if (LobjKey != null)
                {
                    try
                    {
                        MintDelay = int.Parse(LobjKey.GetValue("").ToString());
                    }
                    catch { } // fail quietely - value likely was not set
                              // and that is OK, we just do not want to fail 

                    // load the values
                   foreach (string LstrValue in LobjKey.GetValueNames())
                    {
                        MobjAddinNames.Add(LstrValue);
                    }

                    return true; // tell the caller success
                }
                else
                {
                    throw new Exception("Registry cannot be accessed or DelayedAddins key is missing.");
                }
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log("readDelayedAddins() failed.");
                return false; // tell caller we failed
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }
        #endregion
    }
}
