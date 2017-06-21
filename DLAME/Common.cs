using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DLAME
{
    public static class Common
    {
        public const string MCstrNAME = "DLAME";
        public static string MstrLogPath = "";

        /// <summary>
        /// EXTENSION METHOD
        /// Log any exceptions to a temp file
        /// </summary>
        /// <param name="PobjEx"></param>
        /// <param name="PstrMessage"></param>
        public static void Log(this Exception PobjEx, string PstrMessage = "")
        {
            try
            {
                // get a temp path if this is the first time
                if (string.IsNullOrEmpty(MstrLogPath))
                {
                    MstrLogPath = Path.Combine(Path.GetTempPath(), MCstrNAME + DateTime.Now.ToFileString() + ".log");
                }
                // write to the file
                StreamWriter LobjWriter = new StreamWriter(MstrLogPath, true);
                LobjWriter.WriteLine(DateTime.Now + "\t" + PstrMessage + "\t" + PobjEx.ToString().CleanString());
                LobjWriter.Close();
            }
            catch { } // fail quietly
        }

        /// <summary>
        /// EXTENSION METHOD
        /// Cleans the string of new line characters and tabs, replacing them
        /// with the space (" ") character
        /// </summary>
        /// <param name="PstrString"></param>
        /// <returns></returns>
        public static string CleanString(this string PstrString)
        {
            try
            {
                return PstrString.Replace("\t", " ")
                                 .Replace("\r\n", " ")
                                 .Replace("\r", " ")
                                 .Replace("\n", " ");
            }
            catch
            {
                return ""; // failed
            }
        }

        /// <summary>
        /// EXTENSION METHOD
        /// Returns a string representing the specified date and time in the
        /// format of: YYYYMMDDHHnnss
        /// </summary>
        /// <param name="PobjTime"></param>
        /// <returns></returns>
        public static string ToFileString(this DateTime PobjTime)
        {
            try
            {
                return DateTime.Now.Year.ToString() +
                       DateTime.Now.Month.ToString() +
                       DateTime.Now.Day.ToString() +
                       DateTime.Now.Hour.ToString() +
                       DateTime.Now.Minute.ToString() +
                       DateTime.Now.Second.ToString();
            }
            catch
            {
                return ""; // failed
            }
        }
    }
}
