
//Import libraries
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace UIRequirement;

public enum RUN
{
    IHI = 0,
    CYIENT = 1
}
class Utility
{
    public static RUN m_eRUN = RUN.CYIENT;
    static int myCnt = 0;
    public static string m_sTempPath = "";
    public static string m_sBinPath = "";

    public static string m_sToolName = "CTR Form Automation Tool";
    public static string m_sVersion = "3.1";
    public static string m_sReleaseDate = "10-July-2026";
    //*******************************************************
    //Function  : GetTempDir
    //Purpose   : Get the temporaory directory
    //*******************************************************
    public static void GetTempDir()
    {
        //Get temp Directory
        string path = null;
        try
        {
            //Get the TEMP Path
            path = System.Environment.GetEnvironmentVariable("TEMP");
            if (path == null || path.Length < 1)//Not found
            {
                //Get the TMP path
                path = System.Environment.GetEnvironmentVariable("TMP");
                if (path == null || path.Length < 1)//Not found
                {
                    path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
                }
            }
            m_sTempPath = path + @"\";
        }
        catch { }
    }
    //*******************************************************
    //Function  : WriteErrorLog
    //Purpose   : Write the error data to log
    //*******************************************************
    public static void WriteErrorLog(string sWarning, string sStackTrace, string sMsg)
    {
        try
        {
            if (sWarning.Trim().Length > 0)
                MessageBox.Show(sWarning.Trim());
            string sLogFilepath = m_sTempPath + "ErrorLog.dat";
            if (sLogFilepath != null && sLogFilepath.Length > 0)
            {
                try
                {
                    //Write the Error Log
                    StreamWriter sw;
                    if (myCnt == 0)//First Error
                    {
                        sw = System.IO.File.CreateText(sLogFilepath);
                        DateTime date = DateTime.Now;
                        sw.WriteLine("********** " + Utility.m_sToolName + " V" + Utility.m_sVersion + " *********");
                        sw.WriteLine("Release Date          : " + m_sReleaseDate);
                        sw.WriteLine("Time                  : " + date.ToString());
                        sw.WriteLine("********************************************");
                    }
                    else//Next Errors Append
                    {
                        sw = System.IO.File.AppendText(sLogFilepath);
                    }
                    //----- This is for error
                    sw.WriteLine("--------------------------------------------");
                    if (sWarning.Length > 0)
                        sw.WriteLine("Warning     : " + sWarning);//Write the Error Message
                    if (sStackTrace.Length > 0)
                        sw.WriteLine("StackTrace  : " + sStackTrace);//Write the StackTrace
                    if (sMsg != null && sMsg.Length > 0)
                        sw.WriteLine("Message  : " + sMsg);//Write the StackTrace
                    //close
                    sw.Close();
                    myCnt++;
                }
                catch
                { }
            }
        }
        catch { }
    }
    public static void WriteErrorLog(Exception ee)
    {
        try
        {
            if (ee != null)
            {
                WriteErrorLog(ee.Message, ee.StackTrace, "");
            }
        }
        catch { }
    }
    public static void WriteErrorLog(string msg)
    {
        try
        {
            WriteErrorLog("", "", msg);
        }
        catch { }
    }
    //*******************************************************
    //Function  : WriteErrorLog
    //Purpose   : Write the error data to log
    //*******************************************************
    public static void WarnUser(string smsg)
    {
        try
        {
            //show
            //MessageBox.Show(smsg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
    }
    //*******************************************************
    //Function  : InformationUser
    //Purpose   : information message box to user
    //*******************************************************
    public static void InformationUser(string smsg)
    {
        try
        {
            //show
            //MessageBox.Show(smsg, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
    }
    //*******************************************************
    //Function  : DeleteFile
    //Purpose   : delete the file
    //*******************************************************
    public static void DeleteFile(string sFile)
    {
        try
        {
            if (System.IO.File.Exists(sFile) == true)
            {
                System.IO.File.Delete(sFile);
            }
        }
        catch
        {
            // Utility.WriteErrorLog(ee);
        }
    }
    //*******************************************************
    //Function  : GetInteger
    //Purpose   : convert string to integer
    //*******************************************************
    public static int GetInteger(string sValue)
    {
        int cnt = 0;
        try
        {
            sValue = sValue.Trim();
            if (sValue.Length > 0)
            {
                int n1 = 0;
                if (int.TryParse(sValue, out n1))
                {
                    cnt = n1;
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return cnt;
    }
    //*******************************************************
    //Function  : GetReal
    //Purpose   : Get the real value from string
    //*******************************************************
    public static double GetReal(string sValue)
    {
        double cnt = 0;
        try
        {
            sValue = sValue.Trim();
            if (sValue.Length > 0)
            {
                double n1 = 0;
                if (double.TryParse(sValue, out n1))
                {
                    cnt = n1;
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return cnt;
    }
    public static bool IsReal(string sValue)
    {
        try
        {
            sValue = sValue.Trim();
            if (sValue.Length > 0)
            {
                double n1 = 0;
                if (double.TryParse(sValue, out n1))
                {
                    return true;
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return false;
    }
    //*******************************************************
    //Function  : SubString
    //Purpose   : sub string string from another string
    //*******************************************************
    public static string SubString(string str, string substr)
    {
        try
        {
            if (str.Length > 0 && substr.Length > 0)
            {
                return str.Substring(substr.Length, str.Length - substr.Length).Trim();
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return str;
    }
    //*******************************************************
    //Function  : GetInteger
    //Purpose   : get the integer value from string
    //*******************************************************
    public static bool GetInteger(string sValue, ref int nVal)
    {
        try
        {
            sValue = sValue.Trim();
            if (sValue.Length > 0)
            {
                if (int.TryParse(sValue, out nVal))
                {
                    return true;
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return false;
    }
    //*******************************************************
    //Function  : ReadFile
    //Purpose   : read the file
    //*******************************************************
    public static List<string> ReadFile(string sFile)
    {
        List<string> lstOp = new List<string>();
        try
        {
            string[] sArr = System.IO.File.ReadAllLines(sFile);
            foreach (string ss in sArr)
            {
                if (ss.Trim().Length > 0)
                {
                    lstOp.Add(ss.Trim());
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return lstOp;
    }
    //*******************************************************
    //Function  : SplitString
    //Purpose   : split the string by delimeter
    //*******************************************************
    public static List<string> SplitString(string sValue, string schar)
    {
        List<string> lst = new List<string>();
        try
        {
            string[] sArr = sValue.Split(new string[1] { schar }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string ss in sArr)
            {
                if (ss.Trim().Length > 0)
                {
                    lst.Add(ss.Trim());
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return lst;
    }
    public static List<string> SplitString(string sValue, string[] schars)
    {
        List<string> lst = new List<string>();
        try
        {
            string[] sArr = sValue.Split(schars, StringSplitOptions.RemoveEmptyEntries);
            foreach (string ss in sArr)
            {
                if (ss.Trim().Length > 0)
                {
                    lst.Add(ss.Trim());
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return lst;
    }
    //*******************************************************
    //Function  : trimString
    //Purpose   : trim the string
    //*******************************************************
    public static string trimString(string str)
    {
        try
        {
            if (str.EndsWith(":"))
            {
                return str.Substring(0, str.Length - 1);
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return str;
    }
    //*******************************************************
    //Function  : CombineString
    //Purpose   : combined list of strings by delimeter
    //*******************************************************
    public static string CombineString(List<string> lst, string schar)
    {
        string sval = "";
        try
        {
            foreach (string ss in lst)
            {
                sval = sval + ss + schar;
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return sval.Trim();
    }
    //*******************************************************
    //Function  : SplitString
    //Purpose   : split the string by delimeters
    //*******************************************************

    public static List<string> RemoveDuplicatedFromList(List<string> lstGiven)
    {
        List<string> lst = new List<string>();
        try
        {
            foreach (string ss in lstGiven)
            {
                if (lst.Contains(ss) == false)
                    lst.Add(ss);
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return lst;
    }

    //*******************************************************
    //Function  : MyStringComparison
    //Purpose   : Compare 2 strings
    //*******************************************************
    public static bool MyStringComparison(string str1, string str2, int typeCompare, bool isToupperCompare)
    {
        try
        {
            //use correct spacing
            string strnew1 = CombineString(SplitString(str1, " "), " ");
            string strnew2 = CombineString(SplitString(str2, " "), " ");
            //convert to upper
            if (isToupperCompare)
            {
                strnew1 = strnew1.ToUpper();
                strnew2 = strnew2.ToUpper();
            }
            //---->>> compare
            //1-Contains, 2-Endswith, 3-startwith
            if (typeCompare == 1 && strnew1.Contains(strnew2))
                return true;
            else if (typeCompare == 2 && strnew1.EndsWith(strnew2))
                return true;
            else if (typeCompare == 3 && strnew1.StartsWith(strnew2))
                return true;
            else if (typeCompare == 4 && strnew1 == strnew2)
                return true;
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return false;
    }
    public static bool MyStringComparison(string str1, string str2, bool isToupperCompare)
    {
        try
        {
            //use correct spacing
            string strnew1 = CombineString(SplitString(str1, " "), " ");
            string strnew2 = CombineString(SplitString(str2, " "), " ");
            //convert to upper
            if (isToupperCompare)
            {
                strnew1 = strnew1.ToUpper();
                strnew2 = strnew2.ToUpper();
            }
            if (strnew1 == strnew2)
                return true;
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return false;
    }
    public static bool IsInRange(double min, double max, double dmingiven, double dmaxgiven)
    {
        try
        {
            //bool byes = false;
            double dratio = (max - min) / 10.0;
            for (int j = 0; j <= 10; j++)
            {
                double dv = min + dratio * j;
                if (dv >= dmingiven && dv <= dmaxgiven)
                {
                    return true;
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return false;
    }
    public static string GetDate(string str)
    {
        string month = "", year = "";
        try
        {
            str = str.ToUpper();
            //Month
            if (str.Contains("JAN")) month = "Jan";
            else if (str.Contains("FEB")) month = "Feb";
            else if (str.Contains("MAR")) month = "Mar";
            else if (str.Contains("APR")) month = "Apr";
            else if (str.Contains("MAY")) month = "May";
            else if (str.Contains("JUN")) month = "Jun";
            else if (str.Contains("JUL")) month = "Jul";
            else if (str.Contains("AUG")) month = "Aug";
            else if (str.Contains("SEP")) month = "Sep";
            else if (str.Contains("OCT")) month = "Oct";
            else if (str.Contains("NOV")) month = "Nov";
            else if (str.Contains("DEC")) month = "Dec";
            //Year
            for (int x = 1950; x < 2100; x++)
            {
                if (str.Contains(x.ToString()))
                {
                    year = x.ToString(); break;
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return month + "." + year;
    }
    public static void releaseObject(object obj)
    {
        try
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            obj = null;
        }
        catch (Exception ee)
        {
            obj = null;
            Utility.WriteErrorLog(ee);
        }
        finally
        {
            GC.Collect();
        }
    }
    public static string TrimWordTableCell(string str)
    {
        try
        {
            return str.Replace("\r\a", "").Trim();
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return str.Trim();
    }
    public static string CaclulateMinute(string str)
    {
        try
        {
            //":" means hours..No need of conversion
            //"." means minutes..Need conversion
            if (str.Contains(":") == false)
            {
                str = str.Replace(":", ".").Replace(",", "").Replace(" ", "").Trim();
                if (str.Contains("."))
                {
                    List<string> lst1 = SplitString(str, ".");
                    if (lst1.Count == 2)
                    {
                        int n1 = 0;
                        if (GetInteger(lst1[1], ref n1))
                        {
                            double d1 = n1 / Math.Pow(10, lst1[1].Length);
                            string s2 = Math.Round(d1 * 0.6, 2).ToString();
                            List<string> lst2 = SplitString(s2, ".");
                            if (lst2.Count == 2)
                            {
                                return lst1[0] + "." + lst2[1];
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return str;
    }
    public static string GetPDFAttribute(string str, string type)
    {
        string val = "";
        try
        {
            List<string> lst1 = Utility.SplitString("cyient " + str, type);
            if (lst1.Count >= 2)
            {
                List<string> lst2 = Utility.SplitString(lst1[1], " ");
                if (lst2.Count >= 2 && (lst2[0].StartsWith("-") || lst2[0].StartsWith(":")))
                {
                    val = lst2[1];
                }
                else
                {
                    val = lst2[0];
                }
                if (val.StartsWith("-") || val.StartsWith(":"))
                {
                    val = val.Substring(1, val.Length - 1).Trim();
                }
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return val;
    }
    public static string GetDateString(DateTime time)
    {
        try
        {
            return GetDateString(time.Month, time.Day, time.Year);
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return "";
    }
    public static string GetDateString(int m, int d, int y)
    {
        try
        {
            return DateTimeFormatInfo.CurrentInfo.GetAbbreviatedMonthName(m) + "-" + d.ToString().PadLeft(2, '0') + "-" + y.ToString();
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return "";
    }
    public static string CopyFileToTempPath(string sTemplate)
    {
        string sop = "";
        try
        {
            if (System.IO.File.Exists(sTemplate))
            {
                int cnt = 1;
                while (true)
                {
                    if (cnt == 100)
                    {
                        sop = "";
                        break;
                    }

                    sop = Utility.m_sTempPath + System.IO.Path.GetFileNameWithoutExtension(sTemplate) + "_" + cnt.ToString() + System.IO.Path.GetExtension(sTemplate);
                    try
                    {
                        System.IO.File.Copy(sTemplate, sop, true);
                        break;
                    }
                    catch
                    {
                        cnt = cnt + 1;
                    }
                }
            }
            else
            {
                Utility.WarnUser("File not exists...\n" + sTemplate);
            }
        }
        catch (Exception ee)
        {
            Utility.WriteErrorLog(ee);
        }
        return sop;
    }
}
