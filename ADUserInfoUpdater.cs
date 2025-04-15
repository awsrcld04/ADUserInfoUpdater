using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using Microsoft.Win32;
using System.Management;
using System.Reflection;
using System.Diagnostics;

namespace ADUserInfoUpdater
{
    class AUIUMain
    {
        struct OUSections
        {
            public List<OUInfo> lstOUSections;
        }

        struct OUInfo
        {
            public string strOUPath;
            public string strCompany;
            public string strOffice;
            public string strStreet;
            public string strPOBox;
            public string strCity;
            public string strState;
            public string strZipCode;
            public string strCountry;
            public string strDepartment;
            public string strTelephoneNumber;
        }

        struct CMDArguments
        {
            public bool bParseCmdArguments;
        }

        static void funcPrintParameterWarning()
        {
            Console.WriteLine("Parameters must be specified properly to run ADUserInfoUpdater.");
            Console.WriteLine("Run ADUserInfoUpdater -? to get the parameter syntax.");
        }

        static void funcPrintParameterSyntax()
        {
            Console.WriteLine("ADUserInfoUpdater v1.0");
            Console.WriteLine();
            Console.WriteLine("Parameter syntax:");
            Console.WriteLine();
            Console.WriteLine("Use the following required parameters in the following order:");
            Console.WriteLine("-run                     required parameter");
            Console.WriteLine();
            Console.WriteLine("Example:");
            Console.WriteLine("ADUserInfoUpdater -run");
        }

        static CMDArguments funcParseCmdArguments(string[] cmdargs)
        {
            CMDArguments objCMDArguments = new CMDArguments();

            try
            {
                if (cmdargs[0] == "-run" & cmdargs.Length == 1)
                {
                    objCMDArguments.bParseCmdArguments = true;
                }
                else
                {
                    objCMDArguments.bParseCmdArguments = false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                objCMDArguments.bParseCmdArguments = false;
            }

            return objCMDArguments;
        }

        static void funcProgramExecution(CMDArguments objCMDArguments2)
        {
            try
            {
                // [DebugLine] Console.WriteLine("Entering funcProgramExecution");
                if (funcCheckForFile("configADUserInfoUpdater.txt"))
                {
                    if (funcValidateConfigSections())
                    {
                        funcToEventLog("ADUserInfoUpdater", "ADUserInfoUpdater started.", 1001);

                        funcProgramRegistryTag("ADUserInfoUpdater");

                        bool bOUSectionFlag = false;

                        OUSections newOUSections = new OUSections();
                        newOUSections.lstOUSections = new List<OUInfo>();
                        OUInfo currentOUInfo = new OUInfo();

                        //[DebugLine] Console.WriteLine();

                        TextReader trConfigFile = new StreamReader("configADUserInfoUpdater.txt");

                        using (trConfigFile)
                        {
                            string strNewLine = "";

                            while ((strNewLine = trConfigFile.ReadLine()) != null)
                            {

                                if (strNewLine == "OUSectionBegin" & !bOUSectionFlag)
                                {
                                    bOUSectionFlag = true;
                                    //Console.WriteLine("OUSectionBegin");
                                }
                                if (bOUSectionFlag)
                                {

                                    if (strNewLine.StartsWith("OUPath="))
                                    {
                                        currentOUInfo.strOUPath = strNewLine.Substring(7);
                                        //Console.WriteLine(strNewLine.Substring(7));
                                    }

                                    if (strNewLine.StartsWith("Company="))
                                    {
                                        currentOUInfo.strCompany = strNewLine.Substring(8);
                                        //Console.WriteLine(strNewLine.Substring(8));
                                    }

                                    if (strNewLine.StartsWith("Office="))
                                    {
                                        currentOUInfo.strOffice = strNewLine.Substring(7);
                                        //Console.WriteLine(strNewLine.Substring(7));
                                    }

                                    if (strNewLine.StartsWith("Street="))
                                    {
                                        currentOUInfo.strStreet = strNewLine.Substring(7);
                                        //Console.WriteLine(strNewLine.Substring(7));
                                    }

                                    if (strNewLine.StartsWith("POBox="))
                                    {
                                        currentOUInfo.strPOBox = strNewLine.Substring(6);
                                        //Console.WriteLine(strNewLine.Substring(6));
                                    }

                                    if (strNewLine.StartsWith("City="))
                                    {
                                        currentOUInfo.strCity = strNewLine.Substring(5);
                                        //Console.WriteLine(strNewLine.Substring(5));
                                    }

                                    if (strNewLine.StartsWith("State="))
                                    {
                                        currentOUInfo.strState = strNewLine.Substring(6);
                                        //Console.WriteLine(strNewLine.Substring(6));
                                    }

                                    if (strNewLine.StartsWith("ZipCode="))
                                    {
                                        currentOUInfo.strZipCode = strNewLine.Substring(8);
                                        //Console.WriteLine(strNewLine.Substring(8));
                                    }

                                    if (strNewLine.StartsWith("Country="))
                                    {
                                        currentOUInfo.strCountry = strNewLine.Substring(8);
                                        //Console.WriteLine(strNewLine.Substring(8));
                                    }

                                    if (strNewLine.StartsWith("Department="))
                                    {
                                        currentOUInfo.strDepartment = strNewLine.Substring(11);
                                        //Console.WriteLine(strNewLine.Substring(11));
                                    }

                                    if (strNewLine.StartsWith("TelephoneNumber="))
                                    {
                                        currentOUInfo.strTelephoneNumber = strNewLine.Substring(16);
                                        //Console.WriteLine(strNewLine.Substring(16));
                                    }

                                }
                                if (strNewLine == "OUSectionEnd")
                                {
                                    OUInfo newOUInfo = currentOUInfo;
                                    newOUSections.lstOUSections.Add(newOUInfo);
                                    bOUSectionFlag = false;
                                    //Console.WriteLine("OUSectionEnd");
                                }
                            }
                        }

                        //Console.WriteLine(newOUSections.lstOUSections.Count.ToString());

                        trConfigFile.Close();

                        foreach (OUInfo sectionOUInfo in newOUSections.lstOUSections)
                        {
                            if (funcCheckForOU(sectionOUInfo.strOUPath))
                            {
                                PrincipalContext currentctx = funcCreateScopedPrincipalContext(sectionOUInfo.strOUPath);
                                // [DebugLine] Console.WriteLine(currentctx.Container);

                                // Create the principal user object from the context
                                UserPrincipal usr = new UserPrincipal(currentctx);

                                // Create a PrincipalSearcher object.
                                PrincipalSearcher ps = new PrincipalSearcher(usr);
                                PrincipalSearchResult<Principal> psr = ps.FindAll();

                                TextWriter twCurrent = funcOpenOutputLog();
                                string strOutputMsg = "";

                                foreach (UserPrincipal u in psr)
                                {
                                    //[DebugLine] Console.WriteLine("{0} \t {1}",u.Name,u.DistinguishedName);
                                    //[DebugLine] Console.WriteLine();
                                    DirectoryEntry newDE = new DirectoryEntry("LDAP://" + u.DistinguishedName);
                                    newDE.Properties["physicalDeliveryOfficeName"].Value = sectionOUInfo.strOffice;
                                    newDE.Properties["telephoneNumber"].Value = sectionOUInfo.strTelephoneNumber;
                                    newDE.Properties["streetAddress"].Value = sectionOUInfo.strStreet;
                                    newDE.Properties["postOfficeBox"].Value = sectionOUInfo.strPOBox;
                                    newDE.Properties["l"].Value = sectionOUInfo.strCity;
                                    newDE.Properties["st"].Value = sectionOUInfo.strState;
                                    newDE.Properties["postalCode"].Value = sectionOUInfo.strZipCode;
                                    newDE.Properties["co"].Value = sectionOUInfo.strCountry;
                                    if (sectionOUInfo.strCountry == "United States" | sectionOUInfo.strCountry == "US")
                                    {
                                        newDE.Properties["c"].Value = "US";
                                    }
                                    newDE.Properties["department"].Value = sectionOUInfo.strDepartment;
                                    newDE.Properties["company"].Value = sectionOUInfo.strCompany;
                                    newDE.CommitChanges();

                                    strOutputMsg = "Attributes for " + u.Name + " have been updated. " + "(" + u.DistinguishedName + ")";

                                    funcWriteToOutputLog(twCurrent, strOutputMsg);
                                }

                                twCurrent.Close();
                            }
                            else
                            {
                                Console.WriteLine("Invalid OUPath: {0}", sectionOUInfo.strOUPath);
                            }
                        }

                        funcToEventLog("ADUserInfoUpdater", "ADUserInfoUpdater stopped.", 1002);
                    }
                    else
                    {
                        Console.WriteLine("The config file is not valid.");
                    }
                }
                else
                {
                    Console.WriteLine("Config file configADUserInfoUpdater.txt could not be found.");
                }

            }
            catch (Exception ex)
            {
                // [DebugLine] Console.WriteLine(ex.Source);
                // [DebugLine] Console.WriteLine(ex.Message);
                // [DebugLine] Console.WriteLine(ex.StackTrace);

                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

        }

        static DirectorySearcher funcCreateDSSearcher()
        {
            try
            {
                System.DirectoryServices.DirectorySearcher objDSSearcher = new DirectorySearcher();
                // [Comment] Get local domain context

                string rootDSE;

                System.DirectoryServices.DirectorySearcher objrootDSESearcher = new System.DirectoryServices.DirectorySearcher();
                rootDSE = objrootDSESearcher.SearchRoot.Path;
                //Console.WriteLine(rootDSE);

                // [Comment] Construct DirectorySearcher object using rootDSE string
                System.DirectoryServices.DirectoryEntry objrootDSEentry = new System.DirectoryServices.DirectoryEntry(rootDSE);
                objDSSearcher = new System.DirectoryServices.DirectorySearcher(objrootDSEentry);
                //Console.WriteLine(objDSSearcher.SearchRoot.Path);

                return objDSSearcher;
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return null;
            }
        }

        static PrincipalContext funcCreatePrincipalContext()
        {
            PrincipalContext newctx = new PrincipalContext(ContextType.Machine);

            try
            {
                //Console.WriteLine("Entering funcCreatePrincipalContext");
                Domain objDomain = Domain.GetComputerDomain();
                string strDomain = objDomain.Name;
                DirectorySearcher tempDS = funcCreateDSSearcher();
                string strDomainRoot = tempDS.SearchRoot.Path.Substring(7);
                // [DebugLine] Console.WriteLine(strDomainRoot);
                // [DebugLine] Console.WriteLine(strDomainRoot);

                newctx = new PrincipalContext(ContextType.Domain,
                                    strDomain,
                                    strDomainRoot);

                // [DebugLine] Console.WriteLine(newctx.ConnectedServer);
                // [DebugLine] Console.WriteLine(newctx.Container);



                //if (strContextType == "Domain")
                //{

                //    PrincipalContext newctx = new PrincipalContext(ContextType.Domain,
                //                                    strDomain,
                //                                    strDomainRoot);
                //    return newctx;
                //}
                //else
                //{
                //    PrincipalContext newctx = new PrincipalContext(ContextType.Machine);
                //    return newctx;
                //}
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

            if (newctx.ContextType == ContextType.Machine)
            {
                Exception newex = new Exception("The Active Directory context did not initialize properly.");
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, newex);
            }

            return newctx;
        }

        static PrincipalContext funcCreateScopedPrincipalContext(string strOUPath)
        {
            PrincipalContext newctx = new PrincipalContext(ContextType.Machine);

            try
            {
                //Console.WriteLine("Entering funcCreatePrincipalContext");
                Domain objDomain = Domain.GetComputerDomain();
                string strDomain = objDomain.Name;

                newctx = new PrincipalContext(ContextType.Domain,
                                    strDomain,
                                    strOUPath);

                // [DebugLine] Console.WriteLine(newctx.ConnectedServer);
                // [DebugLine] Console.WriteLine(newctx.Container);



                //if (strContextType == "Domain")
                //{

                //    PrincipalContext newctx = new PrincipalContext(ContextType.Domain,
                //                                    strDomain,
                //                                    strDomainRoot);
                //    return newctx;
                //}
                //else
                //{
                //    PrincipalContext newctx = new PrincipalContext(ContextType.Machine);
                //    return newctx;
                //}
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

            if (newctx.ContextType == ContextType.Machine)
            {
                Exception newex = new Exception("The Active Directory context did not initialize properly.");
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, newex);
            }

            return newctx;
        }

        //static bool funcCheckNameExclusion(string strName, DisabledAccountsParams listParams)
        //{
        //    try
        //    {
        //        bool bNameExclusionCheck = false;

        //        //List<string> listExclude = new List<string>();
        //        //listExclude.Add("Guest");
        //        //listExclude.Add("SUPPORT_388945a0");
        //        //listExclude.Add("krbtgt");

        //        if (listParams.lstExclude.Contains(strName))
        //            bNameExclusionCheck = true;

        //        //string strMatch = listExclude.Find(strName);
        //        foreach (string strNameTemp in listParams.lstExcludePrefix)
        //        {
        //            if (strName.StartsWith(strNameTemp))
        //            {
        //                bNameExclusionCheck = true;
        //                break;
        //            }
        //        }

        //        return bNameExclusionCheck;
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mb1 = MethodBase.GetCurrentMethod();
        //        funcGetFuncCatchCode(mb1.Name, ex);
        //        return false;
        //    }
        //}

        static void funcToEventLog(string strAppName, string strEventMsg, int intEventType)
        {
            try
            {
                string strLogName;

                strLogName = "Application";

                if (!EventLog.SourceExists(strAppName))
                    EventLog.CreateEventSource(strAppName, strLogName);

                //EventLog.WriteEntry(strAppName, strEventMsg);
                EventLog.WriteEntry(strAppName, strEventMsg, EventLogEntryType.Information, intEventType);
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static bool funcCheckForOU(string strOUPath)
        {
            try
            {
                string strDEPath = "";

                if (!strOUPath.Contains("LDAP://"))
                {
                    strDEPath = "LDAP://" + strOUPath;
                }
                else
                {
                    strDEPath = strOUPath;
                }

                if (DirectoryEntry.Exists(strDEPath))
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static bool funcCheckForFile(string strInputFileName)
        {
            try
            {
                if (System.IO.File.Exists(strInputFileName))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static void funcGetFuncCatchCode(string strFunctionName, Exception currentex)
        {
            string strCatchCode = "";

            Dictionary<string, string> dCatchTable = new Dictionary<string, string>();
            dCatchTable.Add("funcGetFuncCatchCode", "f0");
            dCatchTable.Add("funcPrintParameterWarning", "f2");
            dCatchTable.Add("funcPrintParameterSyntax", "f3");
            dCatchTable.Add("funcParseCmdArguments", "f4");
            dCatchTable.Add("funcProgramExecution", "f5");
            dCatchTable.Add("funcCreateDSSearcher", "f7");
            dCatchTable.Add("funcCreatePrincipalContext", "f8");
            dCatchTable.Add("funcCreateScopedPrincipalContext", "f9");
            dCatchTable.Add("funcCheckNameExclusion", "f10");
            dCatchTable.Add("funcFindAccountsToDisable", "f11");
            dCatchTable.Add("funcCheckLastLogin", "f12");
            dCatchTable.Add("funcRemoveUserFromGroup", "f13");
            dCatchTable.Add("funcToEventLog", "f14");
            dCatchTable.Add("funcCheckForFile", "f15");
            dCatchTable.Add("funcCheckForOU", "f16");
            dCatchTable.Add("funcWriteToErrorLog", "f17");
            dCatchTable.Add("funcValidateConfigSections", "f18");

            if (dCatchTable.ContainsKey(strFunctionName))
            {
                strCatchCode = "err" + dCatchTable[strFunctionName] + ": ";
            }

            //[DebugLine] Console.WriteLine(strCatchCode + currentex.GetType().ToString());
            //[DebugLine] Console.WriteLine(strCatchCode + currentex.Message);

            funcWriteToErrorLog(strCatchCode + currentex.GetType().ToString());
            funcWriteToErrorLog(strCatchCode + currentex.Message);

        }

        static void funcWriteToErrorLog(string strErrorMessage)
        {
            try
            {
                FileStream newFileStream = new FileStream("Err-ADUserInfoUpdater.log", FileMode.Append, FileAccess.Write);
                TextWriter twErrorLog = new StreamWriter(newFileStream);

                DateTime dtNow = DateTime.Now;

                string dtFormat = "MMddyyyy HH:mm:ss";

                twErrorLog.WriteLine("{0} \t {1}", dtNow.ToString(dtFormat), strErrorMessage);

                twErrorLog.Close();
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

        }

        static TextWriter funcOpenOutputLog()
        {
            try
            {
                DateTime dtNow = DateTime.Now;

                string dtFormat2 = "MMddyyyy"; // for log file directory creation

                string strPath = Directory.GetCurrentDirectory();

                string strLogFileName = strPath + "\\ADUserInfoUpdater" + dtNow.ToString(dtFormat2) + ".log";

                FileStream newFileStream = new FileStream(strLogFileName, FileMode.Append, FileAccess.Write);
                TextWriter twOuputLog = new StreamWriter(newFileStream);

                return twOuputLog;
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return null;
            }

        }

        static void funcWriteToOutputLog(TextWriter twCurrent, string strOutputMessage)
        {
            try
            {
                DateTime dtNow = DateTime.Now;

                string dtFormat = "MMddyyyy HH:mm:ss";

                twCurrent.WriteLine("{0} \t {1}", dtNow.ToString(dtFormat), strOutputMessage);
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void funcCloseOutputLog(TextWriter twCurrent)
        {
            try
            {
                twCurrent.Close();
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static bool funcValidateConfigSections()
        {
            try
            {
                TextReader trConfigFile = new StreamReader("configADUserInfoUpdater.txt");
                bool bOUSectionBegin = false;
                bool bInvalidConfigFile = false;

                using (trConfigFile)
                {
                    string strNewLine = "";

                    while ((strNewLine = trConfigFile.ReadLine()) != null)
                    {
                        //[DebugLine] Console.WriteLine(strNewLine);
                        if (strNewLine == "OUSectionBegin" & bOUSectionBegin)
                        {
                            bInvalidConfigFile = true;
                            break;
                        }
                        if (strNewLine == "OUSectionBegin" & !bOUSectionBegin)
                        {
                            bInvalidConfigFile = false;
                            bOUSectionBegin = true;
                        }
                        if (strNewLine == "OUSectionEnd")
                        {
                            bOUSectionBegin = false;
                        }
                    }
                }

                trConfigFile.Close();

                if (bInvalidConfigFile)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                // [DebugLine] Console.WriteLine(ex.Source);
                // [DebugLine] Console.WriteLine(ex.Message);
                // [DebugLine] Console.WriteLine(ex.StackTrace);

                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                {
                    funcPrintParameterWarning();
                }
                else
                {
                    if (args[0] == "-?")
                    {
                        funcPrintParameterSyntax();
                    }
                    else
                    {
                        string[] arrArgs = args;
                        CMDArguments objArgumentsProcessed = funcParseCmdArguments(arrArgs);

                        if (objArgumentsProcessed.bParseCmdArguments)
                        {
                            funcProgramExecution(objArgumentsProcessed);
                        }
                        else
                        {
                            funcPrintParameterWarning();
                        } // check objArgumentsProcessed.bParseCmdArguments
                    } // check args[0] = "-?"
                } // check args.Length == 0
            }
            catch (Exception ex)
            {
                Console.WriteLine("errm0: {0}", ex.Message);
            }
        }
    }
}
