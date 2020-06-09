using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security;
using System.Web;


namespace SOPManagement.Models
{
    public static class  Utility
    {


        static Logger oLogger = new Logger();

        /// <summary>

        /// This method is used to encrypt the particular section in the application config

        /// </summary>

        /// <param name="section">Section to encrypt</param>

        //EncryptAppSettings
        public static void ProtectConfiguration()

        {

            // oLogger.LogFileName = HttpContext.Current.Server.MapPath("~/Content/DocFiles/LogFile/ ")+"ProcessLog.txt";

            oLogger.LogFileName = HttpContext.Current.Server.MapPath(ConfigurationManager.AppSettings["logfilepathnm"]);

            // Get the application configuration file.
            //Configuration config =
            //        ConfigurationManager.OpenExeConfiguration(
            //        ConfigurationUserLevel.None);

            Configuration config = null;
            if (HttpContext.Current != null)
            {
                config =
                    System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration("~");
            }
            else
            {
                config =
                    ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            }


            // Define the Rsa provider name.
            string provider =
                "RsaProtectedConfigurationProvider";

            // Get the section to protect.
            ConfigurationSection connStrings =
                config.AppSettings;
            // connStrings = null;

            if (connStrings != null)
            {
                if (!connStrings.SectionInformation.IsProtected)
                {
                    if (!connStrings.ElementInformation.IsLocked)
                    {
                        // Protect the section.
                        connStrings.SectionInformation.ProtectSection(provider);

                        connStrings.SectionInformation.ForceSave = true;
                        config.Save(ConfigurationSaveMode.Full);


                        //Console.WriteLine("Section {0} is now protected by {1}",
                        //    connStrings.SectionInformation.Name,
                        //    connStrings.SectionInformation.ProtectionProvider.Name);

                    }
                    else
                    {
                        //Console.WriteLine(
                        //     "Can't protect, section {0} is locked",
                        //     connStrings.SectionInformation.Name);


                        oLogger.UpdateLogFile(DateTime.Now.ToString() + "Can't protect, appsetting is locked");
                    }
                }
                else
                {
                    //Console.WriteLine(
                    //    "Section {0} is already protected by {1}",
                    //    connStrings.SectionInformation.Name,
                    //    connStrings.SectionInformation.ProtectionProvider.Name);


                }

            }
            else
            {

                //Console.WriteLine("Can't get the section {0}",
                //    connStrings.SectionInformation.Name);

            }


        }

        public static void UnProtectConfiguration()
        {

            // Get the application configuration file.
            //System.Configuration.Configuration config =
            //        ConfigurationManager.OpenExeConfiguration(
            //        ConfigurationUserLevel.None);

            Configuration config = null;
            if (HttpContext.Current != null)
            {
                config =
                    System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration("~");
            }
            else
            {
                config =
                    ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            }



            // Get the section to unprotect.
            ConfigurationSection connStrings =
                config.AppSettings;



            if (connStrings != null)
            {
                if (connStrings.SectionInformation.IsProtected)
                {
                    if (!connStrings.ElementInformation.IsLocked)
                    {
                        // Unprotect the section.
                        connStrings.SectionInformation.UnprotectSection();

                        connStrings.SectionInformation.ForceSave = true;
                        config.Save(ConfigurationSaveMode.Full);

                        Console.WriteLine("Section {0} is now unprotected.",
                            connStrings.SectionInformation.Name);

                    }
                    else
                        Console.WriteLine(
                             "Can't unprotect, section {0} is locked",
                             connStrings.SectionInformation.Name);
                }
                else
                    Console.WriteLine(
                        "Section {0} is already unprotected.",
                        connStrings.SectionInformation.Name);

            }
            else
                Console.WriteLine("Can't get the section {0}",
                    connStrings.SectionInformation.Name);

        }

        public static SecureString GetSecureString(String Password)
        {
            SecureString oSecurePassword = new SecureString();

            foreach (Char c in Password.ToCharArray())
            {
                oSecurePassword.AppendChar(c);

            }
            return oSecurePassword;
        }

        //We'll define expired session as situation when Session.IsNewSession is true 
        // (it is a new session), but  session cookie already exists on visitor's computer 
        //from previous session.Here is a procedure that returns true if session is expired and returns false if not.

        //Session.IsNewSession property tells us if session is created during current request or not.
        //If value is true, it is a new session.If value is false, it is existing active session created before.

        //public static bool IsSessionExpired()
        //{
        //    if (System.Web.HttpContext.Current.Session != null)
        //    {
        //        if (System.Web.HttpContext.Current.Session.IsNewSession)
        //        {
        //            string CookieHeaders = System.Web.HttpContext.Current.Request.Headers["Cookie"];

        //            if ((null != CookieHeaders) && (CookieHeaders.IndexOf("ASP.NET_SessionId") >= 0))
        //            {
        //                // IsNewSession is true, but session cookie exists,
        //                // so, ASP.NET session is expired
        //                return true;
        //            }
        //        }
        //    }

        //    // Session is not expired and function will return false,
        //    // could be new session, or existing active session
        //    return false;
        //}

        public static bool IsSessionExpired()
        {
            bool isSessionOut = false;

            if (HttpContext.Current.Session["UserFullName"] == null)
                isSessionOut = true;

            return isSessionOut;

        }



        public static bool IsNumeric(string value)
        {
            return value.All(char.IsNumber);
        }

        public static string GetSiteAdminUserName()
        {
            string tmpval = "";
            tmpval = ConfigurationManager.AppSettings["siteadmnusereml"];

            return tmpval;

        }

        public static string GetDocLibraryName()
        {
            string tmpval = "";
            tmpval = ConfigurationManager.AppSettings["doclibraryname"];

            return tmpval;

        }

        public static string GetSiteAdminPassowrd()
        {
            string tmpval = "";
            tmpval = ConfigurationManager.AppSettings["siteadmnpassword"];

            return tmpval;

        }


        public static string GetCurrentLoggedInUserEmail()
        {

            string loggedinuser = "";

            string loggedinuseremail = "";

            string clddomainnm = "";

            clddomainnm = ConfigurationManager.AppSettings["clouddomainname"];

         
            loggedinuser = HttpContext.Current.User.Identity.Name;

            loggedinuser = loggedinuser.Split('\\').Last();

            loggedinuseremail=loggedinuser + "@"+ clddomainnm;

            return loggedinuseremail;


        }


        public static int GetLoggedInUserID()
        {
            string useremail="";
            int userid = 0;

            useremail = GetCurrentLoggedInUserEmail().Trim();

          //  useremail = "mschmidt@radiantdelivers.com";

            using (var dbctx = new RadiantSOPEntities())
            {

                userid = dbctx.users.Where(u => u.useremailaddress.Trim().ToLower() == useremail.Trim().ToLower() && u.userstatuscode==1).Select(u => u.userid).FirstOrDefault();
            }

            return userid;

        }

        public static short GetLoggedInUserSOPDeptCode()
        {
            string useremail = GetCurrentLoggedInUserEmail();

          //  string useremail = "mschmidt@radiantdelivers.com";

            int userid = 0;
            short deptcode = 0;

            short sopdeptcode = 0;

            userid = GetLoggedInUserID();


            using (var dbctx = new RadiantSOPEntities())
            {

                deptcode = Convert.ToInt16(dbctx.users.Where(u => u.useremailaddress.Trim().ToLower() == useremail.Trim().ToLower() && u.userstatuscode==1).Select(u => u.departmentcode).FirstOrDefault());
                sopdeptcode= Convert.ToInt16(dbctx.codesdepartments.Where(u => u.departmentcode == deptcode).Select(u => u.sopdeptcode).FirstOrDefault());

            }

            return sopdeptcode;

        }


        public static string GetLoggedInUserSOPDeptName()
        {
            string useremail = GetCurrentLoggedInUserEmail();

            //  string useremail = "mschmidt@radiantdelivers.com";

            string sopdeptname = "";


            using (var dbctx = new RadiantSOPEntities())
            {
                sopdeptname = dbctx.vwUsers.Where(u => u.useremailaddress.Trim().ToLower() == useremail.Trim().ToLower()).Select(u => u.departmentname).FirstOrDefault();

            }

            return sopdeptname;

        }



        public static string GetTempLocalDirPath()
        {

            string tlocaldir = "";
            tlocaldir = ConfigurationManager.AppSettings["templocaldir"];
          //  tlocaldir = HttpContext.Current.Server.MapPath(tlocaldir);

            return tlocaldir;


        }

        public static string GetLogFilePath()
        {

            string tlocalpath = "";
            tlocalpath = ConfigurationManager.AppSettings["logfilepathnm"];
            //  tlocaldir = HttpContext.Current.Server.MapPath(tlocaldir);

            return tlocalpath;


        }


        public static string GetTemplateFileName()
        {

            string templatefile = "";
            templatefile = ConfigurationManager.AppSettings["templatefilename"];
           

            return templatefile;


        }

        public static string GetSiteUrl()
        {

            string tempval = "";
            tempval = ConfigurationManager.AppSettings["siteurl"];

            return tempval;


        }

        public static string GetDashBoardUrl()
        {

            string tempval = "";
            tempval = ConfigurationManager.AppSettings["dashboardurl"];

            return tempval;


        }

        public static string GetLoggedInUserFullName()
        {


            string fullname = "";
            RadiantSOPEntities ctx = new RadiantSOPEntities();

            //lsopno = foldername + "-001";

            fullname = ctx.getUserFullNameByEmailUserID(GetCurrentLoggedInUserEmail(), 0).FirstOrDefault().ToString();

            return fullname;

        }

        

        public static List<SOPClass> GetFolders()
        {

            List<SOPClass> folderlist;

            using (var ctx = new RadiantSOPEntities())
            {

                var folders = ctx.deptsopfiles.Select(x => new SOPClass()
                {
                    FileID = x.FileID,
                    FileName = x.DeptFileName,
                    FilePath = x.SPFilePath,
                    FileLink = x.SPFileLink,
                    SOPNo = x.SOPNo,
                    FileStatuscode = x.filestatuscode

                }).Where(s => s.FilePath == "SOP/" && s.FileID!=193 && s.FileStatuscode == 3).OrderBy(s=>s.FileName);


                folderlist = folders.ToList();


            }


            return folderlist;


        }


     

    }



}