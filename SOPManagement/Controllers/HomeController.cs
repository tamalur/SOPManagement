using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.SharePoint.Client;
using System.Security;
using Microsoft.SharePoint.Client.Search.Query;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop;
using System.Reflection;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ListItem = Microsoft.SharePoint.Client.ListItem;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using System.Xml.Linq;
using System.Globalization;
using System.Collections;
using SOPManagement.Models;
using Newtonsoft.Json;
using System.Text.RegularExpressions;
using Group = Microsoft.SharePoint.Client.Group;
using System.Data.Entity.Infrastructure;
using System.Data.Entity.Core.Objects;
using System.Threading;

//this controller was developed by Tamalur from April 10 to April 22, 2020
namespace SOPManagement.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        // string siteurl;

        string siteurl = "https://radiantdelivers.sharepoint.com/sites/watercooler";

        RadiantSOPEntities ctx = new RadiantSOPEntities();

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult LogOff()
        {
            Session["UserID"] = null; //it's my session variable
            Session.Clear();
            Session.Abandon();
        //    FormsAuthentication.SignOut(); //you write this when you use FormsAuthentication
            return RedirectToAction("Sessionouterr", "Home");
        }
        public ActionResult CreateUploadSOPAuth()
        {

            Session["SOPMsg"] = "";

            if (IsSessionExpired())
            {

                ViewBag.ErrorMessage = "SOP Application: Session not Timed out";
                Session["SOPMsg"] = "SOP Application: Session not Timed out";
                return RedirectToAction("SOPMessage");


            }


            //check admin user access, if not admin redirect to error page

            //string user = "";
            //string loggedinusereml = "";
            //bool isadmin = false;
            //Employee emp = new Employee();

            //user = System.Web.HttpContext.Current.User.Identity.Name;

            //user = user.Split('\\').Last();

            //Session["UserID"] = user;
            //loggedinusereml = user + "@radiantdelivers.com";
            //isadmin = emp.AuthenticateUser("admin", loggedinusereml);

            //if (!isadmin)
            //{

            //    Session["ErrorMsg"] = "SOP Application: Only admin user has access to this page. You are not a admin user. Please contact SOP team for access!";
            //    return RedirectToAction("RedirectForErr");


            //}

            //emp = null;


            return RedirectToAction("CreateUploadSOP");

        }

        public ActionResult ApproveSOP(int? id)
        {

            ViewBag.employees = (from c in ctx.users select new { c.useremailaddress, c.userfullname, c.userstatuscode }).Where(x => x.userstatuscode == 1).Distinct();
            Employee model = new Employee();

            ViewBag.FileID = id;

            model.HasSignedSOP = true;
     
            return View(model);

            // return View();
        }

        
        public ActionResult PublishFile(int? id)
        {
            //give this url to Elhadj 
           // http://localhost:58639/Home/PublishFile/293?chngreqid=18


            string changereqid = Request.QueryString["chngreqid"];


            int sopfileid = 0;
            int validchngreq = 0;

            Session["ErrorMsg"] = "";

            Session["Dashboardlink"] = "http://camis1-bioasp01/Reports/Pages/Report.aspx?ItemPath=%2fSOP+Reports%2fSOP+Dashboard";


            if (IsSessionExpired())
            {

               // ViewBag.ErrorMessage = "SOP Application: Session not Timed out";

                Session["ErrorMsg"] = "SOP Application: Session not Timed out";

                return RedirectToAction("SOPMessage");

            }



            if (id == null || changereqid == null || changereqid == "" || !IsNumeric(changereqid))     //no file provided
            {

                Session["ErrorMsg"] = "Error: Valid File ID and Cahneg Request ID is required to publish the file!";

                return RedirectToAction("SOPMessage");

            }

            else
            {
                validchngreq = Convert.ToInt32(changereqid);
                sopfileid = Convert.ToInt16(id);

                
            }


            Employee oEmp = new Employee();

            string loggedinuser = "";

            // var usr = System.Environment.UserName;

            loggedinuser = System.Web.HttpContext.Current.User.Identity.Name;

            loggedinuser = loggedinuser.Split('\\').Last();

            Session["UserID"] = loggedinuser;

            oEmp.useremailaddress = loggedinuser + "@radiantdelivers.com";

    
            if (oEmp.AuthenticateUser("approver", sopfileid))

            {

                SOPClass oSOP = new SOPClass();

                oSOP.SiteUrl = siteurl;

                oSOP.FileID = sopfileid;

                oSOP.FileChangeRqstID = validchngreq;

                oSOP.GetSOPInfoByFileID();  //get file path, name etc.


                if (oSOP.PublishFile())
                {


                    Session["SOPMsg"] = "SOP "+oSOP.FileName +" has been successfully published!";

                    return RedirectToAction("SOPMessage");

                }

                else
                {
                    Session["SOPMsg"] = "Failed to publish SOP " + oSOP.FileName + oSOP.ErrorMessage+ ".Please contact IT!";
                    return RedirectToAction("SOPMessage");
                }


            }
            else  //failed to authenticate logged in user as approver
            {
                Session["SOPMsg"] = "Failed to authenticate user as an approver of the file.Please contact IT!";
                return RedirectToAction("SOPMessage");
            }

        

            //return View();
        }

        [HttpPost]
        //public ActionResult PublishFile(SOPManagement.Models.SOPClass sopmodel)
        //{

        //    SOPClass sop = new SOPClass();

        //    sop.SiteUrl = siteurl;

        //  // sop.FileID = fileid;

        //    if (sop.PublishFile())

        //    {
        //        Session["SOPMsg"] = "SOP File " + sop.FileName + " was successfully published!";

        //    }
        //    else
        //        Session["SOPMsg"] = "Failed to publis SOP File " + sop.FileName + ".Please contact IT!";


        //    sop = null;

        //    //return RedirectToAction("SOPMessage");

        //    return View();


        //}

        public ActionResult SOPDashboard()
        {
            // do your logging here
            return Redirect("http://camis1-bioasp01/Reports/Pages/Report.aspx?ItemPath=%2fSOP+Reports%2fSOP+Dashboard");
        }

        public ActionResult UploadSOPFile()
        {
            ViewBag.Message = "Upload SOP File";



            ViewBag.ddlDeptFolders = new SelectList(GetFolders(), "FileName", "FileName");

            ViewBag.employees = (from c in ctx.users select new { c.useremailaddress, c.userfullname, c.userstatuscode,c.jobtitle }).Where(x => x.userstatuscode == 1).Distinct();

            ViewBag.departments = (from c in ctx.codesdepartments select new { c.departmentname, c.departmentcode }).Distinct();

            ViewBag.updfrequnits= (from c in ctx.codesUnits select new { c.Unitname, c.unitcode,c.UnitType }).Where(x=>x.UnitType == "UpdateFrequency").Distinct();

            return View();
        }

        [Authorize(Roles = "TransfloWheelsAdmnUsers")]
        //[Authorize(Roles = "TransfloARUsers")]
        //  [RoleFilter] with form authentication in web.cofig use this custom filter to redirect to custom page. make sure you don't use any role in authorize 
        [HttpGet]
        public ActionResult CreateUploadSOP()
        {

           // ViewBag.Title = "Upload or Create SOP";  //I assigned in cshtml file

            ViewBag.ddlDeptFolders = new SelectList(GetFolders(), "FileName", "FileName");

            ViewBag.employees = (from c in ctx.users select new { c.useremailaddress, c.userfullname, c.userstatuscode}).Where(x=>x.userstatuscode==1).Distinct();

            Session["employees"] = ViewBag.employees;

            ViewBag.departments = (from c in ctx.codesdepartments select new { c.departmentname, c.departmentcode }).Distinct();

            ViewBag.updfrequnits = (from c in ctx.codesUnits select new { c.Unitname, c.unitcode, c.UnitType }).Where(x => x.UnitType == "UpdateFrequency").Distinct();



            return View();
        }


        //We'll define expired session as situation when Session.IsNewSession is true 
        // (it is a new session), but  session cookie already exists on visitor's computer 
        //from previous session.Here is a procedure that returns true if session is expired and returns false if not.

        //Session.IsNewSession property tells us if session is created during current request or not.
        //If value is true, it is a new session.If value is false, it is existing active session created before.

        public static bool IsSessionExpired()
        {
            if (System.Web.HttpContext.Current.Session != null)
            {
                if (System.Web.HttpContext.Current.Session.IsNewSession)
                {
                    string CookieHeaders = System.Web.HttpContext.Current.Request.Headers["Cookie"];

                    if ((null != CookieHeaders) && (CookieHeaders.IndexOf("ASP.NET_SessionId") >= 0))
                    {
                        // IsNewSession is true, but session cookie exists,
                        // so, ASP.NET session is expired
                        return true;
                    }
                }
            }

            // Session is not expired and function will return false,
            // could be new session, or existing active session
            return false;
        }

        public ActionResult SOPMessage()
        {
            
            return View();

        }


        [HttpPost]
        public ActionResult CreateUploadSOP(SOPManagement.Models.SOPClass sop)
        {
        

            ViewBag.employees = (from c in ctx.users select new { c.useremailaddress, c.userfullname, c.userstatuscode, c.jobtitle }).Where(x => x.userstatuscode == 1).Distinct();


            //  if (!ModelState.IsValid)   //we are supposed to use ModelState but we validated data through javascript so we don't use this



            //start processing uploaded or new sop file 

            bool bProcessCompleted = false;

            if (sop.SubFolderName.Trim() == "Please select a subfolder")
            {

                sop.SubFolderName = "";

            }


            //1. [Authorized] attribute at the top of action authorizes the user by sopadmin role in domain 
            // then check if session is expired, if so redirect to session timeout page. 

            string user = "";
            string loggedinusereml = "";

            Session["ErrorMsg"] = "";

            if (IsSessionExpired())
             {

                ViewBag.ErrorMessage = "SOP Application: Session not Timed out";
                Session["ErrorMsg"] = "SOP Application: Session not Timed out";
                return Json(new { redirecturl = "/Home/RedirectForErr" }, JsonRequestBehavior.AllowGet);

            }


            // var usr = System.Environment.UserName;

            user = System.Web.HttpContext.Current.User.Identity.Name;

            user = user.Split('\\').Last();

            Session["UserID"] = user;
            loggedinusereml = user + "@radiantdelivers.com";

            // we turend off this code as I am authorizing through [Authorize]
            //check admin user access, if not admin redirect to error page


            //bool isadmin = false;

            //isadmin = emp.AuthenticateUser("admin", loggedinusereml);

            //if (!isadmin)
            //{

            //    Session["ErrorMsg"] = "SOP Application: Only admin user has access to this page. You are not a admin user. Please contact SOP team for access!";
            //    return Json(new { redirecturl = "/Home/RedirectForErr" }, JsonRequestBehavior.AllowGet);

            //}

            //emp = null;

            //2. if the doc file is new then copy sop template with new file name
            //or if user uploads exsiting doc file with new template is uplaoded, then copy the uploaded
            //file into project folder

            //log DateTime:sop.SOPNO: start collecting user email


            Employee emp = new Employee();

            Employee[] rvwrItems = JsonConvert.DeserializeObject<Employee[]>(sop.FilereviewersArr[0]);

            Employee[] vwrItems;

            SOPClass oSop = new SOPClass();

            Employee oEmp = new Employee();

            oEmp.useremailaddress = loggedinusereml;

            oEmp.GetUserByEmail();

            oSop.FileChangeRqsterID = oEmp.userid;

            oSop.DocumentLibName = "SOP";
            oSop.SOPNo = sop.SOPNo;

            //log DateTime:sop.SOPNO: end collecting user email

            ViewBag.Message = "Upload SOP File";

            //   bool fileloaded = false;

            //log DateTime:sop.SOPNO: start saving new or updloaded to temp project folder

            string docpath = Server.MapPath("~/Content/DocFiles/");
            
            if (sop.FileName!=null && sop.FileName.Trim() != "")
            {

                //for new file copy from template to temp file
                System.IO.File.Copy(Server.MapPath("~/Content/docfiles/SOPTemplate.docx"), Server.MapPath("~/Content/docfiles/SOPTemp.docx"), true);

                oSop.FileName = sop.SOPNo + " " + sop.FileName.Trim() + ".docx";

                oSop.FileTitle = sop.FileName.Trim();

                bProcessCompleted = true;
            }


            else if (sop.UploadedFile != null)

            {

                //for uploaded file copy it from posted file to temp file

                oSop.FileName = Path.GetFileName(sop.UploadedFile.FileName);

                oSop.FileTitle = Path.ChangeExtension(oSop.FileName, null);

                oSop.FileName = oSop.SOPNo + " " + oSop.FileName;

                if (!Directory.Exists(docpath))

                {

                    Directory.CreateDirectory(docpath);

                }

                sop.UploadedFile.SaveAs(docpath + "SOPTemp.docx");

                bProcessCompleted = true;

 
            }

            //end step 2 saving new file or uploaded file in projetc folder


            //3. Update cover sheet and revision history with file name, SOP No, reviewers, owner etc.
            if (bProcessCompleted)
            {
                //DateTime:sop.SOPNo:file was saved in project temp folder successfully

                //DateTime:sop.SOPNo:start updating covert sheet and rev history successfully

                short supdfreq = Convert.ToInt16(sop.Updatefreq);

                oSop.FileApproverEmail = sop.FileApproverEmail;
                oSop.FileOwnerEmail =sop.FileOwnerEmail;
                oSop.Reviewers = rvwrItems; 
                oSop.Updatefreq = supdfreq;
                oSop.Updatefrequnit = sop.Updatefrequnit;
                oSop.Updfrequnitcode = sop.Updfrequnitcode;
                oSop.SOPEffectiveDate = Convert.ToDateTime(sop.SOPEffectiveDate);


                oSop.FilePath = docpath + oSop.FileName;

                FileRevision[] oRevarr = new FileRevision[1];

                FileRevision rev1 = new FileRevision();

                rev1.RevisionNo = "1.0";
                rev1.RevisionDate = DateTime.Now;
                rev1.Description = "New SOP";

                oRevarr[0] = rev1;

                //FileRevision rev2 = new FileRevision();

                //rev2.RevisionNo = "2.0";
                //rev2.RevisionDate = DateTime.Now;
                //rev2.Description = "Newly Created";

                //oRevarr[1] = rev2;

                oSop.FileRevisions = oRevarr;

                oSop.SiteUrl = siteurl;
                oSop.FileCurrVersion = "1.0";

                oSop.UpdateCoverRevhistPage();

                bProcessCompleted = true;

            }
            else    

            {

                bProcessCompleted = false;
                //log it
                //DateTime:sop.SOPNo:failed to save new or uploaded file in project temp folder
            }

            //end step 3 updating cover sheet


            //4. Upload the updated file into sharepoint online SOP doc libray in correct department folder
            //and sub folder enetred by user

            if (bProcessCompleted)
            {
                //log it
                //DateTime:sop.SOPNo:that successfully updated coversheet and rev history

                //DateTime:sop.SOPNo:start uploading file in sharepoint online SOP doc library

                Thread.Sleep(6000);

                oSop.FolderName =sop.FolderName;
                oSop.SubFolderName =sop.SubFolderName;

                if (oSop.SubFolderName == "")
                    oSop.FileUrl = "SOP/" + oSop.FolderName + "/";
                else
                    oSop.FileUrl = "SOP/" + oSop.FolderName + "/" + oSop.SubFolderName + "/";

                oSop.FileStream = System.IO.File.ReadAllBytes(oSop.FilePath);

                oSop.UploadDocument();

                bProcessCompleted = true;

            }

            else
            {

                //log this failure
                //DateTime:sop.SOPNo:that failed updating coversheet and rev history

                bProcessCompleted = false;

            }

            // end step 4 uploading file into sharepoint online sop doc lib

            //5. Update SQL server tables with all info like, change request, reviewers, approvers etc.

            if (bProcessCompleted)
            {

                //log 
                //DateTime:sop.SOPNo:successfully uploaded file in SharePoint online SOP doc lib

                //DateTime:sop.SOPNo:start updating SQL tables
                oSop.FileID = oSop.FileID;

                oSop.AddChangeRequest();
                oSop.AddFileReviewers();
                oSop.AddFileApprover();
                oSop.AddFileOwner();
                oSop.AddUpdateFreq();

                bProcessCompleted = true;
            }

            else
            {
                //log this failure
                //DateTime:sop.SOPNo:failed uploading SOP in sharepoint online SOP doc lib


                bProcessCompleted = false;
            }
            //end step 5 SQL table upate


            //6. assign proper permission for owner with full permission, reviewers with contribute permission, 
            //and read permission to users according to admin users entry of users for read permission i.e read 
            //permission to all, or a departement or custom users.

            if (bProcessCompleted)
            {
                //log 
                //DateTime:sop.SOPNo:successfully updated SQL tables
                //DateTime:sop.SOPNo:start assigning permission to SOP file in sharepoint 

                if (sop.AllUsersReadAcc)   //by default all users have read permission
                {

                    oSop.ViewAccessType = "All Users";

                    oSop.AddViewerAccessType();    // add new view type in SQL table 

                }

                else    //if All users are not permitted to view then customize the read permission according to either department or custom users

                {
                    //prepare viewers array


                    if (sop.DepartmentCode>0)  //if department is selected then preference is to get employees by department code
                    {

                        short sdeptcode = sop.DepartmentCode;
                        oEmp.departmentcode = sdeptcode;


                        oEmp.GetEmployeesByDeptCode();

                        vwrItems = oEmp.employees;

                        //first remove existing permission from the file, default is Watercooler Visitors

                        oSop.RemoveAllFilePermissions();

                        //give read permission to all users who are in the selected department

                        oSop.AssignFilePermission("add", "read", vwrItems);

                        //now add view access info by department in SQL Table
                        //we need this to retrieve and change in admin page

                        oSop.DepartmentCode = sdeptcode;
                        oSop.ViewAccessType = "By Department";
                        oSop.AddViewerAccessType();


                    }

                    else if (sop.FileviewersArr.Count() > 0)   //get employees from custom user list
                    {
                        vwrItems = JsonConvert.DeserializeObject<Employee[]>(sop.FileviewersArr[0]);

                        //first remove existing permission from the file, default is Watercooler Visitors

                        oSop.RemoveAllFilePermissions();

                        //give read permission to all custom viewers

                        oSop.AssignFilePermission("add", "read", vwrItems);

                        //now add view access info by custom users in SQL table
                        //we need this to retrieve and change in admin page

                        oSop.Viewers = vwrItems;
                        oSop.ViewAccessType = "By Users";
                        oSop.AddViewerAccessType();
                        oSop.AddFileViewers();



                    }

                }


                //give contribute permission to all reviewers

                oSop.AssignFilePermission("add", "contribute", rvwrItems);


                //give edit permission to approver


                oSop.AssignFilePermission("add", "edit", sop.FileApproverEmail);

                //give full permission to owner

                oSop.AssignFilePermission("add", "full control", sop.FileOwnerEmail);


                bProcessCompleted = true;

            }

            else
            {

                bProcessCompleted = false;
                //log this failure
                //DateTime:sop.SOPNo:failed updating SQL table

            }

            //end step 6 assiging permission to file in sharepoint

            //7. once all above steps are successfully completed, then send success jason message to MVC view html

            //8. if it fails at any step and cannot reach upto step 6 then pinpoint the error and the step 
            //that caused the error and log it in the background, then do reverse engineering to roll back 
            //any changes if possible. If roll back is not possible then send error message to user 
            //with json message to the mvc view advising him/her to contact IT with keeping screen shot of the 
            //error message. You must log the error in the backgroud so you can trace the error to rollback 
            //or complete it manually.

            //if error can be rolled back then send failed jason message to MVC view
            //with failed/error message and advise them to try again or contact IT to resolve this if error happend
            //again


            //server sends OK 200 response to the client

            if (bProcessCompleted)  
            {


                //log 
                //DateTime:sop.SOPNo:successfully assigend permission and completed all SOP processing

                //  Send "Success" to ajax call back in view

                Session["SOPMsg"] = "The SOP " + sop.FileName + " has been successfully processed!";

                return Json(new { success = true, responseText = "The SOP " + sop.FileName + " has been successfully processed!" }, JsonRequestBehavior.AllowGet);

            }
            else
            {
                //log 
                //DateTime:sop.SOPNo:failed assigend permission and SOP processing

                //Send failed

                Session["SOPMsg"] = "Failed processing SOP " + sop.FileName + " , please contact IT!";

                // sop.ErrorMessage = "Failed to process SOP with error:" + sop.ErrorMessage + ". Please try again or Contact IT";
                return Json(new { success = false, responseText = "Failed processing SOP " + sop.FileName + " , please contact IT!" }, JsonRequestBehavior.AllowGet);
            }

            //if any server failutre to send requested response other than OK 200 code then ajax will raise error event


            // return View();
            //return Json(sop);

        }


        //https://www.entityframeworktutorial.net/Querying-with-EDM.aspx

        public List<SOPClass> GetFolders()
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
                    FileStatuscode=x.filestatuscode

                }).Where(s => s.FilePath == "SOP/" && s.FileStatuscode==3);


                folderlist = folders.ToList();


            }


            return folderlist;


        }

        [HttpGet]
        public JsonResult CheckIfExists(string FileName)
        {
            bool isExist = false;
            if (FileName.Equals("abc@gmail.com"))
            {
                isExist = true;
            }
            return Json(isExist, JsonRequestBehavior.AllowGet);
        }

        //public JsonResult CheckIfInt(string Updatefreq)
        //{
        //    bool isExist = false;
        //    if (IsNumeric(Updatefreq))
        //    {
        //        isExist = true;
        //    }
        //    return Json(isExist, JsonRequestBehavior.AllowGet);
        //}



        public ActionResult GetSubFolderList(string foldername)
        {
            List<SOPClass> subfolderlist;

            using (RadiantSOPEntities ctx = new RadiantSOPEntities())

            {
                var subfolders = ctx.deptsopfiles.Select(x => new SOPClass()
                {
                    FileID = x.FileID,
                    FileName = x.DeptFileName,
                    FilePath = x.SPFilePath,
                    FileLink = x.SPFileLink,
                    SOPNo = x.SOPNo,
                    FileStatuscode=x.filestatuscode
                }).Where(s => s.FilePath == "SOP/" + foldername + "/" && !s.FileName.Contains(".docx") && s.FileStatuscode==3);


                subfolderlist = subfolders.ToList();

                ViewBag.ddlSubFolders = new SelectList(subfolderlist, "FileID", "FileName");

            }

            return PartialView("DisplaySubfolders");


        }

        public JsonResult GetSOPNO(string foldername, string subfoldername)
        {

            string lsopno = "";
            SOPClass oSOP = new SOPClass();
            oSOP.FolderName = foldername;
            oSOP.SubFolderName = subfoldername;
            oSOP.GetSOPNo();
            lsopno = oSOP.SOPNo;

            if (lsopno != "")
                return Json(new { success = true, sopno = lsopno });
            else
                return Json(new { success = false });

        }


        //public void UploadFile(HttpPostedFileBase postedFile, string deptfoldername, string subfoldername, string sopno)


        public bool IsNumeric(string value)
        {
            return value.All(char.IsNumber);
        }

        //JsonResult

        [HttpPost]
        public JsonResult UploadCreateFile(HttpPostedFileBase postedFile, string newfilename, string[] reviewers, string[] viewers, string sopno, 
            string approver, string owner, string allvwrs,string vwrdptcode, string deptfoldername,
            string deptsubfoldername, string sopeffdate,string sopupdfreq,string sopupdfrequnit)

        {


            //validate data first

            string user="";
            user= System.Web.HttpContext.Current.User.Identity.Name;

            user= user.Split('\\').Last();

            Session["UserID"] = user;

           // var usr = System.Environment.UserName;


            if (deptsubfoldername== "--Select Subfolder--")
            {

                deptsubfoldername = "";

            }


            Employee[] rvwrItems = JsonConvert.DeserializeObject<Employee[]>(reviewers[0]);

            Employee[] vwrItems;

            //documentlistname = "SOP";

            SOPClass oSop = new SOPClass();

            Employee oEmp = new Employee();

            if (user != "")
                oEmp.useremailaddress = user.Trim() + "@radiantdelivers.com";

            oEmp.GetUserByEmail();

            oSop.FileChangeRqsterID = oEmp.userid;

            oSop.DocumentLibName = "SOP";
            oSop.SOPNo = sopno;


            ViewBag.Message = "Upload SOP File";

            bool fileloaded=false;


            //1. load file first

            string docpath = Server.MapPath("~/Content/DocFiles/");

            if (newfilename.Trim()!="")
            {

                //for new file copy from template to temp file
                System.IO.File.Copy(Server.MapPath("~/Content/docfiles/SOPTemplate.docx"), Server.MapPath("~/Content/docfiles/SOPTemp.docx"), true);

                oSop.FileName = sopno+" "+newfilename + ".docx";

                oSop.FileTitle = newfilename ;

                fileloaded = true;
            }


            else if (postedFile != null)

            {

                //for uploaded file copy it from posted file to temp file

                oSop.FileName = Path.GetFileName(postedFile.FileName);

                oSop.FileTitle = Path.ChangeExtension(oSop.FileName, null);

                oSop.FileName= oSop.SOPNo+" "+oSop.FileName;

                if (!Directory.Exists(docpath))

                {

                    Directory.CreateDirectory(docpath);

                }

                // postedFile.SaveAs(path + Path.GetFileName(postedFile.FileName));

                postedFile.SaveAs(docpath + "SOPTemp.docx");

                fileloaded = true;

                ViewBag.Message = "File uploaded successfully.";

            }

            


            if (fileloaded == true)  //file saved locally
            {

                //2. update coversheet and revision history 

                //update top sheet and revision history first

                short supdfreq = Convert.ToInt16(sopupdfreq);

                oSop.FileApproverEmail = approver;
                oSop.FileOwnerEmail = owner;
                oSop.Reviewers = rvwrItems;
                oSop.Updatefreq = supdfreq;
                oSop.Updatefrequnit = sopupdfrequnit;
                oSop.SOPEffectiveDate = Convert.ToDateTime(sopeffdate);

            
                oSop.FilePath = docpath + oSop.FileName;

                FileRevision[] oRevarr= new FileRevision[1];

                FileRevision rev1 = new FileRevision();

                rev1.RevisionNo = "1.0";
                rev1.RevisionDate = DateTime.Now;
                rev1.Description = "New SOP";

                oRevarr[0] = rev1;

                //FileRevision rev2 = new FileRevision();

                //rev2.RevisionNo = "2.0";
                //rev2.RevisionDate = DateTime.Now;
                //rev2.Description = "Newly Created";

                //oRevarr[1] = rev2;

                oSop.FileRevisions = oRevarr;

                oSop.SiteUrl = siteurl;
                oSop.FileCurrVersion = "1.0";

                oSop.UpdateCoverRevhistPage();


                //3. upload the processed doc file to sharepoint online in SOP doc library

                oSop.FolderName = deptfoldername;
                oSop.SubFolderName = deptsubfoldername;

                if  (oSop.SubFolderName=="")
                    oSop.FileUrl = "SOP/" + oSop.FolderName + "/" ;
                else
                    oSop.FileUrl = "SOP/" + oSop.FolderName + "/" + oSop.SubFolderName + "/";

                oSop.FileStream= System.IO.File.ReadAllBytes(oSop.FilePath);

                oSop.UploadDocument();

                //4. update SQL Data table with reviewers, approver and owner

                oSop.FileID = oSop.FileID;

                oSop.AddChangeRequest();
                oSop.AddFileReviewers();
                oSop.AddFileApprover();
                oSop.AddFileOwner();
                oSop.AddUpdateFreq();
           


                //5. assign permissions to the uploaded SOP file in SP


                if (allvwrs.ToUpper() == "TRUE")   //by default all users have read permission
                {
 
                    oSop.ViewAccessType = "All Users";

                    oSop.AddViewerAccessType();    // add new view type in SQL table 
                    
                }

               if (allvwrs.ToUpper() == "FALSE")   //if All users are not permitted to view then customize the read permission according to either department or custom users

                {
                    //prepare viewers array


                    if (vwrdptcode != "")  //if department is selected then preference is to get employees by department code
                    {

                        short sdeptcode = Convert.ToInt16(vwrdptcode);
                        oEmp.departmentcode = sdeptcode;


                        oEmp.GetEmployeesByDeptCode();

                        vwrItems = oEmp.employees;

                       //first remove existing permission from the file, default is Watercooler Visitors

                        oSop.RemoveAllFilePermissions();

                        //give read permission to all users who are in the selected department

                        oSop.AssignFilePermission("add", "read", vwrItems);

                        //now add view access info by department in SQL Table
                        //we need this to retrieve and change in admin page

                        oSop.DepartmentCode = Convert.ToInt16(vwrdptcode);
                        oSop.ViewAccessType = "By Department";
                        oSop.AddViewerAccessType();


                    }

                    else if (viewers.Count() > 0)   //get employees from custom user list
                    {
                        vwrItems = JsonConvert.DeserializeObject<Employee[]>(viewers[0]);

                        //first remove existing permission from the file, default is Watercooler Visitors

                        oSop.RemoveAllFilePermissions();

                        //give read permission to all custom viewers

                        oSop.AssignFilePermission("add", "read", vwrItems);

                        //now add view access info by custom users in SQL table
                        //we need this to retrieve and change in admin page

                        oSop.Viewers = vwrItems;
                        oSop.ViewAccessType = "By Users";
                        oSop.AddViewerAccessType();
                        oSop.AddFileViewers();
                        


                    }



                }


                //give contribute permission to all reviewers

 
                oSop.AssignFilePermission("add", "contribute", rvwrItems);


                //give edit permission to approver

 
                oSop.AssignFilePermission("add", "edit", approver);

                //give full permission to owner

                oSop.AssignFilePermission("add", "full control", owner);

                

            }






            return Json(fileloaded);
            

        }




        //private ArrayList getFileVersions(string siteurl,string filerelpath)
        //{

        //    ArrayList fversions = new ArrayList();

        //    //SOP / Warehouse Operations /


        //    using (ClientContext clientContext = new ClientContext(siteurl))
        //    {

        //        string userName = "tshaikh@radiantdelivers.com";
        //        string password = "bagerhat79&";


        //        SecureString SecurePassword = GetSecureString(password);
        //        clientContext.Credentials = new SharePointOnlineCredentials(userName, SecurePassword);

        //        Web site = clientContext.Web;
        //        clientContext.Load(site);
        //        // File file = site.GetFileByServerRelativeUrl("/Shared Documents/mydocument.doc");


        //        //FileVersionCollection versions;
        //        Microsoft.SharePoint.Client.File file = site.GetFileByServerRelativeUrl(filerelpath);

        //        clientContext.Load(file);

        //        clientContext.ExecuteQuery();


        //        string id;

        //        FileVersionCollection versions = file.Versions;

        //        clientContext.Load(versions);

        //        PropertyValues fi = file.Properties;
                
        //        clientContext.Load(fi);

            
        //        clientContext.ExecuteQuery();

        //        string lv = file.MajorVersion.ToString();


        //        id = fi["ID"].ToString();


         

        //        if (versions != null)
        //        {
        //            foreach (FileVersion version in versions)
        //            {
        //                Console.WriteLine("Version : {0}", version.VersionLabel);

        //                clientContext.Load(version);
        //                clientContext.ExecuteQuery();


        //                if ((Convert.ToDouble(version.VersionLabel) % 1) == 0)
        //                {
        //                    //You can get all major versions here.

                            
        //                    fversions.Add(version.VersionLabel);

        //                }


        //            }
        //        }


        //    }


           

        //    return fversions;
        //}


        //private void GetFileVersions(string siteURL, string documentListName, string documentListURL, string documentName)
        //{

        //    ClientContext clientContext = new ClientContext(siteURL);

        //    string userName = "tshaikh@radiantdelivers.com";
        //    string password = "bagerhat79&";

        //    SecureString SecurePassword = GetSecureString(password);
        //    clientContext.Credentials = new SharePointOnlineCredentials(userName, SecurePassword);


        //    Web web = clientContext.Web;
        //    clientContext.Load(web);
        //    clientContext.Load(web.Lists);
        //    clientContext.Load(web, wb => wb.ServerRelativeUrl);
        //    clientContext.ExecuteQuery();

        //    Microsoft.SharePoint.Client.List list = web.Lists.GetByTitle(documentListName);
        //    clientContext.Load(list);
        //    clientContext.ExecuteQuery();

        //    Folder folder = web.GetFolderByServerRelativeUrl(web.ServerRelativeUrl + documentListURL);
        //    clientContext.Load(folder);
        //    clientContext.ExecuteQuery();

        //    CamlQuery camlQuery = new CamlQuery();

  
        //    //TO GET ONLY FILE ITEM
        //    camlQuery.ViewXml = "<View Scope='Recursive'> " +
        //                           "  <Query> " +

        //                          " + <Where> " +
        //                               "  <Contains>" +
        //                                    " <FieldRef Name='FileLeafRef'/> " +
        //                                        " <Value Type='File'>" + documentName + "</Value>" +
        //                                   " </Contains> " +
        //                               " </Where> " +

        //                            " </Query> " +
        //                        " </View>";

  
        //    camlQuery.FolderServerRelativeUrl = folder.ServerRelativeUrl;
        //    ListItemCollection listItems = list.GetItems(camlQuery);
        //    clientContext.Load(listItems);
        //    clientContext.ExecuteQuery();


        //    string fid;

        //    foreach (ListItem item in listItems)
        //    {
        //        //item.FileSystemObjectType;

        //        if (item.FileSystemObjectType == FileSystemObjectType.File)
        //        {
        //            // This is the File

        //            Microsoft.SharePoint.Client.File file = item.File;

        //            FileVersionCollection versions = file.Versions;

        //            fid=file.Properties["ID"].ToString();

        //            clientContext.Load(file);
        //            clientContext.Load(versions);
        //            clientContext.ExecuteQuery();


        //            //$file = $item.File
        //            //versions = $file.Versions
        //            //$ctx.Load($file)
        //            //$ctx.Load($versions)
        //            //$ctx.ExecuteQuery()


        //            foreach(FileVersion v in versions)
        //            {

        //                clientContext.Load(v);
        //                clientContext.ExecuteQuery();

        //                User modifiedBy = v.CreatedBy;
        //                clientContext.Load(modifiedBy);

        //                clientContext.ExecuteQuery();

        //                string loginnm =modifiedBy.LoginName;
        //                string title = modifiedBy.Title;


        //            }




        //        }
        //        else if (item.FileSystemObjectType == FileSystemObjectType.Folder)
        //        {
        //            // This is a  Folder
        //        }




        //    }



        //}

                
        public class vwDepartmentFolders
        {
        }
    }  //end of class HomeController

    public class RoleFilterAttribute : ActionFilterAttribute
    {
        public string Role { get; set; }
        public override void OnActionExecuting(ActionExecutingContext ctx)
        {
            // Assume that we have user identity because Authorize is also
            // applied
            var user = ctx.HttpContext.User;
            if (!user.IsInRole(Role))
            {
                ctx.Result = new RedirectResult("url_needed_here");
            }
        }
    }

}