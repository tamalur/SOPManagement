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

namespace SOPManagement.Controllers
{
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

        public ActionResult Sessionouterr()
        {
            if (Session["UserID"] == null)
                ViewBag.SessionOutMsg = "Session timed out. Please enter data again!";

            return View();

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

        public ActionResult CreateFile()
        {
            ViewBag.Message = "Create File Page";

            return View();
        }


        public ActionResult UploadSOPFile()
        {
            ViewBag.Message = "Upload SOP File";



            ViewBag.ddlDeptFolders = new SelectList(GetFolders(), "FileName", "FileName");

            ViewBag.employees = (from c in ctx.users select new { c.useremailaddress, c.userfullname, c.userstatuscode }).Where(x => x.userstatuscode == 1).Distinct();

            ViewBag.departments = (from c in ctx.codesdepartments select new { c.departmentname, c.departmentcode }).Distinct();

            return View();
        }


        public ActionResult CreateUploadSOP()
        {
            ViewBag.Message = "Upload SOP File";

 

            ViewBag.ddlDeptFolders = new SelectList(GetFolders(), "FileName", "FileName");

            ViewBag.employees = (from c in ctx.users select new { c.useremailaddress, c.userfullname, c.userstatuscode}).Where(x=>x.userstatuscode==1).Distinct();

            ViewBag.departments = (from c in ctx.codesdepartments select new { c.departmentname, c.departmentcode }).Distinct();

            return View();
        }

        [HttpPost]
        public ActionResult CreateUploadSOP(SOPClass sop)
        {


            return View();

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


   
}