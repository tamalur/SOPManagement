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

        public ActionResult CreateFile()
        {
            ViewBag.Message = "Create File Page";

            return View();
        }


        public ActionResult UploadSOPFile()
        {
            ViewBag.Message = "Upload SOP File";

            RadiantSOPEntities ctx = new RadiantSOPEntities();

            ViewBag.ddlDeptFolders = new SelectList(GetFolders(), "FileName", "FileName");

            ViewBag.employees = (from c in ctx.users select new { c.useremailaddress, c.userfullname }).Distinct();

            ViewBag.departments = (from c in ctx.codesdepartments select new { c.departmentname, c.departmentcode }).Distinct();

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
                    SOPNo = x.SOPNo
                }).Where(s => s.FilePath == "SOP/");


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
                    SOPNo = x.SOPNo
                }).Where(s => s.FilePath == "SOP/" + foldername + "/" && !s.FileName.Contains(".docx"));


                subfolderlist = subfolders.ToList();

                ViewBag.ddlSubFolders = new SelectList(subfolderlist, "FileID", "FileName");

            }

            return PartialView("DisplaySubfolders");


        }

        public JsonResult GetSOPNO(string foldername, string subfoldername)
        {

            string lsopno = "";

            //RadiantSOPEntities ctx = new RadiantSOPEntities();

            //    //lsopno = foldername + "-001";

            // lsopno = ctx.GetLastSOPNO(foldername, subfoldername).FirstOrDefault().ToString();

            ////  lsopno = ctx.GetLastSOPNO(foldername, "").ToString();
            ///


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


        public JsonResult UploadCreateFile(HttpPostedFileBase postedFile, string newfilename, string[] reviewers, string[] viewers, string sopno, 
            string approver, string owner, string allvwrs,string vwrdptcode, string deptfoldername,
            string deptsubfoldername, string sopeffdate,string sopupdfreq,string sopupdfrequnit)

        {


            //validate data first

            if (deptsubfoldername== "Please select a subfolder")
            {

                deptsubfoldername = "";

            }


            Employee[] rvwrItems = JsonConvert.DeserializeObject<Employee[]>(reviewers[0]);

            Employee[] vwrItems;

            //documentlistname = "SOP";

            SOPClass oSop = new SOPClass();

            Employee oEmp = new Employee();

            oSop.DocumentLibName = "SOP";
            oSop.SOPNo = sopno;


            ViewBag.Message = "Upload SOP File";

            bool fileloaded=false;


            //load file first

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

            

            //if file is uploaded then update top sheet and revision history in newly upaloaded file
            string filepath="";
            string sopfilename;
            string documentlistname;
            int newfileid;



            if (fileloaded == true)
            {

                //once SOP top sheet and revision history is updated then generate name and upload the file with that name



                //sopfilename = Path.GetFileName(postedFile.FileName);

                // UpdateCoverRevhistPage(sopfilename, sopno, rvwrItems, owner, approver, sopeffdate);

                short supdfreq = Convert.ToInt16(sopupdfreq);


                oSop.FileApproverEmail = approver;
                oSop.FileOwnerEmail = owner;
                oSop.Reviewers = rvwrItems;
                oSop.Updatefreq = supdfreq;
                oSop.Updatefrequnit = sopupdfrequnit;
                oSop.SOPEffectiveDate = Convert.ToDateTime(sopeffdate);

            
                oSop.FilePath = docpath + oSop.FileName;

                FileRevision[] oRevarr= new FileRevision[2];

                FileRevision rev1 = new FileRevision();

                rev1.RevisionNo = "1.0";
                rev1.RevisionDate = DateTime.Now;
                rev1.Description = "Newly Created";

                oRevarr[0] = rev1;

                FileRevision rev2 = new FileRevision();

                rev2.RevisionNo = "2.0";
                rev2.RevisionDate = DateTime.Now;
                rev2.Description = "Newly Created";

                oRevarr[1] = rev2;

                oSop.FileRevisions = oRevarr;

                oSop.SiteUrl = siteurl;
                oSop.FileCurrVersion = "2.0";

                //versions[0]= oSop.FileVersions.

                oSop.UpdateCoverRevhistPage();


                //upload the processed doc file to sharepoint

               // filepath = path + sopno+" "+sopfilename;

                //string documentlistUrl = "SOP/" + deptfoldername + "/" + deptsubfoldername + "/";
                //string documentname = Path.ChangeExtension(sopfilename, null);   //"SOPFile";      // Title;

                oSop.FolderName = deptfoldername;
                oSop.SubFolderName = deptsubfoldername;

                oSop.FileUrl = "SOP/" + oSop.FolderName + "/" + oSop.SubFolderName + "/";

                // string filerpath = "/sites/watercooler/SOP/Quality Assurance & Regulatory Affairs (QRA)/QRA  (AIB)/OPS07-01 Training and Personnel.docx";

                
                //byte[] stream = System.IO.File.ReadAllBytes(filepath);

                oSop.FileStream= System.IO.File.ReadAllBytes(oSop.FilePath);

                oSop.UploadDocument();

              //  newfileid =UploadDocument(siteurl, documentlistname, documentlistUrl, documentname, stream, sopno);



                //update SQL Data table with reviewers, approver and owner

                oSop.FileID = oSop.FileID;

                oSop.AddFileReviewers();
                oSop.AddFileApprover();
                oSop.AddFileOwner();
                oSop.AddUpdateFreq();


                //assign permission

        

                if (allvwrs.ToUpper() == "TRUE")
                {
                    //assgnpermcomplete = true;

                    oSop.ViewAccessType = "All Users";

                    oSop.AddViewerAccessType();
                    
                }

                    if (allvwrs.ToUpper() == "FALSE")   //if All users are not permitted to view then customize the read permission according to either department or custom users

                {
                    //prepare viewers array


                    short sdeptcode = Convert.ToInt16(vwrdptcode);

                    oEmp.departmentcode = sdeptcode;


                    if (vwrdptcode != "")  //if department is selected then preference is to get employees by department code
                    {
                        //vwrItems=GetEmployeesByDeptCode(sdeptcode);

                       oEmp.GetEmployeesByDeptCode();

                        vwrItems = oEmp.employees;

                        //  vwrItems

                        //first remove existing permission from the file, default is Watercooler Visitors

                        oSop.RemoveAllFilePermissions();

                        //give read permission to all custom viewers

                        oSop.AssignFilePermission("add", "read", vwrItems);

                        //now add view access type "by department in SQL Table

                        oSop.DepartmentCode = Convert.ToInt16(vwrdptcode);
                        oSop.ViewAccessType = "By Department";
                        oSop.AddViewerAccessType();


                    }

                    else if (viewers.Count() > 0)   //get employees from added list
                    {
                        vwrItems = JsonConvert.DeserializeObject<Employee[]>(viewers[0]);

                        //first remove existing permission from the file, default is Watercooler Visitors

                        oSop.RemoveAllFilePermissions();

                        //give read permission to all custom viewers


                        oSop.AssignFilePermission("add", "read", vwrItems);

                        //now add view access type "by custom users in SQL table

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