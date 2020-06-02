using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.Mvc;
using Newtonsoft.Json;
using SOPManagement.Models;

//this controller was developed by Tamalur from April 10 to April 22, 2020
namespace SOPManagement.Controllers
{
   
   
    //[Authorize]
    public class HomeController : BaseController
    {

        string siteurl = Utility.GetSiteUrl();

        RadiantSOPEntities ctx = new RadiantSOPEntities();

        
        public ActionResult Index()
        {

            if (Utility.IsSessionExpired())
            {

                return RedirectToAction("LogIn");


            }


            return View();
        }
        public ActionResult LogIn()
        {

            return View();
        }

        public ActionResult About()
        {

            if (Utility.IsSessionExpired())
                return RedirectToAction("LogIn");

            ViewBag.Message = "Your application description page.";

            return View();
        }

       
        public ActionResult Contact()
        {

            if (Utility.IsSessionExpired())
            {

                return RedirectToAction("LogIn");


            }

            ViewBag.Message = "Your contact page.";

            return View();
        }


        public ActionResult SOPNoAcessMsg(int? id)
        {
            SOPClass oSOP = new SOPClass();
            oSOP.FileID = Convert.ToInt32(id);

            oSOP.GetSOPInfoByFileID();

            ViewBag.SOPTitle = oSOP.FileTitle;
            ViewBag.SOPNO = oSOP.SOPNo;

            oSOP = null;


            return View();
        }


        public ActionResult CleintServerErr()
        {
            if (Session["UserFullName"] == null)
                return RedirectToAction("LogIn");

            return View();
        }


        public ActionResult LogOff()
        {
            Session["UserFullName"] = null; //it's my session variable
            Session.Clear();
            Session.Abandon();
        //    FormsAuthentication.SignOut(); //you write this when you use FormsAuthentication
            return RedirectToAction("LogIn", "Home");
        }

        public ActionResult DataMaintHRIT()
        {


            return View();
        }

        public ActionResult CreateUploadSOPAuth()
        {


            if (Utility.IsSessionExpired())
            {

                //ViewBag.ErrorMessage = "SOP Application: Session not Timed out";
                //Session["SOPMsg"] = "SOP Application: Session not Timed out";
                return RedirectToAction("LogIn");


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



      //  [Authorize(Roles = "SOPADMIN")]
        public ActionResult AdminSOP(int? id)
        {

            // http://localhost:58639/Home/AdminSOP/298

            TempData["PageTitle"] = "SOP Admin";

            TempData["FileID"] = Convert.ToInt32(id);  //very important field in this context

             SOPClass oSOP = new SOPClass();

             oSOP.FileID = Convert.ToInt32(id); 


            //currently only owner of the SOP can do admin changes

            oSOP.GetSOPInfoByFileID();
            TempData["SOPName"] = oSOP.FileTitle;

            if (!oSOP.AuthenticateUser("admin"))   //only owner can do admin changes

            {
                oSOP = null;
                TempData["SOPMsg"] = "Error:You are not authenticated to do admin changes of the document in question.";
         

                return RedirectToAction("SOPMessage");
            }



            int lastchangerequestid = oSOP.GetLastChngRequestID();

            oSOP.FileChangeRqstID = lastchangerequestid;

            int lastchngreqstatcode = oSOP.GetChngReqSOPStatusCode(); //by file id

            //2=not signed, 1=signed, 3=approved: not allowed to change at this moment since sop is under approval process

            TempData["SOPName"] = oSOP.FileTitle;


            ViewBag.employees = (from c in ctx.vwUsers select new { c.useremailaddress, c.userfullname }).Distinct();

            //(from c in ctx.users select new { c.useremailaddress, c.userfullname, c.userstatuscode }).Where(x => x.userstatuscode == 1).Distinct();

            ViewBag.departments = (from c in ctx.codesSOPDepartments select new {c.sopdeptcode, c.sopdeptname }).Distinct();

            ViewBag.updfrequnits = (from c in ctx.codesUnits select new { c.unitcode, c.Unitname, c.UnitType }).Where(x => x.UnitType == "UpdateFrequency").Distinct();



            return View(oSOP);
        }


        [HttpPost]

        //public ActionResult AdminSOP(SOPClass sop, string archive)
        //if (!string.IsNullOrEmpty(archive))

        public ActionResult AdminSOP(SOPClass sop)

        {


            // do not retrieve data from database with sop object as this obejcts holds all new data entered by user
            //I used helper class

            if (TempData["FileID"] != null)
                sop.FileID = Convert.ToInt32(TempData["FileID"].ToString());

            else
            {

                TempData["SOPMsg"] = "Error:Invalid File ID.";

                return RedirectToAction("SOPMessage");

            }


            Employee[] rvwrItems= JsonConvert.DeserializeObject<Employee[]>(sop.FilereviewersArr[0]);
            Employee[] vwrItems= JsonConvert.DeserializeObject<Employee[]>(sop.FileviewersArr[0]);

            sop.FileReviewers = rvwrItems;

            sop.FileViewers = vwrItems;


            sop.SiteUrl = siteurl;
            sop.DocumentLibName = "SOP";


            SOPClass oSOP = new SOPClass();
            Employee oEmp = new Employee();

            bool admindatachanged;
            bool contributorchngd;
            bool rvwrfound ;
            bool vwrfound ;
            bool updatevwr;


            short lastchngreqstatcode = 0;
            int lastchangerequestid = 0;

            oSOP.FileID = sop.FileID;
            oSOP.SiteUrl = siteurl;
            oSOP.DocumentLibName = sop.DocumentLibName;

            lastchangerequestid = oSOP.GetLastChngRequestID();

            oSOP.FileChangeRqstID = lastchangerequestid;

            lastchngreqstatcode = oSOP.GetChngReqSOPStatusCode(); //by file id


            oSOP.GetSOPInfo();   //get current data with helper class to compare with new data from view, if any change then update it

            TempData["SOPName"] = oSOP.FileTitle.Trim();

            if (oSOP.ViewAccessType.Trim() == "All Users")
                oSOP.AllUsersReadAcc = true;


            sop.FileTitle = oSOP.FileTitle.Trim();
            sop.FileName = oSOP.FileName;
            sop.FilePath = oSOP.FilePath;
            sop.FileUrl = oSOP.FileUrl;

           
            if (sop.AllUsersReadAcc)
                sop.ViewAccessType = "All Users";
            else if (sop.DepartmentCode != null && sop.DepartmentCode > 0)
                sop.ViewAccessType = "By Department";
            else if (sop.FileviewersArr != null && sop.FileviewersArr.Count() > 0)
                sop.ViewAccessType = "By Users";



            try
            {



                // if admin change is made and archive SOP was selected then redirect error page. 

                //first check if there is any admin changes

                admindatachanged = false;
                contributorchngd = false;


                if ((sop.Updatefreq != oSOP.Updatefreq) || (sop.Updfrequnitcode != oSOP.Updfrequnitcode))
                    admindatachanged = true;

                //check owner
                if (sop.FileOwnerEmail.Trim().ToLower() != oSOP.FileOwnerEmail.Trim().ToLower())    //owner was changed
                    contributorchngd = true;

                //check approver
                if (sop.FileApproverEmail.Trim().ToLower() != oSOP.FileApproverEmail.Trim().ToLower())  //approver was changed
                    contributorchngd = true;

                //check reviewers


                if ((sop.FileReviewers!=null & oSOP.FileReviewers!=null) &
                    (sop.FileReviewers.Count() == oSOP.FileReviewers.Count()))
                {

                    foreach (Employee newrvwr in sop.FileReviewers)
                    {
                        rvwrfound = false;
                        foreach (Employee oldrvwr in oSOP.FileReviewers)
                        {
                            if (newrvwr.useremailaddress.Trim().ToLower() == oldrvwr.useremailaddress.Trim().ToLower())
                            {
                                rvwrfound = true;
                                break;
                            }

                        }

                        if (rvwrfound == false)
                        {
                            contributorchngd = true;
                            break;
                        }

                    }

                }
                else
                {
                    contributorchngd = true;
                }


                if (contributorchngd)
                    admindatachanged = true;

                //check viewer access entry 

                if (oSOP.AllUsersReadAcc & !sop.AllUsersReadAcc)   //oSOP is previous an sop is new
                    admindatachanged = true;


                if (!oSOP.AllUsersReadAcc & sop.AllUsersReadAcc)
                    admindatachanged = true;


                updatevwr = false;

                if (!oSOP.AllUsersReadAcc & !sop.AllUsersReadAcc)
                {
                    //check department was changed or not
                    if ((sop.DepartmentCode != null) && oSOP.DepartmentCode != null)
                    {
                        //viewer by department will be treated as always change change becuase there could be  
                        if (sop.DepartmentCode != oSOP.DepartmentCode)
                        {
                            admindatachanged = true;
                            updatevwr = true;
                        }
                        else if (sop.DepartmentCode == oSOP.DepartmentCode)   //if same dept was selected now check if all employees are still same in this department 
                        {

                            Employee[] oldemps;
                            oEmp.departmentcode = Convert.ToInt16(oSOP.DepartmentCode);
                            oEmp.GetEmployeesByDeptCode();
                            oldemps = oEmp.employees;

                            Employee[] newemps;
                            oEmp.departmentcode = Convert.ToInt16(sop.DepartmentCode);
                            oEmp.GetEmployeesByDeptCode();
                            newemps = oEmp.employees;

                            if (newemps != null && oldemps != null)
                            {
                                foreach (Employee newvwr in newemps)
                                {
                                    vwrfound = false;
                                    foreach (Employee oldvwr in oldemps)
                                    {
                                        if (newvwr.useremailaddress.Trim().ToLower() == oldvwr.useremailaddress.Trim().ToLower())
                                        {
                                            vwrfound = true;
                                            break;
                                        }

                                    }

                                    if (vwrfound == false)
                                    {
                                        admindatachanged = true;
                                        updatevwr = true;
                                        break;
                                    }

                                }


                            }
                            else   //if somehow one or both of the old and new employee list is null
                            {
                                admindatachanged = true;
                                updatevwr = true;
                            }
                        }

                    }
                    else if (sop.FileViewers != null && oSOP.FileViewers != null) //check custom viewers were changed or not
                    {
                        //compare custom users
                        //  sop.FileViewers
                        //    oSOP.FileViewers


                        if ((sop.FileViewers != null & oSOP.FileViewers != null) &
                            (sop.FileViewers.Count() == oSOP.FileViewers.Count()))

                        {

                            foreach (Employee newvwr in sop.FileViewers)
                            {
                                vwrfound = false;
                                foreach (Employee oldvwr in oSOP.FileViewers)
                                {
                                    if (newvwr.useremailaddress.Trim().ToLower() == oldvwr.useremailaddress.Trim().ToLower())
                                    {
                                        vwrfound = true;
                                        break;
                                    }

                                }

                                if (vwrfound == false)
                                {
                                    admindatachanged = true;
                                    updatevwr = true;
                                    break;
                                }

                            }


                        }

                        else
                        {
                            admindatachanged = true;
                            updatevwr = true;


                        }

                    }
                    else   //any other viewer combination is a change, i.e. older was by dept now by users vice versa
                    {
                        admindatachanged = true;
                        updatevwr = true;
                    }


                }



                //if there is admin change and archive was selected then redirect to error page
                if (admindatachanged && sop.SOPArchived)

                {

                    TempData["SOPMsg"] = "Error:Failed to submit changes. Cannot make admin changes and archive document in the same submission.";

                    // return RedirectToAction("SOPMessage");

                    return Json(new { success = true, responseText = "Error:Failed to submit changes. Cannot make admin changes and archive document in the same submission." }, JsonRequestBehavior.AllowGet);

                }


                //archive document if there is admin changes and archive file was seleccted

                if (!admindatachanged && sop.SOPArchived)
                {
                    oSOP.ArchiveSOP();
                    oSOP.UpdateSOPFileStatus(4);  //archived

                    TempData["SOPMsg"] = "Success: This SOP has been successfully ARCHIVED.";

                    //return RedirectToAction("SOPMessage");

                    return Json(new { success = true, responseText = "Success: This SOP has been successfully ARCHIVED." }, JsonRequestBehavior.AllowGet);

                }


                //save all admin changes if there is admin change and archive was not selected 

                if (admindatachanged && !sop.SOPArchived)
                {

                    //update frequency of revision if any
                    if ((sop.Updatefreq != oSOP.Updatefreq) || (sop.Updfrequnitcode != oSOP.Updfrequnitcode))
                        sop.UpdateDataByFieldName("Updatefreq");



                    if (contributorchngd)  //create new change request and add new sign activities of current active signatories

                    {

                        sop.FileChangeRqsterID = Utility.GetLoggedInUserID();

                        sop.AddChangeRequest();   //it will get new chngreqid that will be used as follows

                        //add owner sign activities 

                        if (oSOP.FileOwner.signstatuscode == 1)  //signed 
                            oSOP.AssignFilePermissionToUsers("read", "remove", oSOP.FileOwnerEmail.Trim());  //remove last owner read permission 
                        else if (oSOP.FileOwner.signstatuscode == 2) //not signed
                            oSOP.AssignFilePermissionToUsers("contribute", "remove", oSOP.FileOwnerEmail.Trim());  //remove last owner read permission                         


                        if (sop.FileOwnerEmail.Trim().ToLower() != oSOP.FileOwnerEmail.Trim().ToLower())
                        {

                            sop.AssignFilePermissionToUsers("contribute", "add", sop.FileOwnerEmail.Trim());

                            sop.AddFileOwner();

                        }

                        else //owner is not changed but still need new sign actvities for other signatory change
                        {

                            oSOP.FileChangeRqstID = sop.FileChangeRqstID;

                            oSOP.AssignFilePermissionToUsers("contribute", "add", sop.FileOwnerEmail.Trim());

                            oSOP.AddOwneractivities(oSOP.GetOwnershipID());

                            // oSOP.AddFileOwner();


                        }

                        //now add new approver and activities

                        //step 2 approver

                        if (sop.FileOwnerEmail.Trim().ToLower() != oSOP.FileApproverEmail.Trim().ToLower())  //if same then you don't need to delete same person again 
                        {

                            if (oSOP.FileApprover.signstatuscode == 1)  //signed 
                                oSOP.AssignFilePermissionToUsers("read", "remove", oSOP.FileApproverEmail.Trim());  //remove last owner read permission 
                            else if (oSOP.FileApprover.signstatuscode == 2) //not signed
                                oSOP.AssignFilePermissionToUsers("contribute", "remove", oSOP.FileApproverEmail.Trim());  //remove last owner read permission                         

                        }

                        if (sop.FileApproverEmail.Trim().ToLower() != oSOP.FileApproverEmail.Trim().ToLower())  //approver was changed
                        {


                            sop.AssignFilePermissionToUsers("contribute", "add", sop.FileApproverEmail.Trim());

                            sop.AddFileApprover();


                        }

                        else  //create sign activities with old approver
                        {

                            oSOP.FileChangeRqstID = sop.FileChangeRqstID;

                            oSOP.AssignFilePermissionToUsers("contribute", "add", sop.FileApproverEmail.Trim());

                            // oSOP.AddFileOwner();
                            oSOP.AddApproveractivities(oSOP.GetApproveID());


                        }

                        //add new reviewer and actvities

                        //remove all old reviewers' permission and add new reviewers' read permission

                        foreach (Employee oldrvwr in oSOP.FileReviewers)
                        {
                            if (oldrvwr.useremailaddress.Trim().ToLower() != sop.FileOwnerEmail.Trim().ToLower())
                            {
                                if (oldrvwr.useremailaddress.Trim().ToLower() != sop.FileApproverEmail.Trim().ToLower())

                                {
                                    if (oSOP.FileApprover.signstatuscode == 1)  //signed 
                                        oSOP.AssignFilePermissionToUsers("read", "remove", oldrvwr.useremailaddress.Trim());  //remove last owner read permission 
                                    else if (oSOP.FileApprover.signstatuscode == 2) //not signed
                                        oSOP.AssignFilePermissionToUsers("contribute", "remove", oldrvwr.useremailaddress.Trim());  //remove last owner read permission                         


                                }
                            }
                        }

                        //add reviewers in sql and add contribute permission with new change request

                        sop.AddFileReviewers();

                        sop.AssignFilePermissionToUsers("contribute", "add", sop.FileReviewers);

                        //update sop doc file with new signatories with SOP helper class
                        //do not use sop becuase it will lose new input data such vieweraccess type etc. that we need in 
                        //managing veiw access

                        SOPClass oSOPHelper = new SOPClass();

                        string templocaldirpath;
                        templocaldirpath = Server.MapPath(Utility.GetTempLocalDirPath());

                        oSOPHelper.FileID = sop.FileID;
                        oSOPHelper.FileChangeRqstID = sop.FileChangeRqstID;
                        oSOPHelper.FileLocalPath = templocaldirpath + sop.FileName;
                        oSOPHelper.SiteUrl = sop.SiteUrl;
                        oSOPHelper.DocumentLibName = sop.DocumentLibName;

                        oSOPHelper.GetSOPInfo();  //get latest sop info from database including newly changed signatories info

                        oSOPHelper.DownloadFileFromSharePoint(templocaldirpath);

                        oSOPHelper.UpdateCoverRevhistPage(true);

                        //Thread.Sleep(4000);

                        oSOPHelper.FileStream = System.IO.File.ReadAllBytes(oSOPHelper.FileLocalPath);

                        oSOPHelper.UploadDocument();

                        oSOPHelper = null;   //thanks bye bye 

                        // sop.FileLocalPath = templocaldirpath + sop.FileName;
                        //sop.DownloadFileFromSharePoint(templocaldirpath);
                        // sop.UpdateCoverRevhistPage(true);     

                        //Thread.Sleep(4000);

                        //sop.FileStream = System.IO.File.ReadAllBytes(sop.FileLocalPath);

                        // sop.UploadDocument();


                    }  //end checking signatory/contributor change



                    //now process viewers changes that does not need change request


                    if (oSOP.AllUsersReadAcc & !sop.AllUsersReadAcc)   //oSOP is previois an sop is new
                    {

                        //remove all group readers permissions

                        oSOP.AssignFilePermissionToGroup("read", "remove", "Watercooler Visitors");
                        oSOP.AssignFilePermissionToGroup("read", "remove", "SOPAllReaders");
                        oSOP.AssignFilePermissionToGroup("read", "remove", "SEC_Everyone_RadiantCanada");

                        //give custom permissions with new changes


                        //start checking by departement

                        if (sop.ViewAccessType.Trim() == "By Department")
                        {

                            //give read permission to custom users by department


                            if (sop.DepartmentCode > 0)  //if department is selected then preference is to get employees by department code
                            {

                                vwrItems = null;

                                sop.ViewAccessType = "By Department";

                                short sdeptcode = Convert.ToInt16(sop.DepartmentCode);
                                oEmp.departmentcode = sdeptcode;


                                oEmp.GetEmployeesByDeptCode();

                                vwrItems = oEmp.employees;

                                //first remove existing permission from the file, default is Watercooler Visitors


                                sop.AssignFilePermissionToUsers("read", "add", vwrItems);

                                //now add view access info by department in SQL Table
                                //we need this to retrieve and change in admin page

                                sop.DepartmentCode = sdeptcode;
                                sop.AddViewerAccessType();


                            }

                        }  //end checking by department

                        //check what old viewer access was there 
                        if (sop.ViewAccessType.Trim() == "By Users")
                        {

                            //give read permission to custom users
                            if (sop.FileviewersArr.Count() > 0)   //get employees from custom user list
                            {

                                vwrItems = null;

                                sop.ViewAccessType = "By Users";

                                vwrItems = JsonConvert.DeserializeObject<Employee[]>(sop.FileviewersArr[0]);

                                sop.AssignFilePermissionToUsers("read", "add", vwrItems);


                                //now add view access info by custom users in SQL table
                                //we need this to retrieve and change in admin page

                                sop.FileViewers = vwrItems;
                                sop.AddViewerAccessType();
                                sop.AddFileViewers();

                            }


                        }  //end checking by users




                    }  //end checking allusers old and new

                    if (!oSOP.AllUsersReadAcc & sop.AllUsersReadAcc)
                    {


                        //delete old readers

                        string rvwremail = "";

                     //   vwrItems = null;

                        if (oSOP.ViewAccessType == "By Users")
                        {

                            //oSOP.AssignFilePermissionToUsers("read", "remove", oSOP.FileReviewers);

                            foreach (Employee viewer in oSOP.FileViewers)
                            {
                                rvwremail = "";
                                if (viewer.useremailaddress.Trim().ToLower() != sop.FileOwnerEmail.Trim().ToLower())
                                {
                                    if (viewer.useremailaddress.Trim().ToLower() != sop.FileApproverEmail.Trim().ToLower())
                                    {
                                        foreach (Employee rvwr in sop.FileReviewers)
                                        {
                                            if (viewer.useremailaddress.Trim().ToLower() == rvwr.useremailaddress.Trim().ToLower())
                                            {
                                                rvwremail = viewer.useremailaddress.Trim().ToLower();
                                                break;
                                            }

                                        }

                                        if (rvwremail.Trim() == "")
                                            oSOP.AssignFilePermissionToUsers("read", "remove", viewer.useremailaddress.Trim());

                                    }

                                }

                            }

                        }


                        if (oSOP.ViewAccessType == "By Department")
                        {

                            vwrItems = null;

                            short sdeptcode = Convert.ToInt16(oSOP.DepartmentCode);
                            oEmp.departmentcode = sdeptcode;

                            oEmp.GetEmployeesByDeptCode();

                            vwrItems = oEmp.employees;

                            // oSOP.AssignFilePermissionToUsers("read", "remove", vwrItems);


                            foreach (Employee viewer in vwrItems)
                            {

                                rvwremail = "";
                                if (viewer.useremailaddress.Trim().ToLower() != sop.FileOwnerEmail.Trim().ToLower())
                                {
                                    if (viewer.useremailaddress.Trim().ToLower() != sop.FileApproverEmail.Trim().ToLower())
                                    {
                                        foreach (Employee rvwr in sop.FileReviewers)
                                        {
                                            if (viewer.useremailaddress.Trim().ToLower() == rvwr.useremailaddress.Trim().ToLower())
                                            {
                                                rvwremail = viewer.useremailaddress.Trim().ToLower();
                                                break;
                                            }

                                        }

                                        if (rvwremail.Trim() == "")
                                            oSOP.AssignFilePermissionToUsers("read", "remove", viewer.useremailaddress.Trim());

                                    }

                                }


                            }

                        }

                        //give visitor group permission as reader
                        sop.AssignFilePermissionToGroup("read", "add", "Watercooler Visitors");
                        sop.AssignFilePermissionToGroup("read", "add", "SOPAllReaders");
                        sop.AssignFilePermissionToGroup("read", "add", "SEC_Everyone_RadiantCanada");

                        sop.ViewAccessType = "All Users";

                        sop.AddViewerAccessType();    // add new view type in SQL table 


                    } //end checking all users



                    if ((!oSOP.AllUsersReadAcc & !sop.AllUsersReadAcc) & updatevwr)
                    {


                        //delete old readers

                        string rvwremail = "";

                        if (oSOP.ViewAccessType == "By Users")
                        {

                            //oSOP.AssignFilePermissionToUsers("read", "remove", oSOP.FileReviewers);

                            foreach (Employee viewer in oSOP.FileViewers)
                            {
                                rvwremail = "";
                                if (viewer.useremailaddress.Trim().ToLower() != sop.FileOwnerEmail.Trim().ToLower())
                                {
                                    if (viewer.useremailaddress.Trim().ToLower() != sop.FileApproverEmail.Trim().ToLower())
                                    {
                                        foreach (Employee rvwr in sop.FileReviewers)
                                        {
                                            if (viewer.useremailaddress.Trim().ToLower() == rvwr.useremailaddress.Trim().ToLower())
                                            {
                                                rvwremail = viewer.useremailaddress.Trim().ToLower();
                                                break;
                                            }

                                        }

                                        if (rvwremail == "")
                                            oSOP.AssignFilePermissionToUsers("read", "remove", viewer.useremailaddress.Trim());

                                    }

                                }

                            }

                        }


                        if (oSOP.ViewAccessType == "By Department")
                        {

                            vwrItems = null;
                            short sdeptcode = Convert.ToInt16(oSOP.DepartmentCode);
                            oEmp.departmentcode = sdeptcode;

                            oEmp.GetEmployeesByDeptCode();

                            vwrItems = oEmp.employees;

                            // oSOP.AssignFilePermissionToUsers("read", "remove", vwrItems);


                            foreach (Employee viewer in vwrItems)
                            {

                                rvwremail = "";
                                if (viewer.useremailaddress.Trim().ToLower() != sop.FileOwnerEmail.Trim().ToLower())
                                {
                                    if (viewer.useremailaddress.Trim().ToLower() != sop.FileApproverEmail.Trim().ToLower())
                                    {
                                        foreach (Employee rvwr in sop.FileReviewers)
                                        {
                                            if (viewer.useremailaddress.Trim().ToLower() == rvwr.useremailaddress.Trim().ToLower())
                                            {
                                                rvwremail = viewer.useremailaddress.Trim().ToLower();
                                                break;
                                            }

                                        }

                                        if (rvwremail == "")
                                            oSOP.AssignFilePermissionToUsers("read", "remove", viewer.useremailaddress.Trim());

                                    }

                                }


                            }

                        }


                        //add new viewers

                        //if by department then check if new and old dept are same, if not add new dept and remove old dept

                        if (sop.ViewAccessType == "By Department")
                        {
                            if (sop.DepartmentCode > 0)  //if same nothing to 
                            {
                                sop.ViewAccessType = "By Department";

                                vwrItems = null;

                                short sdeptcode = Convert.ToInt16(sop.DepartmentCode);
                                oEmp.departmentcode = sdeptcode;


                                oEmp.GetEmployeesByDeptCode();

                                vwrItems = oEmp.employees;

                                //first remove existing permission from the file, default is Watercooler Visitors


                                sop.AssignFilePermissionToUsers("read", "add", vwrItems);

                                //now add view access info by department in SQL Table
                                //we need this to retrieve and change in admin page

                                sop.DepartmentCode = sdeptcode;
                                sop.AddViewerAccessType();



                            }
                        }

                        //check what old viewer access was there 
                        if (sop.ViewAccessType.Trim() == "By Users")
                        {
                            //insert users in sql tables
                            //give read permission to custom users
                            if (sop.FileviewersArr.Count() > 0)   //get employees from custom user list
                            {

                                vwrItems = null;
                                sop.ViewAccessType = "By Users";

                                //first remove existing permission from the file, default is Watercooler Visitors

                                //give read permission to all custom viewers

                                vwrItems = JsonConvert.DeserializeObject<Employee[]>(sop.FileviewersArr[0]);

                                sop.AssignFilePermissionToUsers("read", "add", vwrItems);


                                //now add view access info by custom users in SQL table
                                //we need this to retrieve and change in admin page

                                sop.FileViewers = vwrItems;
                                sop.AddViewerAccessType();
                                sop.AddFileViewers();

                            }

                        }



                        //end giving viewer permissions

                        //Session["SOPMsg"] = "Admin SOP: You have successfully submitted all admin changes of SOP:" + sop.SOPNo;

                        //return Json(new { success = true, responseText = "The SOP " + sop.FileName + " has been successfully changed!" }, JsonRequestBehavior.AllowGet);




                    }   //end checking publish 


                    TempData["SOPMsg"] = "Success: Submitted administrative changes to the SOP have been updated.";

                    return Json(new { success = true, responseText = "Success: Submitted administrative changes to the SOP have been updated." }, JsonRequestBehavior.AllowGet);



                }  //end checking submit button type

                else  //no valid data was entered
                {

                    TempData["SOPMsg"] = "Error:Valid SOP admin data is required.";

                    return Json(new { success = true, responseText = "Error:Valid SOP admin data is required" }, JsonRequestBehavior.AllowGet);

                    //   return RedirectToAction("SOPMessage");


                }



            }
            catch (Exception ex)
            {

                TempData["SOPMsg"] = "Error:Failed to load SOP Admin page and complete admin changes due to following system error:" + ex.Message;

                // return RedirectToAction("SOPMessage");

                return Json(new { success = true, responseText = "Error:Failed to submit changes. Cannot make admin changes and archive document in the same submission." }, JsonRequestBehavior.AllowGet);


            }

            finally
            {
                oSOP = null;
                oEmp = null;
                vwrItems = null;
                rvwrItems = null;
            }




        }

        public ActionResult SignSOP(int? id)
        {

            //http://localhost:58639/Home/SignSOP/41?chngreqid=8

            string user = "";
            string loggedinusereml = "";
            int loggedinuserid = 0;

            SOPSignatureModel oSM = new SOPSignatureModel();
            SOPClass oSOP = new SOPClass();

            TempData["PageTitle"] = "Signing Off on SOP";

            try

            {


                if (Utility.IsSessionExpired())
                    return RedirectToAction("LogIn");



                loggedinusereml = Utility.GetCurrentLoggedInUserEmail();
                loggedinuserid = Utility.GetLoggedInUserID();

                //loggedinusereml = "rmcdowell@radiantdelivers.com";
                //loggedinuserid = 238;

                //loggedinusereml = "skiernan@radiantdelivers.com";
                //loggedinuserid = 260;


                //loggedinusereml = "gcallaghan@radiantdelivers.com";
                //loggedinuserid = 83;

                //loggedinusereml = "kislam@radiantdelivers.com";
                //loggedinuserid = 145;

                //loggedinusereml = "ediallo@radiantdelivers.com";
                //loggedinuserid = 69;


                //loggedinusereml = "student05@radiantdelivers.com";
                //loggedinuserid = 362;



                string strchangereqid = Request.QueryString["chngreqid"];
                int changereqid = 0;

                if (strchangereqid != null && strchangereqid != "")
                    changereqid = Convert.ToInt32(strchangereqid);

                TempData["FileID"] = id;
                TempData["ChangeIReqID"] = changereqid;

                //need this to keep activity id so we can update sign status after submit
                oSM.FileID = Convert.ToInt32(id);
                oSM.LoggedInUserID = loggedinuserid;
                oSM.ChangeRequestID = changereqid;

                oSOP.FileID = Convert.ToInt32(id);
                oSOP.FileChangeRqstID = Convert.ToInt32(changereqid);
                oSOP.SiteUrl = siteurl;

                oSOP.GetSOPInfo();

                TempData["SOPName"] = oSOP.FileTitle;

                //now verify if he is signatory in any approver group with logged in email address with this file
                //and change request id
                //if not then redirect to unauthonticated error page otherwise get his/her sign status and date to
                //show it in viewer

                oSM.LoggedInUserIsOwner = "";
                oSM.LoggedInUserIsApprover = "";
                oSM.LoggedInUserIsReviewer = "";
                oSM.LoggedInUserAllStatus = "";
                oSM.LoggedInSignDate = DateTime.Today;

                if (loggedinusereml.ToLower().Trim() == oSOP.FileOwner.useremailaddress.ToLower().Trim())
                {
                    oSM.LoggedInUserIsOwner = "yes";
                    oSM.LoggedInUserAllStatus = "Owner";
                    oSM.GetOwnerActivityID();

                    TempData["OwnerActivityID"] = oSM.OwnerActivityID;

                    if (oSOP.FileOwner.signstatus.ToLower() == "signed")
                    {
                        oSM.LoggedInSignedAsOwner = true;
                        oSM.LoggedInSignDate = oSOP.FileOwner.signaturedate;
                    }
                    else
                        oSM.LoggedInSignedAsOwner = false;
                }

                if (loggedinusereml.ToLower().Trim() == oSOP.FileApprover.useremailaddress.ToLower().Trim())
                {
                    oSM.LoggedInUserIsApprover = "yes";
                    if (oSM.LoggedInUserAllStatus != "")
                        oSM.LoggedInUserAllStatus = oSM.LoggedInUserAllStatus + ", " + "Approver";
                    else
                        oSM.LoggedInUserAllStatus = "Approver";

                    oSM.GetApproverActivityID();
                    TempData["ApproverActivityID"] = oSM.ApproverActivityID;

                    if (oSOP.FileApprover.signstatus.ToLower() == "signed")
                    {
                        oSM.LoggedInSignedAsApprover = true;
                        oSM.LoggedInSignDate = oSOP.FileApprover.signaturedate;
                    }
                    else
                    {
                        oSM.LoggedInSignedAsApprover = false;
                    }

                }


                //assign values for this signatory

                oSM.LoggedInUserEmail = loggedinusereml;
                oSM.SOPNo = oSOP.SOPNo;
                oSM.SOPName = oSOP.FileTitle;
                oSM.SOPUrl = oSOP.FileLink;
                oSM.SOPFilePath = oSOP.FilePath;

                oSM.SOPOwnerSignature = oSOP.FileOwner;
                oSM.SOPApprvrSignature = oSOP.FileApprover;
                oSM.SOPRvwerSignatures = oSOP.FileReviewers;

            //    TempData["reviewers"] = oSM.SOPRvwerSignatures;


                foreach (Employee rvwr in oSM.SOPRvwerSignatures)
                {
                    if (loggedinusereml.ToLower().Trim() == rvwr.useremailaddress.ToLower().Trim())
                    {
                        oSM.LoggedInUserIsReviewer = "yes";
                        if (oSM.LoggedInUserAllStatus != "")
                            oSM.LoggedInUserAllStatus = oSM.LoggedInUserAllStatus + ", " + "Reviewer";
                        else
                            oSM.LoggedInUserAllStatus = "Reviewer";

                        oSM.GetReviewerActivityID();
                        TempData["ReviewerActivityID"] = oSM.ReviewerActivityID;


                        if (rvwr.signstatus.ToLower() == "signed")
                        {
                            rvwr.HasSignedSOP = true;
                        }
                        else
                        {
                            rvwr.HasSignedSOP = false;
                            rvwr.signaturedate = DateTime.Today;
                        }

                        break;
                    }
                }

                //logged in user is no where in approvers so redirect him to unauthenticated error page
                if (oSM.LoggedInUserIsOwner == "" && oSM.LoggedInUserIsApprover == "" && oSM.LoggedInUserIsReviewer == "")

                {
                    // ViewBag.ErrorMessage = "SOP Application: Session not Timed out";
                    TempData["SOPMsg"] = "Error: You are not an authenticated approver and cannot sign document in question.";
                    return RedirectToAction("SOPMessage");


                }

                


                return View(oSM);


            }

            catch (Exception ex)
            {

                TempData["SOPMsg"] = "Error:Failed to load the page due to the following system error: "+ex.Message;
                return RedirectToAction("SOPMessage");


            }

            finally
            {
                oSOP = null;
            }

        }


        [HttpPost]
        public ActionResult SignSOP(SOPSignatureModel sm)
        {


            try

            {

                if (Utility.IsSessionExpired())
                    return RedirectToAction("LogIn");



                if (TempData["FileID"] != null && TempData["FileID"].ToString() != "")
                    sm.FileID = Convert.ToInt32(TempData["FileID"]);

                if (TempData["ChangeIReqID"] != null && TempData["ChangeIReqID"].ToString() != "")
                {

                    sm.ChangeRequestID = Convert.ToInt32(TempData["ChangeIReqID"]);
                }



                if (TempData["OwnerActivityID"]!=null && TempData["OwnerActivityID"].ToString() !="")
                {

                    sm.OwnerActivityID = Convert.ToInt32(TempData["OwnerActivityID"]);
                }

                if (TempData["ApproverActivityID"] != null && TempData["ApproverActivityID"].ToString() != "")
                {

                    sm.ApproverActivityID = Convert.ToInt32(TempData["ApproverActivityID"]);
                }

                if (TempData["ReviewerActivityID"] != null && TempData["ReviewerActivityID"].ToString() != "")
                {

                    sm.ReviewerActivityID = Convert.ToInt32(TempData["ReviewerActivityID"]);
                }


                   
                if (sm.UpdateSignatures())
                {
                    sm.UpdateChangeReqstApproval();   //if all signatories sign the SOP change sign status signed in change request activities table

                    TempData["SOPMsg"] = "Document Signed by:"+Utility.GetLoggedInUserFullName();

                    return RedirectToAction("SOPMessage");


                }
                else
                {

                    TempData["SOPMsg"] = "Error:You have already signed off on this SOP.";
                     
                    return RedirectToAction("SOPMessage");


                }

            }

            catch (Exception ex)
            {

                TempData["SOPMsg"] = "Error:Failed to sign off the SOP due to the following system error:" + ex.Message;

                return RedirectToAction("SOPMessage");

            }


            Session["SOPMsg"] = "Signing Off SOP: You already signed off the SOP!";

            return RedirectToAction("SOPMessage");



            //return View(sm);
        }

        public ActionResult SOPDashboard()
        {
            // do your logging here

            if (Utility.IsSessionExpired())
                return RedirectToAction("LogIn");


           return Redirect("http://camis1-bioasp01/Reports/Pages/Report.aspx?ItemPath=%2fSOP+Reports%2fSOP+Dashboard");

         //   return Redirect("http://camis1-bioasd01/Reports_BI_DEV/Pages/Report.aspx?ItemPath=%2fSOP+Reports%2fSOP+Dashboard&ViewMode=Detail");

            
        }

        public ActionResult AccessArchive()
        {


            if (Utility.IsSessionExpired())
                return RedirectToAction("LogIn");

            SOPClass oSop = new SOPClass();

            TempData["PageTitle"] = "Access SOP Archive";

            TempData["SOPName"] = "N/A";

            if (!oSop.AuthenticateUser("accessarchive"))
            {

                TempData["SOPMsg"] = "Error:SOP archive folder can only be accessed by SOP Owners.";

                return RedirectToAction("SOPMessage");

            }

            oSop = null;

            //return Redirect("http://camis1-bioasp01/Reports/Pages/Report.aspx?ItemPath=%2fSOP+Reports%2fSOP+Dashboard");

            return Redirect("https://radiantdelivers.sharepoint.com/:f:/s/Watercooler/Ev0caGyQgLZBl1EaJzxcn4cB8cMsLLlWio-FJt8hq0GTnA?e=j00RuN");



    }


    //public ActionResult UploadSOPFile()
    //{

    //    if (Utility.IsSessionExpired())
    //        return RedirectToAction("LogIn");



    //    ViewBag.Message = "Upload SOP File";



    //    ViewBag.ddlDeptFolders = new SelectList(Utility.GetFolders(), "FileName", "FileName");

    //    ViewBag.employees = ViewBag.employees = (from c in ctx.vwUsers select new { c.useremailaddress, c.userfullname }).Distinct();


    //    ViewBag.departments = (from c in ctx.codesdepartments select new { c.departmentname, c.departmentcode }).Distinct();

    //    ViewBag.updfrequnits= (from c in ctx.codesUnits select new { c.Unitname, c.unitcode,c.UnitType }).Where(x=>x.UnitType == "UpdateFrequency").Distinct();

    //    return View();
    //}

    public ActionResult SOPMessage()
        {
            
            return View();

        }

        

        public ActionResult ProcessPublish()
        {
            int id = 0;  //file id will be provide through dashboad
            int changereqid = 0; //change request id 

            if (TempData["FileID"]!=null)
                id =Convert.ToInt32(TempData["FileID"]);
            if (TempData["ChangeIReqID"]!=null)
                changereqid =Convert.ToInt32(TempData["ChangeIReqID"]);

            //id = ViewBag.FileID;
            //changereqid = ViewBag.ChangeIReqID;

            Session["SOPMsg"] = "";

            //Session["Dashboardlink"] = "http://camis1-bioasp01/Reports/Pages/Report.aspx?ItemPath=%2fSOP+Reports%2fSOP+Dashboard";

            Session["Dashboardlink"] = Utility.GetDashBoardUrl();

            Logger oLogger = new Logger();
            oLogger.LogFileName = HttpContext.Server.MapPath(Utility.GetLogFilePath());

            //Employee oEmp = new Employee();
            SOPClass oSOP = new SOPClass();


            string templocaldirpath; 

            oSOP.SiteUrl = siteurl;
            oSOP.FileID = id;
            oSOP.DocumentLibName = "SOP";
            //oSOP.DocumentLibName = Utility.GetDocLibraryName();
            oSOP.FileChangeRqstID = changereqid;



            try
            {

                templocaldirpath = Server.MapPath(Utility.GetTempLocalDirPath());

                oLogger.UpdateLogFile(DateTime.Now.ToString() + ":Publish SOP Action:Started publishing with File ID:" + id.ToString()+", change request id:"+ changereqid.ToString());


                if (Utility.IsSessionExpired())
                    return RedirectToAction("LogIn");

                oSOP.GetSOPInfo();  //get updated reviewers, approver, owner, version, file name etc.

                TempData["SOPName"] = oSOP.FileTitle;

                //start validating logged in user and SOP

                if (!oSOP.AuthenticateUser("publish"))   //only approver can publish a signed SOP

                {
                    oSOP = null;
                    TempData["SOPMsg"] = "Error:Only Approver of SOP can publish document.";
                    return RedirectToAction("SOPMessage");
                }


                if (changereqid == 0)     //change request is required to publish aganist a change
                {

                    TempData["SOPMsg"] = "Error:Valid change request is required to publish the SOP.";

                    return RedirectToAction("SOPMessage");

                }

                //assign SOP basic info


                oSOP.FileLocalPath = templocaldirpath + oSOP.FileName;

                //We need to check whether the SOP is signed by all parties (approver, reviewer, owner)
                //we will check signed status code in changeactivities table, it must be 1 to publish the sop

                if (oSOP.FileStatuscode == 2)  //not signed
                {
                    TempData["SOPMsg"] = "Error:Operation failed, document has not be signed by all signatories";

                    return RedirectToAction("SOPMessage");

                }


                if (oSOP.FileStatuscode == 3)  //published 
                {
                    TempData["SOPMsg"] = "Error:This SOP has already been published.";

                    return RedirectToAction("SOPMessage");

                }

                oLogger.UpdateLogFile(DateTime.Now.ToString() + ":Publish SOP Action:Successfully validated  SOP for publishing.");


                //just before publishing we need to update the coversheet with signed status of reviewers, approver
                //and owner as well as update version no, revision history etc.


                //string templocaldirpath = Server.MapPath("~/Content/DocFiles/");


                if (oSOP.FileStatuscode == 1)  //signed and ready to publish
                {

                    //download from sharepont online SOP lib so we can update it locally

                    oSOP.DownloadFileFromSharePoint(templocaldirpath);

                    oLogger.UpdateLogFile(DateTime.Now.ToString() + ":Publish SOP Action:Successfully downloaded SOP into local dir from SharePoint Online");


                    //update the cover page and rev history with xceed docx .net library

              //       oSOP.UpdateCoverRevhistPageDocX(true);

                    oSOP.UpdateCoverRevhistPage(true);     //interop com version does not work.

                    oLogger.UpdateLogFile(DateTime.Now.ToString() + ":Publish SOP Action:Successfully updated cover sheet with updated owner, approver etc. and revision history in SOP at local dir");


                    //upload the updated file again to the SOP lib in sharepoint online.


                  //  Thread.Sleep(6000);

                    oSOP.FileStream = System.IO.File.ReadAllBytes(oSOP.FileLocalPath);

                    oSOP.UploadDocument();

                    oLogger.UpdateLogFile(DateTime.Now.ToString() + ":Publish SOP Action:Successfully uoloaded updated SOP in SharePoint Online SOP Library");


                    // at last update status to approve in the so employees with given read access can access it


                    //reassing approvers permission as reader before publishing
                    
                    //remove this on regular schedule as signing a page will remove this contribute permission.
                  //  oSOP.ViewAccessType = "Inherit";

                  //  oSOP.AssignSignatoresReadPermission();


                    if (oSOP.PublishFile())
                    {

                       oLogger.UpdateLogFile(DateTime.Now.ToString() + ":Publish SOP Action:Successfully published (approved) SOP in SharePoint Online SOP Library.");

                        TempData["SOPMsg"] = "Success:The SOP has been published.";

                        return RedirectToAction("SOPMessage");

                    }

                    else
                    {
                        oLogger.UpdateLogFile(DateTime.Now.ToString() + ":Publish SOP Action:Failed to publish (approved) the SOP in SharePoint Online SOP Library with reaosn:"+oSOP.ErrorMessage);

                        TempData["SOPMsg"] = "Error:Failed to publish the SOP due to the following system error: " + oSOP.ErrorMessage ;
                        return RedirectToAction("SOPMessage");
                    }


                } //end checking signed status


                return View();



            }   //end try

            catch (Exception ex)
            {

                oLogger.UpdateLogFile(DateTime.Now.ToString() + ":Publish SOP Action:Failed to publish with error:" + ex.Message);

                TempData["SOPMsg"] = "Failed to publish the SOP due to the following system error:" + ex.Message;
                return RedirectToAction("SOPMessage");

            }
            finally
            {

                oSOP = null;
                //oEmp = null;
                oLogger = null;

                

            }
           

           // return View();

        }


        public ActionResult PublishFile(int? id)
        {
            //give this url to Elhadj to link to dashboard
            // http://localhost:58639/Home/PublishFile/40?chngreqid=6

            if (Utility.IsSessionExpired())
                return RedirectToAction("LogIn");



            string changereqid = Request.QueryString["chngreqid"];

            //   ViewBag.FileID = id;
            //   ViewBag.ChangeIReqID = changereqid;

            TempData["FileID"] = id;
            TempData["ChangeIReqID"] = changereqid;



            return View();
        }


        public ActionResult ProcessChngRqst()
        {


            TempData["PageTitle"] = "SOP Change Request";

            if (Utility.IsSessionExpired())
                return RedirectToAction("LogIn");



            //http://localhost:58639/Home/SOPChngeRequest/92

            int id = 0;  //file id will be provide through dashboad

            if (TempData["FileID"] != null)
                id = Convert.ToInt32(TempData["FileID"]);


            SOPClass oSOP = new SOPClass();
            int ownershipid = 0;
            int aproveid = 0;


            oSOP.FileID = id;   //asign file id
            oSOP.GetSOPInfoByFileID();

            TempData["SOPName"] = oSOP.FileTitle;

            try
            {

                //first check whether logged in user is in owner, approver or reviewer group of this sop
                //if not redirect to unauthenticated message page
                //otherwise proceed with change request

                if (!oSOP.AuthenticateUser("changerequest"))
                {

                    oSOP = null;

                    TempData["SOPMsg"] = "Error:Failed to authenticate user to make a change request.You must be the owner, approver, or reviewer of this SOP to request a change.";

                    return RedirectToAction("SOPMessage");


                }

                int lastchngreqid = oSOP.GetLastChngRequestID();
                oSOP.FileChangeRqstID = lastchngreqid;
                short lastchngstatcode = oSOP.GetChngReqSOPStatusCode();

                if (lastchngstatcode == 1 || lastchngstatcode == 2)
                {

                    TempData["SOPMsg"] = "Error:Failed to complete change request as last change request is not yet published. Note: You can change the SOP until the last change request is published.";

                    return RedirectToAction("SOPMessage");

                }


                if (lastchngstatcode == 3)   //last change request is approved so we can create new change request

                    {

                    oSOP.SiteUrl = siteurl;

                    oSOP.FileChangeRqsterID = Utility.GetLoggedInUserID();

                    oSOP.AddChangeRequest();   //it will get new chngreqid that will be used as follows

                    ownershipid = oSOP.GetOwnershipID();

                    oSOP.AddOwneractivities(ownershipid);

                    aproveid = oSOP.GetApproveID();

                    oSOP.AddApproveractivities(aproveid);

                    oSOP.AddRvwractvtsWithChngRqst();

                    //reassign edit permission permissions to all current signatories with new change request
                    //so they can start editing file


                    oSOP.GetSOPInfo();

                    //we must inheirt permissions so we don't loose previous permissions
                    oSOP.ViewAccessType = "Inherit";

                    oSOP.AssignSigatoriesPermission();

                    TempData["SOPMsg"] = "Success: Change request to the SOP have been submitted.";

                    return RedirectToAction("SOPMessage");

                }

                return View();


            }

            catch (Exception ex)
            {

                TempData["SOPMsg"] = "Error:Failed to submit change request due to the following system error:"+ex.Message;

                return RedirectToAction("SOPMessage");


            }

            finally
            {

                oSOP = null;
                
                GC.Collect();

            }
            

        }
        public ActionResult SOPChngeRequest(int? id)
        {
            //give this url to Elhadj to link to dashboard
            // http://localhost:58639/Home/SOPChngeRequest/97


            if (Utility.IsSessionExpired())
                return RedirectToAction("LogIn");


            TempData["FileID"] = id;

            TempData["PageTitle"] = "SOP Change Request";

            return View();
        }






        //[Authorize(Roles = "SOPADMIN")]
        //[Authorize(Roles = "TransfloARUsers")]
        //  [RoleFilter] with form authentication in web.cofig use this custom filter to redirect to custom page. make sure you don't use any role in authorize 

        [HttpGet]
        public ActionResult CreateUploadSOP()
        {

            if (Utility.IsSessionExpired())
                return RedirectToAction("LogIn");



            //run this protect configuration to encrypt config file so hacker cannot read 
            //sensitive data even they get the config file
            //run this just one time to encrypt or one time to dycript

            //  Utility.ProtectConfiguration();
         //   Utility.UnProtectConfiguration();   //dycrip it when you need to change any data in config file

            // ViewBag.Title = "Upload or Create SOP";  //I assigned in cshtml file

            ViewBag.ddlDeptFolders = new SelectList(Utility.GetFolders(), "FileName", "FileName");

            ViewBag.employees = (from c in ctx.vwUsers select new { c.useremailaddress, c.userfullname }).Distinct();

          //  Session["employees"] = ViewBag.employees;

            ViewBag.departments = (from c in ctx.codesSOPDepartments select new { c.sopdeptname, c.sopdeptcode }).Distinct();

            ViewBag.updfrequnits = (from c in ctx.codesUnits select new { c.Unitname, c.unitcode, c.UnitType }).Where(x => x.UnitType == "UpdateFrequency").Distinct();



            return View();
        }

        [HttpPost]
        public ActionResult CreateUploadSOP(SOPManagement.Models.SOPClass sop)
        {


            //if (Utility.IsSessionExpired())
            //    return RedirectToAction("LogIn");



            //run this protect configuration to encrypt config file so hacker cannot read 
            //sensitive data even they get the config file
            //run this just one time to encrypt or one time to dycript
            //Utility.ProtectConfiguration();
            //Utility.UnProtectConfiguration();


            SOPClass oSop = new SOPClass();

            Employee oEmp = new Employee();

            string user = "";
            //  string loggedinusereml = "";

            TempData.Keep();

            TempData["PageTitle"] = "SOP Create/Upload";

            Logger oLogger = new Logger();
            oLogger.LogFileName= HttpContext.Server.MapPath(Utility.GetLogFilePath());


            try

            {

                //  if (!ModelState.IsValid)   //we are supposed to use ModelState but we validated data through javascript so we don't use this

                if (sop.FileName != null && sop.FileName.Trim() != "")
                {

                    oSop.FileName = sop.FileName.Trim() + ".docx";

                    oSop.FileTitle = Path.ChangeExtension(oSop.FileName, null);

                }


                else if (sop.UploadedFile != null)

                {

                    //for uploaded file copy it from posted file to temp file

                    oSop.FileName = Path.GetFileName(sop.UploadedFile.FileName);

                    oSop.FileTitle = Path.ChangeExtension(oSop.FileName, null);    //without exctension


                }

                TempData["SOPName"] = oSop.FileTitle ;

                oLogger.UpdateLogFile(DateTime.Now.ToString() + ":CreateUpload Action:Started uploading or creating new SOP:"+sop.SOPNo);

                bool bProcessCompleted = true;

                //reload employee in viewbag for loading emplyees ddl 
                ViewBag.employees = ViewBag.employees = (from c in ctx.vwUsers select new { c.useremailaddress, c.userfullname }).Distinct();


                //start processing uploaded or new sop file 

                if (sop.SubFolderName.Trim() == "Please select a subfolder")
                    sop.SubFolderName = "";

                //1. [Authorized] attribute at the top of action authorizes the user by sopadmin role in domain 
                // then check if session is expired, if so redirect to session timeout page. 

                string failurename="";

                if (Utility.IsSessionExpired())
                {

                   // Session["SOPMsg"] = "SOP Create/Upload: Error - SOP Create/Upload: Session timed out";

                    bProcessCompleted = false;

                    failurename = "sessiontimeout";
                }

                if (bProcessCompleted)
                {

                //    loggedinusereml = Utility.GetCurrentLoggedInUserEmail();

                    oSop.FolderName = sop.FolderName;
                    oSop.SubFolderName = sop.SubFolderName;

                    //now authenticate the logged in user by Folder name 


                    //I turned it off as this is managed in client side ajax control in folder change event   

                    //if (!oSop.AuthenticateUser("createupload"))
                    //{
                    //    TempData["SOPMsg"] = "Error:Failed to authenticate user for Department:" + oSop.FolderName + ".You can only create or upload an SOP for a department that you are a part of.";

                    //    bProcessCompleted = false;

                    //    failurename = "accessdenied";


                    //}


                    if (bProcessCompleted)
                    {

                        //oEmp.useremailaddress = loggedinusereml;

                        //oEmp.GetUserByEmail();

                        //oSop.FileChangeRqsterID = oEmp.userid

                        oSop.FileChangeRqsterID = Utility.GetLoggedInUserID();

                        oSop.DocumentLibName = "SOP";

                        //oSop.DocumentLibName = Utility.GetDocLibraryName();         //"SOP";


                        oSop.SOPNo = sop.SOPNo;


                        //check duplicate SOP NO in DB table [deptsopfiles]
                        //it could be happened if in same directory last file was uploaded in 
                        //sharepoint online but data was not updated in DB through ms flow
                        //for delayed or broken process between SP and DB
                        //in this situation send error message that Last file upload is not 
                        //completed successfuly, please refresh Dashboard and check last file 
                        //was uploaed with this SOP NO 

                        if (oSop.HasDuplicateSOPNOInDB())

                        {
                            

                            TempData["SOPMsg"] = "Error:Duplicate SOP found." ;
                            TempData["SOPName"] = "SOP Name:" + oSop.FileTitle + ", SOP No:" + oSop.SOPNo;
                            TempData.Keep();

                            bProcessCompleted = false;

                            failurename = "duplicatesop";

                        }

                    }
                }   //end checking if session is alive

                if (bProcessCompleted)
                {

                    //log DateTime:sop.SOPNO: end collecting user email

                    //   bool fileloaded = false;

                    //log DateTime:sop.SOPNO: start saving new or updloaded to temp project folder


                    string tmpfiledirpathnm = Utility.GetTempLocalDirPath();

                    string tmpfiledirmappath = Server.MapPath(tmpfiledirpathnm);

                    Employee[] vwrItems;

                    Employee[] rvwrItems = JsonConvert.DeserializeObject<Employee[]>(sop.FilereviewersArr[0]);



                    if (sop.FileName != null && sop.FileName.Trim() != "")
                    {


                        //for new file copy from template to temp file

                        string tmpltmapfilepath = Server.MapPath(tmpfiledirpathnm + Utility.GetTemplateFileName());
                        string newmapfilepath = Server.MapPath(tmpfiledirpathnm + oSop.FileName);

                        System.IO.File.Copy(tmpltmapfilepath, newmapfilepath, true);

                    }


                    else if (sop.UploadedFile != null)

                    {

                        if (!Directory.Exists(tmpfiledirmappath))

                        {

                            Directory.CreateDirectory(tmpfiledirmappath);

                        }

                        sop.UploadedFile.SaveAs(tmpfiledirmappath + oSop.FileName);


                    }

                    oLogger.UpdateLogFile(DateTime.Now.ToString() + ":CreateUpload Action:Successfully saved file in local temp directory.");

                    //end step 2 saving new file or uploaded file in projetc folder


                    //3. Update cover sheet and revision history with file name, SOP No, reviewers, owner etc.
                    //DateTime:sop.SOPNo:file was saved in project temp folder successfully

                    //DateTime:sop.SOPNo:start updating covert sheet and rev history successfully

                    short supdfreq = Convert.ToInt16(sop.Updatefreq);

                    oSop.FileApproverEmail = sop.FileApproverEmail;
                    oSop.FileOwnerEmail = sop.FileOwnerEmail;
                    oSop.FileReviewers = rvwrItems;
                    oSop.Updatefreq = supdfreq;
                    oSop.Updatefrequnit = sop.Updatefrequnit;
                    oSop.Updfrequnitcode = sop.Updfrequnitcode;
                    // oSop.SOPEffectiveDate = Convert.ToDateTime(sop.SOPEffectiveDate);

                    //Employee oFileOwner = new Employee();
                    //oFileOwner.useremailaddress = sop.FileOwnerEmail;
                    //oFileOwner.GetUserByEmail();
                    //oSop.FileOwner = oFileOwner;


                    oSop.FileLocalPath = tmpfiledirmappath + oSop.FileName;

                    //we don't have any revision history for new or first time uploaded file
                    //FileRevision[] oRevarr = new FileRevision[1];

                    //FileRevision rev1 = new FileRevision();

                    //rev1.RevisionNo = "";
                    //rev1.RevisionDate = DateTime.Now;
                    //rev1.Description = "";

                    //oRevarr[0] = rev1;

                    //oSop.FileRevisions = oRevarr;


                    oSop.SiteUrl = siteurl;
                    oSop.FileCurrVersion = "1";    //for new file version no is 1 


                    //udpate coverpage with sop basic info and owner, reviewers, approver

               //     oSop.UpdateCoverRevhistPageDocX();

                     oSop.UpdateCoverRevhistPage();

                    oLogger.UpdateLogFile(DateTime.Now.ToString() + ":CreateUpload Action:Successfully updated cover sheet in SOP doc file in local temp dir.");


                    //end step 3 updating cover sheet


                    //4. Upload the updated file into sharepoint online SOP doc libray in correct department folder
                    //and sub folder enetred by user

                    //log it
                    //DateTime:sop.SOPNo:that successfully updated coversheet and rev history

                    //DateTime:sop.SOPNo:start uploading file in sharepoint online SOP doc library

                    //  Thread.Sleep(7000);


                    if (oSop.SubFolderName == "")
                        oSop.FilePath = "SOP/" + oSop.FolderName + "/";
                    else
                        oSop.FilePath = "SOP/" + oSop.FolderName + "/" + oSop.SubFolderName + "/";


                    oSop.FileStream = System.IO.File.ReadAllBytes(oSop.FileLocalPath);

                    oSop.UploadDocument();

                    oLogger.UpdateLogFile(DateTime.Now.ToString() + ":CreateUpload Action:Successfully uploaded SOP file in SharePoint Online SOP doc library.");

                    // end step 4 uploading file into sharepoint online sop doc lib

                    //5. Update SQL server tables with all info like, change request, reviewers, approvers etc.

                    //log 
                    //DateTime:sop.SOPNo:successfully uploaded file in SharePoint online SOP doc lib

                    //DateTime:sop.SOPNo:start updating SQL tables

                    //oSop.FileID got assigned after successfull uplaod in previous step

                    oSop.AddChangeRequest();

                    oSop.AddFileOwner();
                    oSop.AddFileApprover();
                    oSop.AddFileReviewers();
                    oSop.AddUpdateFreq();

                    oLogger.UpdateLogFile(DateTime.Now.ToString() + ":CreateUpload Action:Successfully updated all SQL server table with owner, approver, reviewer etc.");


                    //end step 5 SQL table upate


                    //6. assign proper permission for owner with full permission, reviewers with contribute permission, 
                    //and read permission to users according to admin users entry of users for read permission i.e read 
                    //permission to all, or a departement or custom users.


                    //log 
                    //DateTime:sop.SOPNo:successfully updated SQL tables
                    //DateTime:sop.SOPNo:start assigning permission to SOP file in sharepoint 

                    oSop.ViewAccessType = "";

                    if (sop.AllUsersReadAcc)   //by default all users have read permission
                    {

                        oSop.ViewAccessType = "All Users";

                        oSop.AddViewerAccessType();    // add new view type in SQL table 

                    }

                    else    
                    
                    //if All users are not permitted to view then customize the read permission according to either department or custom users

                    {
                        


                        if (sop.DepartmentCode > 0)  //if department is selected then preference is to get employees by department code
                        {

                            short sdeptcode = Convert.ToInt16(sop.DepartmentCode);
                            oEmp.departmentcode = sdeptcode;

                            oSop.ViewAccessType = "By Department";

                            oEmp.GetEmployeesByDeptCode();

                            vwrItems = oEmp.employees;

                            //first remove existing permission from the file, default is Watercooler Visitors

                            oSop.RemoveAllFilePermissions();

                            //give read permission to all users who are in the selected department

                            // oSop.AssignFilePermission("add", "read", vwrItems);  //this one hits sp server three times in a employee loop

                            oSop.AssignFilePermissionToUsers("read", "add", vwrItems);

                            //now add view access info by department in SQL Table
                            //we need this to retrieve and change in admin page

                            oSop.DepartmentCode = sdeptcode;
                            oSop.AddViewerAccessType();


                        }

                        else if (sop.FileviewersArr.Count() > 0)   //get employees from custom user list
                        {
                            vwrItems = JsonConvert.DeserializeObject<Employee[]>(sop.FileviewersArr[0]);

                            oSop.ViewAccessType = "By Users";
                            //first remove existing permission from the file, default is Watercooler Visitors

                            oSop.RemoveAllFilePermissions();

                            //give read permission to all custom viewers

                            // oSop.AssignFilePermission("add", "read", vwrItems);

                            oSop.AssignFilePermissionToUsers("read", "add", vwrItems);


                            //now add view access info by custom users in SQL table
                            //we need this to retrieve and change in admin page

                            oSop.FileViewers = vwrItems;
                            oSop.AddViewerAccessType();
                            oSop.AddFileViewers();

                        }


                        //then give read permission to sinatories so they can view but cannot modify file before submitting any change request.

                        //oSop.AssignFilePermissionToUsers("read", "add", sop.FileOwnerEmail);
                        //oSop.AssignFilePermissionToUsers("read", "add", sop.FileApproverEmail);
                        //oSop.AssignFilePermissionToUsers("read", "add", rvwrItems);

                    }



                     oSop.AssignSigatoriesPermission();

                    oLogger.UpdateLogFile(DateTime.Now.ToString() + ":CreateUpload Action:Successfully assigned all permissions with SOP file in SharePoint and completed processing SOP:"+sop.SOPNo);


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



                    //log 
                    //DateTime:sop.SOPNo:successfully assigend permission and completed all SOP processing

                    //  Send "Success" to ajax call back in view

                    TempData["SOPMsg"] = "Success: SOP has been loaded in management system.";

                    return Json(new { success = true, responseText = "The SOP " + sop.FileName + " has been successfully processed!" }, JsonRequestBehavior.AllowGet);

                }
                else   //processing failed due to validation i.e. duplicate SOPNO or Session time out 
                {
                    //log 

                    oLogger.UpdateLogFile(DateTime.Now.ToString() + ":CreateUpload Action:Failed to validate SOP entry i.e. session timed out, duplicate SOP.");

                    if (TempData["SOPMsg"]==null || TempData["SOPMsg"].ToString()=="")
                         TempData["SOPMsg"] = "Error:Failed to process SOP, please contact IT.";


                    if (failurename=="sessiontimeout")
                      return Json(new { success = false, responseText = "sessontimeout" }, JsonRequestBehavior.AllowGet);

                    else if (failurename == "duplicatesop")
                        return Json(new { success = false, responseText = "duplicatesop" }, JsonRequestBehavior.AllowGet);
                    else
                        return Json(new { success = false, responseText = "othererror" }, JsonRequestBehavior.AllowGet);
                }

                //if any server failutre to send requested response other than OK 200 code then ajax will raise error event


                // return View();
                //return Json(sop);




            }
            catch (Exception ex)
            {

                oLogger.UpdateLogFile(DateTime.Now.ToString() + ":CreateUpload Action:Failed uplaoding or creating SOP with error:"+ex.Message);

                TempData["SOPMsg"] = "Error:Failed to process the SOP due to the following system error:" + ex.Message;
                return Json(new { success = false, responseText = "Failed processing SOP " + sop.FileName + " , please contact IT!" }, JsonRequestBehavior.AllowGet);

            }

            finally
            {
                oSop = null;
                oEmp = null;
                oLogger = null;

                GC.Collect();
            }
            







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


            //SOPClass oSOP = new SOPClass();

            //oSOP.FolderName = foldername;


            //if (!oSOP.AuthenticateUser("createupload"))
            //{
            //    Session["SOPMsg"] = "SOP Create/Upload: Error - You are not authorized to create/upload SOP. Only owner of any file of the selected department can upload/create SOP.";

            //    return RedirectToAction("SOPMessage");



            //}

            //oSOP = null;

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
                }).Where(s => s.FilePath == "SOP/" + foldername + "/" && !s.FileName.Contains(".docx") && s.FileStatuscode==3).OrderBy(s=>s.FileName);   //valid sub folder


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


        public JsonResult AuthenticateUpload(string foldername)
        {

            SOPClass oSOP = new SOPClass();
            oSOP.FolderName = foldername;


            //now authenticate the logged in user by Folder name 

            bool noauthen=false;
         //  if (noauthen)
           if (!oSOP.AuthenticateUser("createupload"))
            {
                TempData["SOPMsg"] = "Error:You are not authenticated to create or upload SOP in this department.";

                oSOP = null;
                return Json(new { success = false, message = "ERROR: You are not authenticated to create or upload SOP in this department." });
            }
            else
            {
                oSOP = null;
                return Json(new { success = true, message = "yes" });
            }
             

        }




        //public void UploadFile(HttpPostedFileBase postedFile, string deptfoldername, string subfoldername, string sopno)



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

          //  oSop.DocumentLibName = Utility.GetDocLibraryName();

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
                oSop.FileReviewers = rvwrItems;
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

                        oSop.FileViewers = vwrItems;
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


    public class BaseController : Controller

    {

        protected override void OnActionExecuting (ActionExecutingContext filterContext)

        {

            // If session exists

            if (filterContext.HttpContext.Session != null)

            {

                //if new session

                if (filterContext.HttpContext.Session.IsNewSession)

                {

                    //for brand new session IsNewSession will be true and cookie will be null 
                    //for expired session IsNewSession will be true and cookie will not be null for previous session

                    string cookie = filterContext.HttpContext.Request.Headers["Cookie"];

                    //if cookie exists and sessionid index is greater than zero

                    if ((cookie != null) && (cookie.IndexOf("ASP.NET_SessionId") >= 0))

                    {

                        //redirect to desired session 

                        //expiration action and controller

                        filterContext.Result = RedirectToAction("LogIn", "Home");

                        return;

                    }

                    else
                        HttpContext.Session["UserFullName"] = Utility.GetLoggedInUserFullName();

                }

            }

            //otherwise continue with action

            base.OnActionExecuting(filterContext);

        }

    }

}



