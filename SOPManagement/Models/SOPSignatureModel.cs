using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;
namespace SOPManagement.Models
{
    public class SOPSignatureModel
    {

        public int FileID { get; set; }
        public int ChangeRequestID { get; set; }
        public int LoggedInUserID { get; set; }
        public string LoggedInUserEmail { get; set;}
        public string LoggedInUserFullName { get; set;}
        public string LoggedInUserJobTitle { get; set; }

        public string LoggedInUserIsOwner { get; set; }

        public string LoggedInUserIsApprover { get; set; }

        public string LoggedInUserIsReviewer { get; set; }

        public string LoggedInUserAllStatus { get; set; }

        [Display(Name = "Your Signature as Owner")]

        public bool LoggedInSignedAsOwner { get; set; }

        [Display(Name = "Your Signature as Approver")]
        public bool LoggedInSignedAsApprover { get; set; }

        [Display(Name = "Your Signature as Reviewer")]
        public bool LoggedInSignedAsReviewer { get; set; }

        [Display(Name = "Your Signature Date")]
        public DateTime LoggedInSignDate { get; set; }
        
        [Display(Name="SOP No")]
        public string SOPNo { get; set; }

        [Display(Name = "SOP Name")]
        public string SOPName { get; set; }   //with sopno in front SOPNO + " "+ FileTitle

        [Display(Name = "Link to SOP")]
        public string SOPUrl { get; set; }   //with sopno in front SOPNO + " "+ FileTitle

        [Display(Name = "SOP Latest Version")]
        public string SOPLastVersion { get; set; }   //with sopno in front SOPNO + " "+ FileTitle

        [Display(Name = "SOP Department")]
        public string SOPFilePath { get; set; }   //with sopno in front SOPNO + " "+ FileTitle

        [Display(Name = "Department Folder")]
        public string SOPDeptName { get; set; }   //with sopno in front SOPNO + " "+ FileTitle
        public string SOPSubDeptName { get; set; }   //with sopno in front SOPNO + " "+ FileTitle
        public Employee[] SOPRvwerSignatures { get; set; }

        public Employee SOPOwnerSignature { get; set; }

        public Employee SOPApprvrSignature { get; set; }

        public int OwnerActivityID { get; set; }

        public int ApproverActivityID { get; set; }

        public int ReviewerActivityID { get; set; }


        public void GetOwnerActivityID()
        {
            int owneractvityid = 0;
            int ownershipid = 0;

            using (var dbctx = new RadiantSOPEntities())
            {
                //verify logged in user is activer owner
                ownershipid = dbctx.fileowners.Where(o => o.ownerid == LoggedInUserID && o.fileid==FileID && o.ownerstatuscode==1).Select(o => o.ownershipid).FirstOrDefault();

                owneractvityid = dbctx.fileownersactivities.Where(o => o.ownershipid == ownershipid && o.changerequestid == ChangeRequestID).Select(o => o.owneractivityid).FirstOrDefault();
                OwnerActivityID = owneractvityid;


            }

        }

        public void GetApproverActivityID()
        {
            int aprvractvityid = 0;
            int approveid = 0;

            using (var dbctx = new RadiantSOPEntities())
            {
                approveid = dbctx.fileapprovers.Where(o => o.approverid == LoggedInUserID && o.fileid == FileID && o.approverstatuscode==1).Select(o => o.approveid).FirstOrDefault();

                aprvractvityid = dbctx.fileapproversactivities.Where(o => o.approveid == approveid && o.changerequestid == ChangeRequestID).Select(o => o.approveractivityid).FirstOrDefault();
                ApproverActivityID = aprvractvityid;


            }

        }

        public void GetReviewerActivityID()
        {
            int rvwractvityid = 0;
            int reviewid = 0;

            using (var dbctx = new RadiantSOPEntities())
            {
                reviewid = dbctx.filereviewers.Where(o => o.reviewerid == LoggedInUserID && o.fileid == FileID && o.reviewerstatuscode == 1).Select(o => o.reviewid).FirstOrDefault();

                rvwractvityid = dbctx.filereviewersactivities.Where(o => o.reviewid == reviewid && o.changerequestid == ChangeRequestID).Select(o => o.revieweractivityid).FirstOrDefault();
                ReviewerActivityID = rvwractvityid;


            }

        }

        
        public bool UpdateSignatures()
        {
            bool success = false;
            SOPClass oSOP = new SOPClass();

            oSOP.SiteUrl = Utility.GetSiteUrl();

            oSOP.FileID = FileID;
          
             oSOP.FileChangeRqstID = ChangeRequestID;

            // oSOP.GetSOPInfoByFileID();

            oSOP.GetSOPInfo();

            using (var dbcontext = new RadiantSOPEntities())
            {

                if (LoggedInSignedAsOwner)
                {

                    var result = dbcontext.fileownersactivities.SingleOrDefault(b => b.owneractivityid == OwnerActivityID);
                    if (result != null)
                    {
                        result.approvalstatuscode = 1;    //1=signed
                        result.statusdatetime = DateTime.Today;

                        dbcontext.SaveChanges();

                        success = true;
                    }



                    if (success)
                    {

                        int ownerid = 0;
                        int ownershipid = 0;
                        string owneremail = "";
                        bool removecontr = true;

                        ownershipid = Convert.ToInt32(dbcontext.fileownersactivities.Where(o => o.owneractivityid == OwnerActivityID
                        && o.changerequestid == ChangeRequestID).Select(o => o.ownershipid).FirstOrDefault());

                        ownerid = dbcontext.fileowners.Where(o => o.ownershipid == ownershipid && o.ownerstatuscode == 1).Select(o => o.ownerid).FirstOrDefault();
                        // fileid = dbcontext.fileowners.Where(o => o.ownershipid == ownershipid && o.ownerstatuscode == 1).Select(o => o.fileid).FirstOrDefault();
                        owneremail = dbcontext.vwUsers.Where(u => u.userid1 == ownerid).Select(o => o.useremailaddress).FirstOrDefault();

                        //first check if this owener is approver or reviewer and have signed it or not
                        //if this person signs in all group then remove contribute and give read permission
                        //if this person does not sign in any approver group then don't remove contribute permission

                            if (owneremail.Trim().ToLower()==oSOP.FileApproverEmail.Trim().ToLower())
                            {

                                if (oSOP.FileApprover.signstatuscode == 2)    //has not signed as approver so keep contribute permission
                                    removecontr = false;

                            }  
      

                            //if signed as approver then check if same person is reviewer and has not signed, if so do not remove contribute  
                            foreach (Employee rvwr in oSOP.FileReviewers)
                            {
                                if (owneremail.Trim().ToLower() == rvwr.useremailaddress.Trim().ToLower())
                                {
                                    if (rvwr.signstatuscode == 2) //has not signed as reviewer so keep contribute
                                        removecontr = false;

                                    break;
                                }
                            }

                        if (removecontr)
                        {
                            oSOP.AssignFilePermissionToUsers("contribute", "remove", owneremail.Trim().ToLower());
                            // System.Threading.Thread.Sleep(3000);

                            oSOP.AssignFilePermissionToUsers("read", "add", owneremail.Trim().ToLower());
                        }


                    }  //end checking success of signing

                }

                if (LoggedInSignedAsApprover)
                {

                    oSOP.GetSOPInfo();

                    var result = dbcontext.fileapproversactivities.SingleOrDefault(b => b.approveractivityid == ApproverActivityID);
                    if (result != null)
                    {
                        result.approvalstatuscode = 1;    //1=signed
                        result.statusdatetime = DateTime.Today;

                        dbcontext.SaveChanges();

                        success = true;
                    }

                    if (success)
                    {

                        int apporverid = 0;
                        int approveid = 0;
                        string approveremail = "";
                        bool removecontr = true;


                        approveid = Convert.ToInt32(dbcontext.fileapproversactivities.Where(o => o.approveractivityid == ApproverActivityID
                        && o.changerequestid == ChangeRequestID).Select(o => o.approveid).FirstOrDefault());

                        apporverid = dbcontext.fileapprovers.Where(o => o.approveid == approveid && o.approverstatuscode == 1).Select(o => o.approverid).FirstOrDefault();
                        // fileid = dbcontext.fileowners.Where(o => o.ownershipid == ownershipid && o.ownerstatuscode == 1).Select(o => o.fileid).FirstOrDefault();
                        approveremail = dbcontext.vwUsers.Where(u => u.userid1 == apporverid).Select(o => o.useremailaddress).FirstOrDefault();

                        if (approveremail.Trim().ToLower() == oSOP.FileOwnerEmail.Trim().ToLower())
                        {

                            if (oSOP.FileOwner.signstatuscode == 2)    //has not signed as owner so keep contribute permission
                                removecontr = false;

                        }


                        //if signed as approver then check if same person is reviewer and has not signed, if so do not remove contribute  
                        foreach (Employee rvwr in oSOP.FileReviewers)
                        {
                            if (approveremail.Trim().ToLower() == rvwr.useremailaddress.Trim().ToLower())
                            {
                                if (rvwr.signstatuscode == 2) //has not signed as reviewer so keep contribute
                                    removecontr = false;

                                break;
                            }
                        }



                        if (removecontr)
                        {
                            oSOP.AssignFilePermissionToUsers("contribute", "remove", approveremail.Trim().ToLower());
                            // System.Threading.Thread.Sleep(3000);

                            oSOP.AssignFilePermissionToUsers("read", "add", approveremail.Trim().ToLower());
                        }
                    }

                }

      
                if (LoggedInSignedAsReviewer)
                {

                    oSOP.GetSOPInfo();

                    var result = dbcontext.filereviewersactivities.SingleOrDefault(b => b.revieweractivityid == ReviewerActivityID);
                    if (result != null)
                    {
                        result.approvalstatuscode = 1;    //1=signed
                        result.statusdatetime = DateTime.Today;

                        dbcontext.SaveChanges();

                        success = true;
                    }

                    if (success)
                    {
                        int reviewerid = 0;
                        int reviewid = 0;
                        string revieweremail = "";
                        bool removecontr = true;


                        reviewid = Convert.ToInt32(dbcontext.filereviewersactivities.Where(o => o.revieweractivityid == ReviewerActivityID
                        && o.changerequestid == ChangeRequestID).Select(o => o.reviewid).FirstOrDefault());

                        reviewerid = dbcontext.filereviewers.Where(o => o.reviewid == reviewid && o.reviewerstatuscode == 1).Select(o => o.reviewerid).FirstOrDefault();
                        // fileid = dbcontext.fileowners.Where(o => o.ownershipid == ownershipid && o.ownerstatuscode == 1).Select(o => o.fileid).FirstOrDefault();
                        revieweremail = dbcontext.vwUsers.Where(u => u.userid1 == reviewerid).Select(o => o.useremailaddress).FirstOrDefault();

                        if (revieweremail.Trim().ToLower() == oSOP.FileOwnerEmail.Trim().ToLower())
                        {

                            if (oSOP.FileOwner.signstatuscode == 2)    //has not signed as owner so keep contribute permission
                                removecontr = false;
                        }

                        if (revieweremail.Trim().ToLower() == oSOP.FileApproverEmail.Trim().ToLower())
                        {

                            if (oSOP.FileApprover.signstatuscode == 2)    //has not signed as approver so keep contribute permission
                                removecontr = false;
                        }


                        if (removecontr)
                        {
                            oSOP.AssignFilePermissionToUsers("contribute", "remove", revieweremail.Trim().ToLower());
                            // System.Threading.Thread.Sleep(3000);

                            oSOP.AssignFilePermissionToUsers("read", "add", revieweremail.Trim().ToLower());
                        }



                    }

                }


            }

            oSOP = null;


            return success;


        }


        public bool GetSignStatusOfSignatory(string approvertype, string emailaddress)
        {
            bool signstatus = false;
            int userid = 0;
            int approveid = 0;
            int ownershipid = 0;
            int reviewid = 0;

            short signstatuscode = 0;

            using (var dbctx = new RadiantSOPEntities())
            {

                if (approvertype == "onwer")
                {

                    userid = dbctx.vwUsers.Where(u => u.useremailaddress.Trim().ToLower() == emailaddress.Trim().ToLower()).Select(u => u.userid1).FirstOrDefault();
                    ownershipid = dbctx.fileowners.Where(u => u.fileid == FileID && u.ownerid == userid && u.ownerstatuscode == 1).Select(u => u.ownershipid).FirstOrDefault();
                    signstatuscode = Convert.ToInt16(dbctx.fileownersactivities.Where(u => u.ownershipid == ownershipid && u.changerequestid == ChangeRequestID).Select(u => u.approvalstatuscode).FirstOrDefault());


                    if (signstatuscode == 1)   //approver signed
                        signstatus = true;
                }


                if (approvertype == "approver")
                {

                    userid = dbctx.vwUsers.Where(u => u.useremailaddress.Trim().ToLower() == emailaddress.Trim().ToLower()).Select(u => u.userid1).FirstOrDefault();
                    approveid = dbctx.fileapprovers.Where(u => u.fileid == FileID && u.approverid == userid && u.approverstatuscode == 1).Select(u => u.approveid).FirstOrDefault();
                    signstatuscode = Convert.ToInt16(dbctx.fileapproversactivities.Where(u => u.approveid == approveid && u.changerequestid == ChangeRequestID).Select(u => u.approvalstatuscode).FirstOrDefault());


                    if (signstatuscode == 1)   //approver signed
                        signstatus = true;
                }

                if (approvertype == "reviewer")
                {

                    userid = dbctx.vwUsers.Where(u => u.useremailaddress.Trim().ToLower() == emailaddress.Trim().ToLower()).Select(u => u.userid1).FirstOrDefault();
                    reviewid = dbctx.filereviewers.Where(u => u.fileid == FileID && u.reviewerid == userid && u.reviewerstatuscode == 1).Select(u => u.reviewid).FirstOrDefault();
                    signstatuscode = Convert.ToInt16(dbctx.filereviewersactivities.Where(u => u.reviewid == reviewid && u.changerequestid == ChangeRequestID).Select(u => u.approvalstatuscode).FirstOrDefault());


                    if (signstatuscode == 1)   //approver signed
                        signstatus = true;
                }



            }


            return signstatus;
        }

        public void UpdateChangeReqstApproval()
        {

            //find not signed=2 status code in all activities table
            //if we do not find it then all are signed as there will be two status 1=signed 2=not signed

            int ownrnotsgnedactvty = 0;
            int aprvrnotsgnedactvty = 0;
            int rvwrnotsgnedactvty = 0;

            using (var dbctx = new RadiantSOPEntities())
            {
                ownrnotsgnedactvty = dbctx.fileownersactivities.Where(o => o.changerequestid == ChangeRequestID && o.approvalstatuscode == 2).Select(o => o.owneractivityid).FirstOrDefault();
                aprvrnotsgnedactvty = dbctx.fileapproversactivities.Where(o => o.changerequestid == ChangeRequestID && o.approvalstatuscode == 2).Select(o => o.approveractivityid).FirstOrDefault();
                rvwrnotsgnedactvty = dbctx.filereviewersactivities.Where(o => o.changerequestid == ChangeRequestID && o.approvalstatuscode == 2).Select(o => o.revieweractivityid).FirstOrDefault();

                //if there is not signed that is all are signed then update then update changeactivity table with signed status
                //so approver can publish the SOP

                if (ownrnotsgnedactvty==0 && aprvrnotsgnedactvty==0 && rvwrnotsgnedactvty==0)
                {

                    var result = dbctx.filechangerequestactivities.SingleOrDefault(b => b.changerequestid == ChangeRequestID);
                    if (result != null)
                    {
                        result.approvalstatuscode = 1;    //1=signed
                        result.statusdatetime = DateTime.Today;

                        dbctx.SaveChanges();

                   }


                }

            }


        }

    }
}