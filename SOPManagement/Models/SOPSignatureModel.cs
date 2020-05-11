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
        
        [Display(Name="SOP NO")]
        public string SOPNo { get; set; }

        [Display(Name = "SOP Name")]
        public string SOPName { get; set; }   //with sopno in front SOPNO + " "+ FileTitle

        [Display(Name = "SOP Link")]
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
                reviewid = dbctx.filereviewers.Where(o => o.reviewerid == LoggedInUserID && o.fileid == FileID && o.reviewerstatuscode==1).Select(o => o.reviewid).FirstOrDefault();

                rvwractvityid = dbctx.filereviewersactivities.Where(o => o.reviewid == reviewid && o.changerequestid == ChangeRequestID).Select(o => o.revieweractivityid).FirstOrDefault();
                ReviewerActivityID = rvwractvityid;


            }

        }


        public bool UpdateSignatures()
        {
            bool success = false;

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

                }

                if (LoggedInSignedAsApprover)
                {

                    var result = dbcontext.fileapproversactivities.SingleOrDefault(b => b.approveractivityid == ApproverActivityID);
                    if (result != null)
                    {
                        result.approvalstatuscode = 1;    //1=signed
                        result.statusdatetime = DateTime.Today;

                        dbcontext.SaveChanges();

                        success = true;
                    }

                }

      
                if (LoggedInSignedAsReviewer)
                {

                    var result = dbcontext.filereviewersactivities.SingleOrDefault(b => b.revieweractivityid == ReviewerActivityID);
                    if (result != null)
                    {
                        result.approvalstatuscode = 1;    //1=signed
                        result.statusdatetime = DateTime.Today;

                        dbcontext.SaveChanges();

                        success = true;
                    }

                }



            }

            return success;


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