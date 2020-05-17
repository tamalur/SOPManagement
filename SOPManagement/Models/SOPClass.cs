using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop;
using System.Collections;
using Microsoft.SharePoint.Client;
using System.Security;

using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Web.Mvc;
using Microsoft.Online.SharePoint.TenantAdministration;
using System.Threading.Tasks;
using Xceed.Words.NET;
using Xceed.Document.NET;
using Paragraph = Xceed.Document.NET.Paragraph;
using Section = Xceed.Document.NET.Section;
using Table = Xceed.Document.NET.Table;
using System.Drawing;
using System.Runtime.InteropServices;

namespace SOPManagement.Models
{

    [Bind(Exclude = "Id")]

    public class SOPClass
    {

        public int FileID { get; set; }

        public string[] FilereviewersArr { get; set; }

        public string[] FileviewersArr { get; set; }

        [Display(Name = "All Users")]
        public bool AllUsersReadAcc { get; set; }

        public short? FileStatuscode { get; set; }

        [Required(ErrorMessage = "File Owner is Required")]
        [Display(Name = "Select SOP Owner")]
        public string FileOwnerEmail { get; set; }

        public int FileOwnerID { get; set; }

        public int FileChangeRqsterID { get; set; }

        public int FileChangeRqstID { get; set; }

        public int FileApproverID { get; set; }

        [Required(ErrorMessage = "File Approver is Required")]
        [Display(Name = "Select SOP Approver")]
        public string FileApproverEmail { get; set; }


        public string FileTitle { get; set; }   //title is without sopno

        [Display(Name = "SOP File Name")]
        public string FileName { get; set; }   //with sopno in front SOPNO + " "+ FileTitle

        public byte[] FileStream { get; set; }  //with sopno in front SOPNO + " "+ FileTitle

        [Display(Name = "Select Folder")]
        [Required(ErrorMessage = "Folder name is Required")]
        public string FolderName { get; set; }

        [Display(Name = "Select Subfolder")]
        public string SubFolderName { get; set; }

        [Display(Name = "By Department")]
        public string DepartmentName { get; set; }
        public short? DepartmentCode { get; set; }

        public string ViewAccessType { get; set; }

        public int ViewAccessTypeID { get; set; }


        public string SOPNo { get; set; }

        public FileRevision[] FileRevisions { get; set; }

        public string FileCurrVersion { get; set; }


        [Required(ErrorMessage = "Frequency Required")]
        [Display(Name = "Update Frequency")]
        public short Updatefreq { get; set; }

        [Display(Name = "Freq. Unit")]
        public string Updatefrequnit { get; set; }

        public short Updfrequnitcode { get; set; }

        public string FileLink { get; set; }

        public string FilePath { get; set; }

        public string FileLocalPath { get; set; }

        public bool OperationSuccess { get; set; }

        [Display(Name = "Reviewers")]

        public Employee[] FileReviewers { get; set; }

        [Display(Name = "By Users")]

        public Employee[] FileViewers { get; set; }

        public Employee FileOwner { get; set; }
        public Employee FileApprover { get; set; }

        public string DocumentLibName { get; set; }

        public string FileUrl { get; set; }
        public string SiteUrl { get; set; }
        [Display(Name = "SOP Effective Date")]
        public DateTime SOPEffectiveDate { get; set; }

        [Display(Name = "Select File to Upload")]
        //[System.Web.Mvc.Remote("CheckIfExists", "Home", ErrorMessage = "Duplicate File Found")]
        public HttpPostedFileBase UploadedFile { get; set; }

        public bool FileUploaded { get; set; }

        public string ApprovalStatus { get; set; }

        public string AuthorName { get; set; }

        public DateTime SOPCreateDate { get; set; }


        public string ErrorMessage { get; set; }

        public string UserName { get; set; }
        public string Password { get; set; }


        string userName = Utility.GetSiteAdminUserName();  //it is email address of site admin
        string password = Utility.GetSiteAdminPassowrd();

        // all database operation 

        public void UpdateChangeReqID(short statuscode)
        {

            using (var dbcontext = new RadiantSOPEntities())
            {
                var result = dbcontext.filechangerequestactivities.SingleOrDefault(b => b.changerequestid == FileChangeRqstID && b.fileid == FileID);
                if (result != null)
                {
                    result.approvalstatuscode = statuscode;
                    result.statusdatetime = DateTime.Today;

                    dbcontext.SaveChanges();
                }
            }
        }



        public void UpdatePrevOwnerStatus()
        {

            using (var dbcontext = new RadiantSOPEntities())
            {
                foreach (var result in dbcontext.fileowners.Where(b => b.fileid == FileID).ToList())
                {

                    result.ownerstatuscode = 2;    //change status 2 of previous owner of same file if a file is replaced during upload and create
                    result.statusdate = DateTime.Today;
                    result.statusbyuserid = Utility.GetLoggedInUserID();
                }
            
                dbcontext.SaveChanges();
                
            }

        }

        public void UpdatePrevApproverStatus()
        {

            using (var dbcontext = new RadiantSOPEntities())
            {
                foreach (var result in dbcontext.fileapprovers.Where(b => b.fileid == FileID).ToList())
                {

                    result.approverstatuscode = 2;    //change status of previous owner if a file is replaced during upload and create
                    result.statusdate = DateTime.Today;
                    result.statusbyuserid = Utility.GetLoggedInUserID();

                }

                dbcontext.SaveChanges();

            }


        }


        public void UpdatePrevReviewersStatus()
        {
            using (var dbcontext = new RadiantSOPEntities())
            {
                foreach (var result in dbcontext.filereviewers.Where(b => b.fileid == FileID).ToList())
                {

                    result.reviewerstatuscode = 2;    //change status of previous owner if a file is replaced during upload and create
                    result.statusdate = DateTime.Today;
                    result.statusbyuserid = Utility.GetLoggedInUserID();

                }

                dbcontext.SaveChanges();

            }



        }

        public void AddFileReviewers()
        {

            Employee emp = new Employee();

            int rvwrid;

            OperationSuccess = false;

            UpdatePrevReviewersStatus();    //update previous reviewers status with 2 if same file name exists



            foreach (Employee rvwr in FileReviewers)
            {
                emp.useremailaddress = rvwr.useremailaddress;
                emp.GetUserByEmail();
                rvwrid = emp.userid;

                using (var dbcontext = new RadiantSOPEntities())
                {

                    var rvwrtable = new filereviewer()
                    {
                        reviewerid = rvwrid,
                        fileid = FileID,
                        reviewerstatuscode=1,      // 1= Current
                        statusdate = DateTime.Today,
                        statusbyuserid=Utility.GetLoggedInUserID()


                    };
                    dbcontext.filereviewers.Add(rvwrtable);

                    dbcontext.SaveChanges();

                    rvwrid = rvwrtable.reviewid;

                    AddRvwractivities(rvwrid);

                    OperationSuccess = true;
                }

            }

            emp = null;


        }

        public void AddRvwractivities(int previewid)
        {
            using (var dbcontex = new RadiantSOPEntities())
            {
                var rvwractvts = new filereviewersactivity()
                {
                    changerequestid = FileChangeRqstID,
                    reviewid = previewid,
                    approvalstatuscode = 2   //not signed
                    // statusdatetime=DateTime.Today   //no sign no date
                };
                dbcontex.filereviewersactivities.Add(rvwractvts);
                dbcontex.SaveChanges();
            }
        }


        public void AddRvwractvtsWithChngRqst()
        {

            using (var dbctx = new RadiantSOPEntities())
            {
                //only current active reviewers
                var rvrwrs = (from c in dbctx.filereviewers where c.fileid == FileID && c.reviewerstatuscode==1 select c);

                foreach (var rvwr in rvrwrs)
                {
                    AddRvwractivities(rvwr.reviewid);

                }

            }



        }


        public void AddApproveractivities(int papproveid)
        {
            using (var dbcontext = new RadiantSOPEntities())
            {
                var apprvractvs = new fileapproversactivity()
                {
                    changerequestid = FileChangeRqstID,
                    approveid = papproveid,
                    approvalstatuscode = 2   //not signed
                    //statusdatetime = DateTime.Today    // no date should be assigned
                };
                dbcontext.fileapproversactivities.Add(apprvractvs);
                dbcontext.SaveChanges();
            }
        }

        public void AddPublisheractivities(int ppublisherid)
        {
            using (var dbcontext = new RadiantSOPEntities())
            {
                var pblsracvts = new filepublishersactivity()
                {
                    changerequestid = FileChangeRqstID,
                    publisherid = ppublisherid,
                    approvalstatuscode = 8  //8=not approved
                    //statusdatetime= DateTime.Today
                };

                dbcontext.filepublishersactivities.Add(pblsracvts);
                dbcontext.SaveChanges();
            }
        }

        public void AddOwneractivities(int pownershipid)
        {
            using (var dbcontext = new RadiantSOPEntities())
            {
                var owneractvts = new fileownersactivity()
                {
                    changerequestid = FileChangeRqstID,   //got this value during creating change request
                    ownershipid = pownershipid,
                    approvalstatuscode = 2  //not signed
                                            //  statusdatetime = DateTime.Today   //no sign no date

                };
                dbcontext.fileownersactivities.Add(owneractvts);
                dbcontext.SaveChanges();
            }

        }

        public void AddFileApprover()

        {
            //now insert approver into approver table
            Employee emp = new Employee();
            int apprvrid;
            int apprvid;
            OperationSuccess = false;

            emp.useremailaddress = FileApproverEmail;
            emp.GetUserByEmail();
            apprvrid = emp.userid;

            // update prev approver with status code 2 for file replacement
            UpdatePrevApproverStatus();


            using (var dbcontext = new RadiantSOPEntities())
            {

                var aprvrtable = new fileapprover()
                {
                    approverid = apprvrid,
                    fileid = FileID,
                    approverstatuscode = 1,   //current
                    statusdate = DateTime.Today,
                    statusbyuserid = Utility.GetLoggedInUserID()

                };
                dbcontext.fileapprovers.Add(aprvrtable);

                dbcontext.SaveChanges();

                apprvid = aprvrtable.approveid;

                AddApproveractivities(apprvid);

                AddPublisheractivities(apprvrid);   //here publisher id

                OperationSuccess = true;

            }


        }

        public void AddViewerAccessType()
        {
            using (var dbcontext = new RadiantSOPEntities())
            {

                //get old viewaccess records with the FIleID so we can add new viewers access 
                var vwacctype = dbcontext.fileviewaccesstypes.SingleOrDefault(x => x.fileid == FileID);

                if (vwacctype != null)
                {
                    ViewAccessTypeID = vwacctype.viewaccessid;

                    dbcontext.fileviewaccesstypes.Remove(vwacctype);

                    //    dbcontext.SaveChanges();

                    //delete related viewers if any
                    dbcontext.fileviewers.RemoveRange(dbcontext.fileviewers.Where(x => x.viewaccessid == ViewAccessTypeID));

                    dbcontext.SaveChanges();
                }
                //now add new access type and related records

                var fileviewacctype = new fileviewaccesstype()
                {
                    fileid = FileID,
                    viewtypename = ViewAccessType,
                    // departmentname=DepartmentName
                    departmentcode = DepartmentCode

                };
                dbcontext.fileviewaccesstypes.Add(fileviewacctype);
                dbcontext.SaveChanges();

                ViewAccessTypeID = fileviewacctype.viewaccessid;

                OperationSuccess = true;
            }

        }


        public void AddFileViewers()
        {

            Employee emp = new Employee();

            int vwrid;

            OperationSuccess = false;

            //delete all viewers with the file id that was deleted from ViewAccessType table


            foreach (Employee vwr in FileViewers)
            {
                emp.useremailaddress = vwr.useremailaddress;
                emp.GetUserByEmail();
                vwrid = emp.userid;

                using (var dbcontext = new RadiantSOPEntities())
                {

                    var vwrtable = new fileviewer
                    {
                        viewerid = vwrid,
                        viewaccessid = ViewAccessTypeID


                    };
                    dbcontext.fileviewers.Add(vwrtable);

                    dbcontext.SaveChanges();

                    OperationSuccess = true;
                }

            }

            emp = null;


        }

        public void AddFileOwner()

        {

            //now insert file owner into owner table

            Employee emp = new Employee();
            int ownerid;
            int ownershipid;
            OperationSuccess = false;

            emp.useremailaddress = FileOwnerEmail;
            emp.GetUserByEmail();
            ownerid = emp.userid;


            //first update previous owner if any. During upload or new creation, it happens when a file is replaced with same file name 
            //that was created before 

            UpdatePrevOwnerStatus();
                       

            using (var dbcontext = new RadiantSOPEntities())
            {

                var ownertable = new fileowner()
                {
                    ownerid = ownerid,
                    fileid = FileID,
                    ownerstatuscode = 1,  //1=current
                    statusdate = DateTime.Today,
                    statusbyuserid = Utility.GetLoggedInUserID()

                };
                dbcontext.fileowners.Add(ownertable);

                dbcontext.SaveChanges();
                ownershipid = ownertable.ownershipid;

                AddOwneractivities(ownershipid);

                OperationSuccess = true;
            }




        }

        public void AddChangeRequest()
        {



            using (var dbcontex = new RadiantSOPEntities())
            {

                var chngtable = new filechangerequestactivity()
                {
                    fileid = FileID,
                    approvalstatuscode = 2,   //2=not signed
                    statusdatetime = DateTime.Today,
                    requesterid = FileChangeRqsterID
                };
                dbcontex.filechangerequestactivities.Add(chngtable);
                dbcontex.SaveChanges();

                FileChangeRqstID = chngtable.changerequestid;

                OperationSuccess = true;
            }
        }

        public void AddUpdateFreq()

        {
            //now insert update frequence
            OperationSuccess = false;
            DateTime freqschdl = DateTime.Now;

            //switch (Updatefrequnit.Trim().ToLower())
            //{
            //    case "yearly":
            //        freqschdl = freqschdl.AddYears(Updatefreq);
            //        break;
            //    case "monthly":
            //        freqschdl = freqschdl.AddMonths(Updatefreq);
            //        break;
            //    case "weekly":
            //        freqschdl = freqschdl.AddDays(Updatefreq);
            //        break;

            //}


            switch (Updfrequnitcode)
            {
                case 1:    //yearly
                    freqschdl = freqschdl.AddYears(Updatefreq);
                    break;
                case 2:    //monthly
                    freqschdl = freqschdl.AddMonths(Updatefreq);
                    break;
                case 3:    //weekly
                    freqschdl = freqschdl.AddDays(Updatefreq);
                    break;

            }


            using (var dbcontext = new RadiantSOPEntities())
            {

                var updfreqtable = new fileupdateschedule()
                {
                    fileid = FileID,
                    frequencyofrevision = Updatefreq,
                    unitoffrequency = Updatefrequnit,
                    unitcodeupdfreq = Updfrequnitcode,
                    lastrevisionno = "1.0",
                    scheduledatetime = freqschdl


                };
                dbcontext.fileupdateschedules.Add(updfreqtable);

                dbcontext.SaveChanges();

                OperationSuccess = true;

            }


        }

        public void UpdateCoverRevhistPageDocX()
        {



            string tmpfiledirnm = Utility.GetTempLocalDirPath();

            string savePath = HttpContext.Current.Server.MapPath(tmpfiledirnm + FileName);

            object path = savePath;

           

            try
            {


                var wdoc = DocX.Load(savePath);

                

                //  add row in table and data in cell

                Employee emp = new Employee();
                string ownerfullname;
                string approverfullname;
                string approvertitle;


                int totalTables = wdoc.Tables.Count;



                //Add data into reviewer table  - 2nd table in the cover page

                if (totalTables > 0)
                {


                    //update 1st table in cover page, file title, Owner, SOP #, Rev #, Eff date, owner

                    emp.useremailaddress = FileOwnerEmail;
                    emp.GetUserByEmail();
                    ownerfullname = emp.userfullname;


                    //first table tab1 with SOP basic info

                    var tab1 = wdoc.Tables[0];

                    // Select the last row as source row.
                    int selectedRow1 = tab1.Rows.Count;

                    tab1.Rows[0].Cells[1].Paragraphs[0].Remove(false);    //track changes false
                    tab1.Rows[0].Cells[1].Paragraphs[0].Append(FileTitle);

                    tab1.Rows[1].Cells[1].Paragraphs[0].Remove(false);    //track changes false
                    tab1.Rows[1].Cells[1].Paragraphs[0].Append(SOPNo);

                    tab1.Rows[1].Cells[3].Paragraphs[0].Remove(false);    //track changes false
                    tab1.Rows[1].Cells[3].Paragraphs[0].Append(FileCurrVersion);

                    tab1.Rows[2].Cells[1].Paragraphs[0].Remove(false);    //track changes false
                    tab1.Rows[2].Cells[1].Paragraphs[0].Append("");  //reset effective for new file 


                    tab1.Rows[3].Cells[1].Paragraphs[0].Remove(false);    //track changes false
                    tab1.Rows[3].Cells[1].Paragraphs[0].Append(ownerfullname);


                    //2nd table tab2 to add/update reviewers

                    var tab2 = wdoc.Tables[1];

                    int tab2rows = tab2.RowCount - 1;  //exclude first two title row

                    //first remove existing rows except header row

                    int totrvwrs = FileReviewers.Count();

                    int totdiff = tab2rows - totrvwrs;

                    int startindx;

                    //prepare reviewer table first so total reviewers and total table rows become equal

                    if (totdiff > 0)  //remove extar rows when reviewers are less
                    {
                        startindx = tab2.RowCount - totdiff;

                        int r= startindx;

                        
                        for (int i = startindx; i <= tab2rows; i++)
                        {
                            if (r==startindx)

                                 tab2.Rows[i].Remove();
                            else
                                tab2.Rows[i-1].Remove();

                            r = r+1;
                        }
                    }

                    int rowtoadd = Math.Abs(totdiff);  //get positive to loop

                    if (totdiff < 0)  //add rows as reviewers are more than table rows
                    {

                        for (int i = 1; i <= rowtoadd; i++)  //how many rows to add    
                        {
                            tab2.InsertRow();
                            
                            
                        }
                    }

                    //if rows in table and total reviewers are same don't need to do anything

                    //now add reviewers to table as we have equal rows to reviewers

                    int totnewrow = tab2.Rows.Count();  //get updated rows from table
                    int rvwrno = 0;  //for start index of reviewer

                    Xceed.Document.NET.Cell cel1;
                    Xceed.Document.NET.Cell celend;

                    if ((totnewrow - 1) == totrvwrs)
                    {

                        for (int i = 1; i <= totnewrow - 1; i++)   //exclude row[0] for header
                        {

                            emp.useremailaddress = FileReviewers[rvwrno].useremailaddress;
                            emp.GetUserByEmail();

                            tab2.Rows[i].Cells[0].Paragraphs[0].Remove(false);
                            cel1 = tab2.Rows[i].Cells[0];
                            cel1.SetBorder(TableCellBorderType.Left, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));
                            cel1.Paragraphs[0].Append(emp.userfullname);

                            //tab2.Rows[i].Cells[0].Paragraphs[0].Append(emp.userfullname);

                            tab2.Rows[i].Cells[1].Paragraphs[0].Remove(false);
                            tab2.Rows[i].Cells[1].Paragraphs[0].Append(emp.userjobtitle);

                            //resent signature status
                            tab2.Rows[i].Cells[2].Paragraphs[0].Remove(false);
                            tab2.Rows[i].Cells[2].Paragraphs[0].Append("");


                            celend = tab2.Rows[i].Cells[3];
                            celend.SetBorder(TableCellBorderType.Right, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));

                            celend.Paragraphs[0].Remove(false);
                            celend.Paragraphs[0].Append("");   //reset signature date

                            rvwrno = rvwrno + 1;

                        }
                    }
                    //end updating reviewers table

                    // table tab3 to update approver row 

                    emp.useremailaddress = FileApproverEmail;
                    emp.GetUserByEmail();

                    approverfullname = emp.userfullname;
                    approvertitle = emp.userjobtitle;

                    var tab3 = wdoc.Tables[2];

                    // Select the last row as source row.
                    int aprvrRow = tab3.Rows.Count;

                    cel1 = tab3.Rows[1].Cells[0];
                    cel1.SetBorder(TableCellBorderType.Left, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));

                    cel1.Paragraphs[0].Remove(false);    //track changes false
                    cel1.Paragraphs[0].Append(approverfullname);

                    tab3.Rows[1].Cells[1].Paragraphs[0].Remove(false);    //track changes false
                    tab3.Rows[1].Cells[1].Paragraphs[0].Append(approvertitle);

                    tab3.Rows[1].Cells[2].Paragraphs[0].Remove(false);    //track changes false
                    tab3.Rows[1].Cells[2].Paragraphs[0].Append("");  //reset approver signature status

                    celend = tab3.Rows[1].Cells[3];
                    celend.SetBorder(TableCellBorderType.Right, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));

                    celend.Paragraphs[0].Remove(false);
                    celend.Paragraphs[0].Append("");  //reset signature date

                    //start now add/update of revision history table

                    int totrh = 0;

                    if (FileRevisions != null)
                        totrh = FileRevisions.Count();

                    if (totrh>0)
                    {

                        //first remove existing rows except header row

                        int revtabindx = totalTables - 1;

                        var tabrh = wdoc.Tables[revtabindx];

                        int tabrhrows = tabrh.RowCount - 1;  //exclude first one title row


                        int totdiffrh = tabrhrows - totrh;

                        int startindxrh;

                        //prepare revision table first so total revisions and total table rows become equal

                        if (totdiffrh > 0)  //remove extar rows when reviewers are less
                        {
                            startindxrh = tabrh.RowCount - totdiffrh;

                            int r = startindxrh;


                            for (int i = startindxrh; i <= tabrhrows; i++)
                            {
                                if (r == startindxrh)

                                    tabrh.Rows[i].Remove();
                                else
                                    tab2.Rows[i - 1].Remove();

                                r = r + 1;
                            }
                        }

                        int rowtoaddrh = Math.Abs(totdiffrh);  //get positive to loop

                        if (totdiffrh < 0)  //add rows as reh hist are more than table rows
                        {

                            for (int i = 1; i <= rowtoaddrh; i++)  //how many rows to add    
                            {
                                tabrh.InsertRow();


                            }
                        }

                        //if rows in table and total reviewers are same don't need to do anything

                        //now add revisions to table as we have equal rows to revisions

                        int totnewrowrh = tabrh.Rows.Count();  //get updated rows from table
                        int rhno = 0;  //for start index of reviewer

                        if ((totnewrowrh - 1) == totrh)
                        {

                            for (int i = 1; i <= totnewrowrh - 1; i++)   //exclude row[0] for header
                            {


                                cel1 = tabrh.Rows[i].Cells[0];
                                cel1.SetBorder(TableCellBorderType.Left, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));

                                cel1.Paragraphs[0].Remove(false);
                                cel1.Paragraphs[0].Append(FileRevisions[rhno].RevisionNo);

                                tabrh.Rows[i].Cells[1].Paragraphs[0].Remove(false);
                                tabrh.Rows[i].Cells[1].Paragraphs[0].Append(FileRevisions[rhno].RevisionDate.ToString("MMMM dd, yyyy"));

                                celend = tabrh.Rows[i].Cells[2];
                                celend.SetBorder(TableCellBorderType.Right, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));

                                rhno = rhno + 1;

                            }
                        }
                        //end updating revision table



                    } //end checking total revisons

                    //add footer

                    wdoc.AddFooters();


                    // Get the default Footer for this document.
                    Footer footer_default = wdoc.Footers.Odd;

                    Table tabfooter = footer_default.InsertTable(1, 2);

                    //tabfooter.AutoFit = AutoFit.Contents;

                    tabfooter.SetBorder(TableBorderType.Bottom, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));
                    tabfooter.SetBorder(TableBorderType.Top, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));
                    tabfooter.SetBorder(TableBorderType.Left, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));
                    tabfooter.SetBorder(TableBorderType.Right, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));
                    tabfooter.SetBorder(TableBorderType.InsideH, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));
                    tabfooter.SetBorder(TableBorderType.InsideV, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));

                    //footer_default.Tables[0].Rows[0].MergeCells(1, 3);

                    //tabfooter

                    tabfooter.Rows[0].Cells[0].Paragraphs[0].Append(FileTitle);
                    tabfooter.Rows[0].Cells[1].Paragraphs[0].Append("Page ").AppendPageNumber(PageNumberFormat.normal).Append(" of ").AppendPageCount(PageNumberFormat.normal);


                   // tabfooter.SetColumnWidth(0, 100);

                    tabfooter.SetWidthsPercentage(new[] { 50f, 50f }, 800);   //array is percent of each column width, 800 is total table with

                    //footer_default.Tables[0].Rows[0].Cells[0].Paragraphs[0].Append(FileName);
                    // footer_default.Tables[0].Rows[0].Cells[4].Paragraphs[0].Append("Page ").AppendPageNumber(PageNumberFormat.normal).Append(" of ").AppendPageCount(PageNumberFormat.normal);




                    // Insert a Paragraph into the default Footer.


                    //Paragraph p3 = footer_default.InsertParagraph();
                    // p3.Append(FileName).Direction = Direction.LeftToRight;



                    //  Paragraph p4 = footer_default..InsertParagraph();
                    //  p4.Alignment = Alignment.right;

                    //p3.Append("Page ").AppendPageNumber(PageNumberFormat.normal).Alignment=Alignment.right;
                    //p3.Append(" Of ").AppendPageCount(PageNumberFormat.normal).Alignment=Alignment.right;




                    wdoc.Save();

                    wdoc = null;

                }  //if totalTables





            }
    
            catch (Exception ex)
            {
                // ErrorMessage = ex.Message;
                throw ex;

            }

            finally
            {
                //app.Application.Quit();
                

            }





        }
        public void UpdateCoverRevhistPageDocX(bool pUpdSignatureRev)
        {


            //  string newfilename = SOPNo + " " + SOPFileTitle;


            //string savePath = HttpContext.Current.Server.MapPath("~/Content/docfiles/" + FileName);

            string tmpfiledirnm = Utility.GetTempLocalDirPath();

            //string savePath = HttpContext.Current.Server.MapPath("~/Content/docfiles/" + FileName);

            string savePath = HttpContext.Current.Server.MapPath(tmpfiledirnm + FileName);

            object path = savePath;



            try
            {


                var wdoc = DocX.Load(savePath);

                //  add row in table and data in cell


                int totalTables = wdoc.Tables.Count;


                if (totalTables > 0)
                {


                    //update 1st table in cover page, file title, Owner, SOP #, Rev #, Eff date, owner


                    var tab1 = wdoc.Tables[0];

                    // Select the last row as source row.
                    int selectedRow1 = tab1.Rows.Count;

                    tab1.Rows[0].Cells[1].Paragraphs[0].Remove(false);    //track changes false
                    tab1.Rows[0].Cells[1].Paragraphs[0].Append(FileTitle);

                    tab1.Rows[1].Cells[1].Paragraphs[0].Remove(false);    //track changes false
                    tab1.Rows[1].Cells[1].Paragraphs[0].Append(SOPNo);

                    tab1.Rows[2].Cells[1].Paragraphs[0].Remove(false);    //track changes false
                    tab1.Rows[2].Cells[1].Paragraphs[0].Append(DateTime.Today.ToString("MMMM dd, yyyy"));  //publish today


                    tab1.Rows[1].Cells[3].Paragraphs[0].Remove(false);    //track changes false
                    tab1.Rows[1].Cells[3].Paragraphs[0].Append(FileCurrVersion);

                    tab1.Rows[3].Cells[1].Paragraphs[0].Remove(false);    //track changes false
                    tab1.Rows[3].Cells[1].Paragraphs[0].Append(FileOwner.userfullname);


                    //2nd table tab2 to add/update reviewers

                    var tab2 = wdoc.Tables[1];

                    int tab2rows = tab2.RowCount - 1;  //exclude first two title row

                    //first remove existing rows except header row

                    int totrvwrs = FileReviewers.Count();

                    int totdiff = tab2rows - totrvwrs;

                    int startindx;

                    //prepare reviewer table first so total reviewers and total table rows become equal

                    if (totdiff > 0)  //remove extar rows when reviewers are less
                    {
                        startindx = tab2.RowCount - totdiff;

                        int r = startindx;


                        for (int i = startindx; i <= tab2rows; i++)
                        {
                            if (r == startindx)

                                tab2.Rows[i].Remove();
                            else
                                tab2.Rows[i - 1].Remove();

                            r = r + 1;
                        }
                    }

                    int rowtoadd = Math.Abs(totdiff);  //get positive to loop

                    if (totdiff < 0)  //add rows as reviewers are more than table rows
                    {

                        for (int i = 1; i <= rowtoadd; i++)  //how many rows to add    
                        {
                            tab2.InsertRow();


                        }
                    }

                    //if rows in table and total reviewers are same don't need to do anything

                    //now add reviewers to table as we have equal rows to reviewers

                    int totnewrow = tab2.Rows.Count();  //get updated rows from table
                    int rvwrno = 0;  //for start index of reviewer

                    Xceed.Document.NET.Cell cel1;
                    Xceed.Document.NET.Cell celend;

                    if ((totnewrow - 1) == totrvwrs)
                    {

                        for (int i = 1; i <= totnewrow - 1; i++)   //exclude row[0] for header
                        {

                            tab2.Rows[i].Cells[0].Paragraphs[0].Remove(false);
                            cel1 = tab2.Rows[i].Cells[0];
                            cel1.SetBorder(TableCellBorderType.Left, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));
                            cel1.Paragraphs[0].Append(FileReviewers[rvwrno].userfullname);

                            tab2.Rows[i].Cells[1].Paragraphs[0].Remove(false);
                            tab2.Rows[i].Cells[1].Paragraphs[0].Append(FileReviewers[rvwrno].userjobtitle);

                            tab2.Rows[i].Cells[2].Paragraphs[0].Remove(false);
                            tab2.Rows[i].Cells[2].Paragraphs[0].Append(FileReviewers[rvwrno].signstatus);


                            celend = tab2.Rows[i].Cells[3];
                            celend.SetBorder(TableCellBorderType.Right, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));

                            celend.Paragraphs[0].Remove(false);
                            celend.Paragraphs[0].Append(FileReviewers[rvwrno].signaturedate.ToString("MMMM dd, yyyy"));



                            rvwrno = rvwrno + 1;

                        }
                    }
                    //end updating reviewers table

                    // table tab3 to update approver row 


                    var tab3 = wdoc.Tables[2];

                    // Select the last row as source row.
                    int aprvrRow = tab3.Rows.Count;

                    cel1 = tab3.Rows[1].Cells[0];
                    cel1.SetBorder(TableCellBorderType.Left, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));

                    cel1.Paragraphs[0].Remove(false);    //track changes false
                    cel1.Paragraphs[0].Append(FileApprover.userfullname);

                    tab3.Rows[1].Cells[1].Paragraphs[0].Remove(false);    //track changes false
                    tab3.Rows[1].Cells[1].Paragraphs[0].Append(FileApprover.userjobtitle);

                    tab3.Rows[1].Cells[2].Paragraphs[0].Remove(false);    //track changes false
                    tab3.Rows[1].Cells[2].Paragraphs[0].Append(FileApprover.signstatus);


                    celend = tab3.Rows[1].Cells[3];
                    celend.SetBorder(TableCellBorderType.Right, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));

                    celend.Paragraphs[0].Remove(false);    //track changes false
                    celend.Paragraphs[0].Append(FileApprover.signaturedate.ToString("MMMM dd, yyyy"));



                    //start now add/update of revision history table

                    int totrh = 0;

                    if (FileRevisions != null)
                        totrh = FileRevisions.Count();

                    if (totrh > 0)
                    {

                        //first remove existing rows except header row

                        int revtabindx = totalTables - 1;

                        var tabrh = wdoc.Tables[revtabindx];

                        int tabrhrows = tabrh.RowCount - 1;  //exclude first one title row


                        int totdiffrh = tabrhrows - totrh;

                        int startindxrh;

                        //prepare revision table first so total revisions and total table rows become equal

                        if (totdiffrh > 0)  //remove extar rows when reviewers are less
                        {
                            startindxrh = tabrh.RowCount - totdiffrh;

                            int r = startindxrh;


                            for (int i = startindxrh; i <= tabrhrows; i++)
                            {
                                if (r == startindxrh)

                                    tabrh.Rows[i].Remove();
                                else
                                    tab2.Rows[i - 1].Remove();

                                r = r + 1;
                            }
                        }

                        int rowtoaddrh = Math.Abs(totdiffrh);  //get positive to loop

                        if (totdiffrh < 0)  //add rows as reh hist are more than table rows
                        {

                            for (int i = 1; i <= rowtoaddrh; i++)  //how many rows to add    
                            {
                                tabrh.InsertRow();


                            }
                        }

                        //if rows in table and total reviewers are same don't need to do anything

                        //now add revisions to table as we have equal rows to revisions

                        int totnewrowrh = tabrh.Rows.Count();  //get updated rows from table
                        int rhno = 0;  //for start index of reviewer

                        if ((totnewrowrh - 1) == totrh)
                        {

                            for (int i = 1; i <= totnewrowrh - 1; i++)   //exclude row[0] for header
                            {


                                cel1 = tabrh.Rows[i].Cells[0];
                                cel1.SetBorder(TableCellBorderType.Left, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));

                                cel1.Paragraphs[0].Remove(false);
                                cel1.Paragraphs[0].Append(FileRevisions[rhno].RevisionNo);

                                tabrh.Rows[i].Cells[1].Paragraphs[0].Remove(false);
                                tabrh.Rows[i].Cells[1].Paragraphs[0].Append(FileRevisions[rhno].RevisionDate.ToString("MMMM dd, yyyy"));

                                celend = tabrh.Rows[i].Cells[2];
                                celend.SetBorder(TableCellBorderType.Right, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));

                              //  celend.Paragraphs[0].Remove(false);
                              //  celend.Paragraphs[0].Append(FileRevisions[rhno].Description);


                                rhno = rhno + 1;

                            }
                        }
                        //end updating revision table



                    } //end checking total revisons

                    //add footer

                    wdoc.AddFooters();


                    // Get the default Footer for this document.
                    Footer footer_default = wdoc.Footers.Odd;

                    Table tabfooter = footer_default.InsertTable(1, 2);

                    //tabfooter.AutoFit = AutoFit.Contents;

                    tabfooter.SetBorder(TableBorderType.Bottom, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));
                    tabfooter.SetBorder(TableBorderType.Top, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));
                    tabfooter.SetBorder(TableBorderType.Left, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));
                    tabfooter.SetBorder(TableBorderType.Right, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));
                    tabfooter.SetBorder(TableBorderType.InsideH, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));
                    tabfooter.SetBorder(TableBorderType.InsideV, new Xceed.Document.NET.Border(BorderStyle.Tcbs_none, BorderSize.one, 1, Color.Transparent));

                    //footer_default.Tables[0].Rows[0].MergeCells(1, 3);

                    //tabfooter

                    tabfooter.Rows[0].Cells[0].Paragraphs[0].Append(FileTitle);
                    tabfooter.Rows[0].Cells[1].Paragraphs[0].Append("Page ").AppendPageNumber(PageNumberFormat.normal).Append(" of ").AppendPageCount(PageNumberFormat.normal);


                    // tabfooter.SetColumnWidth(0, 100);

                    tabfooter.SetWidthsPercentage(new[] { 50f, 50f }, 800);   //array is percent of each column width, 500 is total table with

                    //footer_default.Tables[0].Rows[0].Cells[0].Paragraphs[0].Append(FileName);
                    // footer_default.Tables[0].Rows[0].Cells[4].Paragraphs[0].Append("Page ").AppendPageNumber(PageNumberFormat.normal).Append(" of ").AppendPageCount(PageNumberFormat.normal);




                    // Insert a Paragraph into the default Footer.


                    //Paragraph p3 = footer_default.InsertParagraph();
                    // p3.Append(FileName).Direction = Direction.LeftToRight;



                    //  Paragraph p4 = footer_default..InsertParagraph();
                    //  p4.Alignment = Alignment.right;

                    //p3.Append("Page ").AppendPageNumber(PageNumberFormat.normal).Alignment=Alignment.right;
                    //p3.Append(" Of ").AppendPageCount(PageNumberFormat.normal).Alignment=Alignment.right;




                    wdoc.Save();


                    wdoc = null;

                    

                }  //if totalTables





            }

            catch (Exception ex)
            {
                ErrorMessage = ex.Message;


            }

            finally
            {
                //app.Application.Quit();


            }





        }
        public void UpdateCoverRevhistPage()
        {



            Logger oLogger = new Logger();

            oLogger.LogFileName = HttpContext.Current.Server.MapPath(Utility.GetLogFilePath());


            string tmpfiledirnm = Utility.GetTempLocalDirPath();

            //string savePath = HttpContext.Current.Server.MapPath("~/Content/docfiles/" + FileName);

            string savePath = HttpContext.Current.Server.MapPath(tmpfiledirnm + FileName);

            object missObj = System.Reflection.Missing.Value;
            object path = savePath;
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

          //  app.DisplayAlerts =WdAlertLevel.wdAlertsNone;

            // Microsoft.Office.Interop.Word.ApplicationClass app = new ApplicationClass();

            Microsoft.Office.Interop.Word.Document wdoc= app.Documents.Open(ref path, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj);

            if (wdoc == null)
                oLogger.UpdateLogFile(DateTime.Now.ToString() + "UpdateCoverRevhistPage:wdoc is null");



            try
            {


              //  System.IO.File.Copy(HttpContext.Current.Server.MapPath("~/Content/docfiles/SOPTemp.docx"), HttpContext.Current.Server.MapPath("~/Content/docfiles/" + FileName), true);
                
               
                wdoc.TrackRevisions =false;

                //  add row in table and data in cell

                Employee emp = new Employee();
                string ownerfullname;
                string approverfullname;
                string approvertitle;

                emp.useremailaddress = FileOwnerEmail;
                emp.GetUserByEmail();
                ownerfullname = emp.userfullname;

                emp.useremailaddress = FileApproverEmail;
                emp.GetUserByEmail();
                approverfullname = emp.userfullname;
                approvertitle = emp.userjobtitle;

                int totalTables = wdoc.Tables.Count;
                bool donotaddrow = false;


                //Add data into reviewer table  - 2nd table in the cover page

                if (totalTables > 0)
                {


                    //update 1st table in cover page, file title, SOP #, Rev #, Eff date, owner

                    Microsoft.Office.Interop.Word.Table tab1 = wdoc.Tables[1];
                    Microsoft.Office.Interop.Word.Range range1 = tab1.Range;

                    // Select the last row as source row.
                    int selectedRow1 = tab1.Rows.Count;

                    // Write new vaules to each cell.
                    tab1.Rows[1].Cells[2].Range.Text = FileTitle;
                    tab1.Rows[2].Cells[2].Range.Text = SOPNo;
                    tab1.Rows[2].Cells[4].Range.Text = FileCurrVersion;
                    tab1.Rows[3].Cells[2].Range.Text = "";   //for new file it will be updated during publishing
                    tab1.Rows[4].Cells[2].Range.Text = ownerfullname;


                    //update 2nd table in  cover page for updating reviewers

                    Microsoft.Office.Interop.Word.Table tab2 = wdoc.Tables[2];
                    Microsoft.Office.Interop.Word.Range range2 = tab2.Range;

                    // Select the last row as source row.
                    int selectedRow2 = tab2.Rows.Count;

                    //keep only 2 rows if there are more than 2 rows in table
                    //int rvrrowcount = Reviewers.Count();

                    int rowstodel;
                    if (selectedRow2 > 2)
                    {
                        rowstodel = selectedRow2 - 2;
                        for (int i = 1; i <= rowstodel; i++)
                        {
                            tab2.Rows[3].Delete();

                        }
                        selectedRow2 = tab2.Rows.Count;
                    }



                    // Select and copy content of the source row.
                    range2.Start = tab2.Rows[selectedRow2].Cells[1].Range.Start;
                    range2.End = tab2.Rows[selectedRow2].Cells[tab2.Rows[selectedRow2].Cells.Count].Range.End;
                    range2.Copy();

                    // Insert a new row after the last row if it is not first row to add data

                    int rvwrcnt = 1;
                    foreach (Employee rvwr in FileReviewers)
                    {

                        if (selectedRow2 == 2 && rvwrcnt == 1)
                        {

                            //if (tab2.Rows[tab2.Rows.Count].Cells[1].Range.Text == "" || tab2.Rows[tab2.Rows.Count].Cells[1].Range.Text == "\r\a")
                            donotaddrow = true;

                        }
                        else
                            donotaddrow = false;

                        if (!donotaddrow)
                            tab2.Rows.Add(ref missObj);


                        // Moves the cursor to the first cell of target row.
                        range2.Start = tab2.Rows[tab2.Rows.Count].Cells[1].Range.Start;
                        range2.End = range2.Start;

                        // Paste values to target row.
                        range2.Paste();

                        // Write new vaules to each cell.

                        emp.useremailaddress = rvwr.useremailaddress;
                        emp.GetUserByEmail();

                        tab2.Rows[tab2.Rows.Count].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        tab2.Rows[tab2.Rows.Count].Cells[1].Range.Text = rvwr.userfullname;

                        tab2.Rows[tab2.Rows.Count].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        tab2.Rows[tab2.Rows.Count].Cells[2].Range.Text = emp.userjobtitle;

                        tab2.Rows[tab2.Rows.Count].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        tab2.Rows[tab2.Rows.Count].Cells[3].Range.Text = "";

                        tab2.Rows[tab2.Rows.Count].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        tab2.Rows[tab2.Rows.Count].Cells[4].Range.Text = "";


                        rvwrcnt = rvwrcnt + 1;


                    }

                    //end updating 2nd reviewers table

                    //update 3rd table for approver

                    //update 1st table in cover page, file title, SOP #, Rev #, Eff date, owner

                    Microsoft.Office.Interop.Word.Table tab3 = wdoc.Tables[3];
                    Microsoft.Office.Interop.Word.Range range3 = tab3.Range;

                    // Write new vaules to each cell of row 3. One row always as there will be one approver
                    tab3.Rows[2].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tab3.Rows[2].Cells[1].Range.Text = approverfullname;

                    tab3.Rows[2].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tab3.Rows[2].Cells[2].Range.Text = approvertitle;

                    tab3.Rows[2].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tab3.Rows[2].Cells[3].Range.Text = "";  //reset approver signature

                    tab3.Rows[2].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tab3.Rows[2].Cells[4].Range.Text = "";  //reset signdate

                    //end updating 3rd table for approver

                    //update last table to add data into Revison history table 

                    Microsoft.Office.Interop.Word.Table tab = wdoc.Tables[totalTables];
                    Microsoft.Office.Interop.Word.Range range = tab.Range;

                    // Select the last row as source row.
                    int selectedRow = tab.Rows.Count;


                    //we don't need revision history for new upload

                    //delete rows if the table has more than three rows 

                    //if (selectedRow > 2)
                    //{
                    //    rowstodel = selectedRow - 2;
                    //    for (int i = 1; i <= rowstodel; i++)
                    //    {
                    //        tab.Rows[3].Delete();

                    //    }
                    //    selectedRow = tab.Rows.Count;
                    //}

                    //// Select and copy content of the source row.
                    //range.Start = tab.Rows[selectedRow].Cells[1].Range.Start;
                    //range.End = tab.Rows[selectedRow].Cells[tab.Rows[selectedRow].Cells.Count].Range.End;
                    //range.Copy();


                    //donotaddrow = false;

                    //int filevercount = 1;

                    //if (FileRevisions!=null)
                    //     foreach (FileRevision rev in FileRevisions)
                    //     {

                    //    if (selectedRow == 2 && filevercount == 1)
                    //    {
                    //        donotaddrow = true;
                    //    }

                    //    else
                    //        donotaddrow = false;

                    //    if (!donotaddrow)
                    //        tab.Rows.Add(ref missObj);

                    //    // Moves the cursor to the first cell of target row.
                    //    range.Start = tab.Rows[tab.Rows.Count].Cells[1].Range.Start;
                    //    range.End = range.Start;

                    //// Paste values to target row.
                    //    range.Paste();


                    //    // Write new vaules to each cell.
                    //    tab.Rows[tab.Rows.Count].Cells[1].Range.Text = rev.RevisionNo;
                    //    tab.Rows[tab.Rows.Count].Cells[2].Range.Text = rev.RevisionDate.ToString("M/d/yyyy");
                    //    tab.Rows[tab.Rows.Count].Cells[3].Range.Text = rev.Description;

                    //    filevercount = filevercount + 1;


                }


                // Set footers
                foreach (Microsoft.Office.Interop.Word.Section wordSection in wdoc.Sections)
                {
                    Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                    footerRange.Tables[1].Cell(1, 1).Range.Text = FileTitle;

                }

                wdoc.TrackRevisions = true;

                wdoc.SaveAs2(savePath);   //save in actual file from tempalte


                //wdoc.Close();


                //    oLogger.UpdateLogFile(DateTime.Now.ToString() + ":Successfully updated cover page");

            }

            catch (Exception ex)
            {
                ErrorMessage = ex.Message;

                oLogger.UpdateLogFile(DateTime.Now.ToString() + ":UpdateCoverRevhistPage:Error:" + ex.Message);

            }

            finally
            {



                if (wdoc != null)
                {
                    wdoc.Close(false); // Close the Word Document.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wdoc);

                }
                if (app != null)
                {
                    app.Quit(false); // Close Word Application.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);

                }

                wdoc = null;
                app = null;

                oLogger.UpdateLogFile(DateTime.Now.ToString() + ":UpdateCoverRevhistPage():Application was closed successfully during cover page update:");

                oLogger = null;

                GC.Collect();


            }

        }

        public void UpdateCoverRevhistPageBack(bool pUpdSignatureRev)
        {

            //this version replaces all version history from begining


            Logger oLogger = new Logger();
            oLogger.LogFileName = HttpContext.Current.Server.MapPath(Utility.GetLogFilePath());

            string tmpfiledirnm = Utility.GetTempLocalDirPath();
            string savePath = HttpContext.Current.Server.MapPath(tmpfiledirnm + FileName);

            object missObj = System.Reflection.Missing.Value;
            object path = savePath;

            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wdoc = app.Documents.Open(ref path, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj);


            try
            {


                wdoc.TrackRevisions = false;

                int totalTables = wdoc.Tables.Count;
                bool donotaddrow = false;


                //Add data into reviewer table  - 2nd table in the cover page

                if (totalTables > 0)
                {


                    //update 1st table in cover page, file title, SOP #, Rev #, Eff date, owner

                    Microsoft.Office.Interop.Word.Table tab1 = wdoc.Tables[1];
                    Microsoft.Office.Interop.Word.Range range1 = tab1.Range;

                    // Select the last row as source row.
                    int selectedRow1 = tab1.Rows.Count;

                    // Write new vaules to each cell.
                    tab1.Rows[1].Cells[2].Range.Text = FileTitle;
                    //   tab1.Rows[2].Cells[2].Range.Text = SOPNo;
                    tab1.Rows[2].Cells[4].Range.Text = FileCurrVersion;
                    tab1.Rows[3].Cells[2].Range.Text = DateTime.Today.ToString("MMMM dd, yyyy"); //current bcs it will publish now
                    tab1.Rows[4].Cells[2].Range.Text = FileOwner.userfullname;

                    //update 2nd table in  cover page for updating reviewers

                    Microsoft.Office.Interop.Word.Table tab2 = wdoc.Tables[2];
                    Microsoft.Office.Interop.Word.Range range2 = tab2.Range;

                    // Select the last row as source row.
                    int selectedRow2 = tab2.Rows.Count;

                    //keep only 2 rows if there are more than 2 rows in table

                    int rowstodel;
                    if (selectedRow2 > 2)
                    {
                        rowstodel = selectedRow2 - 2;
                        for (int i = 1; i <= rowstodel; i++)
                        {
                            tab2.Rows[3].Delete();

                        }
                        selectedRow2 = tab2.Rows.Count;
                    }

                    // Select and copy content of the source row.
                    range2.Start = tab2.Rows[selectedRow2].Cells[1].Range.Start;
                    range2.End = tab2.Rows[selectedRow2].Cells[tab2.Rows[selectedRow2].Cells.Count].Range.End;
                    range2.Copy();

                    // Insert a new row after the last row if it is not first row to add data

                    //Get reviewers with signatures of this file and request

                    int rvwrcnt = 1;

                    using (var ctx = new RadiantSOPEntities())
                    {

                        var rvrwrs = (from c in ctx.vwRvwrsSignatures where c.fileid == FileID && c.changerequestid == FileChangeRqstID select c);

                        foreach (var r in rvrwrs)
                        {

                            if (selectedRow2 == 2 && rvwrcnt == 1)
                            {

                                donotaddrow = true;

                            }
                            else
                                donotaddrow = false;

                            if (!donotaddrow)
                                tab2.Rows.Add(ref missObj);


                            // Moves the cursor to the first cell of target row.
                            range2.Start = tab2.Rows[tab2.Rows.Count].Cells[1].Range.Start;
                            range2.End = range2.Start;

                            // Paste values to target row.
                            range2.Paste();

                            // Write new vaules to each cell.

                            tab2.Rows[tab2.Rows.Count].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            tab2.Rows[tab2.Rows.Count].Cells[1].Range.Text = r.reviewername;

                            tab2.Rows[tab2.Rows.Count].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            tab2.Rows[tab2.Rows.Count].Cells[2].Range.Text = r.reviewertitle;

                            tab2.Rows[tab2.Rows.Count].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            tab2.Rows[tab2.Rows.Count].Cells[3].Range.Text = r.SignedStatus;

                            tab2.Rows[tab2.Rows.Count].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            tab2.Rows[tab2.Rows.Count].Cells[4].Range.Text = Convert.ToDateTime(r.signeddate).ToString("MMMM dd, yyyy");

                            rvwrcnt = rvwrcnt + 1;


                        }
                    }

                    //end updating 2nd reviewers table

                    //update 3rd table for approver

                    Microsoft.Office.Interop.Word.Table tab3 = wdoc.Tables[3];
                    Microsoft.Office.Interop.Word.Range range3 = tab3.Range;


                    tab3.Rows[2].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tab3.Rows[2].Cells[1].Range.Text = FileApprover.userfullname;

                    tab3.Rows[2].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tab3.Rows[2].Cells[2].Range.Text = FileApprover.userjobtitle;

                    tab3.Rows[2].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tab3.Rows[2].Cells[3].Range.Text = FileApprover.signstatus;

                    tab3.Rows[2].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tab3.Rows[2].Cells[4].Range.Text = FileApprover.signaturedate.ToString("MMMM dd, yyyy");


                    //end updating 3rd table for approver

                    //update last table to add data into Revison history table 

                    Microsoft.Office.Interop.Word.Table tab = wdoc.Tables[totalTables];
                    Microsoft.Office.Interop.Word.Range range = tab.Range;

                    // Select the last row as source row.
                    int selectedRow = tab.Rows.Count;


                    //delete rows if the table has more than three rows 

                    if (selectedRow > 2)
                    {
                        rowstodel = selectedRow - 2;
                        for (int i = 1; i <= rowstodel; i++)
                        {
                            tab.Rows[3].Delete();   //keep first two rows

                        }
                        selectedRow = tab.Rows.Count;
                    }

                    // Select and copy content of the source row.
                    range.Start = tab.Rows[selectedRow].Cells[1].Range.Start;
                    range.End = tab.Rows[selectedRow].Cells[tab.Rows[selectedRow].Cells.Count].Range.End;
                    range.Copy();


                    donotaddrow = false;

                    int filevercount = 1;

                    foreach (FileRevision rev in FileRevisions)
                    {
                        decimal drevno;
                        drevno = Convert.ToDecimal(rev.RevisionNo);

                        if ((drevno % 1) == 0)   //only approved version will show here
                        {

                            if (selectedRow == 2 && filevercount == 1)
                            {
                                donotaddrow = true;
                            }

                            else
                                donotaddrow = false;

                            if (!donotaddrow)
                                tab.Rows.Add(ref missObj);

                            // Moves the cursor to the first cell of target row.
                            range.Start = tab.Rows[tab.Rows.Count].Cells[1].Range.Start;
                            range.End = range.Start;

                            // Paste values to target row.
                            range.Paste();


                            // Write new vaules to each cell.

                            tab.Rows[tab.Rows.Count].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            tab.Rows[tab.Rows.Count].Cells[1].Range.Text = rev.RevisionNo;

                            tab.Rows[tab.Rows.Count].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            tab.Rows[tab.Rows.Count].Cells[2].Range.Text = rev.RevisionDate.ToString("MMMM dd, yyyy");

                            //  tab.Rows[tab.Rows.Count].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            //  tab.Rows[tab.Rows.Count].Cells[3].Range.Text = rev.Description;

                            filevercount = filevercount + 1;



                        }  //end checking approved version

                    }  //end for loop

                }


                // Set footers
                foreach (Microsoft.Office.Interop.Word.Section wordSection in wdoc.Sections)
                {
                    Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                    footerRange.Tables[1].Cell(1, 1).Range.Text = FileTitle;

                }

                wdoc.TrackRevisions = true;

                wdoc.SaveAs2(savePath);   //save in actual file from tempalte

            }

            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
                oLogger.UpdateLogFile(DateTime.Now.ToString() + ":UpdateCoverRevhistPage(publish true):Failed to update cover page with error:" + ex.Message);

                throw ex;
            }

            finally
            {


                if (wdoc != null)
                {
                    wdoc.Close(false); // Close the Word Document.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wdoc);

                }
                if (app != null)
                {
                    app.Quit(false); // Close Word Application.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);

                }

                wdoc = null;
                app = null;

                oLogger.UpdateLogFile(DateTime.Now.ToString() + ":UpdateCoverRevhistPage(Publish true):Application was closed successfully during cover page update:");

                oLogger = null;

                GC.Collect();

            }





        }


        public void UpdateCoverRevhistPage(bool pUpdSignatureRev)
        {


            Logger oLogger = new Logger();
            oLogger.LogFileName = HttpContext.Current.Server.MapPath(Utility.GetLogFilePath());

            string tmpfiledirnm = Utility.GetTempLocalDirPath();
            string savePath = HttpContext.Current.Server.MapPath(tmpfiledirnm + FileName);

            object missObj = System.Reflection.Missing.Value;
            object path = savePath;

            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wdoc = app.Documents.Open(ref path, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj);


            try
            {


                wdoc.TrackRevisions = false;

                int totalTables = wdoc.Tables.Count;
                bool donotaddrow = false;


                //Add data into reviewer table  - 2nd table in the cover page

                if (totalTables > 0)
                {


                    //update 1st table in cover page, file title, SOP #, Rev #, Eff date, owner

                    Microsoft.Office.Interop.Word.Table tab1 = wdoc.Tables[1];
                    Microsoft.Office.Interop.Word.Range range1 = tab1.Range;

                    // Select the last row as source row.
                    int selectedRow1 = tab1.Rows.Count;

                    // Write new vaules to each cell.
                    tab1.Rows[1].Cells[2].Range.Text = FileTitle;
                 //   tab1.Rows[2].Cells[2].Range.Text = SOPNo;
                    tab1.Rows[2].Cells[4].Range.Text = FileCurrVersion;
                    tab1.Rows[3].Cells[2].Range.Text = DateTime.Today.ToString("MMMM dd, yyyy"); //current bcs it will publish now
                    tab1.Rows[4].Cells[2].Range.Text =FileOwner.userfullname;

                    //update 2nd table in  cover page for updating reviewers

                    Microsoft.Office.Interop.Word.Table tab2 = wdoc.Tables[2];
                    Microsoft.Office.Interop.Word.Range range2 = tab2.Range;

                    // Select the last row as source row.
                    int selectedRow2 = tab2.Rows.Count;

                    //keep only 2 rows if there are more than 2 rows in table

                    int rowstodel;
                    if (selectedRow2 > 2)
                    {
                        rowstodel = selectedRow2 - 2;
                        for (int i = 1; i <= rowstodel; i++)
                        {
                            tab2.Rows[3].Delete();

                        }
                        selectedRow2 = tab2.Rows.Count;
                    }

                    // Select and copy content of the source row.
                    range2.Start = tab2.Rows[selectedRow2].Cells[1].Range.Start;
                    range2.End = tab2.Rows[selectedRow2].Cells[tab2.Rows[selectedRow2].Cells.Count].Range.End;
                    range2.Copy();

                    // Insert a new row after the last row if it is not first row to add data

                    //Get reviewers with signatures of this file and request

                    int rvwrcnt = 1;

                    using (var ctx = new RadiantSOPEntities())
                    {

                        var rvrwrs =(from c in ctx.vwRvwrsSignatures where c.fileid == FileID && c.changerequestid==FileChangeRqstID select c);

                        foreach (var r in rvrwrs)
                        {

                            if (selectedRow2 == 2 && rvwrcnt == 1)
                            {

                                donotaddrow = true;

                            }
                            else
                            {
                                
                               donotaddrow = false;
                            }


                            if (!donotaddrow)
                                tab2.Rows.Add(ref missObj);


                            // Moves the cursor to the first cell of target row.
                            range2.Start = tab2.Rows[tab2.Rows.Count].Cells[1].Range.Start;
                            range2.End = range2.Start;

                            // Paste values to target row.
                            range2.Paste();

                            // Write new vaules to each cell.

                            tab2.Rows[tab2.Rows.Count].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            tab2.Rows[tab2.Rows.Count].Cells[1].Range.Text = r.reviewername;

                            tab2.Rows[tab2.Rows.Count].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            tab2.Rows[tab2.Rows.Count].Cells[2].Range.Text = r.reviewertitle;

                            tab2.Rows[tab2.Rows.Count].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            tab2.Rows[tab2.Rows.Count].Cells[3].Range.Text = r.SignedStatus;

                            tab2.Rows[tab2.Rows.Count].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            tab2.Rows[tab2.Rows.Count].Cells[4].Range.Text = Convert.ToDateTime(r.signeddate).ToString("MMMM dd, yyyy");

                            rvwrcnt = rvwrcnt + 1;


                        }
                    }

                    //end updating 2nd reviewers table

                    //update 3rd table for approver

                    Microsoft.Office.Interop.Word.Table tab3 = wdoc.Tables[3];
                    Microsoft.Office.Interop.Word.Range range3 = tab3.Range;


                    tab3.Rows[2].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tab3.Rows[2].Cells[1].Range.Text = FileApprover.userfullname;

                    tab3.Rows[2].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tab3.Rows[2].Cells[2].Range.Text =FileApprover.userjobtitle;

                    tab3.Rows[2].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tab3.Rows[2].Cells[3].Range.Text = FileApprover.signstatus;

                    tab3.Rows[2].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    tab3.Rows[2].Cells[4].Range.Text = FileApprover.signaturedate.ToString("MMMM dd, yyyy");


                    //end updating 3rd table for approver

                    //update last table to add new revision data into Revison history table 

                    Microsoft.Office.Interop.Word.Table tab = wdoc.Tables[totalTables];
                    Microsoft.Office.Interop.Word.Range range = tab.Range;

                    // Select the last row as source row.
                    int selectedRow = tab.Rows.Count;


                    //delete rows if the table has more than three rows 

                    //if (selectedRow > 2)
                    //{
                    //    rowstodel = selectedRow - 2;
                    //    for (int i = 1; i <= rowstodel; i++)
                    //    {
                    //        tab.Rows[3].Delete();   //keep first two rows

                    //    }
                    //    selectedRow = tab.Rows.Count;
                    //}

                    
                    // Select and copy content of the source row.
                    
                    range.Start = tab.Rows[selectedRow].Cells[1].Range.Start;
                    range.End = tab.Rows[selectedRow].Cells[tab.Rows[selectedRow].Cells.Count].Range.End;
                    range.Copy();


                    donotaddrow = false;

                    bool lastrowtoupd ;

                   // int filevercount = 1;

                    foreach (FileRevision rev in FileRevisions)
                    {
                        decimal drevno;
                        drevno = Convert.ToDecimal(rev.RevisionNo);

                        //check whether this revision is already in table. if not then only
                        // add new revision so we do not replace historical revision description

                        donotaddrow = false;
                        lastrowtoupd = false;

                        // this loop ensures we do not replace existing revision history 
                        for (int i = 2; i <= tab.Rows.Count; i++)
                        {
                            if (rev.RevisionNo.Trim() == tab.Rows[i].Cells[1].Range.Text.Replace("\r\a", "").Trim())
                            {
                                donotaddrow = true;
                                break;
                            }

                        }

                        //if this is first revision history rev no 1, then do not add row  
                        if (selectedRow ==2 && tab.Rows[2].Cells[1].Range.Text.Replace("\r\a", "").Trim()=="" && rev.RevisionNo.Trim()=="1")
                        {
                            donotaddrow = true;
                            lastrowtoupd = true;  //ensure new data
                        }

                        if (!donotaddrow)
                        {
                            tab.Rows.Add(ref missObj);
                            lastrowtoupd = true;   //ensure new data
                        }

                        if (lastrowtoupd) //update last row only if it is new data
                        {
                            // Moves the cursor to the first cell of target row.
                            range.Start = tab.Rows[tab.Rows.Count].Cells[1].Range.Start;
                            range.End = range.Start;

                            // Paste values to target row.
                            range.Paste();


                            // Write new vaules to each cell.

                            tab.Rows[tab.Rows.Count].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            tab.Rows[tab.Rows.Count].Cells[1].Range.Text = rev.RevisionNo;

                            tab.Rows[tab.Rows.Count].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            tab.Rows[tab.Rows.Count].Cells[2].Range.Text = rev.RevisionDate.ToString("MMMM dd, yyyy");

                            tab.Rows[tab.Rows.Count].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                            tab.Rows[tab.Rows.Count].Cells[3].Range.Text = rev.Description;


                        }

                   //     filevercount = filevercount + 1;

                    }  //end for loop of revision history

                }


                // Set footers
                foreach (Microsoft.Office.Interop.Word.Section wordSection in wdoc.Sections)
                {
                    Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                    footerRange.Tables[1].Cell(1, 1).Range.Text = FileTitle;

                }

                wdoc.TrackRevisions = true;

                wdoc.SaveAs2(savePath);   //save in actual file from tempalte

            }

            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
                oLogger.UpdateLogFile(DateTime.Now.ToString() + ":UpdateCoverRevhistPage(publish true):Failed to update cover page with error:" + ex.Message);

                throw ex;
            }

            finally
            {


                if (wdoc != null)
                {
                    wdoc.Close(false); // Close the Word Document.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wdoc);

                }
                if (app != null)
                {
                    app.Quit(false); // Close Word Application.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);

                }

                wdoc = null;
                app = null;

                oLogger.UpdateLogFile(DateTime.Now.ToString() + ":UpdateCoverRevhistPage(Publish true):Application was closed successfully during cover page update:");

                oLogger = null;

                GC.Collect();

            }





        }

        public short GetLastChngReqSOPStatusCode()
        {
            short laststatcode=0;

            using (var dbctx = new RadiantSOPEntities())
            {
                laststatcode= Convert.ToInt16(dbctx.filechangerequestactivities.Where(f=>f.fileid==FileID).OrderByDescending(f=>f.changerequestid).Select(f => f.approvalstatuscode).FirstOrDefault());

            }

            return laststatcode;
        }


        public void GetSOPInfoByFileID()
        {

            using (var ctx = new RadiantSOPEntities())
            {
                FilePath = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.SPFilePath).FirstOrDefault();
                FileName = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.DeptFileName).FirstOrDefault();
                SOPNo = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.SOPNo).FirstOrDefault();
                FileTitle = Path.ChangeExtension(FileName, null);
                FileLink = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.SPFileLink).FirstOrDefault();
                FileCurrVersion = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.VersionNo).FirstOrDefault();
                Updatefreq = Convert.ToInt16(ctx.fileupdateschedules.Where(d => d.fileid == FileID).Select(d => d.frequencyofrevision).FirstOrDefault());
                Updatefrequnit =ctx.fileupdateschedules.Where(d => d.fileid == FileID).Select(d => d.unitoffrequency).FirstOrDefault();
                Updfrequnitcode= Convert.ToInt16(ctx.fileupdateschedules.Where(d => d.fileid == FileID).Select(d => d.unitcodeupdfreq).FirstOrDefault());

                if (FileCurrVersion != null && FileCurrVersion != "")
                {
                    FileCurrVersion = Math.Round(Math.Ceiling(Convert.ToDecimal(FileCurrVersion))).ToString();

                }

                AuthorName = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.CreatedBy).FirstOrDefault();
                SOPCreateDate = Convert.ToDateTime(ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.CreateDate).FirstOrDefault());
                ViewAccessType = ctx.fileviewaccesstypes.Where(v => v.fileid == FileID).Select(v => v.viewtypename).FirstOrDefault();

                if (ViewAccessType == null)
                    ViewAccessType = "";
                if (ViewAccessType.Trim() == "By Department")
                    DepartmentCode = Convert.ToInt16(ctx.fileviewaccesstypes.Where(v => v.fileid == FileID).Select(v => v.departmentcode).FirstOrDefault());
                if (ViewAccessType.Trim() == "All Users")
                    AllUsersReadAcc = true;


                    var owner = (from f in ctx.fileowners
                                  join u in ctx.users
                                  on f.ownerid equals u.userid
                                  where f.fileid == FileID &&  f.ownerstatuscode==1
                                  select new Employee
                                  {
                                      useremailaddress=u.useremailaddress,
                                      userid=f.ownerid
                                     
                                  }).ToList();


                foreach(Employee emp in owner)
                {

                    FileOwnerEmail = emp.useremailaddress;
                    FileOwnerID = emp.userid;
                }
               

                var approver = (from f in ctx.fileapprovers
                             join u in ctx.users
                             on f.approverid equals u.userid
                             where f.fileid == FileID && f.approverstatuscode == 1
                             select new Employee
                             {
                                 useremailaddress = u.useremailaddress,
                                 userid = f.approverid
                                 
                             }).ToList();

                foreach (Employee emp in approver)
                {

                    FileApproverEmail = emp.useremailaddress;
                    FileApproverID = emp.userid;
                }




            }



        }

        public void GetSOPInfo()
        {

            using (var ctx = new RadiantSOPEntities())
            {
                //basic sop info related to file id
                FilePath = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.SPFilePath).FirstOrDefault();
                FileName = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.DeptFileName).FirstOrDefault();
                SOPNo = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.SOPNo).FirstOrDefault();
                FileTitle = Path.ChangeExtension(FileName, null);
                FileLink = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.SPFileLink).FirstOrDefault();
                FileCurrVersion = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.VersionNo).FirstOrDefault();

                if (FileCurrVersion != null && FileCurrVersion != "")
                {
                    FileCurrVersion = Math.Round(Math.Ceiling(Convert.ToDecimal(FileCurrVersion))).ToString();

                }

                    // ApprovalStatus = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.ApprovalStatus).FirstOrDefault();
                AuthorName = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.CreatedBy).FirstOrDefault();
                SOPCreateDate = Convert.ToDateTime(ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.CreateDate).FirstOrDefault());
                ViewAccessType = ctx.fileviewaccesstypes.Where(v => v.fileid == FileID).Select(v => v.viewtypename).FirstOrDefault();

                if (ViewAccessType == null)
                    ViewAccessType = "";


                    //data related to change request

                Employee oSOPOwner = new Employee();

                oSOPOwner.useremailaddress = ctx.vwOwnerSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.owneremail).FirstOrDefault();
                oSOPOwner.GetUserByEmail();
                oSOPOwner.signaturedate = Convert.ToDateTime(ctx.vwOwnerSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.ownersigneddate).FirstOrDefault());
                oSOPOwner.signstatus= ctx.vwOwnerSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.ownerSignedStatus).FirstOrDefault();

                FileOwner = oSOPOwner;
                FileOwnerEmail = oSOPOwner.useremailaddress;


                Employee oSOPApprover = new Employee();

                oSOPApprover.useremailaddress = ctx.vwApprvrsSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.approveremail).FirstOrDefault();
                oSOPApprover.GetUserByEmail();
                oSOPApprover.signaturedate = Convert.ToDateTime(ctx.vwApprvrsSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.Aprvrsigneddate).FirstOrDefault());
                oSOPApprover.signstatus = ctx.vwApprvrsSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.AprvrSignedStatus).FirstOrDefault();


                FileApprover = oSOPApprover;
                FileApproverEmail = oSOPApprover.useremailaddress;

                FileStatuscode = ctx.filechangerequestactivities.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.approvalstatuscode).FirstOrDefault();
                
               


            }


            //get reviewers

            GetReviewers();

            GetFileRevisions();


        }

        public bool HasDuplicateSOPNOInDB()
        {
            bool hasdupsop = false;

            string dupsop = "";

            using (var dbctx = new RadiantSOPEntities())
            {

                dupsop = dbctx.deptsopfiles.Where(d => d.SOPNo.Trim().ToUpper() == SOPNo.Trim().ToUpper() && d.filestatuscode == 3).Select(d => d.SOPNo).FirstOrDefault();

                if (dupsop != null && dupsop.Trim() != "")
                    hasdupsop = true;
            }

            return hasdupsop;

        }

        public bool AuthenticateUser(string pAuthType)
        {
            bool authensop = false;

            int ownerid = 0;
            int approverid = 0;
            int reviewerid = 0;
            int loggedinuserid = Utility.GetLoggedInUserID();
            short usersopdeptcode = Utility.GetLoggedInUserSOPDeptCode();

            short sopfolderdeptcode=0;


            if (pAuthType == "publish" || pAuthType == "changerequest" || pAuthType == "admin")
            {

                using (var dbctx = new RadiantSOPEntities())
                {
                    ownerid = dbctx.fileowners.Where(o => o.ownerid == loggedinuserid && o.fileid == FileID && o.ownerstatuscode == 1).Select(o => o.ownerid).FirstOrDefault();

                    approverid = dbctx.fileapprovers.Where(a => a.approverid == loggedinuserid && a.fileid == FileID && a.approverstatuscode == 1).Select(a => a.approverid).FirstOrDefault();

                    reviewerid = dbctx.filereviewers.Where(r => r.reviewerid == loggedinuserid && r.fileid == FileID && r.reviewerstatuscode == 1).Select(r => r.reviewerid).FirstOrDefault();

                }


                if (pAuthType == "publish")  // looged in user must be a approver to publish sop 
                {
                    if (approverid > 0)
                        authensop = true;
                }


                if (pAuthType == "changerequest")   //looged user must be a signatory 
                {
                    if (ownerid > 0 || approverid > 0 || reviewerid > 0)
                    {
                        authensop = true;

                    }
                }

                if (pAuthType == "admin")  // looged in user must be a approver to publish sop 
                {
                    if (ownerid > 0)
                        authensop = true;
                }




            }  //end checking autype pf publish or chng req

            if (pAuthType=="createupload")
            {
                //get loggedin user sop department code
                //if his/her department is not same as the deaprtmetn he/selected to create file 
                //then deny access, if same then go to next check


                using (var dbctx = new RadiantSOPEntities())
                {
                    sopfolderdeptcode = dbctx.codesSOPDepartments.Where(d => d.sopdeptname.Trim().ToUpper() == FolderName.Trim().ToUpper()).Select(d => d.sopdeptcode).FirstOrDefault();

                    if (sopfolderdeptcode == usersopdeptcode)  //user belongs to same sop department as the dept of the selected folder name 
                                                               //then find if he/she is owner of any file in that sop department, if so then authentoicate
                    {
                        ownerid = dbctx.vwOwnrsSOPDeptCodes.Where(o => o.ownerid == loggedinuserid && o.sopdeptcode == usersopdeptcode).Select(o => o.ownerid).FirstOrDefault();

                        if (ownerid > 0)
                            authensop = true;
                    }
                    

                }


            }



            return authensop;
        }

        public void GetSOPNo()
        {

            using (var ctx = new RadiantSOPEntities())
            {

                SOPNo = ctx.GetLastSOPNO(FolderName, SubFolderName).FirstOrDefault().ToString();
            }


        }

        public int GetOwnershipID()
        {

            int ownershipid = 0;
            int fileownerid = 0;

            using (var ctx = new RadiantSOPEntities())
            {

                // only active current owner can make change request
                fileownerid = ctx.fileowners.Where(o => o.fileid == FileID && o.ownerstatuscode==1).Select(o => o.ownerid).FirstOrDefault();
                ownershipid = ctx.fileowners.Where(o=>o.fileid == FileID && o.ownerid == fileownerid).Select(o=>o.ownershipid).FirstOrDefault();
            }


            return ownershipid;
        }


        public int GetApproveID()
        {

            int approveid = 0;
            int approverid = 0;

            using (var ctx = new RadiantSOPEntities())
            {

                approverid = ctx.fileapprovers.Where(a => a.fileid == FileID && a.approverstatuscode==1).Select(a => a.approverid).FirstOrDefault();

                approveid = ctx.fileapprovers.Where(o => o.fileid == FileID && o.approverid == approverid).Select(o => o.approveid).FirstOrDefault();
            }


            return approveid;
        }




        //this is for making password secured through channel
        private static SecureString GetSecureString(String Password)
        {
            SecureString oSecurePassword = new SecureString();

            foreach (Char c in Password.ToCharArray())
            {
                oSecurePassword.AppendChar(c);

            }
            return oSecurePassword;
        }

        //all sharepoint operation
        //upload document to sharepoint
        public void UploadDocument()
        {



            using (ClientContext clientContext = new ClientContext(SiteUrl))
            {

                //string userName = Utility.GetSiteAdminUserName();  //it is email address of site admin
                //string password = Utility.GetSiteAdminPassowrd();


                SecureString SecurePassword = GetSecureString(password);
                clientContext.Credentials = new SharePointOnlineCredentials(userName, SecurePassword);

                //ICredentials credentials = new NetworkCredential(userName, password, domain);
                //clientContext.Credentials = credentials;

                Microsoft.SharePoint.Client.List documentsList = clientContext.Web.Lists.GetByTitle(DocumentLibName);

                FileCreationInformation fileCreationInformation = new FileCreationInformation();

                //Assign to content byte[] i.e. documentStream
               // fileCreationInformation.Content = FileStream;

                //fileCreationInformation.Content limits to less than 2 MB file
                
                //use ContentStream instead of Contentas it has no limit to file size
                //however, recommended for: - SharePoint Server 2013. - SharePoint Online when the file is smaller than 10 MB.

                fileCreationInformation.ContentStream = new MemoryStream(FileStream);

                //Allow owerwrite of document
                fileCreationInformation.Overwrite = true;


                //Upload URL
                fileCreationInformation.Url = SiteUrl +"/"+ FilePath + FileName;
                Microsoft.SharePoint.Client.File uploadFile = documentsList.RootFolder.Files.Add(fileCreationInformation);

                //  Update the metadata for a field having name "SOPNO", file owner

                // User theUser = clientContext.Web.SiteUsers.GetByEmail(FileOwner.useremailaddress);

                User theUser = clientContext.Web.SiteUsers.GetByEmail(FileOwnerEmail);


                uploadFile.ListItemAllFields["Owner"] = theUser;



               uploadFile.ListItemAllFields["SOPNO"] = SOPNo;


                uploadFile.ListItemAllFields.Update();
                clientContext.ExecuteQuery(); //upload file


                clientContext.Load(uploadFile, f => f.ListItemAllFields);
                clientContext.ExecuteQuery();
                //Print List Item Id
                Console.WriteLine("List Item Id: {0}", uploadFile.ListItemAllFields.Id);

                FileID = uploadFile.ListItemAllFields.Id;


            }


        }

        //remove existing sharepoint user permissions from file
        public void RemoveAllFilePermissions()
        {


            OperationSuccess = false;

            ClientContext clientContext = new ClientContext(SiteUrl);

            //string userName = Utility.GetSiteAdminUserName();  //it is email address of site admin
            //string password = Utility.GetSiteAdminPassowrd();


            SecureString SecurePassword = GetSecureString(password);
            clientContext.Credentials = new SharePointOnlineCredentials(userName, SecurePassword);


            Web web = clientContext.Web;


            clientContext.Load(web);
            clientContext.Load(web.Lists);
            clientContext.Load(web, wb => wb.ServerRelativeUrl);
            clientContext.ExecuteQuery();

            Microsoft.SharePoint.Client.List list = web.Lists.GetByTitle(DocumentLibName);
            clientContext.Load(list);
            clientContext.ExecuteQuery();

            Folder folder = web.GetFolderByServerRelativeUrl(web.ServerRelativeUrl + FileUrl);
            clientContext.Load(folder);
            clientContext.ExecuteQuery();

            CamlQuery camlQuery = new CamlQuery();


            //TO GET ONLY FILE ITEM
            camlQuery.ViewXml = "<View Scope='Recursive'> " +
                                   "  <Query> " +

                                  " + <Where> " +
                                       "  <Contains>" +
                                            " <FieldRef Name='FileLeafRef'/> " +
                                                " <Value Type='File'>" + FileName + "</Value>" +
                                           " </Contains> " +
                                       " </Where> " +

                                    " </Query> " +
                                " </View>";


            camlQuery.FolderServerRelativeUrl = folder.ServerRelativeUrl;
            ListItemCollection listItems = list.GetItems(camlQuery);
            clientContext.Load(listItems);
            clientContext.ExecuteQuery();


            foreach (ListItem item in listItems)
            {
                //item.FileSystemObjectType;

                if (item.FileSystemObjectType == FileSystemObjectType.File)
                {
                    // This is a File

                    item.BreakRoleInheritance(false, false);


                    RoleAssignmentCollection roleAssCol = item.RoleAssignments;

                    clientContext.Load(roleAssCol);
                    clientContext.ExecuteQuery();


                    int iRoles = 0;
                    while (iRoles < roleAssCol.Count)
                    {
                        //delete the existing permissions

                        //item.RoleAssignments[iRoles].DeleteObject();

                        item.RoleAssignments[iRoles].RoleDefinitionBindings.RemoveAll();
                        //add the reader permission

                        iRoles++;
                    }


                    item.Update();

                    clientContext.ExecuteQuery();
                    OperationSuccess = true;

                }
                else if (item.FileSystemObjectType == FileSystemObjectType.Folder)
                {
                    // This is a  Folder
                }




            }



        }

        //assign right permissions to requested the user with file in sharepoint by email
        public void AssignFilePermission(string operation, string plabel, string useremail)
        {

            ClientContext clientContext = new ClientContext(SiteUrl);


            //string userName = Utility.GetSiteAdminUserName();  //it is email address of site admin
            //string password = Utility.GetSiteAdminPassowrd();

            SecureString SecurePassword = GetSecureString(password);
            clientContext.Credentials = new SharePointOnlineCredentials(userName, SecurePassword);


            Web web = clientContext.Web;
            clientContext.Load(web);
            clientContext.Load(web.Lists);
            clientContext.Load(web, wb => wb.ServerRelativeUrl);
            clientContext.ExecuteQuery();

            Microsoft.SharePoint.Client.List list = web.Lists.GetByTitle(DocumentLibName);
            clientContext.Load(list);
            clientContext.ExecuteQuery();

            Folder folder = web.GetFolderByServerRelativeUrl(web.ServerRelativeUrl + FileUrl);
            clientContext.Load(folder);
            clientContext.ExecuteQuery();

            CamlQuery camlQuery = new CamlQuery();


            //TO GET ONLY FILE ITEM
            camlQuery.ViewXml = "<View Scope='Recursive'> " +
                                   "  <Query> " +

                                  " + <Where> " +
                                       "  <Contains>" +
                                            " <FieldRef Name='FileLeafRef'/> " +
                                                " <Value Type='File'>" + FileName + "</Value>" +
                                           " </Contains> " +
                                       " </Where> " +

                                    " </Query> " +
                                " </View>";

            camlQuery.FolderServerRelativeUrl = folder.ServerRelativeUrl;
            ListItemCollection listItems = list.GetItems(camlQuery);
            clientContext.Load(listItems);
            clientContext.ExecuteQuery();


            string loginname = useremail;


            foreach (ListItem item in listItems)
            {

                if (item.FileSystemObjectType == FileSystemObjectType.File)
                {
                    // This is a File

                    RoleDefinitionBindingCollection rd = new RoleDefinitionBindingCollection(clientContext);


                    rd.Add(clientContext.Web.RoleDefinitions.GetByName(plabel));
                    Principal user = clientContext.Web.EnsureUser(loginname);

                    if (ViewAccessType == "All Users")
                        item.BreakRoleInheritance(true, false);   //inherit permission for all users selection
                    else
                        item.BreakRoleInheritance(false, false);   //do not inherit if all users are not selected


                  //  item.BreakRoleInheritance(false, false);


                    // Microsoft.SharePoint.Client.GroupCollection groupCollection = web.SiteGroups;

                    // Group grpvisitor = groupCollection.GetByName("Watercooler Visitors");
                    // clientContext.Load(grpvisitor);


                    if (operation == "add")
                    {
                        item.RoleAssignments.Add(user, rd);
                    }
                    else if (operation == "remove")
                    {

                        item.RoleAssignments.GetByPrincipal(user).DeleteObject();

                    }


                    item.Update();

                    // item.RoleAssignments.Groups.Remove(grpvisitor);

                    clientContext.ExecuteQuery();

                }
                else if (item.FileSystemObjectType == FileSystemObjectType.Folder)
                {
                    // This is a  Folder
                }




            }



        }


        public void AssignFilePermission(string operation, string plabel, Employee[] employees)
        {

            ClientContext clientContext = new ClientContext(SiteUrl);

            SecureString SecurePassword = GetSecureString(password);
            clientContext.Credentials = new SharePointOnlineCredentials(userName, SecurePassword);

            var users = clientContext.LoadQuery(clientContext.Web.SiteUsers.Where(u => u.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.User && u.UserId.NameIdIssuer == "urn:federation:microsoftonline"));

            clientContext.ExecuteQuery();


            bool emailfound = false;

            foreach (Employee emp in employees)
            {
                //check if the user exists in the site
                // var chkuser = clientContext.LoadQuery(clientContext.Web.SiteUsers.Where(u => u.LoginName == emp.useremailaddress));

                emailfound = false;

                foreach (User u in users)
                {
                    if (u.Email.Trim().ToLower() == emp.useremailaddress.Trim().ToLower())

                    {
                        //userfullname = u.Title;
                        emailfound = true;
                        break;
                    }
                }

                if (emailfound)
                {
                    AssignFilePermission(operation, plabel, emp.useremailaddress.Trim());
                }


            }

        }

            //assign right permissions to requested group of users with file in sharepoint by email
            //public void AssignFilePermission(string operation, string plabel, Employee[] employees)
            //{

            //    ClientContext clientContext = new ClientContext(SiteUrl);


            //    //string userName = Utility.GetSiteAdminUserName();  //it is email address of site admin
            //    //string password = Utility.GetSiteAdminPassowrd();


            //    SecureString SecurePassword = GetSecureString(password);
            //    clientContext.Credentials = new SharePointOnlineCredentials(userName, SecurePassword);



            //    Web web = clientContext.Web;


            //    clientContext.Load(web);
            //    clientContext.Load(web.Lists);
            //    clientContext.Load(web, wb => wb.ServerRelativeUrl);
            //    clientContext.ExecuteQuery();

            //    Microsoft.SharePoint.Client.List list = web.Lists.GetByTitle(DocumentLibName);
            //    clientContext.Load(list);
            //    clientContext.ExecuteQuery();

            //    Folder folder = web.GetFolderByServerRelativeUrl(web.ServerRelativeUrl + FileUrl);
            //    clientContext.Load(folder);
            //    clientContext.ExecuteQuery();

            //    CamlQuery camlQuery = new CamlQuery();


            //    //TO GET ONLY FILE ITEM
            //    camlQuery.ViewXml = "<View Scope='Recursive'> " +
            //                           "  <Query> " +

            //                          " + <Where> " +
            //                               "  <Contains>" +
            //                                    " <FieldRef Name='FileLeafRef'/> " +
            //                                        " <Value Type='File'>" + FileName + "</Value>" +
            //                                   " </Contains> " +
            //                               " </Where> " +

            //                            " </Query> " +
            //                        " </View>";


            //    camlQuery.FolderServerRelativeUrl = folder.ServerRelativeUrl;
            //    ListItemCollection listItems = list.GetItems(camlQuery);
            //    clientContext.Load(listItems);

            //    var users = clientContext.LoadQuery(clientContext.Web.SiteUsers.Where(u => u.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.User && u.UserId.NameIdIssuer == "urn:federation:microsoftonline"));

            //    clientContext.ExecuteQuery();


            //    bool emailfound = false;

            //    foreach (ListItem item in listItems)
            //    {
            //        //item.FileSystemObjectType;

            //        if (item.FileSystemObjectType == FileSystemObjectType.File)
            //        {
            //            // This is a File

            //            RoleDefinitionBindingCollection rd = new RoleDefinitionBindingCollection(clientContext);
            //            rd.Add(clientContext.Web.RoleDefinitions.GetByName(plabel));


            //            // Microsoft.SharePoint.Client.GroupCollection groupCollection = web.SiteGroups;
            //            Principal user;

            //            // Group grpvisitor = groupCollection.GetByName("Watercooler Visitors");
            //            // clientContext.Load(grpvisitor);

            //            //get all site users to find user

            //            //clientContext.ExecuteQuery();

            //          //  emailfound = false;

            //            foreach (Employee emp in employees)
            //            {
            //                //check if the user exists in the site
            //                // var chkuser = clientContext.LoadQuery(clientContext.Web.SiteUsers.Where(u => u.LoginName == emp.useremailaddress));

            //                emailfound = false;

            //                foreach (User u in users)
            //                {
            //                    if (u.Email.Trim().ToLower() == emp.useremailaddress.Trim().ToLower())

            //                    {
            //                        //userfullname = u.Title;
            //                        emailfound = true;
            //                        break;
            //                    }
            //                }

            //                if (emailfound)
            //                {

            //                    user = clientContext.Web.EnsureUser(emp.useremailaddress);

            //                    if (ViewAccessType == "All Users")
            //                        item.BreakRoleInheritance(true, false);   //inherit permission for all users selection
            //                    else
            //                        item.BreakRoleInheritance(false, false);   //do not inherit if all users are not selected


            //                   // rd.Add(clientContext.Web.RoleDefinitions.GetByName(plabel));

            //                    if (operation == "add")
            //                    {
            //                        item.RoleAssignments.Add(user, rd);
            //                    }
            //                    else if (operation == "remove")
            //                    {

            //                        item.RoleAssignments.GetByPrincipal(user).DeleteObject();

            //                    }

            //                    item.Update();


            //                }  //end checking employee exists in the site

            //                clientContext.ExecuteQuery();

            //            }  //end looping employees

            //        }
            //        else if (item.FileSystemObjectType == FileSystemObjectType.Folder)
            //        {
            //            // This is a  Folder
            //        }




            //    }



            //}


            //get file version history from sharepoint

       public void GetFileRevisions()
        {

            ErrorMessage = "";

            ClientContext ctx = new ClientContext(SiteUrl);

            //string userName = Utility.GetSiteAdminUserName();  //it is email address of site admin
            //string password = Utility.GetSiteAdminPassowrd();

            SecureString spassword = GetSecureString(password);

            ctx.Credentials = new SharePointOnlineCredentials(userName, spassword);

            ctx.Load(ctx.Web);

            Web web = ctx.Web;

            //The ServerRelativeUrl property returns a string in the following form, which excludes the name of
            //    the server or root folder: / Site_Name / Subsite_Name / Folder_Name / File_Name.

            ctx.Load(web, wb => wb.ServerRelativeUrl);
            ctx.ExecuteQuery();

            // string filerelurl = web.ServerRelativeUrl + "/SOP/Information Technology (IT)/" + "IT-07 OperationTestFile.docx";

            string filerelurl = web.ServerRelativeUrl + "/" + FilePath.Trim() + FileName;

            Microsoft.SharePoint.Client.File file = web.GetFileByServerRelativeUrl(filerelurl);

            //CheckIn the file
            // file.CheckIn(String.Concat("File CheckingIn at ", DateTime.Now.ToLongDateString()), SP.CheckinType.MajorCheckIn);

            //CheckOut the File
            // file.CheckOut();

            //Publish the file

            ctx.Load(file);
            ctx.ExecuteQuery();

            FileVersionCollection fvc = file.Versions;

            ctx.Load(fvc);

            ctx.ExecuteQuery();


            //get major versions, first count this to define array
            int totrev = 0;
            foreach (FileVersion fv in fvc)
            {

                decimal drevno;
                drevno = Convert.ToDecimal(fv.VersionLabel);

                if ((drevno % 1) == 0)   //only approved version will show here

                {
                    totrev = totrev + 1;
                }
            }

            FileRevision[] oRVArr=new FileRevision[totrev] ;

            int i = 0;

            FileRevision oRiv;

            foreach (FileVersion fv in fvc)
            {

                decimal drevno;
                drevno = Convert.ToDecimal(fv.VersionLabel);

                if ((drevno % 1) == 0)   //only approved version will show here

                {
                    oRiv = new FileRevision();

                    oRiv.FileID = FileID;
                    oRiv.RevisionID = fv.ID;
                    oRiv.RevisionNo = Math.Round(Convert.ToDecimal(fv.VersionLabel)).ToString();
                    oRiv.RevisionDate = fv.Created.ToLocalTime();   //utc to local time
                    oRiv.Description = fv.CheckInComment;
                    oRiv.VersionUrl = fv.Url;

                    oRVArr[i] = oRiv;

                    i = i + 1;
                }
            }

            FileRevisions = oRVArr;

        }


        public void GetReviewers()
        {

            using (var dbctx = new RadiantSOPEntities())
            {
                var rvrwrs = (from c in dbctx.vwRvwrsSignatures where c.fileid == FileID && c.changerequestid == FileChangeRqstID select c);

                Employee[] oRreviewersArr = new Employee[rvrwrs.Count()];
                int i = 0;
                Employee oRvwr;
                foreach (var r in rvrwrs)
                {

                    oRvwr = new Employee();

                    oRvwr.useremailaddress = r.revieweremail;
                    oRvwr.userfullname = r.reviewername;
                    oRvwr.userid = r.reviewerid;
                    oRvwr.userjobtitle = r.reviewertitle;
                    oRvwr.signaturedate = Convert.ToDateTime(r.signeddate);
                    oRvwr.signstatus = r.SignedStatus;
                    oRvwr.signstatuscode = r.SignStatusCode;

                    oRreviewersArr[i] = oRvwr;

                    i = i + 1;

                }

                FileReviewers = oRreviewersArr;
            }


        }

        public void AssignFilePermissionToUsers(string plabel, string addremove, Employee[] employees)
        {

            bool pdone = false;
            ErrorMessage = "";

            SecureString spassword = GetSecureString(password);

            using (var ctx = new ClientContext(SiteUrl))
            {

                ctx.Credentials = new SharePointOnlineCredentials(userName, spassword);
                ctx.Load(ctx.Web);

                Web web = ctx.Web;

                //The ServerRelativeUrl property returns a string in the following form, which excludes the name of
                //    the server or root folder: / Site_Name / Subsite_Name / Folder_Name / File_Name.

                ctx.Load(web, wb => wb.ServerRelativeUrl);
                ctx.ExecuteQuery();

                // string filerelurl = web.ServerRelativeUrl + "/SOP/Information Technology (IT)/" + "IT-07 OperationTestFile.docx";

                string filerelurl = web.ServerRelativeUrl + "/" + FilePath.Trim() + FileName;

                Microsoft.SharePoint.Client.File file = web.GetFileByServerRelativeUrl(filerelurl);

                //need this valid users from sp to verify email that we get from employee list from watercooler
                //if user is no longer in sp then it fails that we don't want
                
                var users = ctx.LoadQuery(ctx.Web.SiteUsers.Where(u => u.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.User && u.UserId.NameIdIssuer == "urn:federation:microsoftonline"));

                ctx.ExecuteQuery();

                bool emailfound;

                RoleDefinitionBindingCollection rd = new RoleDefinitionBindingCollection(ctx);
                rd.Add(ctx.Web.RoleDefinitions.GetByName(plabel));

                Principal user;

                if (ViewAccessType == "All Users" || ViewAccessType == "Inherit")
                    file.ListItemAllFields.BreakRoleInheritance(true, false);   //inherit permission for all users selection
                else
                    file.ListItemAllFields.BreakRoleInheritance(false, false);   //do not inherit if all users are not selected

                foreach (Employee emp in employees)
                {

                    emailfound = false;

                    foreach (User u in users)
                    {
                        if (u.Email.Trim().ToLower() == emp.useremailaddress.Trim().ToLower())

                        {
                            //userfullname = u.Title;
                            emailfound = true;
                            break;
                        }
                    }

                    if (emailfound)
                    {
                        user = ctx.Web.EnsureUser(emp.useremailaddress);   //ensure breaks the process if user is not valid. that's why we checked valid users


                        if (addremove == "add")
                        {
                            file.ListItemAllFields.RoleAssignments.Add(user, rd);
                        }
                        else if (addremove == "remove")
                        {

                            file.ListItemAllFields.RoleAssignments.GetByPrincipal(user).DeleteObject();

                        }

                        file.ListItemAllFields.Update();


                    }


                }  //end looping all employees

                ctx.ExecuteQuery();   //update all permission in one shot. this is time saver here

            } //end using site contex
            
        }

        public void AssignFilePermissionToUsers(string plabel, string addremove, string empemail)
        {

            bool pdone = false;
            ErrorMessage = "";

            SecureString spassword = GetSecureString(password);

            using (var ctx = new ClientContext(SiteUrl))
            {

                ctx.Credentials = new SharePointOnlineCredentials(userName, spassword);
                ctx.Load(ctx.Web);

                Web web = ctx.Web;

                //The ServerRelativeUrl property returns a string in the following form, which excludes the name of
                //    the server or root folder: / Site_Name / Subsite_Name / Folder_Name / File_Name.

                ctx.Load(web, wb => wb.ServerRelativeUrl);
                ctx.ExecuteQuery();

                // string filerelurl = web.ServerRelativeUrl + "/SOP/Information Technology (IT)/" + "IT-07 OperationTestFile.docx";

                string filerelurl = web.ServerRelativeUrl + "/" + FilePath.Trim() + FileName;

                Microsoft.SharePoint.Client.File file = web.GetFileByServerRelativeUrl(filerelurl);

                //need this valid users from sp to verify email that we get from employee list from watercooler
                //if user is no longer in sp then it fails that we don't want

                var users = ctx.LoadQuery(ctx.Web.SiteUsers.Where(u => u.PrincipalType == Microsoft.SharePoint.Client.Utilities.PrincipalType.User && u.UserId.NameIdIssuer == "urn:federation:microsoftonline"));

                ctx.ExecuteQuery();

                bool emailfound;

                RoleDefinitionBindingCollection rd = new RoleDefinitionBindingCollection(ctx);
                rd.Add(ctx.Web.RoleDefinitions.GetByName(plabel));   //get permission label, i.e. read , contribute etc.

                Principal user;

                if (ViewAccessType == "All Users" || ViewAccessType == "Inherit")
                    file.ListItemAllFields.BreakRoleInheritance(true, false);   //inherit permission for all users selection
                else
                    file.ListItemAllFields.BreakRoleInheritance(false, false);   //do not inherit if all users are not selected

                emailfound = false;

                foreach (User u in users)
                {
                    if (u.Email.Trim().ToLower() == empemail.Trim().ToLower())

                    {
                        //userfullname = u.Title;
                        emailfound = true;
                        break;
                    }
                }

                if (emailfound)
                {
                    user = ctx.Web.EnsureUser(empemail.Trim());   //ensure breaks the process if user is not valid. that's why we checked valid users

                    if (addremove == "add")
                    {
                        file.ListItemAllFields.RoleAssignments.Add(user, rd);
                    }
                    else if (addremove == "remove")
                    {

                        file.ListItemAllFields.RoleAssignments.GetByPrincipal(user).DeleteObject();

                    }

                    file.ListItemAllFields.Update();


                }

                ctx.ExecuteQuery();   //update all permission in one shot. this is time saver here

            } //end using site contex

        }


        public void ArchiveSOP()
        {

            //shell comand

            //Connect - SPOnline - url[yoururl]
            //$ctx = Get - SPOContext
            //$web = Get - SPOWeb
            //$fileUrl = "/sites/contosobeta/Shared Documents/test.csv"
            //$newfileUrl = "/sites/contosobeta/Shared Documents/test_rename.csv"
            //$file = $web.GetFileByServerRelativeUrl("$fileUrl")
            //$file.MoveTo("$newfileUrl", 'Overwrite')
            //$ctx.ExecuteQuery()


            bool pdone = false;
            ErrorMessage = "";

            ClientContext ctx = new ClientContext(SiteUrl);

            SecureString spassword = GetSecureString(password);

            ctx.Credentials = new SharePointOnlineCredentials(userName, spassword);

            ctx.Load(ctx.Web);

            Web web = ctx.Web;

            //The ServerRelativeUrl property returns a string in the following form, which excludes the name of
            //    the server or root folder: / Site_Name / Subsite_Name / Folder_Name / File_Name.

            ctx.Load(web, wb => wb.ServerRelativeUrl);
            ctx.ExecuteQuery();

            // string filerelurl = web.ServerRelativeUrl + "/SOP/Information Technology (IT)/" + "IT-07 OperationTestFile.docx";

            string filerelurl = web.ServerRelativeUrl + "/" + FilePath.Trim() + FileName;

            string filenewrelurl = web.ServerRelativeUrl + "/SOP/Archive/"+ SOPNo+" "+FileName;

            Microsoft.SharePoint.Client.File file = web.GetFileByServerRelativeUrl(filerelurl);

            ctx.Load(file);
            ctx.ExecuteQuery();

            file.MoveTo(filenewrelurl, MoveOperations.Overwrite);

            ctx.ExecuteQuery();
        }

        //publish/approve file in sharepoint so all users having read permission can view the sop 
        public bool PublishFile()
        {

            bool pdone = false;
            ErrorMessage = "";

            ClientContext ctx = new ClientContext(SiteUrl);

            SecureString spassword = GetSecureString(password);

            ctx.Credentials = new SharePointOnlineCredentials(userName, spassword);
                        
            ctx.Load(ctx.Web);

            Web web = ctx.Web;

            //The ServerRelativeUrl property returns a string in the following form, which excludes the name of
            //    the server or root folder: / Site_Name / Subsite_Name / Folder_Name / File_Name.

            ctx.Load(web, wb => wb.ServerRelativeUrl);
            ctx.ExecuteQuery();

            // string filerelurl = web.ServerRelativeUrl + "/SOP/Information Technology (IT)/" + "IT-07 OperationTestFile.docx";

            string filerelurl = web.ServerRelativeUrl + "/"+FilePath.Trim() +FileName;

            Microsoft.SharePoint.Client.File file = web.GetFileByServerRelativeUrl(filerelurl);

            //CheckIn the file
            // file.CheckIn(String.Concat("File CheckingIn at ", DateTime.Now.ToLongDateString()), SP.CheckinType.MajorCheckIn);

            //CheckOut the File
            // file.CheckOut();

            //Publish the file

            ctx.Load(file);
            ctx.ExecuteQuery();

            

            if (file.Level.ToString() == "Draft")    //level enum Published=1, Draft=2, Checkout=255 
            {
                file.Approve(String.Concat("File Publishing at ", DateTime.Now.ToLongDateString()));
                ctx.ExecuteQuery();


                //update change request table with published status code 3 
                // we need this because when user will request a change we need to check whether previous
                //request was published or not. if not published, then during change request we will advise
                //user to wait until the current request is approved.

                UpdateChangeReqID(3);

                pdone = true;
            }

            if (file.Level.ToString() == "Published")

            {
                pdone = false;
                ErrorMessage = "It is already published";

            }


                //UnPublish the file
                // file.UnPublish(String.Concat("File UnPublishing at ", DateTime.Now.ToLongDateString()));



                return pdone;
        }


        public void AssignSigatoriesPermission()
        {


            //give contribute permission to all reviewers

            AssignFilePermissionToUsers("contribute", "add", FileReviewers);

            //give edit permission to approver

            AssignFilePermissionToUsers("edit", "add", FileApproverEmail);

            //give full permission to owner

            AssignFilePermissionToUsers("full control", "add", FileOwnerEmail);



        }


        public void AssignSignatoresReadPermission()
        {


            //remove all edit permission as they might edit after publishing that we don't want
            //we will give edit permission to approvers during change request

    
            AssignFilePermissionToUsers("contribute", "remove", FileReviewers);

            AssignFilePermissionToUsers("edit", "remove", FileApproverEmail);
 
            AssignFilePermissionToUsers("full control", "remove", FileOwnerEmail);

            //now reassign them as reader

            AssignFilePermissionToUsers("read", "add", FileReviewers);

            AssignFilePermissionToUsers("read", "add", FileApproverEmail);

            AssignFilePermissionToUsers("read", "add", FileOwnerEmail);




        }


        //download file from sharepoint so system can modify cover page and revision history
        public void DownloadFileFromSharePoint(string tempLocation)
        {


            bool pdone = false;
            ErrorMessage = "";

            ClientContext ctx = new ClientContext(SiteUrl);


            SecureString spassword = GetSecureString(password);

            ctx.Credentials = new SharePointOnlineCredentials(userName, spassword);

            ctx.Load(ctx.Web);

            Web web = ctx.Web;

            //The ServerRelativeUrl property returns a string in the following form, which excludes the name of
            //    the server or root folder: / Site_Name / Subsite_Name / Folder_Name / File_Name.

            ctx.Load(web, wb => wb.ServerRelativeUrl);
            ctx.ExecuteQuery();

            // string filerelurl = web.ServerRelativeUrl + "/SOP/Information Technology (IT)/" + "IT-07 OperationTestFile.docx";

            string filerelurl = web.ServerRelativeUrl + "/" + FilePath.Trim() + FileName;

            Microsoft.SharePoint.Client.File file = web.GetFileByServerRelativeUrl(filerelurl);

            ctx.Load(file);
            ctx.ExecuteQuery();

            FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, file.ServerRelativeUrl);

            var filePath = tempLocation + file.Name;

            using (var fileStream = new System.IO.FileStream(filePath, System.IO.FileMode.Create))
            {
                fileInfo.Stream.CopyTo(fileStream);
            }



            //---------------------------following code is to downlaod all files in a folder
            //DownloadFilesFromSharePoint("https://tenant.sharepoint.com", "/SharedDocuments", @"c:\downloads");

            // ClientContext ctx = new ClientContext(SiteUrl);

            // SecureString SecurePassword = GetSecureString(Password);
            // ctx.Credentials = new SharePointOnlineCredentials(UserName, SecurePassword);


            //// ctx.Credentials = new SharePointOnlineCredentials(UserName, Password);

            // FileCollection files = ctx.Web.GetFolderByServerRelativeUrl(FolderName).Files;

            // ctx.Load(files);
            // ctx.ExecuteQuery();

            // foreach (Microsoft.SharePoint.Client.File file in files)
            // {
            //     FileInformation fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, file.ServerRelativeUrl);
            //     ctx.ExecuteQuery();

            //     var filePath = tempLocation + file.Name;
            //     using (var fileStream = new System.IO.FileStream(filePath, System.IO.FileMode.Create))
            //     {
            //         fileInfo.Stream.CopyTo(fileStream);
            //     }
            // }


        }

   

    } //end of class

    internal class FileReviewers
    {
    }
}  //end of namespace