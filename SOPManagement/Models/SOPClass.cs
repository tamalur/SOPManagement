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

namespace SOPManagement.Models
{

    [Bind(Exclude = "Id")]

    public class SOPClass
    {

        public int FileID { get; set; }

        public string[] FilereviewersArr { get; set; }

        public string[] FileviewersArr { get; set; }

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

        public string DepartmentName { get; set; }
        public short DepartmentCode { get; set; }

        public string ViewAccessType { get; set; }

        public int ViewAccessTypeID { get; set; }


        public string SOPNo { get; set; }

        public FileRevision[] FileRevisions { get; set; }

        public string FileCurrVersion { get; set; }


        [Required(ErrorMessage = "Frequency Required")]
        [Display(Name = "Update Frequency")]
        public short Updatefreq { get; set; }

        [Display(Name = "Select Frequency Unit")]
        public string Updatefrequnit { get; set; }

        public short Updfrequnitcode { get; set; }

        public string FileLink { get; set; }

        public string FilePath { get; set; }

        public string FileLocalPath { get; set; }

        public bool OperationSuccess { get; set; }

        public Employee[] FileReviewers { get; set; }

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


        public void AddFileReviewers()
        {

            Employee emp = new Employee();

            int rvwrid;

            OperationSuccess = false;

            foreach (Employee rvwr in Reviewers)
            {
                emp.useremailaddress = rvwr.useremailaddress;
                emp.GetUserByEmail();
                rvwrid = emp.userid;

                using (var dbcontext = new RadiantSOPEntities())
                {

                    var rvwrtable = new filereviewer()
                    {
                        reviewerid = rvwrid,
                        fileid = FileID

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

        private void AddRvwractivities(int previewid)
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

        private void AddApproveractivities(int papproveid)
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

        private void AddPublisheractivities(int ppublisherid)
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

        private void AddOwneractivities(int pownershipid)
        {
            using (var dbcontext = new RadiantSOPEntities())
            {
                var owneractvts = new fileownersactivity()
                {
                    changerequestid = FileChangeRqstID,
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

            using (var dbcontext = new RadiantSOPEntities())
            {

                var aprvrtable = new fileapprover()
                {
                    approverid = apprvrid,
                    fileid = FileID

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


            foreach (Employee vwr in Viewers)
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

            using (var dbcontext = new RadiantSOPEntities())
            {

                var ownertable = new fileowner()
                {
                    ownerid = ownerid,
                    fileid = FileID

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


            //  string newfilename = SOPNo + " " + SOPFileTitle;


            //string savePath = HttpContext.Current.Server.MapPath("~/Content/docfiles/" + FileName);

            string tmpfiledirnm = Utility.GetTempLocalDirPath();

            //string savePath = HttpContext.Current.Server.MapPath("~/Content/docfiles/" + FileName);

            string savePath = HttpContext.Current.Server.MapPath(tmpfiledirnm + FileName);

            object path = savePath;

            

            try
            {


                //  Microsoft.Office.Interop.Word.Document wdoc = app.Documents.Open(ref path, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj);

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

                    if ((totnewrow - 1) == totrvwrs)
                    {

                        for (int i = 1; i <= totnewrow - 1; i++)   //exclude row[0] for header
                        {

                            emp.useremailaddress = FileReviewers[rvwrno].useremailaddress;
                            emp.GetUserByEmail();

                            tab2.Rows[i].Cells[0].Paragraphs[0].Remove(false);
                            tab2.Rows[i].Cells[0].Paragraphs[0].Append(emp.userfullname);

                            tab2.Rows[i].Cells[1].Paragraphs[0].Remove(false);
                            tab2.Rows[i].Cells[1].Paragraphs[0].Append(emp.userjobtitle);


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

                    tab3.Rows[1].Cells[0].Paragraphs[0].Remove(false);    //track changes false
                    tab3.Rows[1].Cells[0].Paragraphs[0].Append(approverfullname);

                    tab3.Rows[1].Cells[1].Paragraphs[0].Remove(false);    //track changes false
                    tab3.Rows[1].Cells[1].Paragraphs[0].Append(approvertitle);


                    //start now add/update of revision history table



                    int totrh = FileRevisions.Count();

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



                                tabrh.Rows[i].Cells[0].Paragraphs[0].Remove(false);
                                tabrh.Rows[i].Cells[0].Paragraphs[0].Append(FileRevisions[rhno].RevisionNo);

                                tabrh.Rows[i].Cells[1].Paragraphs[0].Remove(false);
                                tabrh.Rows[i].Cells[1].Paragraphs[0].Append(FileRevisions[rhno].RevisionDate.ToShortDateString());


                                rhno = rhno + 1;

                            }
                        }
                        //end updating revision table



                    } //end checking total revisons

                    //add footer

                    wdoc.AddFooters();


                    // Get the default Footer for this document.
                    Footer footer_default = wdoc.Footers.Odd;

                    // Insert a Paragraph into the default Footer.
                    Paragraph p3 = footer_default.InsertParagraph();

                    p3.Append(FileTitle).Bold();


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


            //  string newfilename = SOPNo + " " + SOPFileTitle;


            //string savePath = HttpContext.Current.Server.MapPath("~/Content/docfiles/" + FileName);

            string tmpfiledirnm = Utility.GetTempLocalDirPath();

            //string savePath = HttpContext.Current.Server.MapPath("~/Content/docfiles/" + FileName);

            string savePath = HttpContext.Current.Server.MapPath(tmpfiledirnm + FileName);

            object missObj = System.Reflection.Missing.Value;
            object path = savePath;
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            try
            {


              //  System.IO.File.Copy(HttpContext.Current.Server.MapPath("~/Content/docfiles/SOPTemp.docx"), HttpContext.Current.Server.MapPath("~/Content/docfiles/" + FileName), true);

                Microsoft.Office.Interop.Word.Document wdoc = app.Documents.Open(ref path, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj);


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
                    //tab1.Rows[3].Cells[2].Range.Text = SOPEffectiveDate.ToShortDateString();   //for new file it will be updated during publishing
                    tab1.Rows[4].Cells[2].Range.Text = ownerfullname;


                    //update 2nd table in  cover page for updating reviewers

                    Microsoft.Office.Interop.Word.Table tab2 = wdoc.Tables[2];
                    Microsoft.Office.Interop.Word.Range range2 = tab2.Range;

                    // Select the last row as source row.
                    int selectedRow2 = tab2.Rows.Count;

                    //keep only 3 rows if there are more than 3 rows in table
                    //int rvrrowcount = Reviewers.Count();

                    int rowstodel;
                    if (selectedRow2 > 3)
                    {
                        rowstodel = selectedRow2 - 3;
                        for (int i = 1; i <= rowstodel; i++)
                        {
                            tab2.Rows[4].Delete();

                        }
                        selectedRow2 = tab2.Rows.Count;
                    }



                    // Select and copy content of the source row.
                    range2.Start = tab2.Rows[selectedRow2].Cells[1].Range.Start;
                    range2.End = tab2.Rows[selectedRow2].Cells[tab2.Rows[selectedRow2].Cells.Count].Range.End;
                    range2.Copy();

                    // Insert a new row after the last row if it is not first row to add data

                    int rvwrcnt = 1;
                    foreach (Employee rvwr in Reviewers)
                    {

                        if (selectedRow2 == 3 && rvwrcnt == 1)
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

                        tab2.Rows[tab2.Rows.Count].Cells[1].Range.Text = rvwr.userfullname;
                        tab2.Rows[tab2.Rows.Count].Cells[2].Range.Text = emp.userjobtitle;
                        //tab2.Rows[tab2.Rows.Count].Cells[3].Range.Text = "cell 3";

                        rvwrcnt = rvwrcnt + 1;


                    }

                    //end updating 2nd reviewers table

                    //update 3rd table for approver

                    //update 1st table in cover page, file title, SOP #, Rev #, Eff date, owner

                    Microsoft.Office.Interop.Word.Table tab3 = wdoc.Tables[3];
                    Microsoft.Office.Interop.Word.Range range3 = tab3.Range;

                    // Write new vaules to each cell of row 3. One row always as there will be one approver
                    tab3.Rows[3].Cells[1].Range.Text = approverfullname;
                    tab3.Rows[3].Cells[2].Range.Text = approvertitle;

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
                            tab.Rows[3].Delete();

                        }
                        selectedRow = tab.Rows.Count;
                    }

                    // Select and copy content of the source row.
                    range.Start = tab.Rows[selectedRow].Cells[1].Range.Start;
                    range.End = tab.Rows[selectedRow].Cells[tab.Rows[selectedRow].Cells.Count].Range.End;
                    range.Copy();


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

                    //}

                }


                // Set footers
                foreach (Microsoft.Office.Interop.Word.Section wordSection in wdoc.Sections)
                {
                    Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                    footerRange.Tables[1].Cell(1, 1).Range.Text = FileTitle;



                }

                wdoc.SaveAs2(savePath);   //save in actual file from tempalte


            }

            catch (Exception ex)
            {
                ErrorMessage = ex.Message;

            }

            finally
            {
                app.Application.Quit();

            }





        }

        public void UpdateCoverRevhistPage(bool pUpdSignatureRev)
        {



            string tmpfiledirnm = Utility.GetTempLocalDirPath();

            //string savePath = HttpContext.Current.Server.MapPath("~/Content/docfiles/" + FileName);

            string savePath = HttpContext.Current.Server.MapPath(tmpfiledirnm + FileName);

            //string savePath = HttpContext.Current.Server.MapPath("~/Content/docfiles/" + FileName);

            object missObj = System.Reflection.Missing.Value;
            object path = savePath;
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            try
            {


                //  System.IO.File.Copy(HttpContext.Current.Server.MapPath("~/Content/docfiles/SOPTemp.docx"), HttpContext.Current.Server.MapPath("~/Content/docfiles/" + FileName), true);

                Microsoft.Office.Interop.Word.Document wdoc = app.Documents.Open(ref path, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj);


                //  add row in table and data in cell

                // Employee emp = new Employee();

                string ownerfullname;
                string ownertitle;
                string ownrsigstatus;
                DateTime ownrsigndate;


                string approverfullname;
                string approvertitle;
                string approversignstat;
                DateTime apprvsigndate;


                using (var ctx = new RadiantSOPEntities())
                {
                    approverfullname = ctx.vwApprvrsSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.approvername).FirstOrDefault();
                    approvertitle = ctx.vwApprvrsSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.approvertitle).FirstOrDefault();
                    approversignstat= ctx.vwApprvrsSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.AprvrSignedStatus).FirstOrDefault();
                    apprvsigndate= Convert.ToDateTime(ctx.vwApprvrsSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.Aprvrsigneddate).FirstOrDefault());

                    ownerfullname = ctx.vwOwnerSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.ownerrname).FirstOrDefault();
                    ownertitle = ctx.vwOwnerSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.ownertitle).FirstOrDefault();
                    ownrsigstatus= ctx.vwOwnerSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.ownerSignedStatus).FirstOrDefault();
                    ownrsigndate= Convert.ToDateTime(ctx.vwOwnerSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.ownersigneddate).FirstOrDefault());

                }


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
                    int filecurversion=1;
                    //if FileCurrVersion

                     decimal fd=0 ;
                    
                    bool result = decimal.TryParse(FileCurrVersion, out fd); //i now = 108  

                    if (result)
                    {
                        filecurversion = Convert.ToInt16(Math.Ceiling(fd));
                    }

        

                    // Write new vaules to each cell.
                    tab1.Rows[1].Cells[2].Range.Text = FileTitle;
                    tab1.Rows[2].Cells[2].Range.Text = SOPNo;
                    tab1.Rows[2].Cells[4].Range.Text = filecurversion.ToString();
                    tab1.Rows[3].Cells[2].Range.Text = DateTime.Today.ToShortDateString(); //current bcs it will publish now
                    //SOPEffectiveDate.ToShortDateString();
                    tab1.Rows[4].Cells[2].Range.Text = ownerfullname;


                    //update 2nd table in  cover page for updating reviewers

                    Microsoft.Office.Interop.Word.Table tab2 = wdoc.Tables[2];
                    Microsoft.Office.Interop.Word.Range range2 = tab2.Range;

                    // Select the last row as source row.
                    int selectedRow2 = tab2.Rows.Count;

                    //keep only 3 rows if there are more than 3 rows in table
                    //int rvrrowcount = Reviewers.Count();

                    int rowstodel;
                    if (selectedRow2 > 3)
                    {
                        rowstodel = selectedRow2 - 3;
                        for (int i = 1; i <= rowstodel; i++)
                        {
                            tab2.Rows[4].Delete();

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

                       // var rvrwrs = ctx.vwRvwrsSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d);

                        var rvrwrs =(from c in ctx.vwRvwrsSignatures where c.fileid == FileID && c.changerequestid==FileChangeRqstID select c);

                        foreach (var r in rvrwrs)
                        {
                            // Console.WriteLine(r.reviewername);

                            if (selectedRow2 == 3 && rvwrcnt == 1)
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


                            tab2.Rows[tab2.Rows.Count].Cells[1].Range.Text = r.reviewername;
                            tab2.Rows[tab2.Rows.Count].Cells[2].Range.Text = r.reviewertitle;
                            tab2.Rows[tab2.Rows.Count].Cells[3].Range.Text = r.SignedStatus;
                            if (Convert.ToDateTime(r.signeddate).Year>70)
                                tab2.Rows[tab2.Rows.Count].Cells[4].Range.Text = Convert.ToDateTime(r.signeddate).ToShortDateString();

                            rvwrcnt = rvwrcnt + 1;


                        }
                    }



                    //foreach (Employee rvwr in Reviewers)
                    //{


                    //}

                    //end updating 2nd reviewers table

                    //update 3rd table for approver

                    //update 1st table in cover page, file title, SOP #, Rev #, Eff date, owner

                    Microsoft.Office.Interop.Word.Table tab3 = wdoc.Tables[3];
                    Microsoft.Office.Interop.Word.Range range3 = tab3.Range;

                    // Write new vaules to each cell of row 3. One row always as there will be one approver


                    tab3.Rows[3].Cells[1].Range.Text = approverfullname;
                    tab3.Rows[3].Cells[2].Range.Text = approvertitle;
                    tab3.Rows[3].Cells[3].Range.Text = approversignstat;

                    if (Convert.ToDateTime(apprvsigndate).Year > 70)
                        tab3.Rows[3].Cells[4].Range.Text = apprvsigndate.ToShortDateString();


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
                            tab.Rows[3].Delete();

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
                            tab.Rows[tab.Rows.Count].Cells[1].Range.Text = Math.Round(Convert.ToDecimal(rev.RevisionNo)).ToString();
                            tab.Rows[tab.Rows.Count].Cells[2].Range.Text = rev.RevisionDate.ToString("M/d/yyyy");
                            tab.Rows[tab.Rows.Count].Cells[3].Range.Text = rev.Description;

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

                wdoc.SaveAs2(savePath);   //save in actual file from tempalte


            }

            catch (Exception ex)
            {
                ErrorMessage = ex.Message;

            }

            finally
            {
                app.Application.Quit();

            }





        }



        public void GetSOPInfo()
        {

            using (var ctx = new RadiantSOPEntities())
            {
                //basic sop info related to file id
                FilePath = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.SPFilePath).FirstOrDefault();
                FileName = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.DeptFileName).FirstOrDefault();
                FileTitle = Path.ChangeExtension(FileName, null);
                SOPNo = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.SOPNo).FirstOrDefault();
                FileLink = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.SPFileLink).FirstOrDefault();
                FileCurrVersion = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.VersionNo).FirstOrDefault();
               // ApprovalStatus = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.ApprovalStatus).FirstOrDefault();
                AuthorName = ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.CreatedBy).FirstOrDefault();
                SOPCreateDate = Convert.ToDateTime(ctx.deptsopfiles.Where(d => d.FileID == FileID).Select(d => d.CreateDate).FirstOrDefault());

                //data related to change request

                FileOwner.useremailaddress = ctx.vwOwnerSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.owneremail).FirstOrDefault();
                FileOwner.GetUserByEmail();
                FileOwner.signaturedate = Convert.ToDateTime(ctx.vwOwnerSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.ownersigneddate).FirstOrDefault());
                FileOwner.signstatus= ctx.vwOwnerSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.ownerSignedStatus).FirstOrDefault();

                FileApprover.useremailaddress = ctx.vwApprvrsSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.approveremail).FirstOrDefault();
                FileApprover.GetUserByEmail();
                FileApprover.signaturedate = Convert.ToDateTime(ctx.vwApprvrsSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.Aprvrsigneddate).FirstOrDefault());
                FileApprover.signstatus = ctx.vwApprvrsSignatures.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.AprvrSignedStatus).FirstOrDefault();

                FileStatuscode = ctx.filechangerequestactivities.Where(d => d.fileid == FileID && d.changerequestid == FileChangeRqstID).Select(d => d.approvalstatuscode).FirstOrDefault();
                
            }


            //get reviewers

            GetReviewers();

            GetFileRevisions();


        }

        public void GetSOPNo()
        {

            RadiantSOPEntities ctx = new RadiantSOPEntities();

            SOPNo = ctx.GetLastSOPNO(FolderName, SubFolderName).FirstOrDefault().ToString();


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

                fileCreationInformation.Content = FileStream;
                //Allow owerwrite of document

                fileCreationInformation.Overwrite = true;
                //Upload URL

                fileCreationInformation.Url = SiteUrl +"/"+ FilePath + FileName;
                Microsoft.SharePoint.Client.File uploadFile = documentsList.RootFolder.Files.Add(fileCreationInformation);

                //Update the metadata for a field having name "SOPNO"
                //  string loginname = "tshaikh@radiantdelivers.com";

                //string loginname = userName;   //email address of site admin

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

            FileRevision[] oRVArr=new FileRevision[fvc.Count] ;

            int i = 0;

            FileRevision oRiv;

            foreach (FileVersion fv in fvc)
            {
                oRiv = new FileRevision();

                oRiv.FileID = FileID;
                oRiv.RevisionID = fv.ID;
                oRiv.RevisionNo = fv.VersionLabel;
                oRiv.RevisionDate = fv.Created;
                oRiv.Description = fv.CheckInComment;
                oRiv.VersionUrl = fv.Url;

                oRVArr[i] = oRiv;

                i = i + 1;
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

                    oRreviewersArr[i] = oRvwr;

                    i = i + 1;

                }

                FileReviewers = oRreviewersArr;
            }


        }

        //publish/approve file in sharepoint so all users having read permission can view the sop 
        public bool PublishFile()
        {

            bool pdone = false;
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
                ErrorMessage = ".It is already published";

            }


                //UnPublish the file
                // file.UnPublish(String.Concat("File UnPublishing at ", DateTime.Now.ToLongDateString()));



                return pdone;
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