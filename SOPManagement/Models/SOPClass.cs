﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop;
using System.Collections;
using Microsoft.SharePoint.Client;
using System.Security;

namespace SOPManagement.Models
{
    public class SOPClass
    {

        public int FileID { get; set; }

        public int FileOwnerID { get; set; }

        public int FileApproverID { get; set; }

        public string FileApproverEmail { get; set; }

        public string FileOwnerEmail { get; set; }

        public string FileTitle { get; set; }   //title is without sopno

        public string FileName { get; set; }  //with sopno in front SOPNO + " "+ FileTitle

        public byte[] FileStream { get; set; }  //with sopno in front SOPNO + " "+ FileTitle

        public string FolderName { get; set; }

        public string SubFolderName { get; set; }

        public string SOPNo { get; set; }

        public FileRevision[] FileRevisions { get; set; }

        public string FileCurrVersion { get; set; }

        public short Updatefreq { get; set; }

        public string Updatefrequnit { get; set; }

        public string FileLink { get; set; }

        public string FilePath { get; set; }

        public bool OperationSuccess { get; set; }

        public Employee[] Reviewers { get; set; }

        public Employee[] Viewers { get; set; }

        public string DocumentLibName { get; set; }

        public string FileUrl { get; set; }
        public string SiteUrl { get; set; }

        public DateTime SOPEffectiveDate { get; set; }

        public HttpPostedFileBase UploadedFile { get; set; }

        public bool FileUploaded { get; set; }

        public string ErrorMessage { get; set; }


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

                    OperationSuccess = true;
                }

            }

            emp = null;


        }


        public void AddFileApprover()

        {
            //now insert approver into approver table
            Employee emp = new Employee();
            int apprvrid;
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

                OperationSuccess = true;

            }


        }



        public void AddFileOwner()



        {

            //now insert file owner into owner table

            Employee emp = new Employee();
            int apprvrid;
            OperationSuccess = false;

            emp.useremailaddress = FileOwnerEmail;
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
                OperationSuccess = true;
            }




        }

        public void AddUpdateFreq()

        {
            //now insert update frequence
            OperationSuccess = false;
            DateTime freqschdl = DateTime.Now;

            switch (Updatefrequnit)
            {
                case "Yearly":
                    freqschdl = freqschdl.AddYears(Updatefreq);
                    break;
                case "Monthly":
                    freqschdl = freqschdl.AddMonths(Updatefreq);
                    break;
                case "Weekly":
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
                    lastrevisionno = "1.0",
                    scheduledatetime = freqschdl


                };
                dbcontext.fileupdateschedules.Add(updfreqtable);

                dbcontext.SaveChanges();

                OperationSuccess = true;

            }


        }

        public void UpdateCoverRevhistPage()
        {

                       
            //  string newfilename = SOPNo + " " + SOPFileTitle;


            string savePath = HttpContext.Current.Server.MapPath("~/Content/docfiles/" + FileName);
            object missObj = System.Reflection.Missing.Value;
            object path = savePath;
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            try
            {


                System.IO.File.Copy(HttpContext.Current.Server.MapPath("~/Content/docfiles/SOPTemp.docx"), HttpContext.Current.Server.MapPath("~/Content/docfiles/" + FileName), true);

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
                    tab1.Rows[3].Cells[2].Range.Text = SOPEffectiveDate.ToShortDateString();
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


                    donotaddrow = false;

                    int filevercount = 1;

                    foreach (FileRevision rev in FileRevisions)
                    {

                        if (selectedRow == 2 && filevercount == 1)
                        {
                            //if (tab.Rows[tab.Rows.Count].Cells[1].Range.Text == "" || tab.Rows[tab.Rows.Count].Cells[1].Range.Text == "\r\a")
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
                        tab.Rows[tab.Rows.Count].Cells[1].Range.Text = rev.RevisionNo;
                        tab.Rows[tab.Rows.Count].Cells[2].Range.Text = rev.RevisionDate.ToString("M/d/yyyy");
                        tab.Rows[tab.Rows.Count].Cells[3].Range.Text = rev.Description;

                        filevercount = filevercount + 1;

                    }

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


        public void UploadDocument()
        {



            using (ClientContext clientContext = new ClientContext(SiteUrl))
            {

                string userName = "tshaikh@radiantdelivers.com";
                string password = "bdkbg88#";


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

                fileCreationInformation.Url = SiteUrl + FileUrl + FileName;
                Microsoft.SharePoint.Client.File uploadFile = documentsList.RootFolder.Files.Add(fileCreationInformation);

                //Update the metadata for a field having name "SOPNO"
                string loginname = "tshaikh@radiantdelivers.com";

                User theUser = clientContext.Web.SiteUsers.GetByEmail(loginname);


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

        private static SecureString GetSecureString(String Password)
        {
            SecureString oSecurePassword = new SecureString();

            foreach (Char c in Password.ToCharArray())
            {
                oSecurePassword.AppendChar(c);

            }
            return oSecurePassword;
        }


        public void RemoveAllFilePermissions()
        {


            OperationSuccess = false;

            ClientContext clientContext = new ClientContext(SiteUrl);

            string userName = "tshaikh@radiantdelivers.com";
            string password = "bdkbg88#";



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

        public void AssignFilePermission(string operation, string plabel, string useremail)
        {

            ClientContext clientContext = new ClientContext(SiteUrl);

            string userName = "tshaikh@radiantdelivers.com";
            string password = "bdkbg88#";

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
                    item.BreakRoleInheritance(false, false);


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

            string userName = "tshaikh@radiantdelivers.com";
            string password = "bdkbg88#";

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

                    RoleDefinitionBindingCollection rd = new RoleDefinitionBindingCollection(clientContext);
                    rd.Add(clientContext.Web.RoleDefinitions.GetByName(plabel));


                    // Microsoft.SharePoint.Client.GroupCollection groupCollection = web.SiteGroups;
                    Principal user;

                    // Group grpvisitor = groupCollection.GetByName("Watercooler Visitors");
                    // clientContext.Load(grpvisitor);



                    foreach (Employee emp in employees)
                    {

                        user = clientContext.Web.EnsureUser(emp.useremailaddress);
                        item.BreakRoleInheritance(false, false);

                        if (operation == "add")
                        {
                            item.RoleAssignments.Add(user, rd);
                        }
                        else if (operation == "remove")
                        {

                            item.RoleAssignments.GetByPrincipal(user).DeleteObject();

                        }

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
        public void GetSOPNo()
        {

            RadiantSOPEntities ctx = new RadiantSOPEntities();

            SOPNo = ctx.GetLastSOPNO(FolderName, SubFolderName).FirstOrDefault().ToString();


        }

    } //end of class
}  //end of namespace