using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using Microsoft.SharePoint.Client;
using System.Security;


namespace SOPManagement.Controllers
{
    public class HomeController : Controller
    {
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


        public void ExportToWord(string Title, string SOPNO)
        {

            

            string savePath = Server.MapPath("~/Content/docfiles/SOPFile.docx");
            string templatePath = Server.MapPath("~/Content/DocFiles/SOPTmplt.docx");
            Microsoft.Office.Interop.Word.Application wapp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();
            doc = wapp.Documents.Open(templatePath);
            doc.Activate();
            // doc.Tables[1].Rows.Add();


            //doc.Tables[1].Rows.Add(
            //doc.Tables[1].Rows[4]);

            //get value from sql database

            if (doc.Bookmarks.Exists("filetitle"))
            {
                doc.Bookmarks["filetitle"].Range.Text = Title;
            }
            if (doc.Bookmarks.Exists("sopno"))
            {
                doc.Bookmarks["sopno"].Range.Text = SOPNO;
            }
            if (doc.Bookmarks.Exists("revno"))
            {
                // doc.Bookmarks["Time"].Range.Text = DateTime.Now.ToString("yyyy-MM-dd");
                doc.Bookmarks["revno"].Range.Text = "1.0";
            }

        
            //this.Application.ActiveDocument.Tables[1].Rows[1]);



            doc.SaveAs2(savePath);
            wapp.Application.Quit();


          //  add row in table and data in cell

            object missObj = System.Reflection.Missing.Value;
            object path = savePath;
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wdoc = app.Documents.Open(ref path, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj);

            // Select the last table.
            // For this demo, the table has size of 3*4 (row * column).
            int totalTables = wdoc.Tables.Count;
            // Microsoft.Office.Interop.Word.Table tab = wdoc.Tables[totalTables];
            Microsoft.Office.Interop.Word.Table tab = wdoc.Tables[totalTables];
            Microsoft.Office.Interop.Word.Range range = tab.Range;

            // Select the last row as source row.
            int selectedRow = tab.Rows.Count;

            // Select and copy content of the source row.
            range.Start = tab.Rows[selectedRow].Cells[1].Range.Start;
            range.End = tab.Rows[selectedRow].Cells[tab.Rows[selectedRow].Cells.Count].Range.End;
            range.Copy();

            // Insert a new row after the last row if it is not first row to add data

            if (selectedRow>2 )
                tab.Rows.Add(ref missObj);

            // Moves the cursor to the first cell of target row.
            range.Start = tab.Rows[tab.Rows.Count].Cells[1].Range.Start;
            range.End = range.Start;

            // Paste values to target row.
            range.Paste();

            // Write new vaules to each cell.
            tab.Rows[tab.Rows.Count].Cells[1].Range.Text = "new row";
            tab.Rows[tab.Rows.Count].Cells[2].Range.Text = "cell 2";
            tab.Rows[tab.Rows.Count].Cells[2].Range.Text = "cell 3";


            wdoc.SaveAs2(savePath);
            app.Application.Quit();


            //1. upload file to sharepoint 
            //2. update persmission in sql server . first add reviewers, owner, approver of file in sql server table
            //3. add file level permissiom in sharepoint i.e. owner full, reviewer and approver edit permission
            //if there is error in 1 then don't proceed. If there is error in 2 then give retry and update file permission from sql table.
            //if still error then rollback in 1 and notify user to redo or contact admin. 
            
            

            //1.uploading the processed doc file to sharepoint

            string siteurl = "https://radiantdelivers.sharepoint.com/sites/watercooler";
            string documentlistname = "SOP";
            string documentlistUrl = "SOP/Warehouse Operations/";
            string documentname = Title;
            byte[] stream = System.IO.File.ReadAllBytes(savePath);
            UploadDocument(siteurl, documentlistname, documentlistUrl, documentname, stream,SOPNO);

           



            Response.Write("Success");
        }



        private void UploadDocument(string siteURL, string documentListName, string documentListURL, string documentName, byte[] documentStream,string pSOPNO)
        {


         //   ClientContext clientContext = new ClientContext(siteurl);
           // SecureString SecurePassword = GetSecureString(Password);
         //   clientContext.Credentials = new SharePointOnlineCredentials(username, SecurePassword);


            using (ClientContext clientContext = new ClientContext(siteURL))
            {

                string userName = "tshaikh@radiantdelivers.com";
                string password = "bagerhat79&";
                

                SecureString SecurePassword = GetSecureString(password);
                clientContext.Credentials = new SharePointOnlineCredentials(userName, SecurePassword);

                //ICredentials credentials = new NetworkCredential(userName, password, domain);
                //clientContext.Credentials = credentials;

                List documentsList = clientContext.Web.Lists.GetByTitle(documentListName);

                var fileCreationInformation = new FileCreationInformation();
                //Assign to content byte[] i.e. documentStream

                fileCreationInformation.Content = documentStream;
                //Allow owerwrite of document

                fileCreationInformation.Overwrite = true;
                //Upload URL
                documentName = documentName + ".docx";
                fileCreationInformation.Url = siteURL + documentListURL + documentName;
                Microsoft.SharePoint.Client.File uploadFile = documentsList.RootFolder.Files.Add(fileCreationInformation);

                //Update the metadata for a field having name "SOPNO"
                string loginname= "tshaikh@radiantdelivers.com";

                User theUser = clientContext.Web.SiteUsers.GetByEmail(loginname);


                uploadFile.ListItemAllFields["Owner"] = theUser;

                

                uploadFile.ListItemAllFields["SOPNO"] = pSOPNO;


                uploadFile.ListItemAllFields.Update();
                clientContext.ExecuteQuery();


                //assign permission to the file

                var list = clientContext.Web.Lists.GetByTitle("TestDocLibrary");






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

    }  //end of class HomeController


   
}