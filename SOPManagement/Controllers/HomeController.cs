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

            //1.create requested word document from tempate with cover page
            //2. upload file to sharepoint 
            //3. update persmission in sql server . first add reviewers, owner, approver of file in sql server table
            //4. add file level permissiom in sharepoint i.e. owner full, reviewer and approver edit permission
            //if there is error in 1 then don't proceed. If there is error in 2 then give retry and update file permission from sql table.
            //if still error then rollback in 1 and notify user to redo or contact admin. 


            string savePath = Server.MapPath("~/Content/docfiles/SOPFile.docx");
            string templatePath = Server.MapPath("~/Content/DocFiles/SOPTmplt.docx");


            //Merge file

             string originaldoc = Server.MapPath("~/Content/docfiles/SOPFile.docx"); ;

            string outputfilepath;
            string[] filestomerge = new string[2];
            System.IO.File.Copy(Server.MapPath("~/Content/docfiles/SOPCoverTmplt.docx"), Server.MapPath("~/Content/docfiles/SOPFile.docx"),true);

            filestomerge[0]= Server.MapPath("~/Content/docfiles/SOPFile.docx");
            filestomerge[1]= Server.MapPath("~/Content/docfiles/SOPBodyTmplt.docx");
            outputfilepath= Server.MapPath("~/Content/docfiles/MergedFile.docx");

               MergeDoc(filestomerge[0], filestomerge[1], outputfilepath);


            //MergeWordFiles(filestomerge);

            //1.create requested word document from tempate with cover page

            //CreateDocFromTemplate(Title, SOPNO);


            //2.uploading the processed doc file to sharepoint

            string siteurl = "https://radiantdelivers.sharepoint.com/sites/watercooler";
            string documentlistname = "SOP";
            string documentlistUrl = "SOP/Warehouse Operations/";
            string documentname = "SOPTfffeeee.docx";      // Title;

            //byte[] stream = System.IO.File.ReadAllBytes(savePath);

          //  UploadDocument(siteurl, documentlistname, documentlistUrl, documentname, stream,SOPNO);

            //3. insert permision in sql server

            //4.assign file permission
         //   AssignFilePermission(siteurl, documentlistname, documentlistUrl, documentname,"add");


            Response.Write("Success");
        }




        public virtual Byte[] MergeWordFiles(string[] sourceFiles)
        {
            int f = 0;
            // If only one Word document then skip merge.
            if (sourceFiles.Count() == 1)
            {
                return System.IO.File.ReadAllBytes(sourceFiles[0]);
            }
            else
            {
                MemoryStream destinationFile = new MemoryStream();

                // Add first file
                var firstFile = sourceFiles[0];

                destinationFile.Write(System.IO.File.ReadAllBytes(firstFile), 0, firstFile.Length);
                destinationFile.Position = 0;

                int pointer = 1;
                byte[] ret;

                // Add the rest of the files
                try
                {
                    using (WordprocessingDocument mainDocument = WordprocessingDocument.Open(new MemoryStream(System.IO.File.ReadAllBytes(firstFile)), true))
                    {
                        System.Xml.Linq.XElement newBody = XElement.Parse(mainDocument.MainDocumentPart.Document.Body.OuterXml);

                        for (pointer = 1; pointer < sourceFiles.Count(); pointer++)
                        {

                            var sFile = sourceFiles[pointer];
                            WordprocessingDocument tempDocument = WordprocessingDocument.Open(new MemoryStream(System.IO.File.ReadAllBytes(sFile)), true);
                            XElement tempBody = XElement.Parse(tempDocument.MainDocumentPart.Document.Body.OuterXml);
                            newBody.Add(XElement.Parse(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run(new Break { Type = BreakValues.Page })).OuterXml));
                            newBody.Add(tempBody);

                            mainDocument.MainDocumentPart.Document.Body = new Body(newBody.ToString());
                            mainDocument.MainDocumentPart.Document.Save();
                            mainDocument.Package.Flush();
                        }
                    }
                }
                catch (OpenXmlPackageException oxmle)
                {
                    throw new Exception(string.Format(CultureInfo.CurrentCulture, "Error while merging files. Document index {0}", pointer), oxmle);
                }
                catch (Exception e)
                {
                    throw new Exception(string.Format(CultureInfo.CurrentCulture, "Error while merging files. Document index {0}", pointer), e);
                }
                finally
                {
                    ret = destinationFile.ToArray();
                    destinationFile.Close();
                    destinationFile.Dispose();
                }

                return ret;
            }
        }


        private void MergeDoc(string coverdoc, string bodydoc, string stroutdoc)
        {


            //byte[] word1 = System.IO.File.ReadAllBytes(coverdoc);
            //byte[] word2 = System.IO.File.ReadAllBytes(bodydoc);

            //byte[] result = Merge(word1, word2);

            //System.IO.File.WriteAllBytes(stroutdoc, result);


            using (WordprocessingDocument myDoc =
                WordprocessingDocument.Open(coverdoc, true))
            {
                string altChunkId = "AltChunkId" + DateTime.Now.Ticks.ToString();
                MainDocumentPart mainPart = myDoc.MainDocumentPart;
                AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(
                    AlternativeFormatImportPartType.WordprocessingML, altChunkId);

                using (FileStream fileStream = System.IO.File.Open(bodydoc, FileMode.Open))
                    chunk.FeedData(fileStream);
                AltChunk altChunk = new AltChunk();
                altChunk.Id = altChunkId;
                mainPart.Document
                    .Body
                    .InsertAfter(altChunk, mainPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Last());
                mainPart.Document.Save();
            }



//            Using myDoc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open("D:\\Test.docx", True)
//        Dim altChunkId = "AltChunkId" + DateTime.Now.Ticks.ToString().Substring(0, 2)
//        Dim mainPart = myDoc.MainDocumentPart
//        Dim chunk = mainPart.AddAlternativeFormatImportPart(
//            DocumentFormat.OpenXml.Packaging.AlternativeFormatImportPartType.WordprocessingML, altChunkId)
//        Using fileStream As IO.FileStream = IO.File.Open("D:\\Test1.docx", IO.FileMode.Open)
//            chunk.FeedData(fileStream)
//        End Using
//        Dim altChunk = New DocumentFormat.OpenXml.Wordprocessing.AltChunk()
//        altChunk.Id = altChunkId
//        mainPart.Document.Body.InsertAfter(altChunk, mainPart.Document.Body.Elements(Of DocumentFormat.OpenXml.Wordprocessing.Paragraph).Last())
//        mainPart.Document.Save()
//End Using


        }


        private static byte[] Merge(byte[] dest, byte[] src)
        {
            string altChunkId = "AltChunkId" + DateTime.Now.Ticks.ToString();

            var memoryStreamDest = new MemoryStream();
            memoryStreamDest.Write(dest, 0, dest.Length);
            memoryStreamDest.Seek(0, SeekOrigin.Begin);
            var memoryStreamSrc = new MemoryStream(src);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStreamDest, true))
            {
                MainDocumentPart mainPart = doc.MainDocumentPart;
                AlternativeFormatImportPart altPart =
                    mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, altChunkId);
                altPart.FeedData(memoryStreamSrc);
                var altChunk = new AltChunk();
                altChunk.Id = altChunkId;
                DocumentFormat.OpenXml.OpenXmlElement lastElem = mainPart.Document.Body.Elements<AltChunk>().LastOrDefault();
                if (lastElem == null)
                {
                    lastElem = mainPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Last();
                }


                //Page Brake einfügen
                DocumentFormat.OpenXml.Wordprocessing.Paragraph pageBreakP = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
                Run pageBreakR = new Run();
                Break pageBreakBr = new Break() { Type = BreakValues.Page };

                pageBreakP.Append(pageBreakR);
                pageBreakR.Append(pageBreakBr);

                return memoryStreamDest.ToArray();
            }




        }



            private void CreateDocFromTemplate(string filetitle,string sopno)
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
                doc.Bookmarks["filetitle"].Range.Text = filetitle;
            }

            if (doc.Bookmarks.Exists("filetitleh"))
            {
                doc.Bookmarks["filetitleh"].Range.Text = filetitle;
            }


            if (doc.Bookmarks.Exists("sopno"))
            {
                doc.Bookmarks["sopno"].Range.Text = sopno;  //sopno + " ";
            }

            if (doc.Bookmarks.Exists("sopnoh"))
            {
                doc.Bookmarks["sopnoh"].Range.Text = sopno;
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

            if (selectedRow > 2)
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



        }

        private void AssignFilePermission(string siteURL, string documentListName, string documentListURL, string documentName, string operation)
        {

            var clientContext = new ClientContext(siteURL);

            string userName = "tshaikh@radiantdelivers.com";
            string password = "bagerhat79&";

            SecureString SecurePassword = GetSecureString(password);
            clientContext.Credentials = new SharePointOnlineCredentials(userName, SecurePassword);


            Web web = clientContext.Web;
            clientContext.Load(web);
            clientContext.Load(web.Lists);
            clientContext.Load(web, wb => wb.ServerRelativeUrl);
            clientContext.ExecuteQuery();

            Microsoft.SharePoint.Client.List list = web.Lists.GetByTitle(documentListName);
            clientContext.Load(list);
            clientContext.ExecuteQuery();

            Folder folder = web.GetFolderByServerRelativeUrl(web.ServerRelativeUrl + documentListURL);
            clientContext.Load(folder);
            clientContext.ExecuteQuery();

            CamlQuery camlQuery = new CamlQuery();

            ////TO GET ONLY FILE ITEM
            //camlQuery.ViewXml = @"<View Scope='Recursive'>
            //                         <Query>

            //                        <Where>
            //                             <Contains>
            //                                 <FieldRef Name='FileLeafRef'/>
            //                                     <Value Type='File'>SOPStudent01.docx</Value>
            //                                </Contains>
            //                            </Where>

            //                         </Query>
            //                     </View>";


            //TO GET ONLY FILE ITEM
            camlQuery.ViewXml = "<View Scope='Recursive'> "+
                                   "  <Query> "+

                                  " + <Where> " +
                                       "  <Contains>"+
                                            " <FieldRef Name='FileLeafRef'/> "+
                                                " <Value Type='File'>"+ documentName + "</Value>"+
                                           " </Contains> " +
                                       " </Where> " +

                                    " </Query> "+
                                " </View>";

            //TO GET ALL FOLDERS AND FILE ITEMS
            //camlQuery.ViewXml = @"<View Scope='RecursiveAll'>
            //                         <Query>
            //                         </Query>
            //                     </View>";


            camlQuery.FolderServerRelativeUrl = folder.ServerRelativeUrl;
            ListItemCollection listItems = list.GetItems(camlQuery);
            clientContext.Load(listItems);
            clientContext.ExecuteQuery();


            string loginname = "student05@radiantdelivers.com";

            foreach (ListItem item in listItems)
            {
                //item.FileSystemObjectType;

                if (item.FileSystemObjectType == FileSystemObjectType.File)
                {
                    // This is a File



                    RoleDefinitionBindingCollection rd = new RoleDefinitionBindingCollection(clientContext);

                    //if (!item.HasUniqueRoleAssignments)
                    //{
                    //    item.BreakRoleInheritance(false, false);
                    //}

                    rd.Add(clientContext.Web.RoleDefinitions.GetByName("Contribute"));
                    Principal user = clientContext.Web.EnsureUser(loginname);
                    item.BreakRoleInheritance(false, false);


                    //Get the list of Role Assignments to list item and remove one by one.

                    //RoleAssignmentCollection SPRoleAssColn = item.RoleAssignments;

                    //clientContext.ExecuteQuery();

                    ////   for (int i = SPRoleAssColn.Count - 1; i >= 0; i--)

                    //foreach (RoleAssignment ri in SPRoleAssColn)
                    //{

                    //    ri.DeleteObject();

                    //}

                    if (operation == "add")
                    {
                        item.RoleAssignments.Add(user, rd);
                    }
                    else if (operation=="remove")
                    {
                        item.RoleAssignments.GetByPrincipal(user).DeleteObject();

                    }

                    item.Update();
                    clientContext.ExecuteQuery();

                }
                else if (item.FileSystemObjectType == FileSystemObjectType.Folder)
                {
                    // This is a  Folder
                }




            }



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

                Microsoft.SharePoint.Client.List documentsList = clientContext.Web.Lists.GetByTitle(documentListName);

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
                clientContext.ExecuteQuery(); //upload file

     

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