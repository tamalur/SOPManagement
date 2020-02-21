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



            //if (postedFile != null)

            //{

            //    string path = Server.MapPath("~/Content/DocFiles");

            //    if (!Directory.Exists(path))

            //    {

            //        Directory.CreateDirectory(path);

            //    }



            //    postedFile.SaveAs(path + Path.GetFileName(postedFile.FileName));

            //    ViewBag.Message = "File uploaded successfully.";

            //}



            ////Create db context object here 

            //RadiantYYZEntities2 dbContext = new RadiantYYZEntities2();


            ////Get the value from database and then set it to ViewBag to pass it View
            //IEnumerable<SelectListItem> ItemFolders = dbContext.vwDepartmentFolders.Select(c => new SelectListItem
            //{
            //    Value = c.SPFilePath,
            //    Text = c.DeptFileName

            //}).Where(q=>q.Value == "SOP/");

            //ViewBag.ddlDeptFolders = ItemFolders;

            ////ViewBag.ddlDeptFolders = dbContext.vwDepartmentFolders.Where(i=>i.FileID==170).FirstOrDefault();

            ////string strSelFolder = Request.Form["ddlDeptFolders"].ToString();

            ////Get the value from database and then set it to ViewBag to pass it View
            //IEnumerable<SelectListItem> itemSubFolders = dbContext.vwDepartmentSubFolders.Select(c => new SelectListItem
            //{
            //    Value = c.SPFilePath,
            //    Text = c.DeptFileName

            //}).Where(q =>q.Text == "OPS");

            //ViewBag.ddlSubFolders = itemSubFolders;


            RadiantYYZEntities2 edm = new RadiantYYZEntities2();

            ViewBag.ddlDeptFolders = new SelectList(GetFolders(), "DeptFileName", "DeptFileName");
            return View();
        }


        public List<deptsopfile> GetFolders()
        {

            List<deptsopfile> folderlist;

           

            using (var ctx = new RadiantYYZEntities2())
            {
                var folders = ctx.deptsopfiles
                                .Where(s => s.SPFilePath == "SOP/");

                folderlist = folders.ToList();

            }


            return folderlist;


            //List<deptsopfile> folders = (from d in deptsopfile
            //                             where d.Tags.All(t => _tags.Contains(t))
            //                    select d.id).ToList<int>();

            //List<deptsopfile> folders = new List<deptsopfile>
            //    return folders;


        }

        public ActionResult GetSubFolderList(string foldername)
        {


            List<deptsopfile> subfolderlist;

            using (var ctx = new RadiantYYZEntities2())
            {
                var subfolders = ctx.deptsopfiles
                                .Where(s => s.SPFilePath == "SOP/"+ foldername +"/" && !s.DeptFileName.Contains(".docx"));

                subfolderlist = subfolders.ToList();

                ViewBag.ddlSubFolders= new SelectList(subfolderlist, "FileID", "DeptFileName");

            }

               return PartialView("DisplaySubfolders");




        }

        //public ViewResult 

        public void UploadFile(HttpPostedFileBase postedFile, string deptfoldername, string subfoldername, string sopno)

        {



            ViewBag.Message = "Upload SOP File";

            


            if (postedFile != null)

            {

                string path = Server.MapPath("~/Content/DocFiles/");

                if (!Directory.Exists(path))

                {

                    Directory.CreateDirectory(path);

                }



                postedFile.SaveAs(path + Path.GetFileName(postedFile.FileName));

                ViewBag.Message = "File uploaded successfully.";


                ArrayList arrvers;

                string filerpath = "/sites/watercooler/SOP/Quality Assurance & Regulatory Affairs (QRA)/QRA  (AIB)/OPS07-01 Training and Personnel.docx" ;
                arrvers = getFileVersions(siteurl, filerpath);



            }








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
            string templatePath = Server.MapPath("~/Content/DocFiles/SOPTemplate.docx");


            //Merge file

            // string originaldoc = Server.MapPath("~/Content/docfiles/SOPFile.docx"); ;

            //string outputfilepath;
            //string[] filestomerge = new string[2];
            //System.IO.File.Copy(Server.MapPath("~/Content/docfiles/SOPCoverTmplt.docx"), Server.MapPath("~/Content/docfiles/SOPFile.docx"),true);

            //filestomerge[0]= Server.MapPath("~/Content/docfiles/SOPFile.docx");
            //filestomerge[1]= Server.MapPath("~/Content/docfiles/SOPBodyTmplt.docx");
            //outputfilepath= Server.MapPath("~/Content/docfiles/MergedFile.docx");

            //string iChunkId = "AltChunkId" + DateTime.Now.Ticks.ToString();

            //   MergeDocumentWithPagebreak(filestomerge[0], filestomerge[1], iChunkId);
            //  MergeDoc(filestomerge[0], filestomerge[1], outputfilepath);


            //MergeWordFiles(filestomerge);


            string permissionlabel = "Contribute";  // this will come from query string
            string useremail = "tshaikh@radiantdelivers.com"; //this come from windows logged name

            //1.create requested word document from tempate with cover page




            CreateDocFromTemplate(Title, SOPNO);


            //2.uploading the processed doc file to sharepoint

            string documentlistname = "SOP";
            string documentlistUrl = "SOP/Quality Assurance & Regulatory Affairs (QRA)/QRA  (AIB)/";
            string documentname = Title;   //"SOPFile";      // Title;

           // string filerpath = "/sites/watercooler/SOP/Quality Assurance & Regulatory Affairs (QRA)/QRA  (AIB)/OPS07-01 Training and Personnel.docx";


            byte[] stream = System.IO.File.ReadAllBytes(savePath);

             UploadDocument(siteurl, documentlistname, documentlistUrl, documentname, stream,SOPNO);

            ArrayList arrvers;
            string filerpath = "/sites/watercooler/SOP/Warehouse Operations/" + Title + ".docx";
            arrvers = getFileVersions(siteurl, filerpath);


            //3. insert permision in sql server

            //4.assign file permission
           //  AssignFilePermission(siteurl, documentlistname, documentlistUrl, documentname,"add",permissionlabel, useremail);


            //     GetFileVersions(siteurl, documentlistname, documentlistUrl, documentname);

            Response.Write("Success");
        }



        /// <summary>
        /// Merge word document with page break in-between
        /// </summary>
        /// <param name=”sourceFile”></param>
        /// <param name=”destinationFile”></param>
        /// <param name=”AltChunkID”></param>
        private void MergeDocumentWithPagebreak(string sourceFile, string destinationFile, string AltChunkID)
        {
            using (WordprocessingDocument myDoc =
            WordprocessingDocument.Open(sourceFile, true))
            {

                string altChunkId = AltChunkID;
                MainDocumentPart mainPart = myDoc.MainDocumentPart;

                //Append page break
                DocumentFormat.OpenXml.Wordprocessing.Paragraph para = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run((new Break() { Type = BreakValues.Page })));
                mainPart.Document.Body.InsertAfter(para, mainPart.Document.Body.LastChild);

                //Append file
                AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(
                 AlternativeFormatImportPartType.WordprocessingML, altChunkId);
                using (FileStream fileStream = System.IO.File.Open(destinationFile, FileMode.Open))
                    chunk.FeedData(fileStream);
                AltChunk altChunk = new AltChunk();
                altChunk.Id = altChunkId;
                mainPart.Document
                .Body
                .InsertAfter(altChunk, mainPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Last());
                mainPart.Document.Save();

            }
        }



        private void MergeDoc(string coverdoc, string bodydoc, string stroutdoc)
        {


            using (WordprocessingDocument myDoc =
                WordprocessingDocument.Open(coverdoc, true))
            {
                string altChunkId = "AltChunkId" + DateTime.Now.Ticks.ToString();
                MainDocumentPart mainPart = myDoc.MainDocumentPart;


                //Append page break
                DocumentFormat.OpenXml.Wordprocessing.Paragraph para = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run((new Break() { Type = BreakValues.Page })));
                mainPart.Document.Body.InsertAfter(para, mainPart.Document.Body.LastChild);

                //append file

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



        }


        //private void DownlodFile(string url,string listTitle)
        //{
        //    using (var clientContext = new ClientContext(url))
        //    {

        //        var list = clientContext.Web.Lists.GetByTitle(listTitle);
        //        //var listItem = list.GetItemById(listItemId);
        //        clientContext.Load(list);
        //       // clientContext.Load(listItem, i => i.File);
        //        clientContext.ExecuteQuery();

        //        //var fileRef = listItem.File.ServerRelativeUrl;
        //        var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, fileRef);
        //        //var fileName = Path.Combine(filePath, (string)listItem.File.Name);
        //        using (var fileStream = System.IO.File.Create(fileName)) ;
        //        //{
        //        //    fileInfo.Stream.CopyTo(fileStream);
        //        //}
        //    }

        //}

 
            private void CreateDocFromTemplate(string filetitle,string sopno)
        {


            //     System.IO.File.Copy(Server.MapPath("~/Content/docfiles/SOPTemplate.docx"), Server.MapPath("~/Content/docfiles/SOPFile.docx"), true);

            //    System.IO.File.Copy(Server.MapPath("~/Content/docfiles/SOPFile.docx"), Server.MapPath("~/Content/docfiles/SOPTemp.docx"), true);

            //     string savePath = Server.MapPath("~/Content/docfiles/SOPFile.docx");
            //    string templatePath = Server.MapPath("~/Content/DocFiles/SOPTemp.docx");


            //  Microsoft.Office.Interop.Word.Application wapp = new Microsoft.Office.Interop.Word.Application();
            //Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();

            //doc = wapp.Documents.Open(templatePath);   //open template and update it
            //doc.Activate();
            //// doc.Tables[1].Rows.Add();


            ////doc.Tables[1].Rows.Add(
            ////doc.Tables[1].Rows[4]);

            ////get value from sql database

            //if (doc.Bookmarks.Exists("filetitle"))
            //{

            //    //var start = doc.Bookmarks["filetitle"].Start;
            //    //var end = doc.Bookmarks["filetitle"].End;

            //    //Microsoft.Office.Interop.Word.Range range = doc.Range(start, end);

            //    doc.Bookmarks["filetitle"].Range.Text = filetitle;


            //    //object missO = System.Reflection.Missing.Value;

            //    //Microsoft.Office.Interop.Word.Range range = wapp.ActiveDocument.Content;

            //    //Microsoft.Office.Interop.Word.Find find = range.Find;
            //    //find.Text = filetitle;
            //    //find.ClearFormatting();
            //    //find.Execute(ref missO, ref missO, ref missO, ref missO, ref missO,
            //    //    ref missO, ref missO, ref missO, ref missO, ref missO,
            //    //    ref missO, ref missO, ref missO, ref missO, ref missO);



            //   // Microsoft.Office.Interop.Word.Range range = doc.Range(start, end);

            //    //doc.Bookmarks.Add("filetitle", range);


            //}

            //if (doc.Bookmarks.Exists("filetitleh"))
            //{
            //    doc.Bookmarks["filetitleh"].Range.Text = filetitle;
            //}


            //if (doc.Bookmarks.Exists("sopno"))
            //{
            //    doc.Bookmarks["sopno"].Range.Text = sopno;  //sopno + " ";
            //}

            //if (doc.Bookmarks.Exists("sopnoh"))
            //{
            //    doc.Bookmarks["sopnoh"].Range.Text = sopno;
            //}

            //if (doc.Bookmarks.Exists("revno"))
            //{
            //    // doc.Bookmarks["Time"].Range.Text = DateTime.Now.ToString("yyyy-MM-dd");
            //    doc.Bookmarks["revno"].Range.Text = "1.0";
            //}


            ////this.Application.ActiveDocument.Tables[1].Rows[1]);



            //doc.SaveAs2(savePath);
            //wapp.Application.Quit();




          //  System.IO.File.Copy(Server.MapPath("~/Content/docfiles/SOPTemplate.docx"), Server.MapPath("~/Content/docfiles/SOPFile.docx"), true);

          System.IO.File.Copy(Server.MapPath("~/Content/docfiles/SOPTemp.docx"), Server.MapPath("~/Content/docfiles/SOPFile.docx"), true);

            string savePath = Server.MapPath("~/Content/docfiles/SOPFile.docx");


            //  add row in table and data in cell

            object missObj = System.Reflection.Missing.Value;
            object path = savePath;
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wdoc = app.Documents.Open(ref path, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj);

            
            
            int totalTables = wdoc.Tables.Count;
            bool donotaddrow = false;


            //Add data into reviewer table  - 2nd table in the cover page

            if (totalTables>0 )
            {


                //update first table in  cover page

                Microsoft.Office.Interop.Word.Table tab1 = wdoc.Tables[1];
                Microsoft.Office.Interop.Word.Range range1 = tab1.Range;

                // Select the last row as source row.
                int selectedRow1 = tab1.Rows.Count;

                // Write new vaules to each cell.
                tab1.Rows[1].Cells[2].Range.Text = filetitle;
                tab1.Rows[2].Cells[2].Range.Text = sopno;
                tab1.Rows[2].Cells[4].Range.Text = "1.0";
                //  tab1.Rows[tab1.Rows.Count].Cells[3].Range.Text = "cell 3";


                //update 2nd table in  cover page for updating reviewers

                Microsoft.Office.Interop.Word.Table tab2 = wdoc.Tables[2];          
                Microsoft.Office.Interop.Word.Range range2 = tab2.Range;

                // Select the last row as source row.
                int selectedRow2 = tab2.Rows.Count;

                // Select and copy content of the source row.
                range2.Start = tab2.Rows[selectedRow2].Cells[1].Range.Start;
                range2.End = tab2.Rows[selectedRow2].Cells[tab2.Rows[selectedRow2].Cells.Count].Range.End;
                range2.Copy();

                // Insert a new row after the last row if it is not first row to add data

                if (selectedRow2 >= 3)
                {
                    if (tab2.Rows[tab2.Rows.Count].Cells[1].Range.Text == "" || tab2.Rows[tab2.Rows.Count].Cells[1].Range.Text == "\r\a")
                        donotaddrow = true;

                    if (!donotaddrow)
                        tab2.Rows.Add(ref missObj);

                }

                // Moves the cursor to the first cell of target row.
                range2.Start = tab2.Rows[tab2.Rows.Count].Cells[1].Range.Start;
                range2.End = range2.Start;

                // Paste values to target row.
                range2.Paste();

                // Write new vaules to each cell.
                tab2.Rows[tab2.Rows.Count].Cells[1].Range.Text = "new row";
                tab2.Rows[tab2.Rows.Count].Cells[2].Range.Text = "cell 2";
                tab2.Rows[tab2.Rows.Count].Cells[3].Range.Text = "cell 3";



                //Add data into Revison history table - last table in last page

                Microsoft.Office.Interop.Word.Table tab = wdoc.Tables[totalTables];
                Microsoft.Office.Interop.Word.Range range = tab.Range;

                // Select the last row as source row.
                int selectedRow = tab.Rows.Count;

                // Select and copy content of the source row.
                range.Start = tab.Rows[selectedRow].Cells[1].Range.Start;
                range.End = tab.Rows[selectedRow].Cells[tab.Rows[selectedRow].Cells.Count].Range.End;
                range.Copy();

                // Insert a new row after the last row if it is not first row to add data

                //if (selectedRow >= 4)
                //    tab.Rows.Add(ref missObj);

                donotaddrow = false;

                if (selectedRow >= 2)
                {
                    if (tab.Rows[tab.Rows.Count].Cells[1].Range.Text == "" || tab.Rows[tab.Rows.Count].Cells[1].Range.Text == "\r\a")
                        donotaddrow = true;

                    if (!donotaddrow)
                        tab.Rows.Add(ref missObj);

                }


                // Moves the cursor to the first cell of target row.
                range.Start = tab.Rows[tab.Rows.Count].Cells[1].Range.Start;
                range.End = range.Start;

                // Paste values to target row.
                range.Paste();


                
                ArrayList vers=new ArrayList(new string[] { "1.0" });

                
                string filerpath = "SOP/Warehouse Operations/" + filetitle +".docx";
                //  vers = getFileVersions(siteurl, filerpath);

                //vers.Add(arr);

                foreach (string s in vers)
                {

                    // Write new vaules to each cell.
                    tab.Rows[tab.Rows.Count].Cells[1].Range.Text = s;
                    //tab.Rows[tab.Rows.Count].Cells[2].Range.Text = "cell 2";
                    //tab.Rows[tab.Rows.Count].Cells[3].Range.Text = "cell 3";

                }

                //// Write new vaules to each cell.
                //tab.Rows[tab.Rows.Count].Cells[1].Range.Text = "new row";
                //tab.Rows[tab.Rows.Count].Cells[2].Range.Text = "cell 2";
                //tab.Rows[tab.Rows.Count].Cells[3].Range.Text = "cell 3";



                //object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
                //foreach (Microsoft.Office.Interop.Word.Section section in wdoc.Sections)
                //{
                //    Microsoft.Office.Interop.Word.Range footerRange = section.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                //    footerRange.Find.Text = "<Title>";
                //    footerRange.Find.Replacement.Text = filetitle;
                //    footerRange.Find.Execute(ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref missObj, ref replaceAll, ref missObj, ref missObj, ref missObj, ref missObj);
                //}


            }


            //update footer

            ////Add the footers into the document
            //foreach (Microsoft.Office.Interop.Word.Section wordSection in wdoc.Sections)
            //{
            //    //Get the footer range and add the footer details.
            //    Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            //    //  footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
            //    //  footerRange.Font.Size = 10;

            //    footerRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);




            //    footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            //    footerRange.Text = filetitle +" "+sopno;

            //    //footerRange.Fields.Add(footerRange)





            //}


            //foreach (Microsoft.Office.Interop.Word.Section sec in wdoc.Sections)

            //{

            //  //  Microsoft.Office.Interop.Word.WdSeekView.wdSeekPrimaryFooter

            //    Microsoft.Office.Interop.Word.Range rng = sec.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            //    //  rng.Font.Name = "Arial";
            //    //  rng.Font.Size = 8;
            //    rng.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
            //    rng.Fields.Add(rng, Type: Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
            //}



            // Set footers
            foreach (Microsoft.Office.Interop.Word.Section wordSection in wdoc.Sections)
            {
                Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                footerRange.Tables[1].Cell(1, 1).Range.Text = sopno + " " + filetitle;



                // wdoc.Tables.Add(footerRange, 1, 2);
                //Object oMissing = System.Reflection.Missing.Value;



                // Object TotalPages = Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages;

                // Object CurrentPage = Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage;





                //  footerRange.Tables[1].Cell(1, 2).Range.Fields.Add(footerRange, ref CurrentPage, ref oMissing, ref oMissing);




                //footerRange.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd);
                //footerRange.Paragraphs.TabStops.Add(app.InchesToPoints(3.25F), Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabCenter);
                //footerRange.Paragraphs.TabStops.Add(app.InchesToPoints(6.5F), Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight);

                //footerRange.Fields.Add(footerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage, "\t", true);



                ////footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                ////footerRange.Font.Size = 20;
                //footerRange.Text = "    \t";
                ////footerRange.InsertBefore("01-DEC-18");



                //footerRange.InsertBefore(filetitle);


            }



            wdoc.SaveAs2(savePath);   //save in actual file from tempalte
            app.Application.Quit();



        }


        private ArrayList getFileVersions(string siteurl,string filerelpath)
        {

            ArrayList fversions = new ArrayList();

            //SOP / Warehouse Operations /


            using (ClientContext clientContext = new ClientContext(siteurl))
            {

                string userName = "tshaikh@radiantdelivers.com";
                string password = "bagerhat79&";


                SecureString SecurePassword = GetSecureString(password);
                clientContext.Credentials = new SharePointOnlineCredentials(userName, SecurePassword);

                Web site = clientContext.Web;
                clientContext.Load(site);
                // File file = site.GetFileByServerRelativeUrl("/Shared Documents/mydocument.doc");


                //FileVersionCollection versions;
                Microsoft.SharePoint.Client.File file = site.GetFileByServerRelativeUrl(filerelpath);

                clientContext.Load(file);

                clientContext.ExecuteQuery();


                string id;

                FileVersionCollection versions = file.Versions;

                clientContext.Load(versions);

                PropertyValues fi = file.Properties;
                
                clientContext.Load(fi);

            
                clientContext.ExecuteQuery();

                var lv = file.MajorVersion.ToString();


                id = fi["ID"].ToString();


         

                if (versions != null)
                {
                    foreach (FileVersion version in versions)
                    {
                        Console.WriteLine("Version : {0}", version.VersionLabel);

                        clientContext.Load(version);
                        clientContext.ExecuteQuery();


                        if ((Convert.ToDouble(version.VersionLabel) % 1) == 0)
                        {
                            //You can get all major versions here.

                            
                            fversions.Add(version.VersionLabel);

                        }


                    }
                }


            }


           

            return fversions;
        }

        private void AssignFilePermission(string siteURL, string documentListName, string documentListURL, string documentName, string operation, string plabel,string useremail)
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


            string loginname = useremail;

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

                    rd.Add(clientContext.Web.RoleDefinitions.GetByName(plabel));
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

        private void GetFileVersions(string siteURL, string documentListName, string documentListURL, string documentName)
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

  
            //TO GET ONLY FILE ITEM
            camlQuery.ViewXml = "<View Scope='Recursive'> " +
                                   "  <Query> " +

                                  " + <Where> " +
                                       "  <Contains>" +
                                            " <FieldRef Name='FileLeafRef'/> " +
                                                " <Value Type='File'>" + documentName + "</Value>" +
                                           " </Contains> " +
                                       " </Where> " +

                                    " </Query> " +
                                " </View>";

  
            camlQuery.FolderServerRelativeUrl = folder.ServerRelativeUrl;
            ListItemCollection listItems = list.GetItems(camlQuery);
            clientContext.Load(listItems);
            clientContext.ExecuteQuery();


            string fid;

            foreach (ListItem item in listItems)
            {
                //item.FileSystemObjectType;

                if (item.FileSystemObjectType == FileSystemObjectType.File)
                {
                    // This is the File

                    Microsoft.SharePoint.Client.File file = item.File;

                    FileVersionCollection versions = file.Versions;

                    fid=file.Properties["ID"].ToString();

                    clientContext.Load(file);
                    clientContext.Load(versions);
                    clientContext.ExecuteQuery();


                    //$file = $item.File
                    //versions = $file.Versions
                    //$ctx.Load($file)
                    //$ctx.Load($versions)
                    //$ctx.ExecuteQuery()


                    foreach(FileVersion v in versions)
                    {

                        clientContext.Load(v);
                        clientContext.ExecuteQuery();

                        var modifiedBy = v.CreatedBy;
                        clientContext.Load(modifiedBy);

                        clientContext.ExecuteQuery();

                        var loginnm=modifiedBy.LoginName;
                        var title = modifiedBy.Title;


                    }




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


                clientContext.Load(uploadFile, f => f.ListItemAllFields);
                clientContext.ExecuteQuery();
                //Print List Item Id
                Console.WriteLine("List Item Id: {0}", uploadFile.ListItemAllFields.Id);



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

        public class vwDepartmentFolders
        {
        }
    }  //end of class HomeController


   
}