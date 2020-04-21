using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Diagnostics;
using System.IO;

namespace SOPManagement.Models
{

         public class Logger {

         private TextWriter _twLogger;

         private string sEventLogName;

         private EventLog EventLog1;

         private string sLogFileName;
         public string LogFileName
         {
             get
             {
                 return sLogFileName;
             }
             set
             {
                 sLogFileName = value;
             }
         }
         
         public void WriteToEventLog(string sMessageToWrite, string sEventLogname) 
         
         {
             sEventLogName = sEventLogname;
             if (!EventLog.SourceExists("SOP File Process"))
             {
                 EventLog.CreateEventSource("SOP File Process", sEventLogName);
             }
             EventLog1 = new EventLog();
             EventLog1.Source = "SOP File Process";
             EventLog1.Log = sEventLogName;

             EventLog1.WriteEntry(sMessageToWrite);
         }
         
         public void LogEvent(string description)
  

         {
 
             if ((_twLogger == null)) {
                 
             }
             else {
                 
             }
         }

         public void UpdateLogFile(string sMessageToWrite)
         {
             TextWriter sw;
             try
             {
                 sw = File.AppendText(LogFileName);
                 if ((sMessageToWrite.Trim() != ""))
                 {
                     sw.WriteLine(sMessageToWrite);
                     sw.Flush();
                 }
                 sw.Close();
                 sw.Dispose();

             }
             catch (Exception ex)
             {
                 throw new Exception("Error source-UpdateLogFile in Logger class with error: " + ex.Message);
             }
             finally
             {


                 sw = null;


             }
         }


     }

 

}
