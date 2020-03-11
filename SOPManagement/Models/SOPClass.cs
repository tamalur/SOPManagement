using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SOPManagement.Models
{
    public class SOPClass
    {

        public int FileID { get; set; }

        public int FileOwnerID { get; set; }

        public int FileApproverID { get; set; }

        public string FileApproverEmail { get; set; }

        public string FileOwnerEmail { get; set; }

        public Employee[] FileReviewers { get; set; }

        public string DeptFileName { get; set;}

        public string SOPNo { get; set; }

        public string SPFileLink { get; set; }

        public string SPFilePath { get; set; }

        public bool OperationSuccess { get; set;}


        public void AddFileReviewers()
        {

            Employee emp = new Employee();

            int rvwrid;

            OperationSuccess = false;

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
    }
}