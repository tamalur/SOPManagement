using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SOPManagement.Models
{
    public class Employee
    {

        public int userid { get; set; }

        public string userfullname { get; set; }

        public string useremailaddress { get; set; }

        public string userjobtitle { get; set; }

        public short departmentcode { get; set; }

        public string departmentname { get; set; }

        public short userstatuscode { get; set; }

         public Employee[] employees { get; set; }
        public bool HasSignedSOP { get; set; }

        public DateTime signaturedate { get; set; }

        public string signstatus { get; set; }

        public short signstatuscode { get; set; }

        public bool AuthenticateUserWithSOP(string pAuthtype, int pFileid)
        {
            bool bAuthenticate = false;

            GetUserByEmail();

       
            //check activities database to find him/her as this file approver 
            if (pAuthtype.ToLower().Trim() == "approver")
            {
                using (var ctx = new RadiantSOPEntities())
                {
                    var aprvr = ctx.vwApprvrsSignatures.Where(u => u.fileid == pFileid 
                    && u.approverid == userid).Select(u => u.approverid).FirstOrDefault();

                    if (aprvr > 0)
                    {
                        bAuthenticate = true;

                    }

                }
            }


            //check activities database to find him/her as reviewer 
            if (pAuthtype.ToLower().Trim() == "reviewer")
            {
                using (var ctx = new RadiantSOPEntities())
                {
                    var rvwr = ctx.vwRvwrsSignatures.Where(u => u.fileid == pFileid && 
                    u.reviewerid == userid).Select(u => u.reviewerid).FirstOrDefault();

                    if (rvwr > 0)
                    {
                        bAuthenticate = true;

                    }

                }
            }


            //check activities database to find him/her as owner 
            if (pAuthtype.ToLower().Trim() == "owner")
            {
                using (var ctx = new RadiantSOPEntities())
                {
                    var ownr = ctx.vwOwnerSignatures.Where(u => u.fileid == pFileid &&
                    u.ownerid == userid).Select(u => u.ownerid).FirstOrDefault();

                    if (ownr > 0)
                    {
                        bAuthenticate = true;

                    }

                }
            }


            return bAuthenticate;
        }


        //public bool AuthenticateUser(string authtype, string useremail)
        //{
        //    bool bAuthenticate = false;

        //    //check user database sopadminuser if authtype=admin
        //    if (authtype.ToLower().Trim() == "admin")
        //    {
        //        using (var ctx = new RadiantSOPEntities())
        //        {
        //            var admnu = ctx.users.Where(u => u.sopadminuser == true && u.useremailaddress == useremail).Select(u => u.userid).FirstOrDefault();

        //            if (admnu >0)
        //            {
        //                bAuthenticate = true;

        //            }

        //        }
        //    }


        //    return bAuthenticate;
        //}
        //public void GetUserInfoByEmail()
        //{

        //    RadiantSOPEntities ctx = new RadiantSOPEntities();

        //    //lsopno = foldername + "-001";

        //    userfullname = ctx.getUserFullNameByEmailUserID(useremailaddress, 0).FirstOrDefault().ToString();

        //    userjobtitle = ctx.GetUserJobTitleByEmailUserID(useremailaddress, 0).FirstOrDefault().ToString();


        //}

        public void GetEmployeesByDeptCode()
        {
            //ctx.vwUsers.Where(i => i.departmentcode == departmentcode).FirstOrDefault();

            //Query Entity Framework by using type of query- LINQ - Entities 


            Employee[] empllist;


            using (var ctx = new RadiantSOPEntities())
            {

                var employees = ctx.vwUsers.Select(x => new Employee()
                {

                    userid = x.userid1,
                    useremailaddress = x.useremailaddress,
                    userfullname = x.userfullname,
                    userjobtitle = x.jobtitle,
                    departmentcode = (short)x.departmentcode,
                    departmentname = x.departmentname

                }).Where(q => q.departmentcode == departmentcode || q.departmentcode==12);

                //empllist = employees.ToList();

                empllist = new Employee[employees.Count()];
                int i = 0;

                foreach (Employee emp in employees)
                {
                    empllist[i] = emp;

                    i++;

                }

            }

            employees= empllist;


        }

        public void GetUserByEmail()
        {
            //ctx.vwUsers.Where(i => i.departmentcode == departmentcode).FirstOrDefault();

            //Query Entity Framework by using type of query- LINQ - Entities 


            using (var ctx = new RadiantSOPEntities())
            {

                //first LINQ query through class
                var employee = ctx.vwUsers.Select(x => new Employee()
                {

                    userid = x.userid1,
                    useremailaddress = x.useremailaddress,
                    userfullname = x.userfullname,
                    userjobtitle = x.jobtitle,
                    departmentcode = (short)x.departmentcode,
                    departmentname = x.departmentname

                }).Where(q => q.useremailaddress == useremailaddress);


                //assign property from returned object from linq above
                foreach (Employee emp in employee)
                {
                    userid = emp.userid;
                    userfullname = emp.userfullname;
                    userjobtitle = emp.userjobtitle;
                    departmentcode = emp.departmentcode;
                    departmentname = emp.departmentname;

                }


            }


        }



    }
}