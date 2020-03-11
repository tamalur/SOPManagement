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


        public void GetUserInfoByEmail()
        {

            RadiantSOPEntities ctx = new RadiantSOPEntities();

            //lsopno = foldername + "-001";

            userfullname = ctx.getUserFullNameByEmailUserID(useremailaddress, 0).FirstOrDefault().ToString();

            userjobtitle = ctx.GetUserJobTitleByEmailUserID(useremailaddress, 0).FirstOrDefault().ToString();


        }

        public void GetUserByEmail()
        {
            //ctx.vwUsers.Where(i => i.departmentcode == departmentcode).FirstOrDefault();

            //Query Entity Framework by using type of query- LINQ - Entities 


            using (var ctx = new RadiantSOPEntities())
            {

                var employee = ctx.vwUsers.Select(x => new Employee()
                {

                    userid = x.userid,
                    useremailaddress = x.useremailaddress,
                    userfullname = x.userfullname,
                    userjobtitle = x.jobtitle,
                    departmentcode = (short)x.departmentcode,
                    departmentname = x.departmentname

                }).Where(q => q.useremailaddress == useremailaddress);


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