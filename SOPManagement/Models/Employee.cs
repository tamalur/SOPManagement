using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SOPManagement.Models
{
    public class Employee
    {

        public int userid { get; set; }

        public string userfullname { get; set;}

        public string useremailaddress { get; set; }

        public string userjobtitle { get; set; }

        public short departmentcode { get; set; }

        public short userstatuscode { get; set; }

        
        public void GetUserInfoByEmail()
        {

            RadiantSOPEntities ctx = new RadiantSOPEntities();

            //lsopno = foldername + "-001";

            userfullname = ctx.getUserFullNameByEmailUserID(useremailaddress,0).FirstOrDefault().ToString();

            userjobtitle = ctx.GetUserJobTitleByEmailUserID(useremailaddress, 0).FirstOrDefault().ToString();

        }

    }

}