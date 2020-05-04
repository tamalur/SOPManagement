using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;
namespace SOPManagement.Models
{
    public class SOPSignatureModel
    {
        public string LoggedInUserEmail { get; set;}
        public string LoggedInUserFullName { get; set;}
        public string LoggedInUserJobTitle { get; set; }

        public string LoggedInUserIsOwner { get; set; }

        public string LoggedInUserIsApprover { get; set; }

        public string LoggedInUserIsReviewer { get; set; }

        [Display(Name = "Your Signature")]
        public bool LoggedInSigned { get; set; }

        [Display(Name = "Your Signature Date")]
        public DateTime LoggedInSignDate { get; set; }
        
        [Display(Name="SOP NO")]
        public string SOPNo { get; set; }

        [Display(Name = "SOP Name")]
        public string SOPName { get; set; }   //with sopno in front SOPNO + " "+ FileTitle

        [Display(Name = "SOP Link")]
        public string SOPUrl { get; set; }   //with sopno in front SOPNO + " "+ FileTitle

        [Display(Name = "SOP Latest Version")]
        public string SOPLastVersion { get; set; }   //with sopno in front SOPNO + " "+ FileTitle

        [Display(Name = "SOP Department")]
        public string SOPFilePath { get; set; }   //with sopno in front SOPNO + " "+ FileTitle

        [Display(Name = "Department Folder")]
        public string SOPDeptName { get; set; }   //with sopno in front SOPNO + " "+ FileTitle
        public string SOPSubDeptName { get; set; }   //with sopno in front SOPNO + " "+ FileTitle
        public Employee[] SOPRvwerSignatures { get; set; }

        public Employee SOPOwnerSignature { get; set; }

        public Employee SOPApprvrSignature { get; set; }

    }
}