//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace SOPManagement.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class vwDepartmentSubFolder
    {
        public int FileID { get; set; }
        public string DeptFileName { get; set; }
        public string SOPNo { get; set; }
        public string SPFileLink { get; set; }
        public Nullable<System.DateTime> CreateDate { get; set; }
        public string CreatedBy { get; set; }
        public Nullable<System.DateTime> LastModifiedDate { get; set; }
        public string ModifiedBy { get; set; }
        public string VersionNo { get; set; }
        public string ApprovalStatus { get; set; }
        public string Approvedby { get; set; }
        public string DeptFileOwner { get; set; }
        public string SPFilePath { get; set; }
        public Nullable<short> prioritycode { get; set; }
        public Nullable<short> filestatuscode { get; set; }
    }
}