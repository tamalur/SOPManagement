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
    
    public partial class fileapprover
    {
        public int approveid { get; set; }
        public int approverid { get; set; }
        public int fileid { get; set; }
        public Nullable<short> approverstatuscode { get; set; }
    }
}