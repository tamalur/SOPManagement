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
    
    public partial class vwRvwrsSignature
    {
        public int fileid { get; set; }
        public int changerequestid { get; set; }
        public Nullable<int> requesterid { get; set; }
        public string reviewername { get; set; }
        public string reviewertitle { get; set; }
        public string SignedStatus { get; set; }
        public short SignStatusCode { get; set; }
        public Nullable<System.DateTime> signeddate { get; set; }
        public int reviewerid { get; set; }
        public string revieweremail { get; set; }
    }
}
