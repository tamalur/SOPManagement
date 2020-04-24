using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SOPManagement.Models
{
    public class FileRevision
    {

        public int FileID { get; set; }

        public int RevisionID { get; set; }

        public string RevisionNo { get; set; }

        public string Description { get; set; }

        public DateTime RevisionDate { get; set; }

        public string VersionUrl { get; set; }
    }
}