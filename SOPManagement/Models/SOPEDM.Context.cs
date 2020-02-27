﻿//------------------------------------------------------------------------------
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
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    using System.Data.Entity.Core.Objects;
    using System.Linq;
    
    public partial class RadiantSOPEntities : DbContext
    {
        public RadiantSOPEntities()
            : base("name=RadiantSOPEntities")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public virtual DbSet<codesdepartment> codesdepartments { get; set; }
        public virtual DbSet<deptsopfile> deptsopfiles { get; set; }
        public virtual DbSet<fileapprover> fileapprovers { get; set; }
        public virtual DbSet<fileapproversactivity> fileapproversactivities { get; set; }
        public virtual DbSet<filechangerequestactivity> filechangerequestactivities { get; set; }
        public virtual DbSet<fileowner> fileowners { get; set; }
        public virtual DbSet<fileownersactivity> fileownersactivities { get; set; }
        public virtual DbSet<filepublisher> filepublishers { get; set; }
        public virtual DbSet<filepublishersactivity> filepublishersactivities { get; set; }
        public virtual DbSet<filereviewer> filereviewers { get; set; }
        public virtual DbSet<filereviewersactivity> filereviewersactivities { get; set; }
        public virtual DbSet<fileupdateschedule> fileupdateschedules { get; set; }
        public virtual DbSet<user> users { get; set; }
        public virtual DbSet<codesapprovalstatu> codesapprovalstatus { get; set; }
        public virtual DbSet<codesfilestatu> codesfilestatus { get; set; }
        public virtual DbSet<codesuserstatu> codesuserstatus { get; set; }
        public virtual DbSet<v_sopreport> v_sopreport { get; set; }
        public virtual DbSet<vwDepartmentFolder> vwDepartmentFolders { get; set; }
        public virtual DbSet<vwDepartmentSubFolder> vwDepartmentSubFolders { get; set; }
        public virtual DbSet<vwSOPReviewer> vwSOPReviewers { get; set; }
    
        public virtual ObjectResult<Nullable<int>> sp_getSOPNo(string deptfolder, string deptsubfolder)
        {
            var deptfolderParameter = deptfolder != null ?
                new ObjectParameter("deptfolder", deptfolder) :
                new ObjectParameter("deptfolder", typeof(string));
    
            var deptsubfolderParameter = deptsubfolder != null ?
                new ObjectParameter("deptsubfolder", deptsubfolder) :
                new ObjectParameter("deptsubfolder", typeof(string));
    
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction<Nullable<int>>("sp_getSOPNo", deptfolderParameter, deptsubfolderParameter);
        }
    
        public virtual ObjectResult<Nullable<int>> GetLastSOPNO(string deptfolder, string deptsubfolder)
        {
            var deptfolderParameter = deptfolder != null ?
                new ObjectParameter("deptfolder", deptfolder) :
                new ObjectParameter("deptfolder", typeof(string));
    
            var deptsubfolderParameter = deptsubfolder != null ?
                new ObjectParameter("deptsubfolder", deptsubfolder) :
                new ObjectParameter("deptsubfolder", typeof(string));
    
            return ((IObjectContextAdapter)this).ObjectContext.ExecuteFunction<Nullable<int>>("GetLastSOPNO", deptfolderParameter, deptsubfolderParameter);
        }
    }
}
