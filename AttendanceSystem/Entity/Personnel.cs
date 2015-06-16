//------------------------------------------------------------------------------
// <auto-generated>
//     此代码已从模板生成。
//
//     手动更改此文件可能导致应用程序出现意外的行为。
//     如果重新生成代码，将覆盖对此文件的手动更改。
// </auto-generated>
//------------------------------------------------------------------------------

namespace AttendanceSystem.Entity
{
    using System;
    using System.Collections.Generic;
    
    public partial class Personnel
    {
        public Personnel()
        {
            this.Attendances = new HashSet<Attendance>();
        }
    
        public int PersonnelId { get; set; }
        public string PersonnelName { get; set; }
        public Nullable<int> DepartmentId { get; set; }
        public Nullable<double> FreeAnnualLeave { get; set; }
        public Nullable<double> UsedAnnualLeave { get; set; }
        public Nullable<double> RemainingAnnualLeave { get; set; }
    
        public virtual ICollection<Attendance> Attendances { get; set; }
        public virtual Department Department { get; set; }
    }
}