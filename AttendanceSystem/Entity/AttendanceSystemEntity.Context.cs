﻿//------------------------------------------------------------------------------
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
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    
    public partial class AttendanceSystemEntity : DbContext
    {
        public AttendanceSystemEntity()
            : base("AttendanceSystemEntity")
        {
        }
    
        public virtual DbSet<Attendance> Attendances { get; set; }
        public virtual DbSet<Department> Departments { get; set; }
        public virtual DbSet<LeaveType> LeaveTypes { get; set; }
        public virtual DbSet<Personnel> Personnels { get; set; }
    }
}
