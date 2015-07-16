using System;

using PropertyChanged;

namespace AttendanceSystem.Models
{
    [ImplementPropertyChanged]
    public class AddRecordModel
    {
        public string PersonnelName { get; set; }
        public string LeaveName { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public double? LeaveDays { get; set; }
        public double? TransformLeave { get; set; }
    }
}
