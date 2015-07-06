using PropertyChanged;

namespace AttendanceSystem.Models
{
    [ImplementPropertyChanged]
    public class PersonnnelModel
    {
        public int PersonnelIndex { get; set; }
        public string PersonnelName { get; set; }
        public string DepartmentName { get; set; }
        public double? FreeAnnualLeave { get; set; }
        public double? UsedAnnualLeave { get; set; }
        public double? RemainingAnnualLeave { get; set; }
        public double? CumulativeSickLeave { get; set; }
        public double? CumulativePrivateLeave { get; set; }
    }
}