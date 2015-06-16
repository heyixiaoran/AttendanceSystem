using PropertyChanged;

namespace AttendanceSystem.Models
{
    [ImplementPropertyChanged]
    public class PersonnnelModel
    {
        public int PersonnelId { get; set; }
        public string PersonnelName { get; set; }
        public int DepartmentId { get; set; }
        public float FreeAnnualLeave { get; set; }
        public float UsedAnnualLeave { get; set; }
        public float RemainingAnnualLeave { get; set; }
    }
}