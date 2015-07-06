using PropertyChanged;

namespace AttendanceSystem.Models
{
    [ImplementPropertyChanged]
    public class AttendanceModel
    {
        public string Name { get; set; }
        public double AttendanceHour { get; set; }
        public int LateTime { get; set; }
        public int Absenteeism { get; set; }
        public double OvertimeHours { get; set; }
    }
}