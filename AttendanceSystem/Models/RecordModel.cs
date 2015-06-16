using System;

namespace AttendanceSystem.Models
{
    public class RecordModel
    {
        public string Name { get; set; }
        public DateTime Date { get; set; }
        public DateTime ArriveTime { get; set; }
        public DateTime LeaveTime { get; set; }
    }
}