using System;
using System.Collections.ObjectModel;
using PropertyChanged;

namespace AttendanceSystem.Models
{
    [ImplementPropertyChanged]
    public class AttendanceRecordModel
    {
        public int Index { get; set; }
        //public int? AttendanceId { get; set; }//考勤Id
        public int? DepartmentId { get; set; }//部门Id
        public ObservableCollection<DepartmentModel> DepartmentCollection { get; set; }
        public int? PersonnelId { get; set; }//人员Id
        public string PersonnelName { get; set; }//人员姓名
        public double? SickLeave { get; set; }//病假（天）
        public double? CumulativeSickLeave { get; set; }//累计病假
        public double? PrivateLeave { get; set; }//事假（天）
        public double? CumulativePrivateLeave { get; set; }//累计事假
        public double? TransformLeave { get; set; }//病事假转换
        public double? FreeAnnualLeave { get; set; }//可休年假
        public double? UsedAnnualLeave { get; set; }//已休年假
        public double? RemainingAnnualLeave { get; set; }//剩余年假
        public int? OtherLeaveTypeId { get; set; }//其他假别
        public ObservableCollection<LeaveTypeModel> LeaveCollection { get; set; }
        public int? LateTime { get; set; }//迟到（次）
        public int? Absenteeism { get; set; }//旷工（天）
        public double? AttendanceHour { get; set; }//本月出勤工时（小时）
        public double? OvertimeHour { get; set; }//本月加班（小时）
        public string Note { get; set; }//备注
        public DateTime Date { get; set; }//日期
    }
}
