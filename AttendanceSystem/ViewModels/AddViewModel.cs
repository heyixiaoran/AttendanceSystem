using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

using AttendanceSystem.Models;

using Caliburn.Micro;

using MahApps.Metro.Controls;

namespace AttendanceSystem.ViewModels
{
    public class AddViewModel : Screen
    {
        public ObservableCollection<AttendanceRecordModel> AttendanceCollection { get; set; }

        public ObservableCollection<AddRecordModel> LeaveRecordCollection { get; set; }

        public ObservableCollection<LeaveTypeModel> LeaveCollection { get; set; }

        public AddViewModel()
        {
            LeaveRecordCollection = new ObservableCollection<AddRecordModel>();

            InitLeaveCollection();
        }

        public void OnDataGridCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            try
            {
                if(e.Column.DisplayIndex == 0)
                {
                    var dataGrid = sender as DataGrid;
                    if(dataGrid != null)
                    {
                        var aa=  dataGrid.SelectedItem as AddRecordModel;
                    }
                }
                //var dataGrid = sender as DataGrid;
                //if(dataGrid != null)
                //{
                //    dataGrid.Focus();
                //    dataGrid.BeginEdit();
                //}

                //var cell = dataGrid.ItemContainerGenerator.ContainerFromIndex(e.Column.DisplayIndex);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void OnContentChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        public void OnStartDateChanged(object sender, SelectionChangedEventArgs e)
        {
            var addRecordModel = sender as AddRecordModel;
            if(addRecordModel != null)
            {
                addRecordModel.StartDate = (e.Source as DatePicker).SelectedDate;
                if(addRecordModel.EndDate != null)
                {
                    addRecordModel.LeaveDays = ((DateTime)addRecordModel.EndDate).Subtract((DateTime)addRecordModel.StartDate).Duration().TotalDays + 1;
                }
                else
                {
                    addRecordModel.LeaveDays = 1;
                }
            }
        }

        public void OnEndDateChanged(object sender, SelectionChangedEventArgs e)
        {
            var addRecordModel = sender as AddRecordModel;
            if(addRecordModel != null)
            {
                addRecordModel.EndDate = (e.Source as DatePicker).SelectedDate;
                if(addRecordModel.StartDate != null)
                {
                    addRecordModel.LeaveDays = ((DateTime)addRecordModel.EndDate).Subtract((DateTime)addRecordModel.StartDate).Duration().TotalDays + 1;
                }
                else
                {
                    addRecordModel.LeaveDays = 1;
                }
            }
        }

        private void InitLeaveCollection()
        {
            LeaveCollection = new ObservableCollection<LeaveTypeModel>
            {
                new LeaveTypeModel {LeaveId = "病假", LeaveName = "病假"},
                new LeaveTypeModel {LeaveId = "事假", LeaveName = "事假"},
                new LeaveTypeModel {LeaveId = "年假", LeaveName = "年假"},
                new LeaveTypeModel {LeaveId = "婚假", LeaveName = "婚假"},
                new LeaveTypeModel {LeaveId = "产假", LeaveName = "产假"},
                new LeaveTypeModel {LeaveId = "丧假", LeaveName = "丧假"}
            };
        }

        public void CancleClick()
        {
            TryClose();
        }

        public void OkClick()
        {
            if(CalculateLeave())
            {
                var eventX = IoC.Get<IEventAggregator>();
                eventX.PublishOnUIThread(AttendanceCollection);

                TryClose();
            }
        }

        public void AddRecord()
        {
            LeaveRecordCollection.Add(new AddRecordModel());
        }

        public void DeleteRecord(Grid grid)
        {
            var dataGrid = grid.FindChild<DataGrid>("AddDataGrid");
            if(dataGrid != null)
            {
                var record = dataGrid.SelectedItem as AddRecordModel;
                if(record != null)
                {
                    LeaveRecordCollection.Remove(record);
                }
            }
        }

        private bool CalculateLeave()
        {
            try
            {
                foreach(var leaveRecords in LeaveRecordCollection.GroupBy(l => l.PersonnelName))
                {
                    var attendance = AttendanceCollection.FirstOrDefault(a => a.PersonnelName == leaveRecords.FirstOrDefault().PersonnelName);
                    if(attendance != null)
                    {
                        foreach(var leaveRecord in leaveRecords)
                        {
                            if(attendance != null && !string.IsNullOrEmpty(leaveRecord.PersonnelName))
                            {
                                switch(leaveRecord.LeaveName)
                                {
                                    case "病假":
                                        attendance.SickLeave = leaveRecord.LeaveDays;
                                        attendance.CumulativeSickLeave += leaveRecord.LeaveDays;
                                        break;
                                    case "事假":
                                        attendance.PrivateLeave = leaveRecord.LeaveDays;
                                        attendance.CumulativePrivateLeave += leaveRecord.LeaveDays;
                                        break;
                                    case "年假":
                                        attendance.LeaveName = leaveRecord.LeaveName;
                                        attendance.FreeAnnualLeave -= leaveRecord.LeaveDays;
                                        attendance.UsedAnnualLeave += leaveRecord.LeaveDays;
                                        if(leaveRecord.LeaveDays > attendance.RemainingAnnualLeave)
                                        {
                                            attendance.TransformLeave = leaveRecord.LeaveDays - attendance.RemainingAnnualLeave;
                                            attendance.RemainingAnnualLeave = 0;
                                        }
                                        else
                                        {
                                            attendance.RemainingAnnualLeave -= leaveRecord.LeaveDays;
                                        }
                                        break;
                                    case "婚假":
                                        attendance.LeaveName = leaveRecord.LeaveName;
                                        break;
                                    case "产假":
                                        attendance.LeaveName = leaveRecord.LeaveName;
                                        break;
                                    case "丧假":
                                        attendance.LeaveName = leaveRecord.LeaveName;
                                        break;
                                }

                                if(leaveRecord.EndDate == null)
                                {
                                    attendance.Note += ((DateTime)leaveRecord.StartDate).ToString("MM-dd") + leaveRecord.LeaveName + leaveRecord.LeaveDays + "天；";
                                }
                                else
                                {
                                    attendance.Note += ((DateTime)leaveRecord.StartDate).ToString("MM-dd") + "—" + ((DateTime)leaveRecord.EndDate).ToString("MM-dd") + leaveRecord.LeaveName + leaveRecord.LeaveDays + "天；";
                                }
                            }
                        }

                        attendance.Note = attendance.Note.TrimEnd('；');
                    }
                }

                return true;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }
    }
}
