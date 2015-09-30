using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Xml;
using AttendanceSystem.Commons;
using AttendanceSystem.Models;
using AttendanceSystem.Views;

using Caliburn.Micro;

using MahApps.Metro;
using MahApps.Metro.Controls;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Application = System.Windows.Application;
using Brush = System.Windows.Media.Brush;
using MenuItem = System.Windows.Controls.MenuItem;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Screen = Caliburn.Micro.Screen;

namespace AttendanceSystem.ViewModels
{
    public class MainViewModel : Screen, IShell, IHandle<ObservableCollection<AttendanceRecordModel>>
    {
        #region

        private const string _configFileSourceUri = "Files/ConfigInfo.xml";

        private const string _personnelRecordFileSourceUri = "Files/Personnel.xlsx";

        private readonly string _attendanceDataUri = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "/AttendanceData";

        private readonly string _configFileUri = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "/AttendanceData/ConfigInfo.xml";

        private readonly string _personnelRecordFileUri = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "/AttendanceData/Personnel.xlsx";

        private readonly string _attendanceRecordFileUri = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "/AttendanceData/{0}";

        #endregion

        #region 属性

        public bool IsWriteRecords { get; set; }

        public int DefaultWorkHours { get; set; }

        public string StatusString { get; set; }

        public string StartWorkTime { get; set; }

        public string DefaultStartWorkTime { get; set; }

        public string DefaultEndWorkTime { get; set; }

        public ObservableCollection<AttendanceRecordModel> AttendanceCollection { get; set; }

        public ObservableCollection<DepartmentModel> DepartmentCollection { get; set; }

        public ObservableCollection<LeaveTypeModel> LeaveCollection { get; set; }

        public List<AccentColorMenuData> AccentColors { get; set; }

        public List<AppThemeMenuData> AppThemes { get; set; }

        #endregion

        public MainViewModel()
        {
            InitAccentColors();
            InitAppThemes();
            CheckEnvironment();
            ReadConfigInfo();
            InitDepartmentCollection();
            InitLeaveCollection();
            //InitialTray();

            AttendanceCollection = new ObservableCollection<AttendanceRecordModel>();
        }

        private NotifyIcon notifyIcon;

        private void InitialTray()
        {
            this.notifyIcon = new NotifyIcon();
            this.notifyIcon.BalloonTipText = "AttendanceSystem";
            this.notifyIcon.ShowBalloonTip(2000);
            this.notifyIcon.Text = "AttendanceSystem";
            this.notifyIcon.Icon = Icon.ExtractAssociatedIcon(System.Windows.Forms.Application.ExecutablePath);
            this.notifyIcon.Visible = true;

        }

        private void InitAppThemes()
        {
            AppThemes = ThemeManager.AppThemes
                                    .Select(a => new AppThemeMenuData() { Name = a.Name, BorderColorBrush = a.Resources["BlackColorBrush"] as Brush, ColorBrush = a.Resources["WhiteColorBrush"] as Brush })
                                    .ToList();
        }

        private void InitAccentColors()
        {
            AccentColors = ThemeManager.Accents
                                       .Select(a => new AccentColorMenuData() { Name = a.Name, ColorBrush = a.Resources["AccentColorBrush"] as Brush })
                                       .ToList();
        }

        public void ChangeTheme(object sender)
        {
            var menuItem = sender as MenuItem;

        }

        public void ImportExcel()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "选择文件",
                Filter = "Excel files (*.xlsx)|*.xlsx",
                FileName = string.Empty,
                FilterIndex = 1,
                RestoreDirectory = true
            };

            bool? showDialog = openFileDialog.ShowDialog();

            if(showDialog != null && (bool)showDialog)
            {
                try
                {
                    AttendanceCollection.Clear();
                    StatusString = "Impoarting";

                    IWorkbook workbook = null;
                    using(var fs = new FileStream(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        if(openFileDialog.FileName.IndexOf(".xlsx", StringComparison.Ordinal) > 0) // 2007版本
                        {
                            workbook = new XSSFWorkbook(fs);
                        }
                        else if(openFileDialog.FileName.IndexOf(".xls", StringComparison.Ordinal) > 0) // 2003版本
                        {
                            workbook = new HSSFWorkbook(fs);
                        }

                        if(workbook != null)
                        {
                            var sheet = workbook.GetSheetAt(0);

                            //获取sheet的首行
                            XSSFRow headerRow = (XSSFRow)sheet.GetRow(0);

                            //总的列数
                            int cellCount = headerRow.LastCellNum;
                            var records = new ObservableCollection<RecordModel>();
                            for(int i = (sheet.FirstRowNum + 1); i < sheet.LastRowNum; i++)
                            {
                                XSSFRow row = (XSSFRow)sheet.GetRow(i);
                                var record = new RecordModel();

                                for(int j = row.FirstCellNum; j < cellCount; j++)
                                {
                                    var cell = row.GetCell(j);

                                    if(cell != null)
                                    {
                                        switch(j)
                                        {
                                            case 0:
                                                record.Name = cell.ToString();
                                                break;
                                            case 1:
                                                if(cell.CellType == CellType.String)
                                                {
                                                    DateTime dateTimeX;
                                                    if(DateTime.TryParse(cell.ToString(), out dateTimeX))
                                                    {
                                                        record.Date = Convert.ToDateTime(cell.ToString());
                                                    }
                                                }
                                                else if(cell.CellType == CellType.Numeric)
                                                {
                                                    record.Date = Convert.ToDateTime(cell.DateCellValue.ToShortDateString());
                                                }
                                                break;
                                            case 2:
                                                if(cell.CellType == CellType.String && cell.ToString() != "")
                                                {
                                                    DateTime dateTimeX;
                                                    if(DateTime.TryParse(cell.ToString(), out dateTimeX))
                                                    {
                                                        record.ArriveTime = Convert.ToDateTime(Convert.ToDateTime(cell.ToString()).ToShortTimeString());
                                                    }
                                                    else
                                                    {
                                                        record.ArriveTime = Convert.ToDateTime(DefaultStartWorkTime);
                                                    }
                                                }
                                                else if(cell.CellType == CellType.Numeric)
                                                {
                                                    record.ArriveTime = Convert.ToDateTime(cell.DateCellValue.ToShortTimeString());
                                                }
                                                break;
                                            case 3:
                                                if(cell.CellType == CellType.String && cell.ToString() != "")
                                                {
                                                    DateTime dateTimeX;
                                                    if(DateTime.TryParse(cell.ToString(), out dateTimeX))
                                                    {
                                                        record.LeaveTime = Convert.ToDateTime(Convert.ToDateTime(cell.ToString()).ToShortTimeString());
                                                    }
                                                    else
                                                    {
                                                        record.ArriveTime = Convert.ToDateTime(DefaultEndWorkTime);
                                                    }
                                                }
                                                else if(cell.CellType == CellType.Numeric)
                                                {
                                                    record.LeaveTime = Convert.ToDateTime(cell.DateCellValue.ToShortTimeString());
                                                }
                                                break;
                                        }
                                    }
                                }
                                records.Add(record);
                            }

                            CalculateAttendance(records);
                        }
                    }
                }
                catch(Exception ex)
                {
                    StatusString = "";
                    MessageBox.Show(ex.Message);
                }
            }
        }

        public void OpenDataFolder()
        {
            try
            {
                Process.Start(_attendanceDataUri);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void ExportToExcel()
        {
            if(IsWriteRecords)
            {
                CreateNewPersonnelFile();
            }

            CreateNewAttendanceFile();
        }

        public void OpenSettingWindow()
        {
            ToggleFlyout(0);
        }

        public void WriteToConfigFile()
        {
            try
            {
                if(File.Exists(_configFileUri))
                {
                    var xmlDocument = new XmlDocument();
                    xmlDocument.Load(_configFileUri);
                    var selectSingleNode = xmlDocument.SelectSingleNode("ConfigInfo");
                    if(selectSingleNode != null)
                    {
                        var defaultWorkHours = selectSingleNode.SelectSingleNode("DefaultWorkHours");
                        if(defaultWorkHours != null)
                        {
                            defaultWorkHours.Attributes["Value"].Value = DefaultWorkHours.ToString();
                        }
                        var startWorkTime = selectSingleNode.SelectSingleNode("StartWorkTime");
                        if(startWorkTime != null)
                        {
                            startWorkTime.Attributes["Value"].Value = Convert.ToDateTime(StartWorkTime).TimeOfDay.ToString();
                        }
                        var defaultStartWorkTime = selectSingleNode.SelectSingleNode("DefaultStartWorkTime");
                        if(defaultStartWorkTime != null)
                        {
                            defaultStartWorkTime.Attributes["Value"].Value = Convert.ToDateTime(DefaultStartWorkTime).TimeOfDay.ToString();
                        }
                        var defaultEndWorkTime = selectSingleNode.SelectSingleNode("DefaultEndWorkTime");
                        if(defaultEndWorkTime != null)
                        {
                            defaultEndWorkTime.Attributes["Value"].Value = Convert.ToDateTime(DefaultEndWorkTime).TimeOfDay.ToString();
                        }

                        xmlDocument.Save(_configFileUri);
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void OnDataGridCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {

        }

        public void AddLeave()
        {
            var addViewModel = IoC.Get<AddViewModel>();
            addViewModel.AttendanceCollection = AttendanceCollection;
            IoC.Get<IWindowManager>().ShowDialog(addViewModel);
        }

        public void Handle(ObservableCollection<AttendanceRecordModel> attendanceCollection)
        {
            AttendanceCollection = attendanceCollection;
        }

        public void ChangeTheme()
        {
            var theme = ThemeManager.DetectAppStyle(Application.Current);

            // now set the Green accent and dark theme
            ThemeManager.ChangeAppStyle(Application.Current,
                                        ThemeManager.GetAccent("Green"),
                                        ThemeManager.GetAppTheme("BaseDark"));
        }

        private void CreateNewPersonnelFile()
        {
            try
            {
                using(var fs = new FileStream(_personnelRecordFileUri, FileMode.Create, FileAccess.Write))
                {
                    var workbook = new XSSFWorkbook();
                    var sheet = workbook.CreateSheet("sheet1");
                    var headerRow = sheet.CreateRow(0);
                    headerRow.CreateCell(0).SetCellValue("序号");
                    headerRow.CreateCell(1).SetCellValue("姓名");
                    headerRow.CreateCell(2).SetCellValue("部门");
                    headerRow.CreateCell(3).SetCellValue("可休年假");
                    headerRow.CreateCell(4).SetCellValue("已休年假");
                    headerRow.CreateCell(5).SetCellValue("剩余年假");
                    headerRow.CreateCell(6).SetCellValue("累计病假");
                    headerRow.CreateCell(7).SetCellValue("累计事假");

                    for(int i = 0; i < AttendanceCollection.Count; i++)
                    {
                        var row = sheet.CreateRow(i + 1);
                        row.CreateCell(0).SetCellValue(AttendanceCollection[i].PersonnelIndex);
                        row.CreateCell(1).SetCellValue(AttendanceCollection[i].PersonnelName);
                        row.CreateCell(2).SetCellValue(AttendanceCollection[i].DepartmentName);
                        row.CreateCell(3).SetCellValue(AttendanceCollection[i].FreeAnnualLeave.ToString());
                        row.CreateCell(4).SetCellValue(AttendanceCollection[i].UsedAnnualLeave.ToString());
                        row.CreateCell(5).SetCellValue(AttendanceCollection[i].RemainingAnnualLeave.ToString());
                        row.CreateCell(6).SetCellValue(AttendanceCollection[i].CumulativeSickLeave.ToString());
                        row.CreateCell(7).SetCellValue(AttendanceCollection[i].CumulativePrivateLeave.ToString());
                    }

                    workbook.Write(fs);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CreateNewAttendanceFile()
        {
            try
            {
                using(var fs = new FileStream(string.Format(_attendanceRecordFileUri, DateTime.Now.ToString("yyyy-MM") + "月考勤表.xlsx"), FileMode.Create, FileAccess.Write))
                {
                    var workbook = new XSSFWorkbook();
                    var sheet = workbook.CreateSheet("sheet1");
                    var headerRow = sheet.CreateRow(0);
                    headerRow.CreateCell(0).SetCellValue("序号");
                    headerRow.CreateCell(1).SetCellValue("部门");
                    headerRow.CreateCell(2).SetCellValue("姓名");
                    headerRow.CreateCell(3).SetCellValue("病假（天）");
                    headerRow.CreateCell(4).SetCellValue("累计病假");
                    headerRow.CreateCell(5).SetCellValue("事假（天）");
                    headerRow.CreateCell(6).SetCellValue("累计事假");
                    headerRow.CreateCell(7).SetCellValue("病事假转换");
                    headerRow.CreateCell(8).SetCellValue("可休年假");
                    headerRow.CreateCell(9).SetCellValue("已休年假");
                    headerRow.CreateCell(10).SetCellValue("剩余年假");
                    headerRow.CreateCell(11).SetCellValue("其他假别");
                    headerRow.CreateCell(12).SetCellValue("迟到（次）");
                    headerRow.CreateCell(13).SetCellValue("旷工（天）");
                    headerRow.CreateCell(14).SetCellValue("本月出勤工时（小时）");
                    headerRow.CreateCell(15).SetCellValue("本月加班(小时）");
                    headerRow.CreateCell(16).SetCellValue("备注");

                    for(int i = 0; i < AttendanceCollection.Count; i++)
                    {
                        var row = sheet.CreateRow(i + 1);
                        row.CreateCell(0).SetCellValue(AttendanceCollection[i].PersonnelIndex);
                        row.CreateCell(1).SetCellValue(AttendanceCollection[i].DepartmentName);
                        row.CreateCell(2).SetCellValue(AttendanceCollection[i].PersonnelName);
                        row.CreateCell(3).SetCellValue(AttendanceCollection[i].SickLeave.ToString());
                        row.CreateCell(4).SetCellValue(AttendanceCollection[i].CumulativeSickLeave.ToString());
                        row.CreateCell(5).SetCellValue(AttendanceCollection[i].PrivateLeave.ToString());
                        row.CreateCell(6).SetCellValue(AttendanceCollection[i].CumulativePrivateLeave.ToString());
                        row.CreateCell(7).SetCellValue(AttendanceCollection[i].TransformLeave.ToString());
                        row.CreateCell(8).SetCellValue(AttendanceCollection[i].FreeAnnualLeave.ToString());
                        row.CreateCell(9).SetCellValue(AttendanceCollection[i].UsedAnnualLeave.ToString());
                        row.CreateCell(10).SetCellValue(AttendanceCollection[i].RemainingAnnualLeave.ToString());
                        row.CreateCell(11).SetCellValue(AttendanceCollection[i].LeaveName);
                        row.CreateCell(12).SetCellValue(AttendanceCollection[i].LateTime.ToString());
                        row.CreateCell(13).SetCellValue(AttendanceCollection[i].Absenteeism.ToString());
                        row.CreateCell(14).SetCellValue(AttendanceCollection[i].AttendanceHour.ToString());
                        row.CreateCell(15).SetCellValue(AttendanceCollection[i].OvertimeHour.ToString());
                        row.CreateCell(16).SetCellValue(AttendanceCollection[i].Note);
                    }

                    workbook.Write(fs);
                }

                MessageBox.Show("Success");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CheckEnvironment()
        {
            try
            {
                StatusString = "Checking Environment";

                if(!Directory.Exists(_attendanceDataUri))
                {
                    Directory.CreateDirectory(_attendanceDataUri);
                }

                if(!File.Exists(_configFileUri))
                {
                    File.Copy(_configFileSourceUri, _configFileUri);
                }

                if(!File.Exists(_personnelRecordFileUri))
                {
                    File.Copy(_personnelRecordFileSourceUri, _personnelRecordFileUri);
                }

                StatusString = "";
            }
            catch(Exception ex)
            {
                StatusString = "";
                MessageBox.Show(ex.Message);
            }
        }

        private void InitLeaveCollection()
        {
            LeaveCollection = new ObservableCollection<LeaveTypeModel>
            {
                new LeaveTypeModel {LeaveId = "", LeaveName = ""},
                new LeaveTypeModel {LeaveId = "年假", LeaveName = "年假"},
                new LeaveTypeModel {LeaveId = "婚假", LeaveName = "婚假"},
                new LeaveTypeModel {LeaveId = "产假", LeaveName = "产假"},
                new LeaveTypeModel {LeaveId = "丧假", LeaveName = "丧假"}
            };
        }

        private void InitDepartmentCollection()
        {
            DepartmentCollection = new ObservableCollection<DepartmentModel>
            {
                new DepartmentModel {DepartmentId = "", DepartmentName = ""},
                new DepartmentModel {DepartmentId = "技术", DepartmentName = "技术"},
                new DepartmentModel {DepartmentId = "人力资源部", DepartmentName = "人力资源部"},
                new DepartmentModel {DepartmentId = "朝阳门", DepartmentName = "朝阳门"},
                new DepartmentModel {DepartmentId = "新员工", DepartmentName = "新员工"}
            };
        }

        private void ReadConfigInfo()
        {
            try
            {
                var xmlDocument = new XmlDocument();

                if(File.Exists(_configFileUri))
                {
                    xmlDocument.Load(_configFileUri);
                    var selectSingleNode = xmlDocument.SelectSingleNode("ConfigInfo");
                    if(selectSingleNode != null)
                    {
                        var defaultWorkHours = selectSingleNode.SelectSingleNode("DefaultWorkHours");
                        if(defaultWorkHours != null)
                        {
                            DefaultWorkHours = Convert.ToInt32(defaultWorkHours.Attributes["Value"].Value);
                        }
                        var startWorkTimeNode = selectSingleNode.SelectSingleNode("StartWorkTime");
                        if(startWorkTimeNode != null)
                        {
                            StartWorkTime = startWorkTimeNode.Attributes["Value"].Value;
                        }
                        var defaultStartWorkTimeNode = selectSingleNode.SelectSingleNode("DefaultStartWorkTime");
                        if(defaultStartWorkTimeNode != null)
                        {
                            DefaultStartWorkTime = defaultStartWorkTimeNode.Attributes["Value"].Value;
                        }
                        var defaultEndWorkTimeNode = selectSingleNode.SelectSingleNode("DefaultEndWorkTime");
                        if(defaultEndWorkTimeNode != null)
                        {
                            DefaultEndWorkTime = defaultEndWorkTimeNode.Attributes["Value"].Value;
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CalculateAttendance(ObservableCollection<RecordModel> records)
        {
            try
            {
                StatusString = "Calculating attendance";

                var personnels = ReadPersonnelRecord();

                var groupRecords = records.GroupBy(r => r.Name);

                int index = 1;
                var attendanceCollection = new ObservableCollection<AttendanceRecordModel>();
                foreach(var group in groupRecords)
                {
                    int absenteeism = 0;
                    int lateCount = 0;
                    double attendanceMinutes = 0;

                    foreach(var record in group)
                    {
                        if((record.ArriveTime.ToString() == "0001/1/1 0:00:00" && record.LeaveTime.ToString() != "0001/1/1 0:00:00") || (record.ArriveTime.ToString() == "0001-01-01 0:00:00" && record.LeaveTime.ToString() != "0001-01-01 0:00:00"))
                        {
                            record.ArriveTime = Convert.ToDateTime(Convert.ToDateTime(DefaultStartWorkTime).ToShortTimeString());
                        }
                        else if((record.ArriveTime.ToString() != "0001/1/1 0:00:00" && record.LeaveTime.ToString() == "0001/1/1 0:00:00") || (record.ArriveTime.ToString() != "0001-01-01 0:00:00" && record.LeaveTime.ToString() == "0001-01-01 0:00:00"))
                        {
                            record.LeaveTime = Convert.ToDateTime(Convert.ToDateTime(DefaultEndWorkTime).ToShortTimeString());
                        }

                        if(DateTime.Compare(record.ArriveTime, Convert.ToDateTime(StartWorkTime)) > 0)
                        {
                            lateCount += 1;
                        }

                        attendanceMinutes += record.LeaveTime.Subtract(record.ArriveTime).Duration().TotalMinutes;
                    }
                    absenteeism += group.Count(r => (r.ArriveTime.ToString() == "0001/1/1 0:00:00" && r.LeaveTime.ToString() == "0001/1/1 0:00:00") || (r.ArriveTime.ToString() == "0001-01-01 0:00:00" && r.LeaveTime.ToString() == "0001-01-01 0:00:00"));

                    var attendanceHour = Math.Floor(attendanceMinutes / 60);

                    var personnelX = personnels.FirstOrDefault(p => p.PersonnelName == group.Key);
                    attendanceCollection.Add(new AttendanceRecordModel
                    {
                        PersonnelIndex = index,
                        DepartmentName = personnelX == null ? "" : personnelX.DepartmentName,//部门
                        PersonnelName = group.Key,//人员姓名
                        SickLeave = 0,//病假（天）
                        CumulativeSickLeave = personnelX == null ? 0 : personnelX.CumulativeSickLeave,//累计病假
                        PrivateLeave = 0,//事假（天）
                        CumulativePrivateLeave = personnelX == null ? 0 : personnelX.CumulativePrivateLeave,//累计事假
                        TransformLeave = 0,//病事假转换
                        FreeAnnualLeave = personnelX == null ? 0 : personnelX.FreeAnnualLeave,//可休年假
                        UsedAnnualLeave = personnelX == null ? 0 : personnelX.UsedAnnualLeave,//已休年假
                        RemainingAnnualLeave = personnelX == null ? 0 : personnelX.RemainingAnnualLeave,//剩余年假
                        LeaveName = "",//其他假别
                        LateTime = lateCount,//迟到（次）
                        Absenteeism = absenteeism,//旷工（天）
                        AttendanceHour = attendanceHour,//本月出勤工时（小时）
                        OvertimeHour = attendanceHour - DefaultWorkHours > 0 ? attendanceHour - DefaultWorkHours : 0,//本月加班（小时）
                        Note = "",//备注
                    });

                    index++;
                }

                AttendanceCollection = attendanceCollection;
                StatusString = "";
            }
            catch(Exception ex)
            {
                StatusString = "";
                MessageBox.Show(ex.Message);
            }
        }

        private ObservableCollection<PersonnnelModel> ReadPersonnelRecord()
        {
            try
            {
                IWorkbook workbook = null;
                var fs = new FileStream(_personnelRecordFileUri, FileMode.Open, FileAccess.Read);
                if(_personnelRecordFileUri.IndexOf(".xlsx", StringComparison.Ordinal) > 0) // 2007版本
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if(_personnelRecordFileUri.IndexOf(".xls", StringComparison.Ordinal) > 0) // 2003版本
                {
                    workbook = new HSSFWorkbook(fs);
                }

                var personnels = new ObservableCollection<PersonnnelModel>();

                if(workbook != null)
                {
                    var sheet = workbook.GetSheetAt(0);

                    //获取sheet的首行
                    XSSFRow headerRow = (XSSFRow)sheet.GetRow(0);

                    //总的列数
                    int cellCount = headerRow.LastCellNum;
                    for(int i = (sheet.FirstRowNum + 1); i < sheet.LastRowNum; i++)
                    {
                        XSSFRow row = (XSSFRow)sheet.GetRow(i);
                        var personnel = new PersonnnelModel();

                        for(int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            var cell = row.GetCell(j);

                            switch(j)
                            {
                                case 0:
                                    personnel.PersonnelIndex = Convert.ToInt32(cell.ToString());
                                    break;
                                case 1:
                                    personnel.PersonnelName = cell.ToString();
                                    break;
                                case 2:
                                    personnel.DepartmentName = cell.ToString();
                                    break;
                                case 3:
                                    personnel.FreeAnnualLeave = cell == null ? 0 : Convert.ToDouble(cell.ToString());
                                    break;
                                case 4:
                                    personnel.UsedAnnualLeave = cell == null ? 0 : Convert.ToDouble(cell.ToString());
                                    break;
                                case 5:
                                    personnel.RemainingAnnualLeave = cell == null ? 0 : Convert.ToDouble(cell.ToString());
                                    break;
                                case 6:
                                    personnel.CumulativeSickLeave = cell == null ? 0 : Convert.ToDouble(cell.ToString());
                                    break;
                                case 7:
                                    personnel.CumulativePrivateLeave = cell == null ? 0 : Convert.ToDouble(cell.ToString());
                                    break;
                            }
                        }
                        personnels.Add(personnel);
                    }
                }

                return personnels;
            }
            catch(Exception ex)
            {
                StatusString = "";
                MessageBox.Show(ex.Message);
                return new ObservableCollection<PersonnnelModel>();
            }
        }

        private void ToggleFlyout(int index)
        {
            try
            {
                var mainView = GetView() as MainView;
                var flyout = mainView.Flyouts.Items[index] as Flyout;
                if(flyout == null)
                {
                    return;
                }

                flyout.IsOpen = !flyout.IsOpen;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
