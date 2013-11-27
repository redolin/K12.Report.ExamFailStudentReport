using Aspose.Cells;
using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace K12.Report.ExamFailStudentReport.Forms
{
    public partial class FrmExamFailCondition : BaseForm
    {
        List<DAO.ExamVO> _ExamList = new List<DAO.ExamVO>();
        List<string> _ClassIdList;
        List<DAO.ClassVO> _ClassList;

        Dictionary<string, DAO.StudentVO> _MultiTagStudentDic = new Dictionary<string, DAO.StudentVO>();
        Dictionary<string, string> _AllCourseDic = new Dictionary<string, string>();
        DAO.ConfigureRecord _Configure;
        bool _IsProcessWarningStudent = true;

        private readonly string _DefaultTagName = "預設";
        private readonly decimal _DefaultPassScore = 60;
        private readonly string _SheetListName = "學分未達標準總表";
        private readonly string _SheetDetailName = "成績明細";
        private readonly string _NoScroe = "缺";
        private readonly string _ExcelFileName = "不及格成績單";
        private readonly int _MAX_ROW_COUNT = 65535;
        private readonly string _Title = "評量成績未達標準名單";

        decimal _PassRate = 50;
        DAO.ExamVO _ExamObj;
        // 取得所有資料
        BackgroundWorker _BGW_Step1 = new BackgroundWorker();
        // 計算不及格學生
        BackgroundWorker _BGW_Step2 = new BackgroundWorker();
        int _ListRowIndex = 0;
        int _DetailRowIndex = 0;

        public FrmExamFailCondition()
        {
            InitializeComponent();
            this.Text = _Title;
            _ClassIdList = K12.Presentation.NLDPanels.Class.SelectedSource;
            // TODO, 看能不能改成多執行緒
            _ExamList = DAO.FDQuery.GetDistincExamList(_ClassIdList);

            // set default value
            _Configure = new DAO.ConfigureRecord();
            _Configure.SchoolYear = K12.Data.School.DefaultSchoolYear;
            _Configure.Semester = K12.Data.School.DefaultSemester;
            _Configure.ExamId = "";
            _Configure.PassCreditRate = "50";

            // get configure value if exists
            GetConfigure();
        }

        private void FrmExamFailCondition_Load(object sender, EventArgs e)
        {
            
            SetComponentValue();

            _BGW_Step1.DoWork += new DoWorkEventHandler(BGW_DoWork_Step1);
            _BGW_Step1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted_Step1);

            _BGW_Step2.DoWork += new DoWorkEventHandler(BGW_DoWork_Step2);
            _BGW_Step2.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted_Step2);
        }

        private void cboSchoolYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            ShowExamList();
        }

        private void cboSemester_SelectedIndexChanged(object sender, EventArgs e)
        {
            ShowExamList();
        }


        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            FormComponetEnable(false);

            if (CheckData() == false)
            {
                MsgBox.Show("還有條件尚未選擇!");
                FormComponetEnable(true);
                return;
            }

            SaveConfigure();

            _PassRate = decimal.Parse(this.iiPassRate.Text);
            _ExamObj = this.cboExamList.SelectedItem as DAO.ExamVO;

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "另存新檔";
            saveFileDialog1.FileName = _ExcelFileName + "(" + _ExamObj.SchoolYear + "學年度 第" + _ExamObj.Semester + "學期 " + _ExamObj.ExameName + ")";
            saveFileDialog1.Filter = "Excel (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // 新增背景執行緒來處理資料的輸出, 把檔案儲存的路徑當作參數傳入
                _BGW_Step1.RunWorkerAsync(new object[] { saveFileDialog1.FileName });
            }
            else
            {
                FormComponetEnable(true);
            }
            
        }

        #region 背景執行緒方法

        #region BGW Step1
        // 主要邏輯區塊
        void BGW_DoWork_Step1(object sender, DoWorkEventArgs e)
        {

            string fileName = (string)((object[])e.Argument)[0];

            // 取得學生所有修課的成績, 包括學生自己的特殊計算規則假如有的話
            Dictionary<string, DAO.ClassVO> ClassListDic = DAO.FDQuery.GetAllStudentScore(_ClassIdList, _ExamObj);

            // 取得學生類別
            DAO.FDQuery.GetAllStudentTag(ClassListDic, _ClassIdList);

            // 取得班級計分規則
            DAO.FDQuery.GetAllClassCalRule(ClassListDic, _ClassIdList);

            // 班級排序
            _ClassList = ClassListDic.Values.ToList();
            _ClassList.Sort(SortClass);

            // 排序學生
            foreach (DAO.ClassVO obj in _ClassList)
            {
                obj.StudentList = obj.StudentListDic.Values.ToList();
                obj.StudentList.Sort(SortStudent);
            }

            // 判斷學生及格分數
            #region 判斷學生及格分數
            foreach (DAO.ClassVO ClassObj in _ClassList)
            {
                foreach(DAO.StudentVO StudentObj in ClassObj.StudentList)
                {
                    #region 從學生身上的類別找出及格標準
                    foreach(string TagName in StudentObj.StudentTag)
                    {
                        decimal PassScore = Utility.GetPassScroe(StudentObj.PassCriterionDic, TagName, ClassObj.ClassGradeYear);
                        if(PassScore != -1)
                        {
                            if(StudentObj.PassScore != -1)
                            {
                                // 有多重及格標準
                                StudentObj.IsMultiTag = true;
                                StudentObj.PassScore = Utility.MinPassScroe(StudentObj.PassScore, PassScore);

                                _MultiTagStudentDic.Add(StudentObj.StudentId, StudentObj);
                            }
                            else
                            {
                                StudentObj.PassScore = PassScore;
                            }
                        }
                    }
                    #endregion

                    #region 從"預設"的類別找出及格標準
                    if (StudentObj.PassScore == -1)
                    {
                        decimal PassScore = Utility.GetPassScroe(StudentObj.PassCriterionDic, _DefaultTagName, ClassObj.ClassGradeYear);
                        if(PassScore == -1)
                        {
                            // 假如"預設"的類別還是找不到, 就用60當作及格
                            StudentObj.PassScore = _DefaultPassScore;
                        }
                        else
                        {
                            StudentObj.PassScore = PassScore;
                        }
                    }
                    #endregion
                }
            }
            #endregion

            
            e.Result = new object[] { fileName };
            
        }

        // 當背景程式結束, 就會呼叫method
        void BGW_RunWorkerCompleted_Step1(object sender, RunWorkerCompletedEventArgs e)
        {
            string fileName = (string)((object[])e.Result)[0];
            bool IsContinue = false;
            // 假如有學生有多重及格標準 顯示給使用者看
            if(_MultiTagStudentDic.Count > 0)
            {
                FrmWarningStudent frm = new FrmWarningStudent(_MultiTagStudentDic);
                DialogResult result = frm.ShowDialog();

                if(result == DialogResult.OK)
                {
                    // 繼續產生
                    IsContinue = true;
                    _IsProcessWarningStudent = true;
                }
                else if(result == DialogResult.Ignore)
                {
                    // 加到待處理名單
                    IsContinue = true;
                    _IsProcessWarningStudent = false;
                }
                else if(result == DialogResult.Cancel)
                {
                    // 關閉
                    IsContinue = false;
                }
            }
            else
            {
                IsContinue = true;
                _IsProcessWarningStudent = true;
            }
             

            // 進行Step2
            if (IsContinue == true)
            {
                _BGW_Step2.RunWorkerAsync(new object[] { fileName });
            }
        }
        #endregion

        #region BGW Step2
        void BGW_DoWork_Step2(object sender, DoWorkEventArgs e)
        {
            string fileName = (string)((object[])e.Argument)[0];

            #region 計算學生的學分數
            // 計算學生的學分數
            foreach (DAO.ClassVO ClassObj in _ClassList)
            {
                foreach (DAO.StudentVO StudentObj in ClassObj.StudentList)
                {
                    
                    if(_IsProcessWarningStudent == false && StudentObj.IsMultiTag == true)
                        continue;

                    decimal PassScore = StudentObj.PassScore;
                    decimal TotalCredit = 0;
                    decimal GotCredit = 0;
                    foreach (DAO.CourseVO CourseObj in StudentObj.CourseListDic.Values)
                    {
                        TotalCredit += CourseObj.Credit;

                        if (CourseObj.CourseScore >= PassScore)
                        {
                            GotCredit += CourseObj.Credit;
                            CourseObj.IsPass = true;
                        }
                        else
                        {
                            CourseObj.IsPass = false;
                        }
                    }

                    StudentObj.TotalCredit = TotalCredit;
                    StudentObj.GotCredit = GotCredit;

                    decimal rate = Math.Round((GotCredit / TotalCredit) * 100, 2, MidpointRounding.AwayFromZero);

                    StudentObj.CreditRate = rate;
                    if (_PassRate > rate)
                    {
                        ClassObj.EaxmFailStudentList.Add(StudentObj);
                    }
                    else
                    {
                        StudentObj.IsPass = true;
                    }
                }
            }   // end of 計算學生的學分數
            #endregion

            // 輸出到Excel
            #region 輸出到Excel
            Workbook report = new Workbook();
            Worksheet sheetList = report.Worksheets[0];
            Worksheet sheetDetail = report.Worksheets[report.Worksheets.Add()];

            sheetList.Name = _SheetListName;
            sheetDetail.Name = _SheetDetailName;

            // 輸出List表頭
            _ListRowIndex = 0;
            _DetailRowIndex = 0;

            Cells ListCells = sheetList.Cells;
            Cells DetailCells = sheetDetail.Cells;

            OutListTitle(ListCells);

            foreach (DAO.ClassVO ClassObj in _ClassList)
            {
                // 輸出Detail表頭
                OutDetailTitle(DetailCells, ClassObj);

                foreach (DAO.StudentVO StudentObj in ClassObj.EaxmFailStudentList)
                {
                    // 輸出資料到List
                    OutListData(ListCells, ClassObj, StudentObj);

                    // 輸出資料到Detail
                    OutDetailData(DetailCells, ClassObj, StudentObj);
                }

                // 輸出Deatil分頁標籤
                OutDetailSplitPage(DetailCells);
            }
            #endregion

            // for test
            OutRowData(report);

            // 儲存結果
            e.Result = new object[] { report, fileName, _DetailRowIndex > _MAX_ROW_COUNT };
        }

        void BGW_RunWorkerCompleted_Step2(object sender, RunWorkerCompletedEventArgs e)
        {
            FormComponetEnable(true);

            if (e.Error == null)
            {
                Workbook report = (Workbook)((object[])e.Result)[0];
                bool overLimit = (bool)((object[])e.Result)[2];

                #region 儲存 Excel
                string path = (string)((object[])e.Result)[1];

                if (File.Exists(path))
                {
                    bool needCount = true;
                    try
                    {
                        File.Delete(path);
                        needCount = false;
                    }
                    catch { }
                    int i = 1;
                    while (needCount)
                    {
                        string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                        if (!File.Exists(newPath))
                        {
                            path = newPath;
                            break;
                        }
                        else
                        {
                            try
                            {
                                File.Delete(newPath);
                                path = newPath;
                                break;
                            }
                            catch { }
                        }
                    }
                }
                try
                {
                    File.Create(path).Close();
                }
                catch
                {
                    SaveFileDialog sd = new SaveFileDialog();
                    sd.Title = "另存新檔";
                    sd.FileName = Path.GetFileNameWithoutExtension(path) + ".xls";
                    sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
                    if (sd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            File.Create(sd.FileName);
                            path = sd.FileName;
                        }
                        catch
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
                report.Save(path, FileFormatType.Excel2003);
                #endregion
                if (overLimit)
                    MsgBox.Show("匯出資料已經超過Excel的極限(65536筆)。\n超出的資料無法被匯出。\n\n請減少選取學生人數。");
                System.Diagnostics.Process.Start(path);
            }
            else
            {
                MsgBox.Show(_Title + "發生未預期錯誤。\n" + e.Error.Message);
            }
        }
        #endregion

        #endregion

        #region Excel輸出方法
        private void OutListTitle(Cells cells)
        {
            int columnIndex = 0;
            string title = K12.Data.School.ChineseName + " " + _ExamObj.ExameName + " 未達" + _PassRate + "%學分名單";
            cells[_ListRowIndex, columnIndex++].PutValue(title);

            _ListRowIndex++;
            columnIndex = 0;
            // 班級 	座號	姓名	學號	應得學分數	實得學分數	未達百分比
            cells[_ListRowIndex, columnIndex++].PutValue("班級");
            cells[_ListRowIndex, columnIndex++].PutValue("座號");
            cells[_ListRowIndex, columnIndex++].PutValue("姓名");
            cells[_ListRowIndex, columnIndex++].PutValue("學號");
            cells[_ListRowIndex, columnIndex++].PutValue("應得學分數");
            cells[_ListRowIndex, columnIndex++].PutValue("實得學分數");
            cells[_ListRowIndex, columnIndex++].PutValue("未達百分比");
        }

        private void OutDetailTitle(Cells cells, DAO.ClassVO ClassObj)
        {
            int columnIndex = 0;
            string title = K12.Data.School.ChineseName + " " + ClassObj.ClassName + "  " + _ExamObj.ExameName + " 未達" + _PassRate + "%學分名單";
            cells[_DetailRowIndex, columnIndex++].PutValue(title);

            // 座號	姓名	學號	國文Ⅳ	英文Ⅳ	... 及格標準	應得學分數	實得學分數	未達百分比
            _DetailRowIndex++;
            columnIndex = 0;
            cells[_DetailRowIndex, columnIndex++].PutValue("座號");
            cells[_DetailRowIndex, columnIndex++].PutValue("姓名");
            cells[_DetailRowIndex, columnIndex++].PutValue("學號");
            foreach(DAO.CourseVO CourseObj in ClassObj.AllCourseListDic.Values)
            {
                cells[_DetailRowIndex, columnIndex++].PutValue(CourseObj.CourseName);
            }
            cells[_DetailRowIndex, columnIndex++].PutValue("及格標準");
            cells[_DetailRowIndex, columnIndex++].PutValue("應得學分數");
            cells[_DetailRowIndex, columnIndex++].PutValue("實得學分數");
            cells[_DetailRowIndex, columnIndex++].PutValue("未達百分比");

            _DetailRowIndex++;
            cells[_DetailRowIndex, 0].PutValue("學分數");  // 合併儲存格, TODO
            // 學分數 
            columnIndex = 3;
            foreach (DAO.CourseVO CourseObj in ClassObj.AllCourseListDic.Values)
            {
                cells[_DetailRowIndex, columnIndex++].PutValue(CourseObj.Credit);
            }
        }

        private void OutListData(Cells cells, DAO.ClassVO ClassObj, DAO.StudentVO StudentObj)
        {
            int columnIndex = 0;
            _ListRowIndex++;
            // 班級	座號	姓名	學號	應得學分數	實得學分數	未達百分比
            cells[_ListRowIndex, columnIndex++].PutValue(ClassObj.ClassName);
            cells[_ListRowIndex, columnIndex++].PutValue(StudentObj.SeatNo);
            cells[_ListRowIndex, columnIndex++].PutValue(StudentObj.StudentName);
            cells[_ListRowIndex, columnIndex++].PutValue(StudentObj.StudentNumber);
            cells[_ListRowIndex, columnIndex++].PutValue(StudentObj.TotalCredit);
            cells[_ListRowIndex, columnIndex++].PutValue(StudentObj.GotCredit);
            cells[_ListRowIndex, columnIndex++].PutValue(StudentObj.CreditRate + "%");
        }

        private void OutDetailData(Cells cells, DAO.ClassVO ClassObj, DAO.StudentVO StudentObj)
        {
            int columnIndex = 0;
            _DetailRowIndex++;
            // 座號	姓名	學號	國文Ⅳ	英文Ⅳ	... 及格標準	應得學分數	實得學分數	未達百分比
            cells[_DetailRowIndex, columnIndex++].PutValue(StudentObj.SeatNo);
            cells[_DetailRowIndex, columnIndex++].PutValue(StudentObj.StudentName);
            cells[_DetailRowIndex, columnIndex++].PutValue(StudentObj.StudentNumber);

            #region 輸出每一科的分數
            foreach (DAO.CourseVO CourseObj in ClassObj.AllCourseListDic.Values)
            {
                DAO.CourseVO StuCourseObj = Utility.FindCourse(StudentObj.CourseListDic.Values.ToList(), CourseObj.CourseId);
                if (StuCourseObj == null)
                {
                    // 沒有修這堂課
                    cells[_DetailRowIndex, columnIndex++].PutValue("");
                }
                else
                {
                    if (StuCourseObj.CourseScore == -1)
                    {
                        // 有這堂課, 卻沒有分數
                        cells[_DetailRowIndex, columnIndex++].PutValue(_NoScroe);
                    }
                    else
                    {
                        cells[_DetailRowIndex, columnIndex++].PutValue(StuCourseObj.CourseScore);
                    }

                    // TODO
                    if (StuCourseObj.IsPass == true)
                    {
                        // 黑色的字
                    }
                    else
                    {
                        // 紅色的字
                    }
                    
                }
            }
            #endregion

            cells[_DetailRowIndex, columnIndex++].PutValue(StudentObj.PassScore);
            cells[_DetailRowIndex, columnIndex++].PutValue(StudentObj.TotalCredit);
            cells[_DetailRowIndex, columnIndex++].PutValue(StudentObj.GotCredit);
            cells[_DetailRowIndex, columnIndex++].PutValue(StudentObj.CreditRate + "%");
        }

        private void OutDetailSplitPage(Cells cells)
        {
            // TODO
            _DetailRowIndex+=2;
        }

        private void OutRowData(Workbook wb)
        {
            Worksheet sheet = wb.Worksheets[wb.Worksheets.Add()];
            Cells cells = sheet.Cells;
            int rowIndex = 0;
            int colIndex = 0;

            string[] title = {"班級", "學生ID", "學號", "姓名", "應得學分", "實得學分", "比例", "及格分數"};

            foreach (DAO.ClassVO ClassObj in _ClassList)
            {
                colIndex = 0;
                foreach(string name in title)
                {
                    cells[rowIndex, colIndex++].PutValue(name);
                }
                foreach(DAO.CourseVO CourseObj in ClassObj.AllCourseListDic.Values)
                {
                    cells[rowIndex,colIndex++].PutValue(CourseObj.CourseName);
                }

                foreach(DAO.StudentVO StudentObj in ClassObj.StudentList)
                {
                    colIndex = 0;
                    rowIndex++;

                    // "班級", "學生ID", "學號", "姓名", "應得學分", "實得學分", "比例", "及格分數"
                    cells[rowIndex, colIndex++].PutValue(ClassObj.ClassName);
                    cells[rowIndex, colIndex++].PutValue(StudentObj.StudentId);
                    cells[rowIndex, colIndex++].PutValue(StudentObj.StudentNumber);
                    cells[rowIndex, colIndex++].PutValue(StudentObj.StudentName);
                    cells[rowIndex, colIndex++].PutValue(StudentObj.TotalCredit);
                    cells[rowIndex, colIndex++].PutValue(StudentObj.GotCredit);
                    cells[rowIndex, colIndex++].PutValue(StudentObj.CreditRate);
                    cells[rowIndex, colIndex++].PutValue(StudentObj.PassScore);
                    foreach(DAO.CourseVO CourseObj in ClassObj.AllCourseListDic.Values)
                    {

                        DAO.CourseVO StuCourseObj = Utility.FindCourse(StudentObj.CourseListDic.Values.ToList(), CourseObj.CourseId);
                        if (StuCourseObj == null)
                        {
                            cells[rowIndex, colIndex++].PutValue("");
                        }
                        else
                        {
                            cells[rowIndex, colIndex++].PutValue(StuCourseObj.CourseScore);
                        }
                    }
                }
            }
        }   // end of OutRowData
        #endregion

        #region 自訂的方法
        private void SetComponentValue()
        {
            // 開始填入內容
            foreach (DAO.ExamVO ExamObj in _ExamList)
            {
                if (!this.cboSchoolYear.Items.Contains(ExamObj.SchoolYear))
                    this.cboSchoolYear.Items.Add(ExamObj.SchoolYear);
            }

            this.cboSemester.Items.Add("1");
            this.cboSemester.Items.Add("2");

            // 填入上次使用者的內容
            this.iiPassRate.Text = _Configure.PassCreditRate;

            foreach (string str in this.cboSchoolYear.Items)
            {
                if (str == _Configure.SchoolYear)
                {
                    // will call ShowExamList
                    this.cboSchoolYear.SelectedItem = str;
                    break;
                }
            }

            foreach (string str in this.cboSemester.Items)
            {
                if (str == _Configure.Semester)
                {
                    // will call ShowExamList
                    this.cboSemester.SelectedItem = str;
                    break;
                }
            }

            for (int intI = 1; intI <= this.cboExamList.Items.Count; intI++)
            {
                DAO.ExamVO ExamObj = this.cboExamList.Items[intI] as DAO.ExamVO;

                if (ExamObj.ExamId == _Configure.ExamId)
                {
                    this.cboExamList.SelectedItem = ExamObj;
                    break;
                }
            }

        }

        private void ShowExamList()
        {
            if (this.cboSchoolYear.SelectedIndex < 0 || this.cboSemester.SelectedIndex < 0)
                return;

            FormComponetEnable(false);
            this.cboExamList.Items.Clear();
            this.cboExamList.Items.Add("");
            string SchoolYear = this.cboSchoolYear.SelectedItem.ToString();
            string Semester = this.cboSemester.SelectedItem.ToString();

            foreach (DAO.ExamVO exam in _ExamList)
            {
                if (exam.SchoolYear == SchoolYear && exam.Semester == Semester)
                {
                    this.cboExamList.Items.Add(exam);
                }
            }
            this.cboExamList.DisplayMember = "ExameName";

            FormComponetEnable(true);
            if (this.cboExamList.Items.Count > 0)
                this.cboExamList.SelectedIndex = 0;
            this.cboExamList.Focus();
        }

        private void FormComponetEnable(bool enabled)
        {
            this.cboSchoolYear.Enabled = enabled;
            this.cboSemester.Enabled = enabled;
            this.cboExamList.Enabled = enabled;
            this.iiPassRate.Enabled = enabled;
            this.btnPrint.Enabled = enabled;
        }

        private bool CheckData()
        {
            bool IsPass = true;

            if (this.cboSchoolYear.SelectedIndex < 0)
                return false;

            if (this.cboSemester.SelectedIndex < 0)
                return false;

            if (this.cboExamList.SelectedIndex <= 0)
                return false;

            return IsPass;
        }
        #endregion

        #region 畫面內容的取得跟儲存
        private void GetConfigure()
        {
            // 取得設定
            List<DAO.ConfigureRecord> ConfList = DAO.Configure.SelectConfigure();
            if (ConfList.Count > 0)
            {
                _Configure = ConfList[0];
            }
        }

        private void SaveConfigure()
        {
            _Configure.SchoolYear = this.cboSchoolYear.SelectedItem.ToString();
            _Configure.Semester = this.cboSemester.SelectedItem.ToString();
            _Configure.ExamId = ((DAO.ExamVO)this.cboExamList.SelectedItem).ExamId;
            _Configure.PassCreditRate = this.iiPassRate.Text;
            DAO.Configure.UpdateByRecord(_Configure);
        }
        #endregion

        #region 排序方法
        /// <summary>
        /// 排序:年級/班級序號/班級名稱
        /// </summary>
        /// <param name="obj1"></param>
        /// <param name="obj2"></param>
        private int SortClass(DAO.ClassVO obj1, DAO.ClassVO obj2)
        {
            string seatno1 = obj1.ClassGradeYear.PadLeft(1, '0');   // 年級
            seatno1 += obj1.ClassDisplyOrder.PadLeft(3, '0');       // 班級序號
            seatno1 += obj1.ClassName.PadLeft(20, '0');             // 班級名稱

            string seatno2 = obj2.ClassGradeYear.PadLeft(1, '0');   // 年級
            seatno1 += obj2.ClassDisplyOrder.PadLeft(3, '0');       // 班級序號
            seatno1 += obj2.ClassName.PadLeft(20, '0');             // 班級名稱

            return seatno1.CompareTo(seatno2);
        }

        /// <summary>
        /// 排序:座號/學號/姓名
        /// </summary>
        private int SortStudent(DAO.StudentVO obj1, DAO.StudentVO obj2)
        {
            string seatno1 = obj1.SeatNo.PadLeft(3, '0');
            seatno1 += obj1.StudentNumber.PadLeft(20, '0');
            seatno1 += obj1.StudentName.PadLeft(10, '0');

            string seatno2 = obj2.SeatNo.PadLeft(3, '0');
            seatno2 += obj2.StudentNumber.PadLeft(20, '0');
            seatno2 += obj2.StudentName.PadLeft(10, '0');

            return seatno1.CompareTo(seatno2);
        }
        #endregion
    }
}
