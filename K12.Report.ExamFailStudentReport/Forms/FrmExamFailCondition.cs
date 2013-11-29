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

        // 是否要處理有雙重及格標準
        bool _IsProcessWarningStudent = true;
        // 雙重及格標準學生
        Dictionary<string, DAO.StudentVO> _MultiTagStudentDic = new Dictionary<string, DAO.StudentVO>();
        // 未達百分比
        decimal _PassRate = 50;
        // 目前畫面選的試別
        DAO.ExamVO _ExamObj;
        // 畫面設定
        DAO.ConfigureRecord _Configure;
        // 取得所有學生的所有課程分數
        BackgroundWorker _BGW_Step1 = new BackgroundWorker();
        // 計算不及格學生
        BackgroundWorker _BGW_Step2 = new BackgroundWorker();
        // 取得所有試別
        BackgroundWorker _BGW_ExamList = new BackgroundWorker();

        #region Excel用的
        int _ListRowIndex = 0;
        int _DetailRowIndex = 0;
        private static readonly string _ExcelFileName = "不及格成績單";
        private static readonly string _SheetListName = "學分未達標準總表";
        private static readonly string _SheetDetailName = "成績明細";
        private static readonly string _NoScroe = "缺";
        private static readonly int _MAX_ROW_COUNT = 65535;
        private static readonly int _MAX_List_Column = 7;
        private static readonly int _Default_Detail_Course_Count = 20;
        // 20: 預設版型有20個課程, 3: 班級,座號,姓名 4: 及格分數 應得 實得 比例
        private static readonly int _Default_Detail_Column = _Default_Detail_Course_Count + 3 + 4;
        // 37, 50, 26 是標題的高度, 25是資料的高度
        private static readonly int[] _Detail_Row_Height = { 37, 50, 26, 25 };
        private static readonly int _List_Row_Height = 22;
        private static readonly int _Detail_Column_Weight = 42;

        private static readonly string[] _List_Row_Range_Cell = { "A3", "G3" };
        private Range _List_Row_Range;

        private int _Current_Course_Count;
        private Style _Detail_Style_Normal;
        private Style _Detail_Style_Red;
        private static readonly string[] _Detail_Title_Range_Cell = { "A1", "AA3" };
        private Range _Detail_Title_Range;
        #endregion

        public FrmExamFailCondition()
        {
            InitializeComponent();
            this.Text = Global.Title;
            _ClassIdList = K12.Presentation.NLDPanels.Class.SelectedSource;

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
            _BGW_Step1.WorkerReportsProgress = true;
            _BGW_Step1.DoWork += new DoWorkEventHandler(BGW_DoWork_Step1);
            _BGW_Step1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted_Step1);
            _BGW_Step1.ProgressChanged += new ProgressChangedEventHandler(BGW_ProgressChanged);

            _BGW_Step2.WorkerReportsProgress = true;
            _BGW_Step2.DoWork += new DoWorkEventHandler(BGW_DoWork_Step2);
            _BGW_Step2.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted_Step2);
            _BGW_Step2.ProgressChanged += new ProgressChangedEventHandler(BGW_ProgressChanged);

            _BGW_ExamList.DoWork += new DoWorkEventHandler(BGW_DoWork_ExamList);
            _BGW_ExamList.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted_ExamList);

            FormComponetEnable(false);
            this.Loading.Visible = true;
            this.Loading.IsRunning = true;
            _BGW_ExamList.RunWorkerAsync();

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

            if (CheckData() == false) return;

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
        #region BGW _ExamList
        // 取得不重複的試別列表
        void BGW_DoWork_ExamList(object sender, DoWorkEventArgs e)
        {
            _ExamList = DAO.FDQuery.GetDistincExamList(_ClassIdList);
        }

        // 產生畫面的資料
        void BGW_RunWorkerCompleted_ExamList(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Loading.Visible = false;
            FormComponetEnable(true);
            SetComponentValue();
        }
        #endregion

        #region BGW Step1
        // 儲存所有學生的所有課程分數, 跟及格標準
        void BGW_DoWork_Step1(object sender, DoWorkEventArgs e)
        {
            string fileName = (string)((object[])e.Argument)[0];

            // 取得學生所有修課的成績, 包括學生自己的特殊計算規則假如有的話
            Dictionary<string, DAO.ClassVO> ClassListDic = DAO.FDQuery.GetAllStudentScore(_ClassIdList, _ExamObj);
            _BGW_Step1.ReportProgress(10);

            // 取得學生類別
            DAO.FDQuery.GetAllStudentTag(ClassListDic, _ClassIdList);

            // 取得班級計分規則
            DAO.FDQuery.GetAllClassCalRule(ClassListDic, _ClassIdList);
            _BGW_Step1.ReportProgress(20);

            // 班級排序
            _ClassList = ClassListDic.Values.ToList();
            _ClassList.Sort(SortClass);
            _BGW_Step1.ReportProgress(30);

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

                                if (!_MultiTagStudentDic.Keys.Contains(StudentObj.StudentId))
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
                        decimal PassScore = Utility.GetPassScroe(StudentObj.PassCriterionDic, Global.DefaultTagName, ClassObj.ClassGradeYear);
                        if(PassScore == -1)
                        {
                            // 假如"預設"的類別還是找不到, 就用60當作及格
                            StudentObj.PassScore = Global.DefaultPassScore;
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


            _BGW_Step1.ReportProgress(50);

            e.Result = new object[] { fileName };
            
        }

        // 判斷有沒有兩個以上的及格標準
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
            else
            {
                FormComponetEnable(true);
                FISCA.Presentation.MotherForm.SetStatusBarMessage(Global.Title + "產生取消。", 100);
            }
        }
        #endregion

        #region BGW Step2
        // 計算所有學生應得/實得的學分數跟比例, 並輸出到Excel
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

            _BGW_Step2.ReportProgress(60);

            // 輸出到Excel
            #region 輸出到Excel
            Workbook report = new Workbook();
            report.Open(new MemoryStream(Properties.Resources.不及格成績單_Template));
            Worksheet sheetList = report.Worksheets[_SheetListName];
            Worksheet sheetDetail = report.Worksheets[_SheetDetailName];

            // 先儲存Style
            GetExcelStyle(report);

            // 輸出List表頭
            _ListRowIndex = 0;
            _DetailRowIndex = 0;

            Cells ListCells = sheetList.Cells;
            Cells DetailCells = sheetDetail.Cells;

            OutListTitle(ListCells);
            _BGW_Step2.ReportProgress(70);

            foreach (DAO.ClassVO ClassObj in _ClassList)
            {
                if (ClassObj.EaxmFailStudentList.Count <= 0) continue;

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
                OutDetailSplitPage(sheetDetail);
            }
            #endregion

            // for test
            if (Global.IsDebug) OutRowData(report);

            _BGW_Step2.ReportProgress(90);
            // 儲存結果
            e.Result = new object[] { report, fileName, _DetailRowIndex > _MAX_ROW_COUNT };
        }

        // 儲存Excel
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

                FISCA.Presentation.MotherForm.SetStatusBarMessage( Global.Title + "產生完成。", 100);

                System.Diagnostics.Process.Start(path);
            }
            else
            {
                MsgBox.Show(Global.Title + "發生未預期錯誤。\n" + e.Error.Message);
            }
        }
        #endregion

        void BGW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage( Global.Title + "產生中", e.ProgressPercentage );
        }
        
        #endregion

        #region Excel輸出方法
        // 準備Excel的Style
        private void GetExcelStyle(Workbook wb)
        {
            Cells ListCells = wb.Worksheets[_SheetListName].Cells;
            Cells DetailCells = wb.Worksheets[_SheetDetailName].Cells;

            // 取得 "成績明細" 中正常的Style
            _Detail_Style_Normal = DetailCells[3, 0].Style;
            // 取得 "成績明細" 中不正常的Style (紅色的字體)
            _Detail_Style_Red = DetailCells[3, 1].Style;
            // 取完Style之後, 改為正常的
            DetailCells[3, 1].Style = _Detail_Style_Normal;

            // 複製 "總表" 中第一列, 準備之後輸出時複製Style用
            _List_Row_Range = ListCells.CreateRange(_List_Row_Range_Cell[0], _List_Row_Range_Cell[1]);

            // 複製 "成績明細" 中的標題, 準備之後輸出時複製Style用
            _Detail_Title_Range = DetailCells.CreateRange(_Detail_Title_Range_Cell[0], _Detail_Title_Range_Cell[1]);
        }

        #region 針對"成績明細"的處理
        private void SetDetailColumnValue(Cells cells, int col, string value,  Style style)
        {
            cells[_DetailRowIndex ,col].PutValue(value);
            cells[_DetailRowIndex ,col].Style = style;
            cells.SetColumnWidthPixel(col, _Detail_Column_Weight);
        }

        private void SetDetailColumnValue(Cells cells, int col, decimal value, Style style)
        {
            cells[_DetailRowIndex, col].PutValue(value);
            cells[_DetailRowIndex, col].Style = style;
            cells.SetColumnWidthPixel(col, _Detail_Column_Weight);
        }

        private void OutDetailTitle(Cells cells, DAO.ClassVO ClassObj)
        {
            int columnIndex = 0;
            // 先複製格式
            Range TitleRange = cells.CreateRange(_DetailRowIndex, columnIndex, 3, _Default_Detail_Column);
            TitleRange.CopyStyle(_Detail_Title_Range);

            // 設定高度
            cells.SetRowHeightPixel(_DetailRowIndex, _Detail_Row_Height[0]);
            cells.SetRowHeightPixel((_DetailRowIndex+1), _Detail_Row_Height[1]);
            cells.SetRowHeightPixel((_DetailRowIndex+2), _Detail_Row_Height[2]);
            
            // 合併儲存格
            cells.Merge(_DetailRowIndex, columnIndex, 1, _Default_Detail_Column);

            // 輸出資料
            columnIndex = 0;
            string title = K12.Data.School.ChineseName + " " + ClassObj.ClassName + "  " + _ExamObj.ExameName + " 未達" + _PassRate + "%學分名單";
            cells[_DetailRowIndex, columnIndex++].PutValue(title);

            _DetailRowIndex++;
            columnIndex = 0;
            SetDetailColumnValue(cells, columnIndex++, "座號", _Detail_Style_Normal);
            SetDetailColumnValue(cells, columnIndex++, "姓名", _Detail_Style_Normal);
            SetDetailColumnValue(cells, columnIndex++, "學號", _Detail_Style_Normal);

            // 課程名稱, 國文Ⅳ	英文Ⅳ	...
            int CourseIndex = 0;
            _Current_Course_Count = 0;
            foreach(DAO.CourseVO CourseObj in ClassObj.AllCourseListDic.Values)
            {
                SetDetailColumnValue(cells, columnIndex++, CourseObj.CourseName, _Detail_Style_Normal);
                CourseIndex++;

                _Current_Course_Count++;
            }

            // 假如印出的課程, 超過預設課程數量, 則下一個columnIndex就用超出的課程數量
            // 假如沒有超過預設的課程數量, 就用預設的課程數量當作columnIndex
            // always +3, 因為前面有"座號" "姓名" "學號"
            columnIndex = ( (_Current_Course_Count < _Default_Detail_Course_Count)? _Default_Detail_Course_Count : _Current_Course_Count ) + 3;

            int StartColumn = columnIndex;
            SetDetailColumnValue(cells, columnIndex++, "及格標準", _Detail_Style_Normal);
            SetDetailColumnValue(cells, columnIndex++, "應得學分數", _Detail_Style_Normal);
            SetDetailColumnValue(cells, columnIndex++, "實得學分數", _Detail_Style_Normal);
            SetDetailColumnValue(cells, columnIndex++, "及格百分比", _Detail_Style_Normal);

            // 學分數 
            _DetailRowIndex++;
            columnIndex = 0;

            SetDetailColumnValue(cells, columnIndex, "學分數", _Detail_Style_Normal);
            
            // 合併儲存格
            cells.Merge(_DetailRowIndex, columnIndex, 1, 3);

            columnIndex = 3;
            foreach (DAO.CourseVO CourseObj in ClassObj.AllCourseListDic.Values)
            {
                SetDetailColumnValue(cells, columnIndex++, CourseObj.Credit, _Detail_Style_Normal);
            }

            // 及格標準, 應得學分數, 實得學分數, 及格百分比
            SetDetailColumnValue(cells, StartColumn++, "", _Detail_Style_Normal);
            SetDetailColumnValue(cells, StartColumn++, "", _Detail_Style_Normal);
            SetDetailColumnValue(cells, StartColumn++, "", _Detail_Style_Normal);
            SetDetailColumnValue(cells, StartColumn++, "", _Detail_Style_Normal);

        }

        private void OutDetailData(Cells cells, DAO.ClassVO ClassObj, DAO.StudentVO StudentObj)
        {
            
            if (_DetailRowIndex > _MAX_ROW_COUNT) return;

            int columnIndex = 0;
            _DetailRowIndex++;

            // 設定高度
            cells.SetRowHeightPixel(_DetailRowIndex, _Detail_Row_Height[3]);

            // 座號	姓名	學號
            SetDetailColumnValue(cells, columnIndex++, StudentObj.SeatNo, _Detail_Style_Normal);
            SetDetailColumnValue(cells, columnIndex++, StudentObj.StudentName, _Detail_Style_Normal);
            SetDetailColumnValue(cells, columnIndex++, StudentObj.StudentNumber, _Detail_Style_Normal);

            #region 輸出每一科的分數
            // 國文Ⅳ	英文Ⅳ	... 
            foreach (DAO.CourseVO CourseObj in ClassObj.AllCourseListDic.Values)
            {
                DAO.CourseVO StuCourseObj = Utility.FindCourse(StudentObj.CourseListDic.Values.ToList(), CourseObj.CourseId);
                if (StuCourseObj == null)
                {
                    // 沒有修這堂課
                    SetDetailColumnValue(cells, columnIndex++, "", _Detail_Style_Normal);
                }
                else
                {
                    if (StuCourseObj.CourseScore == -1)
                    {
                        // 有這堂課, 卻沒有分數
                        SetDetailColumnValue(cells, columnIndex++, _NoScroe, _Detail_Style_Red);
                    }
                    else
                    {
                        if (StuCourseObj.IsPass == true)
                            SetDetailColumnValue(cells, columnIndex++, StuCourseObj.CourseScore, _Detail_Style_Normal);
                        else
                            SetDetailColumnValue(cells, columnIndex++, StuCourseObj.CourseScore, _Detail_Style_Red);
                    }
                }
            }
            #endregion

            // 假如印出的課程小於預設的課程數量, 就把剩下的欄位補上Style(邊框)
            for (int intI = 0; intI < (_Default_Detail_Course_Count - _Current_Course_Count); intI++)
            {
                SetDetailColumnValue(cells, columnIndex++, "", _Detail_Style_Normal);
            }

            // 及格標準	應得學分數	實得學分數	未達百分比
            SetDetailColumnValue(cells, columnIndex++, StudentObj.PassScore, _Detail_Style_Normal);
            SetDetailColumnValue(cells, columnIndex++, StudentObj.TotalCredit, _Detail_Style_Normal);
            SetDetailColumnValue(cells, columnIndex++, StudentObj.GotCredit, _Detail_Style_Normal);
            SetDetailColumnValue(cells, columnIndex++, StudentObj.CreditRate + "%", _Detail_Style_Normal);

        }

        private void OutDetailSplitPage(Worksheet sheet)
        {
            _DetailRowIndex += 2;

            sheet.VPageBreaks.Add(_DetailRowIndex, _Default_Detail_Column);
            sheet.HPageBreaks.Add(_DetailRowIndex, _Default_Detail_Column);
        }
        #endregion

        #region 針對"總表"的處理 
        private void OutListTitle(Cells cells)
        {
            int columnIndex = 0;
            string title = K12.Data.School.ChineseName + " " + _ExamObj.ExameName + " 未達" + _PassRate + "%學分名單";
            cells[_ListRowIndex, columnIndex++].PutValue(title);

            _ListRowIndex++;
        }

        private void OutListData(Cells cells, DAO.ClassVO ClassObj, DAO.StudentVO StudentObj)
        {
            int columnIndex = 0;
            _ListRowIndex++;

            if (_ListRowIndex > _MAX_ROW_COUNT)
                return;

            // 先產生空白列
            Range range = cells.CreateRange(_ListRowIndex, columnIndex, 1, _MAX_List_Column);
            range.CopyStyle(_List_Row_Range);

            // 設定高度
            cells.SetRowHeightPixel(_ListRowIndex, _List_Row_Height);

            // 班級	座號	姓名	學號	應得學分數	實得學分數	未達百分比
            cells[_ListRowIndex, columnIndex++].PutValue(ClassObj.ClassName);
            cells[_ListRowIndex, columnIndex++].PutValue(StudentObj.SeatNo);
            cells[_ListRowIndex, columnIndex++].PutValue(StudentObj.StudentName);
            cells[_ListRowIndex, columnIndex++].PutValue(StudentObj.StudentNumber);
            cells[_ListRowIndex, columnIndex++].PutValue(StudentObj.TotalCredit);
            cells[_ListRowIndex, columnIndex++].PutValue(StudentObj.GotCredit);
            cells[_ListRowIndex, columnIndex++].PutValue(StudentObj.CreditRate + "%");
        }
        #endregion

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

            for (int intI = 1; intI < this.cboExamList.Items.Count; intI++)
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
            string msg = "";

            if (this.cboSchoolYear.SelectedIndex < 0)
            {
                msg = "請選擇學年度!";
                IsPass = false;
            }

            if (this.cboSemester.SelectedIndex < 0)
            {
                msg = "請選擇學期!";
                IsPass = false;
            }

            if (this.cboExamList.SelectedIndex <= 0)
            {
                msg = "請選擇試別!";
                IsPass = false;
            }

            if (IsPass == false)
            {
                MsgBox.Show(msg);
                FormComponetEnable(true);
            }

            return IsPass;
        }
        #endregion

        #region 畫面內容的取得跟儲存
        private void GetConfigure()
        {
            // 取得設定
            List<DAO.ConfigureRecord> ConfList = DAO.Configure.SelectConfigure();
            if (ConfList.Count == 0)
            {
                DAO.Configure.InsertByRecord(_Configure);
                ConfList = DAO.Configure.SelectConfigure();
            }
            _Configure = ConfList[0];
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
            seatno2 += obj2.ClassDisplyOrder.PadLeft(3, '0');       // 班級序號
            seatno2 += obj2.ClassName.PadLeft(20, '0');             // 班級名稱

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
