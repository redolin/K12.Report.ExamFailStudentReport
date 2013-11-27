using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace K12.Report.ExamFailStudentReport.Forms
{
    public partial class FrmWarningStudent : BaseForm
    {
        Dictionary<string, DAO.StudentVO> _MultiTagStudentDic;

        public FrmWarningStudent(Dictionary<string, DAO.StudentVO> multiTagStudentDic)
        {
            InitializeComponent();

            this._MultiTagStudentDic = multiTagStudentDic;
        }

        private void btnWait_Click(object sender, EventArgs e)
        {
            // 看學生有無重複加入待處理名單
            List<string> StudentIdList = new List<string>();
            foreach(string StudentId in _MultiTagStudentDic.Keys)
            {
                // 透過學生ID判斷是否已存在學生待處理
                if (!K12.Presentation.NLDPanels.Student.TempSource.Contains(StudentId))
                {
                    StudentIdList.Add(StudentId);
                }
            }

            // 把學生加入待處理名單
            K12.Presentation.NLDPanels.Student.AddToTemp(StudentIdList);

            // 設定選擇結果
            this.DialogResult = DialogResult.Ignore;

            this.Close();
        }

        private void btnContinue_Click(object sender, EventArgs e)
        {
            // 繼續產生, 不做任何事情

            // 設定選擇結果
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            // 離開

            // 設定選擇結果
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void FrmWarningStudent_Load(object sender, EventArgs e)
        {
            this.lvWarningStudent.Items.Clear();
            ListViewItem item = null;
            foreach (DAO.StudentVO StudentObj in _MultiTagStudentDic.Values)
            {
                item = ConvertToListViewItem(StudentObj);

                lvWarningStudent.Items.Add(item);
            }
        }

        private ListViewItem ConvertToListViewItem(DAO.StudentVO rec)
        {
            ListViewItem item = new ListViewItem();

            item.Tag = rec;
            // 學號
            item.Text = rec.StudentNumber;
            // 班級
            item.SubItems.Add(rec.ClassName);
            // 姓名
            item.SubItems.Add(rec.StudentName);
            
            return item;
        }
    }
}
