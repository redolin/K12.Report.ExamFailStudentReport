using FISCA.Permission;
using FISCA.Presentation;
using K12.Presentation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K12.Report.ExamFailStudentReport
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void main()
        {
            CheckUDTExist();

            RibbonBarItem rbRptItem1 = MotherForm.RibbonBarItems["班級", "資料統計"];
            rbRptItem1["報表"]["成績相關報表"]["評量成績未達標準名單"].Enable = Permissions.IsEnableExamFailReport;
            rbRptItem1["報表"]["成績相關報表"]["評量成績未達標準名單"].Click += delegate
            {
                if (NLDPanels.Class.SelectedSource.Count > 0)
                {
                    Forms.FrmExamFailCondition frm = new Forms.FrmExamFailCondition();
                    frm.ShowDialog();
                }
            };

            // 在權限畫面出現"評量成績未達標準名單"權限
            Catalog catalog1 = RoleAclSource.Instance["班級"]["報表"];
            catalog1.Add(new RibbonFeature(Permissions.KeyExamFailReport, "評量成績未達標準名單"));

        }


        private static void CheckUDTExist()
        {
            // 檢查UDT
            BackgroundWorker bkWork;

            bkWork = new BackgroundWorker();
            bkWork.DoWork += new DoWorkEventHandler(_bkWork_DoWork);
            bkWork.RunWorkerAsync();
        }

        static void _bkWork_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                // 檢查並建立UDT Table
                DAO.Configure.CreateConfigureUDTTable();
            }
            catch
            {
                
            }
        }
    }
}
