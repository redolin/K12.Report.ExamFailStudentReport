using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Report.ExamFailStudentReport.DAO
{
    public class StudentVO
    {
        /// <summary>
        /// 系統編號
        /// </summary>
        public string StudentId { get; set; }

        /// <summary>
        /// 班級名稱
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        public string StudentName { get; set; }

        /// <summary>
        /// 學號
        /// </summary>
        public string StudentNumber { get; set; }

        /// <summary>
        /// 座號
        /// </summary>
        public string SeatNo { get; set; }

        /// <summary>
        /// 學生是否有兩個以上的及格標準
        /// </summary>
        public bool IsMultiTag = false;

        /// <summary>
        /// false: 學生的學分比例低於及格比例
        /// true: 學生的學分比例高於及格比例
        /// </summary>
        public bool IsPass = false;

        /// <summary>
        /// 學生類別清單
        /// </summary>
        public List<string> StudentTag= new List<string>();

        /// <summary>
        /// 課程列表
        /// </summary>
        public Dictionary<string, CourseVO> CourseListDic = new Dictionary<string, CourseVO>();

        private string studentCalRule;
        /// <summary>
        /// 學生及格標準(XML), 假如學生身上沒有指定的及格標準, 就會改用班級的及格標準
        /// </summary>
        public string StudentCalRule {
            get
            {
                return this.studentCalRule;
            }
             
            set
            {
                if(!string.IsNullOrEmpty(value))
                {
                    this.studentCalRule = value;
                    this.PassCriterionDic = Utility.ConvertCalRuleXML(value);
                }
            }
        }

        /// <summary>
        /// 學生的及格標準, 假如學生身上沒有指定的及格標準, 就會改用班級的及格標準
        /// </summary>
        public List<PassCriterionVO> PassCriterionDic = new List<PassCriterionVO>();

        /// <summary>
        /// 及格分數
        /// </summary>
        public decimal PassScore = -1;

        /// <summary>
        /// 應得學分
        /// </summary>
        public decimal TotalCredit { get; set; }

        /// <summary>
        /// 實得學分
        /// </summary>
        public decimal GotCredit { get; set; }

        /// <summary>
        /// 學分比例
        /// </summary>
        public decimal CreditRate { get; set; }
    }
}
