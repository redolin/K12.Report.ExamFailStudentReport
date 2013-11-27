using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Report.ExamFailStudentReport.DAO
{
    public class ExamVO
    {
        /// <summary>
        /// 系統編號
        /// </summary>
        public string ExamId { get; set; }

        /// <summary>
        /// 試別名稱
        /// </summary>
        public string ExameName { get; set; }

        /// <summary>
        /// 學年度
        /// </summary>
        public string SchoolYear { get; set; }

        /// <summary>
        /// 學期
        /// </summary>
        public string Semester { get; set; }

        public ExamVO(string examId, string examName, string schoolYear, string semester)
        {
            this.ExamId = examId;
            this.ExameName = examName;
            this.SchoolYear = schoolYear;
            this.Semester = semester;
        }
    }
}
