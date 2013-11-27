using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace K12.Report.ExamFailStudentReport.DAO
{
    public class CourseVO
    {
        /// <summary>
        /// 系統編號
        /// </summary>
        public string CourseId { get; set; }

        /// <summary>
        /// 學分數
        /// </summary>
        public decimal Credit { get; set; }

        /// <summary>
        /// 課程名稱
        /// </summary>
        public string CourseName { get; set; }

        /// <summary>
        /// 課程分數
        /// </summary>
        public decimal CourseScore { get; set; }

        /// <summary>
        /// 課程分數是否高過及格分數
        /// </summary>
        public bool IsPass = false;

        public CourseVO(DataRow row)
        {
            this.CourseId = ("" + row["course_id"]).Trim();
            this.CourseName = ("" + row["course_name"]).Trim();

            string tmp = ("" + row["credit"]).Trim();
            decimal iTmp = 0;
            decimal.TryParse(tmp, out iTmp);
            this.Credit = iTmp;

            tmp = ("" + row["score"]).Trim();
            if(string.IsNullOrEmpty(tmp))
            {
                this.CourseScore = -1;
            }
            else
            {
                iTmp = -1;
                decimal.TryParse(tmp, out iTmp);
                this.CourseScore = iTmp;
            }
        }
    }
}
