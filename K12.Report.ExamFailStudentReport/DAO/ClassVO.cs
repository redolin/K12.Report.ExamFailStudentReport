using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace K12.Report.ExamFailStudentReport.DAO
{
    public class ClassVO
    {
        /// <summary>
        /// 系統編號
        /// </summary>
        public string ClassId { get; set; }

        /// <summary>
        /// 班級名稱
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// 年級
        /// </summary>
        public string ClassGradeYear { get; set; }

        /// <summary>
        /// 班級內所有學生的課程集合
        /// </summary>
        public Dictionary<string, CourseVO> AllCourseListDic = new Dictionary<string, CourseVO>();

        /// <summary>
        /// 學生清單
        /// </summary>
        public Dictionary<string, StudentVO> StudentListDic = new Dictionary<string, StudentVO>();

        /// <summary>
        /// 學生清單 (sorted)
        /// </summary>
        public List<StudentVO> StudentList = new List<StudentVO>();

        /// <summary>
        /// 班級內學分比例低過標準的學生
        /// </summary>
        public List<StudentVO> EaxmFailStudentList = new List<StudentVO>();

        /// <summary>
        /// 班級排序
        /// </summary>
        public string ClassDisplyOrder { get; set; }

        public ClassVO(DataRow row)
        {
            this.ClassId = ("" + row["ref_class_id"]).Trim();
            this.ClassName = ("" + row["class_name"]).Trim();
            this.ClassGradeYear = ("" + row["grade_year"]).Trim();
            this.ClassDisplyOrder = ("" + row["display_order"]).Trim();
        }

        /// <summary>
        /// 新增學生到班級內, 包括此學生的課程
        /// </summary>
        /// <param name="row"></param>
        public void AddStudent(DataRow row)
        {
            string StudentId = ("" + row["student_id"]).Trim();
            if(string.IsNullOrEmpty(StudentId)) return;

            StudentVO StudentObj;

            // 新增學生
            if(!StudentListDic.Keys.Contains(StudentId))
            {
                StudentObj = new StudentVO();
                StudentObj.StudentId = StudentId;
                StudentObj.ClassName = ("" + row["class_name"]).Trim();
                StudentObj.SeatNo = ("" + row["seat_no"]).Trim();
                StudentObj.StudentName = ("" + row["student_name"]).Trim();
                StudentObj.StudentNumber = ("" + row["student_number"]).Trim();
                StudentObj.StudentCalRule = ("" + row["student_calc_rule"]).Trim();
                StudentListDic.Add(StudentId, StudentObj);
            }
            else
            {
                StudentObj = StudentListDic[StudentId];
            }

            // 新增課程
            string CourseId = ("" + row["course_id"]).Trim();
            if(string.IsNullOrEmpty(CourseId)) return;
            
            CourseVO CourseObj = new CourseVO(row);
            if (!StudentObj.CourseListDic.Keys.Contains(CourseId))
            {
                StudentObj.CourseListDic.Add(CourseId, CourseObj);
            }

            if(!AllCourseListDic.Keys.Contains(CourseId))
            {
                AllCourseListDic.Add(CourseId, CourseObj);
                if (Global.IsDebug) Console.WriteLine(StudentObj.ClassName + " " + CourseObj.CourseName);
            }
        }
    }
}
