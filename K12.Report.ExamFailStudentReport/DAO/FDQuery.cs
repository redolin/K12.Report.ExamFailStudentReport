using FISCA.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace K12.Report.ExamFailStudentReport.DAO
{
    /// <summary>
    /// 使用 FISCA.Data Query
    /// </summary>
    class FDQuery
    {
        public const string _StudentSatus = "1";
        /// <summary>
        /// 取得不重複的試別清單
        /// </summary>
        /// <param name="ClassIdList"></param>
        /// <returns></returns>
        public static List<ExamVO> GetDistincExamList(List<string> ClassIdList)
        {
            List<ExamVO> result = new List<ExamVO>();
            StringBuilder sb = new StringBuilder();
            sb.Append("select te_include.ref_exam_id, exam.exam_name, course.school_year, course.semester, exam.display_order");
            sb.Append(" from student");
            sb.Append(" left join class");
            sb.Append(" on student.ref_class_id = class.id");
            sb.Append(" left join sc_attend");
            sb.Append(" on sc_attend.ref_student_id = student.id");
            sb.Append(" left join course");
            sb.Append(" on course.id = sc_attend.ref_course_id");
            sb.Append(" left join te_include");
            sb.Append(" on te_include.ref_exam_template_id = course.ref_exam_template_id");
            sb.Append(" left join exam");
            sb.Append(" on te_include.ref_exam_id = exam.id");
            sb.Append(" where class.id in ('"+string.Join("','", ClassIdList.ToArray())+"')");
            sb.Append(" and course.ref_exam_template_id is not NULL");
            sb.Append(" and te_include.ref_exam_id is not NULL");
            sb.Append(" and student.status = '" + _StudentSatus + "'");
            sb.Append(" group by ref_exam_id, exam.exam_name, course.school_year, course.semester, exam.display_order");
            sb.Append(" order by course.school_year, course.semester, exam.display_order");
            Console.WriteLine("[GetDistincExamList] sql: ["+sb.ToString()+"]");

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(sb.ToString());

            foreach (DataRow row in dt.Rows)
            {
                ExamVO ExamObj = new ExamVO("" + row["ref_exam_id"], "" + row["exam_name"], "" + row["school_year"], "" + row["semester"]);

                result.Add(ExamObj);
            }

            return result;
        }   // end of GetDistincExamList

        /// <summary>
        /// 取得學生所有修課的成績, 包括學生自己的特殊計算規則假如有的話
        /// </summary>
        /// <param name="ClassIdList"></param>
        /// <param name="exam"></param>
        /// <returns></returns>
        public static Dictionary<string, ClassVO> GetAllStudentScore(List<string> ClassIdList, ExamVO exam)
        {
            Dictionary<string, ClassVO> result = new Dictionary<string,ClassVO>();
            StringBuilder sb = new StringBuilder();
            sb.Append("select student.id as student_id,");
            sb.Append("student.ref_class_id,");
            sb.Append("class.class_name,");
            sb.Append("class.grade_year,");
            sb.Append("class.display_order,");
            sb.Append("student.seat_no,");
            sb.Append("student.name as student_name,");
            sb.Append("student.student_number,");
            sb.Append("course.id as course_id,");
            sb.Append("course.credit,");
            sb.Append("course.course_name,");
            sb.Append("sce_take.score,");
            sb.Append("score_calc_rule.content as student_calc_rule");
            sb.Append(" from student");
            sb.Append(" left join class");
            sb.Append(" on student.ref_class_id = class.id");
            sb.Append(" left join sc_attend");
            sb.Append(" on sc_attend.ref_student_id = student.id");
            sb.Append(" left join course");
            sb.Append(" on course.id = sc_attend.ref_course_id");
            sb.Append(" left join te_include");
            sb.Append(" on te_include.ref_exam_template_id = course.ref_exam_template_id");
            sb.Append(" left join sce_take");
            sb.Append(" on (sce_take.ref_sc_attend_id=sc_attend.id and sce_take.ref_exam_id = te_include.ref_exam_id)");
            sb.Append(" left join score_calc_rule");
            sb.Append(" on score_calc_rule.id = student.ref_score_calc_rule_id");
            sb.Append(" where class.id in ('" + string.Join("','", ClassIdList.ToArray()) + "')");
            sb.Append(" and student.status = '" + _StudentSatus + "'");
            sb.Append(" and course.school_year = '" + exam.SchoolYear + "'");
            sb.Append(" and course.semester = '" + exam.Semester + "'");
            sb.Append(" and te_include.ref_exam_id = '" + exam.ExamId + "'");
            Console.WriteLine("[GetAllStudentScore] sql: [" + sb.ToString() + "]");

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(sb.ToString());

            foreach (DataRow row in dt.Rows)
            {
                string ClassId = ("" + row["ref_class_id"]).Trim();
                if (string.IsNullOrEmpty(ClassId)) continue;

                ClassVO ClassObj;
                if(result.Keys.Contains(ClassId))
                {
                    ClassObj = result[ClassId];
                }
                else
                {
                    ClassObj = new ClassVO(row);
                    result.Add(ClassId, ClassObj);
                }

                ClassObj.AddStudent(row);
                
            }

            return result;
        }   // end of GetAllStudentScore

        /// <summary>
        /// 取得所有班級的計算規則
        /// </summary>
        /// <param name="ClassList"></param>
        /// <param name="ClassIdList"></param>
        public static void GetAllClassCalRule(Dictionary<string, ClassVO> ClassList, List<string> ClassIdList)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("select class.id, score_calc_rule.content");
            sb.Append(" from class");
            sb.Append(" left join score_calc_rule");
            sb.Append(" on class.ref_score_calc_rule_id=score_calc_rule.id");
            sb.Append(" where class.id in ('" + string.Join("','", ClassIdList.ToArray()) + "')");
            Console.WriteLine("[GetClassCalRule] sql: [" + sb.ToString() + "]");

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(sb.ToString());

            foreach (DataRow row in dt.Rows)
            {
                string ClassId = ("" + row["id"]).Trim();
                if(string.IsNullOrEmpty(ClassId)) continue;

                if(ClassList.Keys.Contains(ClassId))
                {
                    ClassVO ClassObj = ClassList[ClassId];
                    string ClassCalRule = ("" + row["content"]).Trim();
                    foreach(StudentVO Student in ClassObj.StudentListDic.Values)
                    {
                        // 假如學生身上沒有及格標準, 就改成班級的及格標準
                        if(string.IsNullOrEmpty(Student.StudentCalRule))
                        {
                            Student.StudentCalRule = ClassCalRule;
                        }
                    }
                }
            }
        }   // end of GetAllClassCalRule

        /// <summary>
        /// 取得所有學生的類別
        /// </summary>
        /// <param name="ClassList"></param>
        /// <param name="ClassIdList"></param>
        public static void GetAllStudentTag(Dictionary<string, ClassVO> ClassList, List<string> ClassIdList)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("select student.ref_class_id as class_id,");
            sb.Append("student.id as student_id,");
            sb.Append("tag.prefix,");
            sb.Append("tag.name");
            sb.Append(" from student");
            sb.Append(" left join class");
            sb.Append(" on student.ref_class_id = class.id");
            sb.Append(" left join tag_student");
            sb.Append(" on student.id = tag_student.ref_student_id");
            sb.Append(" left join tag");
            sb.Append(" on tag_student.ref_tag_id = tag.id");
            sb.Append(" where class.id in ('" + string.Join("','", ClassIdList.ToArray()) + "')");
            sb.Append(" and student.status = '" + _StudentSatus + "'");
            Console.WriteLine("[GetAllStudentTag] sql: [" + sb.ToString() + "]");

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(sb.ToString());

            foreach (DataRow row in dt.Rows)
            {
                string ClassId = ("" + row["class_id"]).Trim();
                string StudentId = ("" + row["student_id"]).Trim();

                if(string.IsNullOrEmpty(ClassId)) continue;
                if(string.IsNullOrEmpty(StudentId)) continue;

                ClassVO ClassObj;
                StudentVO StudentObj;
                // 找不到班級就不處理
                if(ClassList.Keys.Contains(ClassId))
                {
                    ClassObj = ClassList[ClassId];

                    // 找不到學生就不處理
                    if(ClassObj.StudentListDic.Keys.Contains(StudentId))
                    {
                        StudentObj = ClassObj.StudentListDic[StudentId];
                        string prefix = ("" + row["prefix"]).Trim();
                        string name = ("" + row["name"]).Trim();

                        StudentObj.StudentTag.Add(Utility.GetTagName(prefix, name));
                    }
                }
            }
        }   // end of GetAllStudentTag
    }
}
