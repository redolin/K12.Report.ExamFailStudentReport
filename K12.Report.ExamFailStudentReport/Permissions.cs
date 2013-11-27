using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Report.ExamFailStudentReport
{
    class Permissions
    {
        public const string KeyExamFailReport = "K12.Report.ExamFailStudentReport.cs";
        
        public static bool IsEnableExamFailReport
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[KeyExamFailReport].Executable;
            }
        }

        
    }
}
