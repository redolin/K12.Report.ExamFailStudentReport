using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Report.ExamFailStudentReport.DAO
{
    [FISCA.UDT.TableName("K12.Report.ExamFailStudentReport.Configure")]
    public class ConfigureRecord : FISCA.UDT.ActiveRecord
    {
        /// <summary>
        /// 學年度
        /// </summary>
        [FISCA.UDT.Field]
        public string SchoolYear { get; set; }
        
        /// <summary>
        /// 學期
        /// </summary>
        [FISCA.UDT.Field]
        public string Semester { get; set; }

        /// <summary>
        /// 試別ID
        /// </summary>
        [FISCA.UDT.Field]
        public string ExamId { get; set; }

        /// <summary>
        /// 及格比例
        /// </summary>
        [FISCA.UDT.Field]
        public string PassCreditRate { get; set; }
    }
}
