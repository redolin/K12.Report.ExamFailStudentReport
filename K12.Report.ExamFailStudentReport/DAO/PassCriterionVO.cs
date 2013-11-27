using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Report.ExamFailStudentReport.DAO
{
    public class PassCriterionVO
    {
        /// <summary>
        /// 學生類別名稱
        /// </summary>
        public string TagName { get; set; }

        /// <summary>
        /// 及格標準清單, key: GradeYear, value: PassScore
        /// </summary>
        public Dictionary<string, decimal> PassScoreListDic = new Dictionary<string, decimal>();

        public PassCriterionVO(string tagName, string gy1Score, string gy2Score, string gy3Score, string gy4Score)
        {
            this.TagName = tagName;

            // 一年級及格分數
            decimal iTemp = 0;
            decimal.TryParse(gy1Score, out iTemp);
            this.PassScoreListDic.Add("1", iTemp);

            // 二年級及格分數
            iTemp = 0;
            decimal.TryParse(gy2Score, out iTemp);
            this.PassScoreListDic.Add("2", iTemp);

            // 三年級及格分數
            iTemp = 0;
            decimal.TryParse(gy3Score, out iTemp);
            this.PassScoreListDic.Add("3", iTemp);

            // 四年級及格分數
            iTemp = 0;
            decimal.TryParse(gy4Score, out iTemp);
            this.PassScoreListDic.Add("4", iTemp);

        }
    }
}
