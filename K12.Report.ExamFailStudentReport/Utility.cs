using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace K12.Report.ExamFailStudentReport
{
    class Utility
    {
        public static readonly string _ContactChar = ":";
        public static readonly string _XML_PassScroe = "及格標準";
        public static readonly string _XML_PassScore_StudentTag = "學生類別";
        public static readonly string[] _XML_PassScore_Attribute = { "類別", "一年級及格標準", "二年級及格標準", "三年級及格標準", "四年級及格標準" };

        /// <summary>
        /// 取得學生類別的完整名稱
        /// </summary>
        /// <param name="prefix"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static string GetTagName(string prefix, string name)
        {
            string result = "";

            if(string.IsNullOrEmpty(prefix))
            {
                result = name;
            }
            else
            {
                result = prefix.Trim() + _ContactChar + name.Trim();
            }

            return result;
        }

        /// <summary>
        /// 把及格標準(string)轉成物件
        /// </summary>
        /// <param name="xml"></param>
        /// <returns></returns>
        public static List<DAO.PassCriterionVO> ConvertCalRuleXML(string xml)
        {
            List<DAO.PassCriterionVO> result = new List<DAO.PassCriterionVO>();

            XElement root = XElement.Parse(xml);

            XElement ScoreList = root.Element(_XML_PassScroe);

            result = (
                        from sss in ScoreList.Elements(_XML_PassScore_StudentTag)
                        select new DAO.PassCriterionVO(
                                                sss.Attribute(_XML_PassScore_Attribute[0]).Value,
                                                sss.Attribute(_XML_PassScore_Attribute[1]).Value,
                                                sss.Attribute(_XML_PassScore_Attribute[2]).Value,
                                                sss.Attribute(_XML_PassScore_Attribute[3]).Value,
                                                sss.Attribute(_XML_PassScore_Attribute[4]).Value)
                     ).ToList();
            return result;
        }

        /// <summary>
        /// 取得及格分數, 假如根據類別跟年級找不到及格分數, 就會回傳-1
        /// </summary>
        /// <param name="PassCriterionList"></param>
        /// <param name="TagName"></param>
        /// <param name="GradeYear"></param>
        /// <returns></returns>
        public static decimal GetPassScroe(List<DAO.PassCriterionVO> PassCriterionList, string TagName, string GradeYear)
        {
            decimal result = -1;

            foreach(DAO.PassCriterionVO obj in PassCriterionList)
            {
                if(obj.TagName == TagName)
                {
                    if (obj.PassScoreListDic.Keys.Contains(GradeYear))
                    {
                        result = obj.PassScoreListDic[GradeYear];
                    }
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// 判斷及格分數那個比較小, 用於學生有多個及格標準時
        /// </summary>
        /// <param name="score1"></param>
        /// <param name="score2"></param>
        /// <returns></returns>
        public static decimal MinPassScroe(decimal score1, decimal score2)
        {
            if(score1 > score2)
                return score2;
            else
                return score1;
        }

        /// <summary>
        /// 在list中找到指定的課程
        /// </summary>
        /// <param name="CourseList"></param>
        /// <param name="CourseId"></param>
        /// <returns></returns>
        public static DAO.CourseVO FindCourse(List<DAO.CourseVO> CourseList, string CourseId)
        {
            DAO.CourseVO result = null;
            foreach(DAO.CourseVO CourseObj in CourseList)
            {
                if(CourseObj.CourseId == CourseId)
                {
                    result = CourseObj;
                    break;
                }
            }

            return result;
        }
    }
}
