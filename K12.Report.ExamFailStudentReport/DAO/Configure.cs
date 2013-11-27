using FISCA.DSAUtil;
using FISCA.UDT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Report.ExamFailStudentReport.DAO
{
    /// <summary>
    /// 處理 UDT 資料
    /// </summary>
    public class Configure
    {
        /// <summary>
        /// 建立使用到的 UDT Table：主要檢查資料庫有沒有建立UDT，沒有建自動建立。
        /// </summary>
        public static void CreateConfigureUDTTable()
        {
            FISCA.UDT.SchemaManager Manager = new SchemaManager(new DSConnection(FISCA.Authentication.DSAServices.DefaultDataSource));

            // 設定
            Manager.SyncSchema(new ConfigureRecord());
        }

        /// <summary>
        /// 新增設定
        /// </summary>
        /// <param name="rec"></param>
        public static void InsertByRecord(ConfigureRecord rec)
        {
            if (rec != null)
            {
                List<ConfigureRecord> insertList = new List<ConfigureRecord>();
                insertList.Add(rec);
                AccessHelper accessHelper = new AccessHelper();
                accessHelper.InsertValues(insertList);
            }
        }

        /// <summary>
        /// 更新設定
        /// </summary>
        /// <param name="rec"></param>
        public static void UpdateByRecord(ConfigureRecord rec)
        {
            if (rec != null)
            {
                List<ConfigureRecord> updateList = new List<ConfigureRecord>();
                updateList.Add(rec);
                AccessHelper accessHelper = new AccessHelper();
                accessHelper.UpdateValues(updateList);
            }
        }

        /// <summary>
        /// 取得設定
        /// </summary>
        /// <returns></returns>
        public static List<ConfigureRecord> SelectConfigure()
        {
            List<ConfigureRecord> dataList = new List<ConfigureRecord>();

            AccessHelper accessHelper = new AccessHelper();
            // 當有 Where 條件寫法
            dataList = accessHelper.Select<ConfigureRecord>();

            if (dataList == null)
                return new List<ConfigureRecord>();

            return dataList;
        }
    }
}
