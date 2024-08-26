using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System
{
    static class NameCheck
    {
        /// <summary>
        /// 報表名稱
        /// </summary>
        public static string ReportName = "Routine Report";

        /// <summary>
        /// 檢查學校名稱是否符合json內容
        /// 並且修改報表名稱
        /// </summary>
        public static bool Check()
        {
            string jsonString = Encoding.UTF8.GetString(Properties.Resources.School);
            Dictionary<string, string> SchoolNameDic = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonString);

            string DSNS = FISCA.Authentication.DSAServices.AccessPoint;
            if (SchoolNameDic.ContainsValue(DSNS))
            {
                ReportName = "Weekly Report";
                return true;
            }
            else
            {
                ReportName = "Routine Report";
                return false;
            }
        }



    }
}
