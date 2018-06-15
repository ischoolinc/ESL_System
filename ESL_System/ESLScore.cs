using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System
{
    class ESLScore
    {

        /// <summary>
        /// 成績ID
        /// </summary>
        public string ID { get; set; }

        /// <summary>
        /// 參考課程ID
        /// </summary>
        public string RefCourseID { get; set; }

        /// <summary>
        /// 參考學生ID
        /// </summary>
        public string RefStudentID { get; set; }

        /// <summary>
        /// 參考教師ID
        /// </summary>
        public string RefTeacherID { get; set; }

        /// <summary>
        /// Term(評量)
        /// </summary>
        public string Term { get; set; }

        /// <summary>
        /// Subject(科目)
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// Assessment(子項目)
        /// </summary>
        public string Assessment { get; set; }

        /// <summary>
        /// Custom_Assessment(教師自定義子項目)
        /// </summary>
        public string Custom_Assessment { get; set; }

        /// <summary>
        /// 成績值
        /// </summary>
        public string Value { get; set; }


        /// <summary>
        /// 成績分數 (經由系統按照設定比例計算後的分數)
        /// </summary>
        public decimal Score { get; set; }


    }
}
