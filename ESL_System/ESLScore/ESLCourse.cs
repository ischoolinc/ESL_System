using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System
{
    class ESLCourse
    {

        /// <summary>
        /// 課程ID
        /// </summary>
        public string ID { get; set; }

        /// <summary>
        /// 課程名稱
        /// </summary>
        public string CourseName { get; set; }
       
        /// <summary>
        /// 子評量項目設定
        /// </summary>
        public string Description { get; set; }
    }
}
