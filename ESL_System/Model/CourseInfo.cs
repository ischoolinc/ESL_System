using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System.Model
{
    /// <summary>
    /// 課程資訊 為了依照領域排序用
    /// </summary>
    class CourseInfo
    {
        public CourseInfo(string subject, string domain, int domainOrder, int subjectOrder)
        {
            this.DomainOrder = 99999; // 初始化 若沒有在領域排序對照表之領域 排在最後面
            this.SubjectOrder = 99999;
            this.Subject = subject;
            this.Domain = domain;
            this.DomainOrder = domainOrder;
            this.SubjectOrder = subjectOrder;
        }
        public string Subject { get; set; }
        public string Domain { get; set; }
        public int DomainOrder { get; set; }
        public int SubjectOrder { get; set; }
    }
}
