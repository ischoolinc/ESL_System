
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.UDT;
using FISCA.Presentation.Controls;
using System.IO;

namespace ESL_System.CourseExtendControls
{
    //[Framework.AccessControl.FeatureCode("Content0200")]    
    [FISCA.Permission.FeatureCode("JHSchool.Course.Detail0000", "基本資料")]
    internal partial class BasicInfoItem : FISCA.Presentation.DetailContent
    {
        public BasicInfoItem()
        {
            InitializeComponent();

            Group = "基本資料";

        }
    }
}