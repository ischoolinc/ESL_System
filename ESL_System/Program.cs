using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA;
using FISCA.Presentation;
using K12.Presentation;
using FISCA.Permission;


namespace ESL_System
{
    public class Program
    {
        //2018/4/11 穎驊因應康橋英文系統、弘文ESL 專案 ，開始建構教務作業ESL 評分樣版設定
        [FISCA.MainMethod()]
        public static void Main()
        {
            FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();

            _AccessHelper.Select<UDT_ReportTemplate>(); // 先將UDT 選起來，如果是第一次開啟沒有話就會新增

            Catalog ribbon = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            ribbon.Add(new RibbonFeature("ESL評分樣版設定", "ESL評分樣版設定"));

            MotherForm.RibbonBarItems["教務作業", "基本設定"]["設定"]["ESL評分樣版設定"].Enable = UserAcl.Current["ESL評分樣版設定"].Executable;
            

            MotherForm.RibbonBarItems["教務作業", "基本設定"]["設定"]["ESL評分樣版設定"].Click += delegate
            {
                Form.ESL_TemplateSetupManager form = new Form.ESL_TemplateSetupManager();

                form.ShowDialog();

            };


        }
    }
}
