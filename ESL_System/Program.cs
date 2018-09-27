using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA;
using FISCA.Presentation;
using K12.Presentation;
using FISCA.Permission;
using JHSchool;


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

            Catalog ribbon2 = RoleAclSource.Instance["課程"]["ESL課程"];
            ribbon2.Add(new RibbonFeature("ESL評量分數計算", "評量成績結算"));

            MotherForm.RibbonBarItems["課程", "ESL課程"]["評量成績結算"].Enable = false;

            K12.Presentation.NLDPanels.Course.SelectedSourceChanged += (sender, e) =>
            {
                if (K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0)
                {
                    MotherForm.RibbonBarItems["課程", "ESL課程"]["評量成績結算"].Enable = UserAcl.Current["ESL評量分數計算"].Executable;
                }
                else
                {
                    MotherForm.RibbonBarItems["課程", "ESL課程"]["評量成績結算"].Enable = false;
                }
            };

            MotherForm.RibbonBarItems["課程", "ESL課程"]["評量成績結算"].Image = Properties.Resources.calc_64;
            MotherForm.RibbonBarItems["課程", "ESL課程"]["評量成績結算"].Size = RibbonBarButton.MenuButtonSize.Medium;

            MotherForm.RibbonBarItems["課程", "ESL課程"]["評量成績結算"].Click += delegate
            {
                Form.CheckCalculateTermForm form = new Form.CheckCalculateTermForm(K12.Presentation.NLDPanels.Course.SelectedSource);

                if (form.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
                {

                    CalculateTermScore cts = new CalculateTermScore(K12.Presentation.NLDPanels.Course.SelectedSource, form.target_exam_id);

                    cts.CalculateESLTermScore(); // 計算ESL 評量 成績
                }

            };

            Catalog ribbon3 = RoleAclSource.Instance["課程"]["ESL報表"];
            ribbon3.Add(new RibbonFeature("ESL期末成績單", "ESL期末成績單"));

            MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["ESL報表"]["ESL期末成績單"].Enable = UserAcl.Current["ESL期末成績單"].Executable && K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0;

            K12.Presentation.NLDPanels.Course.SelectedSourceChanged += delegate
            {
                MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["ESL報表"]["ESL期末成績單"].Enable = UserAcl.Current["ESL期末成績單"].Executable && (K12.Presentation.NLDPanels.Course.SelectedSource.Count > 0);
            };


            MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["ESL報表"]["ESL期末成績單"].Click += delegate
            {
                List<K12.Data.CourseRecord> esl_couse_list = K12.Data.Course.SelectByIDs(K12.Presentation.NLDPanels.Course.SelectedSource);

                ESL_KcbsFinalReportFormNEW form = new ESL_KcbsFinalReportFormNEW(esl_couse_list);

            };

            
            //Catalog ribbon4 = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            //ribbon4.Add(new RibbonFeature("ESL期中分班", "ESL期中分班"));

            //RibbonBarButton group = Course.Instance.RibbonBarItems["教務"]["ESL期中分班"];
            //group.Size = RibbonBarButton.MenuButtonSize.Medium;
            //group.Image = Properties.Resources.meeting_refresh_64;
            //group.Enable = UserAcl.Current["ESL期中分班"].Executable;
            //group.Click += delegate
            //{
            //    if (Course.Instance.SelectedList.Count > 0)
            //        new SwapAttendStudents(Course.Instance.SelectedList.Count).ShowDialog();
            //};


        }
    }
}
