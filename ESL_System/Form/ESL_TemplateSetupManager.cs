using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.Data;
using DevComponents.DotNetBar;
using System.Xml.Linq;


namespace ESL_System.Form
{
    public partial class ESL_TemplateSetupManager : FISCA.Presentation.Controls.BaseForm
    {
        private BackgroundWorker _workder;

        private Dictionary<string, string> hintGuideDict = new Dictionary<string, string>();

        private Dictionary<string, string> typeCovertDict = new Dictionary<string, string>();




        // 現在點在哪一小節
        DevComponents.AdvTree.Node node_now;

        public ESL_TemplateSetupManager()
        {
            InitializeComponent();
            HideNavigationBar();

            hintGuideDict.Add("string", "請輸入文字(ex : mid-term)，不得空白、重覆。");
            hintGuideDict.Add("integer", "請輸入整數數字(ex : 50)");
            hintGuideDict.Add("time", "請輸入日期(ex : 2018/04/21 00:00:00)");
            hintGuideDict.Add("teacherKind", "點選左鍵選取: 教師一、教師二、教師三");
            hintGuideDict.Add("ScoreKind", "點選左鍵選取:分數、指標、評語");
            hintGuideDict.Add("AllowCustom", "點選左鍵選取:是、否");

            typeCovertDict.Add("Score", "分數");
            typeCovertDict.Add("Indicator", "指標");
            typeCovertDict.Add("Comment", "評語");

        }


        /// <summary>
        /// 非同步處理，使用時要小心。
        /// </summary>
        private void LoadAssessmentSetups()
        {
            //_workder = new BackgroundWorker();
            //_workder.DoWork += new DoWorkEventHandler(Workder_DoWork);
            //_workder.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Workder_RunWorkerCompleted);

            BeforeLoadAssessmentSetup();
            //_workder.RunWorkerAsync();
        }

        private void BeforeLoadAssessmentSetup()
        {
            //將設計範例全部清光，開始抓取table 資料
    
            advTree1.Nodes.Clear();

            Loading = true;
            ipList.Items.Clear();
            //panel1.Enabled = false;

            //btnSave.Enabled = false;


            string query = "select * from exam_template where description !=''";
            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(query);

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {

                    ButtonItem item = new ButtonItem();
                    item.Text = "" + dr[1];
                    item.Tag = "" + dr[0];
                    item.OptionGroup = "AssessmentSetup";
                    //item.Click += new EventHandler(AssessmentSetup_Click);
                    //item.DoubleClick += new EventHandler(AssessmentSetup_DoubleClick);
                    ipList.Items.Add(item);
                }
            }



            // 取得ESL 描述 in description

            string selQuery = "select id,description from exam_template where name = 'ESL 科目樣版' ";
            dt = qh.Select(selQuery);
            string xmlStr = "<root>" + dt.Rows[0]["description"].ToString() + "</root>";
            XElement elmRoot = XElement.Parse(xmlStr);

            if (elmRoot != null)
            {
                if (elmRoot.Element("ESLTemplate") != null)
                {
                    foreach (XElement ele_term in elmRoot.Element("ESLTemplate").Elements("Term"))
                    {
                        Term t = new Term();

                        t.Name = ele_term.Attribute("Name").Value;
                        t.Percentage = ele_term.Attribute("Percentage").Value;
                        t.InputStartTime = ele_term.Attribute("InputStartTime").Value;
                        t.InputEndTime = ele_term.Attribute("InputEndTime").Value;

                        t.SubjectList = new List<Subject>();

                        foreach (XElement ele_subject in ele_term.Elements("Subject"))
                        {
                            Subject s = new Subject();

                            s.Name = ele_subject.Attribute("Name").Value;
                            s.Percentage = ele_subject.Attribute("Percentage").Value;

                            s.AssessmentList = new List<Assessment>();

                            foreach (XElement ele_assessment in ele_subject.Elements("Assessment"))
                            {
                                Assessment a = new Assessment();

                                a.Name = ele_assessment.Attribute("Name").Value;
                                a.Percentage = ele_assessment.Attribute("Percentage").Value;
                                a.TeacherRole = ele_assessment.Attribute("TeacherRole").Value;
                                a.Type = ele_assessment.Attribute("Type").Value;
                                a.AllowCustomAssessment = ele_assessment.Attribute("AllowCustomAssessment").Value;

                                a.IndicatorsList = new List<Indicators>();

                                if (ele_assessment.Element("Indicators") != null)
                                {

                                    foreach (XElement ele_Indicator in ele_assessment.Element("Indicators").Elements("Indicator"))
                                    {
                                        Indicators i = new Indicators();

                                        i.Name = ele_Indicator.Value;

                                        a.IndicatorsList.Add(i);
                                    }
                                }
                                
                                s.AssessmentList.Add(a);
                            }


                            t.SubjectList.Add(s);
                        }



                        ParseDBxmlToNodeUI(t);
                    }


                }
            }

        }

        private void HideNavigationBar()
        {
            npLeft.NavigationBar.Visible = false;
        }

        private void AfterLoadAssessmentSetup()
        {
            ipList.RecalcLayout();
            //panel1.Enabled = true;

            btnSave.Enabled = true;

            Loading = false;
        }

        private bool Loading
        {
            get { return loading.Visible; }
            set { loading.Visible = value; }
        }

        private void ESL_TemplateSetupManager_Load(object sender, EventArgs e)
        {
            BeforeLoadAssessmentSetup();
        }


        // 加入新試別
        private void node8_NodeDoubleClick(object sender, EventArgs e)
        {
            DevComponents.AdvTree.Node new_term_node = new DevComponents.AdvTree.Node();

            new_term_node.Text = "請輸入新試別名稱";
            //新試別點擊兩下，可以改變名稱
            new_term_node.NodeDoubleClick += new System.EventHandler(NodeRename);


            DevComponents.AdvTree.Node add_new_subject_node_btn = new DevComponents.AdvTree.Node();

            add_new_subject_node_btn.Text = "<b><font color=\"#ED1C24\">+加入新子項目</font></b>";

            add_new_subject_node_btn.NodeDoubleClick += new System.EventHandler(InsertNewSubject);


            new_term_node.Nodes.Add(add_new_subject_node_btn);

            //不用Add 改用Insert ，是因為可以指定位子，讓加入新試別功能可以永遠在最後一項。
            //advTree1.Nodes.Add(new_term_node);
            advTree1.Nodes.Insert(advTree1.Nodes.Count - 1, new_term_node);
        }



        private void NodeRename(object sender, EventArgs e)
        {
            DevComponents.AdvTree.Node term_node = (DevComponents.AdvTree.Node)sender;

            //term_node.BeginEdit();

            //2018/4/19 穎驊註解，看起來不太需要特別將編輯模式結束。
            //term_node.EndEdit();

            //term_node.Text = "YOYO";
        }

        private void InsertNewSubject(object sender, EventArgs e)
        {
            DevComponents.AdvTree.Node subject_node = (DevComponents.AdvTree.Node)sender;

            DevComponents.AdvTree.Node mother_node = subject_node.Parent;

            DevComponents.AdvTree.Node new_subject_node = new DevComponents.AdvTree.Node();

            new_subject_node.Text = "請輸入新子項目名稱";
            //新子項目點擊兩下，可以改變名稱
            new_subject_node.NodeDoubleClick += new System.EventHandler(NodeRename);

            DevComponents.AdvTree.Node add_new_assessment_node_btn = new DevComponents.AdvTree.Node();

            add_new_assessment_node_btn.Text = "<b><font color=\"#ED1C24\">+加入新子評分項目</font></b>";

            add_new_assessment_node_btn.NodeDoubleClick += new System.EventHandler(InsertNewAssessment);

            new_subject_node.Nodes.Add(add_new_assessment_node_btn);

            mother_node.Nodes.Insert(mother_node.Nodes.Count - 1, new_subject_node);
        }

        private void InsertNewAssessment(object sender, EventArgs e)
        {
            DevComponents.AdvTree.Node assessment_node = (DevComponents.AdvTree.Node)sender;

            DevComponents.AdvTree.Node mother_node = assessment_node.Parent;

            DevComponents.AdvTree.Node new_assessment_node = new DevComponents.AdvTree.Node();

            new_assessment_node.Text = "請輸入新子評分項目名稱";
            //新子項目點擊兩下，可以改變名稱
            new_assessment_node.NodeDoubleClick += new System.EventHandler(NodeRename);

            mother_node.Nodes.Insert(mother_node.Nodes.Count - 1, new_assessment_node);
        }



        private void NodeMouseDown(object sender, MouseEventArgs e)
        {

            node_now = (DevComponents.AdvTree.Node)sender;

            if (node_now.SelectedCell == null || node_now.Cells[1] != node_now.SelectedCell)
            {
                advTree1.ContextMenu = null;
                return;
            }


            MenuItem[] menuItems = new MenuItem[0];

            switch (node_now.Tag)
            {
                case "teacherKind":
                    //Declare the menu items and the shortcut menu.
                    menuItems = new MenuItem[]{new MenuItem("教師一",MenuItemNew_Click),
                new MenuItem("教師二",MenuItemNew_Click), new MenuItem("教師三",MenuItemNew_Click)};
                    LeftMouseClick(menuItems, e);
                    break;

                case "ScoreKind":
                    //Declare the menu items and the shortcut menu.
                    menuItems = new MenuItem[]{new MenuItem("分數",MenuItemNew_Click),
                new MenuItem("指標",MenuItemNew_Click), new MenuItem("評語",MenuItemNew_Click)};
                    LeftMouseClick(menuItems, e);
                    break;
                case "AllowCustom":
                    //Declare the menu items and the shortcut menu.
                    menuItems = new MenuItem[]{new MenuItem("是",MenuItemNew_Click),
                new MenuItem("否",MenuItemNew_Click)};
                    LeftMouseClick(menuItems, e);
                    break;
                default:
                    break;
            }

            //ContextMenuStrip contexMenuuu = new ContextMenuStrip();

            //contexMenuuu.Items.Add("Edit ");
            //contexMenuuu.Items.Add("Delete ");
            //contexMenuuu.Show();
            //contexMenuuu.ItemClicked += new ToolStripItemClickedEventHandler(
            //    contexMenuuu_ItemClicked);




        }

        //void contexMenuuu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        //{
        //    ToolStripItem item = e.ClickedItem;
        //    // your code here
        //}

        //  將右鍵點選的項目(ex: 教師一、教師二、教師三) 指定給 目前所選node 的第二個cell 
        private void MenuItemNew_Click(Object sender, System.EventArgs e)
        {
            System.Windows.Forms.MenuItem mi = (System.Windows.Forms.MenuItem)sender;

            node_now.Cells[1].Text = mi.Text;

            if (mi.Text == "指標")
            {
                DevComponents.AdvTree.Node new_indicator_setting_node = new DevComponents.AdvTree.Node();

                new_indicator_setting_node.Tag = "string";
                new_indicator_setting_node.Text = "設定指標";

                DevComponents.AdvTree.Cell col_1 = new DevComponents.AdvTree.Cell();
                DevComponents.AdvTree.Cell col_2 = new DevComponents.AdvTree.Cell();

                col_2.Text = "請輸入文字(ex: A、B、C)";

                new_indicator_setting_node.Cells.Add(col_1);

                new_indicator_setting_node.Cells.Add(col_2);

                node_now.Nodes.Add(new_indicator_setting_node);

            }
            else
            {
                node_now.Nodes.Clear();
            }

        }

        // 在編輯完之後驗證使用者輸入的資料是否符合型態
        private void advTree1_AfterCellEditComplete(object sender, DevComponents.AdvTree.CellEditEventArgs e)
        {
            DevComponents.AdvTree.AdvTree advt = (DevComponents.AdvTree.AdvTree)sender;

            node_now = advt.SelectedNode;

            switch (node_now.Tag)
            {
                // 整數
                case "integer":
                    if (!int.TryParse(node_now.SelectedCell.Text, out int check_cell_int))
                    {
                        //node_now.Style = DevComponents.AdvTree.NodeStyles.Red;
                        //node_now.StyleSelected = DevComponents.AdvTree.NodeStyles.Red;
                        node_now.Cells[2].Text = "<b><font color=\"#ED1C24\">" + hintGuideDict["" + node_now.Tag] + "</font></b>";
                    }
                    else
                    {
                        //node_now.Style = null;
                        //node_now.StyleSelected = null;
                        node_now.Cells[2].Text = hintGuideDict["" + node_now.Tag];
                    }
                    break;
                // 時間
                case "time":
                    if (!DateTime.TryParse(node_now.SelectedCell.Text, out DateTime check_cell_DateTime))
                    {
                        //node_now.Style = DevComponents.AdvTree.NodeStyles.Red;
                        //node_now.StyleSelected = DevComponents.AdvTree.NodeStyles.Red;
                        node_now.Cells[2].Text = "<b><font color=\"#ED1C24\">" + hintGuideDict["" + node_now.Tag] + "</font></b>";
                    }
                    else
                    {
                        //node_now.Style = null;
                        //node_now.StyleSelected = null;
                        node_now.Cells[2].Text = hintGuideDict["" + node_now.Tag];
                    }
                    break;
                // 字串，名稱必定要輸入
                case "string":
                    if (node_now.SelectedCell.Text == "")
                    {
                        //node_now.Style = DevComponents.AdvTree.NodeStyles.Red;
                        //node_now.StyleSelected = DevComponents.AdvTree.NodeStyles.Red;
                        node_now.Cells[2].Text = "<b><font color=\"#ED1C24\">" + hintGuideDict["" + node_now.Tag] + "</font></b>";
                    }
                    else
                    {
                        //node_now.Style = null;
                        //node_now.StyleSelected = null;
                        node_now.Cells[2].Text = hintGuideDict["" + node_now.Tag];
                    }
                    break;

                default:
                    break;
            }
        }

        private void LeftMouseClick(MenuItem[] menuItems, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ContextMenu buttonMenu = new ContextMenu(menuItems);

                advTree1.ContextMenu = buttonMenu;

                advTree1.ContextMenu.Show(advTree1, e.Location);
            }

        }

        private void ParseDBxmlToNodeUI(Term t)
        {

            DevComponents.AdvTree.Node new_term_node = new DevComponents.AdvTree.Node();

            new_term_node.Text = "試別" + (advTree1.Nodes.Count);
            
            //設定為不能點選編輯，避免使用者誤用
            new_term_node.Cells[0].Editable = false;
            

            DevComponents.AdvTree.Node new_term_node_name = new DevComponents.AdvTree.Node();
            DevComponents.AdvTree.Node new_term_node_percentage = new DevComponents.AdvTree.Node();
            DevComponents.AdvTree.Node new_term_node_inputStartTime = new DevComponents.AdvTree.Node();
            DevComponents.AdvTree.Node new_term_node_inputEndTime = new DevComponents.AdvTree.Node();

            //項目
            new_term_node_name.Text = "名稱";
            new_term_node_percentage.Text = "比例";
            new_term_node_inputStartTime.Text = "成績輸入開始時間:";
            new_term_node_inputEndTime.Text = "成績輸入截止時間:";
            //node Tag
            new_term_node_name.Tag = "string";
            new_term_node_percentage.Tag = "integer";
            new_term_node_inputStartTime.Tag = "time";
            new_term_node_inputEndTime.Tag = "time";

            //值
            new_term_node_name.Cells.Add(new DevComponents.AdvTree.Cell(t.Name));
            new_term_node_percentage.Cells.Add(new DevComponents.AdvTree.Cell(t.Percentage));
            new_term_node_inputStartTime.Cells.Add(new DevComponents.AdvTree.Cell(t.InputStartTime));
            new_term_node_inputEndTime.Cells.Add(new DevComponents.AdvTree.Cell(t.InputEndTime));

            //說明
            new_term_node_name.Cells.Add(new DevComponents.AdvTree.Cell(hintGuideDict["" + new_term_node_name.Tag]));
            new_term_node_percentage.Cells.Add(new DevComponents.AdvTree.Cell(hintGuideDict["" + new_term_node_percentage.Tag]));
            new_term_node_inputStartTime.Cells.Add(new DevComponents.AdvTree.Cell(hintGuideDict["" + new_term_node_inputStartTime.Tag]));
            new_term_node_inputEndTime.Cells.Add(new DevComponents.AdvTree.Cell(hintGuideDict["" + new_term_node_inputEndTime.Tag]));

            //設定為不能點選編輯，避免使用者誤用
            new_term_node_name.Cells[0].Editable = false;
            new_term_node_name.Cells[2].Editable = false;
            new_term_node_percentage.Cells[0].Editable = false;
            new_term_node_percentage.Cells[2].Editable = false;
            new_term_node_inputStartTime.Cells[0].Editable = false;
            new_term_node_inputStartTime.Cells[2].Editable = false;
            new_term_node_inputEndTime.Cells[0].Editable = false;
            new_term_node_inputEndTime.Cells[2].Editable = false;

            new_term_node.Nodes.Add(new_term_node_name);
            new_term_node.Nodes.Add(new_term_node_percentage);
            new_term_node.Nodes.Add(new_term_node_inputStartTime);
            new_term_node.Nodes.Add(new_term_node_inputEndTime);

            // 子項目
            foreach (Subject s in t.SubjectList)
            {
                DevComponents.AdvTree.Node new_subjet_node = new DevComponents.AdvTree.Node();

                new_subjet_node.Text = "子項目" + (new_term_node.Nodes.Count);

                DevComponents.AdvTree.Node new_subject_node_name = new DevComponents.AdvTree.Node();
                DevComponents.AdvTree.Node new_subject_node_percentage = new DevComponents.AdvTree.Node();

                //項目
                new_subject_node_name.Text = "名稱";
                new_subject_node_percentage.Text = "比例";
                //node Tag
                new_subject_node_name.Tag = "string";
                new_subject_node_percentage.Tag = "integer";

                //值
                new_subject_node_name.Cells.Add(new DevComponents.AdvTree.Cell(s.Name));
                new_subject_node_percentage.Cells.Add(new DevComponents.AdvTree.Cell(s.Percentage));

                //說明
                new_subject_node_name.Cells.Add(new DevComponents.AdvTree.Cell(hintGuideDict["" + new_subject_node_name.Tag]));
                new_subject_node_percentage.Cells.Add(new DevComponents.AdvTree.Cell(hintGuideDict["" + new_subject_node_percentage.Tag]));

                //設定為不能點選編輯，避免使用者誤用
                new_subject_node_name.Cells[0].Editable = false;
                new_subject_node_name.Cells[2].Editable = false;
                new_subject_node_percentage.Cells[0].Editable = false;
                new_subject_node_percentage.Cells[2].Editable = false;

                new_subjet_node.Nodes.Add(new_subject_node_name);
                new_subjet_node.Nodes.Add(new_subject_node_percentage);

                //設定為不能點選編輯，避免使用者誤用
                new_subjet_node.Cells[0].Editable = false;
                

                foreach (Assessment a in s.AssessmentList)
                {
                    DevComponents.AdvTree.Node new_assessment_node = new DevComponents.AdvTree.Node();

                    new_assessment_node.Text = "子評量" + (new_subjet_node.Nodes.Count);

                    DevComponents.AdvTree.Node new_assessment_node_name = new DevComponents.AdvTree.Node();
                    DevComponents.AdvTree.Node new_assessment_node_percentage = new DevComponents.AdvTree.Node();
                    DevComponents.AdvTree.Node new_assessment_node_teacherRole = new DevComponents.AdvTree.Node();
                    DevComponents.AdvTree.Node new_assessment_node_type = new DevComponents.AdvTree.Node();
                    DevComponents.AdvTree.Node new_assessment_node_allowCustomAssessment = new DevComponents.AdvTree.Node();


                    //項目
                    new_assessment_node_name.Text = "名稱";
                    new_assessment_node_percentage.Text = "比例";
                    new_assessment_node_teacherRole.Text = "評分老師";
                    new_assessment_node_type.Text = "評分種類";
                    new_assessment_node_allowCustomAssessment.Text = "是否允許自訂項目";

                    //node Tag
                    new_assessment_node_name.Tag = "string";
                    new_assessment_node_percentage.Tag = "integer";
                    new_assessment_node_teacherRole.Tag = "teacherKind";
                    new_assessment_node_type.Tag = "ScoreKind";
                    new_assessment_node_allowCustomAssessment.Tag = "AllowCustom";

                    //值
                    new_assessment_node_name.Cells.Add(new DevComponents.AdvTree.Cell(a.Name));
                    new_assessment_node_percentage.Cells.Add(new DevComponents.AdvTree.Cell(a.Percentage));
                    new_assessment_node_teacherRole.Cells.Add(new DevComponents.AdvTree.Cell(a.TeacherRole));
                    new_assessment_node_type.Cells.Add(new DevComponents.AdvTree.Cell(typeCovertDict[a.Type]));
                    new_assessment_node_allowCustomAssessment.Cells.Add(new DevComponents.AdvTree.Cell(a.AllowCustomAssessment == "true" ? "是" : "否"));

                    //說明
                    new_assessment_node_name.Cells.Add(new DevComponents.AdvTree.Cell(hintGuideDict["" + new_assessment_node_name.Tag]));
                    new_assessment_node_percentage.Cells.Add(new DevComponents.AdvTree.Cell(hintGuideDict["" + new_assessment_node_percentage.Tag]));
                    new_assessment_node_teacherRole.Cells.Add(new DevComponents.AdvTree.Cell(hintGuideDict["" + new_assessment_node_teacherRole.Tag]));
                    new_assessment_node_type.Cells.Add(new DevComponents.AdvTree.Cell(hintGuideDict["" + new_assessment_node_type.Tag]));
                    new_assessment_node_allowCustomAssessment.Cells.Add(new DevComponents.AdvTree.Cell(hintGuideDict["" + new_assessment_node_allowCustomAssessment.Tag]));

                    // 點擊事件 (適用於:teacherKind、ScoreKind、AllowCustom)
                    new_assessment_node_teacherRole.NodeMouseDown += new System.Windows.Forms.MouseEventHandler(NodeMouseDown);
                    new_assessment_node_type.NodeMouseDown += new System.Windows.Forms.MouseEventHandler(NodeMouseDown);
                    new_assessment_node_allowCustomAssessment.NodeMouseDown += new System.Windows.Forms.MouseEventHandler(NodeMouseDown);

                    //設定為不能點選編輯，避免使用者誤用
                    new_assessment_node_name.Cells[0].Editable = false;
                    new_assessment_node_name.Cells[2].Editable = false;
                    new_assessment_node_percentage.Cells[0].Editable = false;
                    new_assessment_node_percentage.Cells[2].Editable = false;
                    new_assessment_node_teacherRole.Cells[0].Editable = false;
                    new_assessment_node_teacherRole.Cells[2].Editable = false;
                    new_assessment_node_type.Cells[0].Editable = false;
                    new_assessment_node_type.Cells[2].Editable = false;
                    new_assessment_node_allowCustomAssessment.Cells[0].Editable = false;
                    new_assessment_node_allowCustomAssessment.Cells[2].Editable = false;


                    //假如有指標型分數 則加入最後一層指標型分數輸入
                    if (a.Type == "Indicator")
                    {
                        foreach (Indicators i in a.IndicatorsList)
                        {
                            DevComponents.AdvTree.Node new_indicators_node = new DevComponents.AdvTree.Node();

                            //項目
                            new_indicators_node.Text = "設定指標" + (new_assessment_node_type.Nodes.Count);
                            //node Tag
                            new_indicators_node.Tag = "string";
                            //值
                            new_indicators_node.Cells.Add(new DevComponents.AdvTree.Cell(i.Name));
                            //說明
                            new_indicators_node.Cells.Add(new DevComponents.AdvTree.Cell(hintGuideDict["" + new_indicators_node.Tag]));

                            //設定為不能點選編輯，避免使用者誤用
                            new_indicators_node.Cells[0].Editable = false;
                            new_indicators_node.Cells[2].Editable = false;

                            new_assessment_node_type.Nodes.Add(new_indicators_node);
                        }

                    }

                    new_assessment_node.Nodes.Add(new_assessment_node_name);
                    new_assessment_node.Nodes.Add(new_assessment_node_percentage);
                    new_assessment_node.Nodes.Add(new_assessment_node_teacherRole);
                    new_assessment_node.Nodes.Add(new_assessment_node_type);
                    new_assessment_node.Nodes.Add(new_assessment_node_allowCustomAssessment);

                    //設定為不能點選編輯，避免使用者誤用
                    new_assessment_node.Cells[0].Editable = false;
                    

                    new_subjet_node.Nodes.Add(new_assessment_node);

                }


                new_term_node.Nodes.Add(new_subjet_node);

            }






            //DevComponents.AdvTree.Node add_new_subject_node_btn = new DevComponents.AdvTree.Node();

            //add_new_subject_node_btn.Text = "<b><font color=\"#ED1C24\">+加入新子項目</font></b>";

            //add_new_subject_node_btn.NodeDoubleClick += new System.EventHandler(InsertNewSubject);

            //new_term_node.Nodes.Add(add_new_subject_node_btn);

            //不用Add 改用Insert ，是因為可以指定位子，讓加入新試別功能可以永遠在最後一項。
            advTree1.Nodes.Add(new_term_node);
            //advTree1.Nodes.Insert(advTree1.Nodes.Count - 1, new_term_node);





        }


    }
}

