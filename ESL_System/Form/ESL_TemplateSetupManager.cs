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
using K12.Data;
using System.Xml;


namespace ESL_System.Form
{
    public partial class ESL_TemplateSetupManager : FISCA.Presentation.Controls.BaseForm
    {
        private BackgroundWorker _workder;

        private Dictionary<string, string> hintGuideDict = new Dictionary<string, string>();

        private Dictionary<string, string> typeCovertDict = new Dictionary<string, string>();
        private Dictionary<string, string> typeCovertRevDict = new Dictionary<string, string>();

        private Dictionary<string, string> teacherRoleCovertDict = new Dictionary<string, string>();
        private Dictionary<string, string> teacherRoleCovertRevDict = new Dictionary<string, string>();


        private Dictionary<string, string> nodeTagCovertDict = new Dictionary<string, string>();


        // 現在點在哪一小節
        DevComponents.AdvTree.Node node_now;

        public ESL_TemplateSetupManager()
        {
            InitializeComponent();
            HideNavigationBar();

            hintGuideDict.Add("string", "請輸入文字，不得空白、重覆。");
            hintGuideDict.Add("integer", "請輸入整數數字");
            hintGuideDict.Add("time", "請輸入日期(ex : 2018/04/21 00:00:00)");
            hintGuideDict.Add("teacherKind", "點選左鍵選取: 教師一、教師二、教師三");
            hintGuideDict.Add("ScoreKind", "點選左鍵選取:分數、指標、評語");
            hintGuideDict.Add("AllowCustom", "點選左鍵選取:是、否");

            typeCovertDict.Add("Score", "分數");
            typeCovertDict.Add("Indicator", "指標");
            typeCovertDict.Add("Comment", "評語");

            typeCovertRevDict.Add("分數", "Score");
            typeCovertRevDict.Add("指標", "Indicator");
            typeCovertRevDict.Add("評語", "Comment");

            teacherRoleCovertDict.Add("1", "教師一");
            teacherRoleCovertDict.Add("2", "教師二");
            teacherRoleCovertDict.Add("3", "教師三");

            teacherRoleCovertRevDict.Add("教師一", "1");
            teacherRoleCovertRevDict.Add("教師二", "2");
            teacherRoleCovertRevDict.Add("教師三", "3");

            nodeTagCovertDict.Add("term", "試別");
            nodeTagCovertDict.Add("subject", "科目");
            nodeTagCovertDict.Add("assessment", "評量");

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
                        t.Weight = ele_term.Attribute("Weight").Value;
                        t.InputStartTime = ele_term.Attribute("InputStartTime").Value;
                        t.InputEndTime = ele_term.Attribute("InputEndTime").Value;

                        t.SubjectList = new List<Subject>();

                        foreach (XElement ele_subject in ele_term.Elements("Subject"))
                        {
                            Subject s = new Subject();

                            s.Name = ele_subject.Attribute("Name").Value;
                            s.Weight = ele_subject.Attribute("Weight").Value;

                            s.AssessmentList = new List<Assessment>();

                            foreach (XElement ele_assessment in ele_subject.Elements("Assessment"))
                            {
                                Assessment a = new Assessment();

                                a.Name = ele_assessment.Attribute("Name").Value;
                                a.Weight = ele_assessment.Attribute("Weight").Value;
                                a.TeacherSequence = ele_assessment.Attribute("TeacherSequence").Value;
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

            // 選擇老師後 更新項目名稱 評量(XXX,教師一)
            if (node_now.TagString == "teacherKind")
            {
                node_now.Parent.Cells[0].Text = nodeTagCovertDict["" + node_now.Parent.Tag] + "(" + node_now.Parent.Nodes[0].Cells[1].Text + "," + node_now.SelectedCell.Text + ")";
            }
                         
            if (mi.Text == "指標")
            {
                // 選擇指標後，將比重設定為0，且disable
                node_now.Parent.Nodes[1].Cells[1].Text = "0";
                node_now.Parent.Nodes[1].Cells[2].Text = "指標型分數無法輸入比例";
                node_now.Parent.Nodes[1].Enabled = false;


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
                // 選非指標，將其還原
                node_now.Parent.Nodes[1].Enabled = true;
                node_now.Parent.Nodes[1].Cells[2].Text = hintGuideDict[node_now.Parent.Nodes[1].TagString];
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
                // 字串，名稱必定要輸入，且不能重覆
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
                        
                        // 更新項目名稱
                        if (node_now.Cells[0].Text=="名稱:"  )
                        {
                            if (node_now.Parent.TagString != "assessment")
                            {
                                node_now.Parent.Cells[0].Text = nodeTagCovertDict["" + node_now.Parent.Tag] + "(" + node_now.SelectedCell.Text + ")";
                            }
                            else
                            {
                                // 加上評分 教師腳色
                                node_now.Parent.Cells[0].Text = nodeTagCovertDict["" + node_now.Parent.Tag] + "(" + node_now.SelectedCell.Text +","+ node_now.Parent.Nodes[2].Cells[1].Text +")";
                            }                            
                        }
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

        //將資料印至畫面上
        private void ParseDBxmlToNodeUI(Term t)
        {

            DevComponents.AdvTree.Node new_term_node = new DevComponents.AdvTree.Node();

            new_term_node.Text = "試別(" +t.Name+ ")";

            new_term_node.TagString = "term";

            //設定為不能點選編輯，避免使用者誤用
            new_term_node.Cells[0].Editable = false;

            //設定為不能拖曳，避免使用者誤用
            new_term_node.DragDropEnabled = false;

            DevComponents.AdvTree.Node new_term_node_name = new DevComponents.AdvTree.Node();
            DevComponents.AdvTree.Node new_term_node_percentage = new DevComponents.AdvTree.Node();
            DevComponents.AdvTree.Node new_term_node_inputStartTime = new DevComponents.AdvTree.Node();
            DevComponents.AdvTree.Node new_term_node_inputEndTime = new DevComponents.AdvTree.Node();

            //項目
            new_term_node_name.Text = "名稱:";
            new_term_node_percentage.Text = "比例:";
            new_term_node_inputStartTime.Text = "成績輸入開始時間:";
            new_term_node_inputEndTime.Text = "成績輸入截止時間:";
            //node Tag
            new_term_node_name.Tag = "string";
            new_term_node_percentage.Tag = "integer";
            new_term_node_inputStartTime.Tag = "time";
            new_term_node_inputEndTime.Tag = "time";

            //值
            new_term_node_name.Cells.Add(new DevComponents.AdvTree.Cell(t.Name));
            new_term_node_percentage.Cells.Add(new DevComponents.AdvTree.Cell(t.Weight));
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

            //設定為不能拖曳，避免使用者誤用
            new_term_node_name.DragDropEnabled = false;
            new_term_node_percentage.DragDropEnabled = false;
            new_term_node_inputStartTime.DragDropEnabled = false;
            new_term_node_inputEndTime.DragDropEnabled = false;
            

            new_term_node.Nodes.Add(new_term_node_name);
            new_term_node.Nodes.Add(new_term_node_percentage);
            new_term_node.Nodes.Add(new_term_node_inputStartTime);
            new_term_node.Nodes.Add(new_term_node_inputEndTime);

            // 科目
            foreach (Subject s in t.SubjectList)
            {
                DevComponents.AdvTree.Node new_subjet_node = new DevComponents.AdvTree.Node();

                new_subjet_node.Text = "科目("+ s.Name+")";

                new_subjet_node.TagString = "subject";

                DevComponents.AdvTree.Node new_subject_node_name = new DevComponents.AdvTree.Node();
                DevComponents.AdvTree.Node new_subject_node_percentage = new DevComponents.AdvTree.Node();

                //項目
                new_subject_node_name.Text = "名稱:";
                new_subject_node_percentage.Text = "比例:";
                //node Tag
                new_subject_node_name.Tag = "string";
                new_subject_node_percentage.Tag = "integer";

                //值
                new_subject_node_name.Cells.Add(new DevComponents.AdvTree.Cell(s.Name));
                new_subject_node_percentage.Cells.Add(new DevComponents.AdvTree.Cell(s.Weight));

                //說明
                new_subject_node_name.Cells.Add(new DevComponents.AdvTree.Cell(hintGuideDict["" + new_subject_node_name.Tag]));
                new_subject_node_percentage.Cells.Add(new DevComponents.AdvTree.Cell(hintGuideDict["" + new_subject_node_percentage.Tag]));

                //設定為不能點選編輯，避免使用者誤用
                new_subject_node_name.Cells[0].Editable = false;
                new_subject_node_name.Cells[2].Editable = false;
                new_subject_node_percentage.Cells[0].Editable = false;
                new_subject_node_percentage.Cells[2].Editable = false;

                //設定為不能拖曳，避免使用者誤用
                new_subject_node_name.DragDropEnabled = false;
                new_subject_node_percentage.DragDropEnabled = false;

                new_subjet_node.Nodes.Add(new_subject_node_name);
                new_subjet_node.Nodes.Add(new_subject_node_percentage);

                //設定為不能點選編輯，避免使用者誤用
                new_subjet_node.Cells[0].Editable = false;


                foreach (Assessment a in s.AssessmentList)
                {
                    DevComponents.AdvTree.Node new_assessment_node = new DevComponents.AdvTree.Node();

                    new_assessment_node.Text = "評量(" + a.Name + ","+ teacherRoleCovertDict[a.TeacherSequence] + ")";

                    new_assessment_node.TagString = "assessment";
                    

                    DevComponents.AdvTree.Node new_assessment_node_name = new DevComponents.AdvTree.Node();
                    DevComponents.AdvTree.Node new_assessment_node_percentage = new DevComponents.AdvTree.Node();
                    DevComponents.AdvTree.Node new_assessment_node_teacherRole = new DevComponents.AdvTree.Node();
                    DevComponents.AdvTree.Node new_assessment_node_type = new DevComponents.AdvTree.Node();
                    DevComponents.AdvTree.Node new_assessment_node_allowCustomAssessment = new DevComponents.AdvTree.Node();


                    //項目
                    new_assessment_node_name.Text = "名稱:";
                    new_assessment_node_percentage.Text = "比例:";
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
                    new_assessment_node_percentage.Cells.Add(new DevComponents.AdvTree.Cell(a.Weight));
                    new_assessment_node_teacherRole.Cells.Add(new DevComponents.AdvTree.Cell(teacherRoleCovertDict[a.TeacherSequence]));
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

                    //設定為不能拖曳，避免使用者誤用
                    new_assessment_node_name.DragDropEnabled = false;
                    new_assessment_node_percentage.DragDropEnabled = false;
                    new_assessment_node_teacherRole.DragDropEnabled = false;
                    new_assessment_node_type.DragDropEnabled = false;
                    new_assessment_node_allowCustomAssessment.DragDropEnabled = false;


                    //假如有指標型分數 則加入最後一層指標型分數輸入
                    if (a.Type == "Indicator")
                    {
                        foreach (Indicators i in a.IndicatorsList)
                        {
                            DevComponents.AdvTree.Node new_indicators_node = new DevComponents.AdvTree.Node();

                            //項目
                            new_indicators_node.Text = "設定指標";
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

                        // 假如其為指標型分數 將比例 設定 0,disable
                        new_assessment_node_percentage.Cells[1].Text = "0";
                        new_assessment_node_percentage.Cells[2].Text = "指標型分數無法輸入比例";
                        new_assessment_node_percentage.Enabled = false;

                        //new_assessment_node_type.Expand(); //預設不展開此項
                    }

                    new_assessment_node.Nodes.Add(new_assessment_node_name);
                    new_assessment_node.Nodes.Add(new_assessment_node_percentage);
                    new_assessment_node.Nodes.Add(new_assessment_node_teacherRole);
                    new_assessment_node.Nodes.Add(new_assessment_node_type);
                    new_assessment_node.Nodes.Add(new_assessment_node_allowCustomAssessment);

                    //設定為不能點選編輯，避免使用者誤用
                    new_assessment_node.Cells[0].Editable = false;


                    //設定為不能拖曳，避免使用者誤用
                    new_subjet_node.DragDropEnabled = false;

                    new_subjet_node.Nodes.Add(new_assessment_node);
                    //展開
                    new_subjet_node.Expand();

                }


                new_term_node.Nodes.Add(new_subjet_node);
                //展開
                new_term_node.Expand();

            }






            //DevComponents.AdvTree.Node add_new_subject_node_btn = new DevComponents.AdvTree.Node();

            //add_new_subject_node_btn.Text = "<b><font color=\"#ED1C24\">+加入新子項目</font></b>";

            //add_new_subject_node_btn.NodeDoubleClick += new System.EventHandler(InsertNewSubject);

            //new_term_node.Nodes.Add(add_new_subject_node_btn);

            //不用Add 改用Insert ，是因為可以指定位子，讓加入新試別功能可以永遠在最後一項。
            advTree1.Nodes.Add(new_term_node);            
            //advTree1.Nodes.Insert(advTree1.Nodes.Count - 1, new_term_node);





        }

        //儲存ESL 樣板
        private void btnSave_Click(object sender, EventArgs e)
        {
            string description_xml = "";

            XmlDocument doc = new XmlDocument();
                        
            XmlElement root = doc.DocumentElement;
            
            //string.Empty makes cleaner code
            XmlElement element_ESLTemplate = doc.CreateElement(string.Empty, "ESLTemplate", string.Empty);
            doc.AppendChild(element_ESLTemplate);

            foreach (DevComponents.AdvTree.Node term_node in advTree1.Nodes)
            {
                XmlElement element_Term = doc.CreateElement(string.Empty, "Term", string.Empty);
                element_Term.SetAttribute("Name", term_node.Nodes[0].Cells[1].Text);
                element_Term.SetAttribute("Weight", term_node.Nodes[1].Cells[1].Text);
                element_Term.SetAttribute("InputStartTime", term_node.Nodes[2].Cells[1].Text);
                element_Term.SetAttribute("InputEndTime", term_node.Nodes[3].Cells[1].Text);

                foreach (DevComponents.AdvTree.Node subject_node in term_node.Nodes)
                {
                    if (subject_node.TagString == "subject")
                    {
                        XmlElement element_Subject = doc.CreateElement(string.Empty, "Subject", string.Empty);
                        element_Subject.SetAttribute("Name", subject_node.Nodes[0].Cells[1].Text);
                        element_Subject.SetAttribute("Weight", subject_node.Nodes[1].Cells[1].Text);


                        foreach (DevComponents.AdvTree.Node assessment_node in subject_node.Nodes)
                        {
                            if (assessment_node.TagString == "assessment")
                            {
                                XmlElement element_Assessment = doc.CreateElement(string.Empty, "Assessment", string.Empty);

                                element_Assessment.SetAttribute("Name", assessment_node.Nodes[0].Cells[1].Text);
                                element_Assessment.SetAttribute("Weight", assessment_node.Nodes[1].Cells[1].Text);
                                element_Assessment.SetAttribute("TeacherSequence", teacherRoleCovertRevDict[assessment_node.Nodes[2].Cells[1].Text]);
                                element_Assessment.SetAttribute("Type", typeCovertRevDict[assessment_node.Nodes[3].Cells[1].Text]);
                                element_Assessment.SetAttribute("AllowCustomAssessment", assessment_node.Nodes[4].Cells[1].Text);

                                //假如 type 有子項目，其代表indicators
                                if (assessment_node.Nodes[3].Nodes.Count > 0)
                                {
                                    XmlElement element_indicators = doc.CreateElement(string.Empty, "Indicators", string.Empty);

                                    foreach (DevComponents.AdvTree.Node indicators_node in assessment_node.Nodes[3].Nodes)
                                    {
                                        XmlElement element_indicator = doc.CreateElement(string.Empty, "Indicator", string.Empty);

                                        XmlText text = doc.CreateTextNode(indicators_node.Cells[1].Text);

                                        element_indicator.AppendChild(text);

                                        element_indicators.AppendChild(element_indicator);
                                    }

                                    element_Assessment.AppendChild(element_indicators);
                                }



                                element_Subject.AppendChild(element_Assessment);
                            }
                        }
                        element_Term.AppendChild(element_Subject);
                    }
                }

                element_ESLTemplate.AppendChild(element_Term);
            }
            
            //XmlElement element3 = doc.CreateElement(string.Empty, "level2", string.Empty);
            //XmlText text1 = doc.CreateTextNode("text");
            //element3.AppendChild(text1);
            //element2.AppendChild(element3);

            //XmlElement element4 = doc.CreateElement(string.Empty, "level2", string.Empty);
            //XmlText text2 = doc.CreateTextNode("other text");
            //element4.AppendChild(text2);
            //element2.AppendChild(element4);


            description_xml = doc.OuterXml;







            UpdateHelper uh = new UpdateHelper();

            string updQuery = "UPDATE exam_template SET description ='"+ description_xml + "' WHERE name ='ESL 科目樣版' ";

            //執行sql，更新
            uh.Execute(updQuery);






        }
    }
}

