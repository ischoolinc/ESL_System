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
using System.IO;
using Aspose.Words;

namespace ESL_System.Form
{
    public partial class PrintStudentESLReportForm : FISCA.Presentation.Controls.BaseForm
    {
        private BackgroundWorker _bw = new BackgroundWorker();

        private string _wordURL = "";

        // 儲放樣板 id 與 樣板名稱的對照
        private Dictionary<string, string> _templateIDNameDict = new Dictionary<string, string>();

        // 儲放學生ESL 成績的dict 其結構為 <studentID,<scoreKey,scoreValue>
        private Dictionary<string, Dictionary<string, string>> _scoreDict = new Dictionary<string, Dictionary<string, string>>();

        // 儲放ESL 成績單 科目、比重設定 的dict 
        private Dictionary<string, Dictionary<string, string>> _itemDict = new Dictionary<string,Dictionary<string, string>>();

        // 紀錄成績 為 指標型indicator 的 key值 ， 作為對照 key 為 courseID_termName_subjectName_assessment_Name
        private List<string> _indicatorList = new List<string>();

        // 紀錄成績 為 評語型comment 的 key值
        private List<string> _commentList = new List<string>();

        private Document _doc;

        private DataTable _mergeDataTable = new DataTable();


        public PrintStudentESLReportForm()
        {
            InitializeComponent();

            _bw.DoWork += new DoWorkEventHandler(_bkWork_DoWork);
            _bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_worker_RunWorkerCompleted);
            _bw.ProgressChanged += new ProgressChangedEventHandler(_worker_ProgressChanged);
            _bw.WorkerReportsProgress = true;
        }

        // 列印
        private void btnPrint_Click(object sender, EventArgs e)
        {
            // 關閉畫面控制項
            btnPrint.Enabled = false;
            btnClose.Enabled = false;
            linklabel3.Enabled = false;


            // 2018/10/29 穎驊註解，目前的作法先每一次都讓使用者列印前，選擇列印樣板
            // 等本次期中考後，再看使用情境 怎麼去做列印樣板設定。
            OpenFileDialog ope = new OpenFileDialog();
            ope.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";

            if (ope.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            else
            {
                _wordURL = ope.FileName;
                _bw.RunWorkerAsync();

            }
        }

        // 離開
        private void btnClose_Click(object sender, EventArgs e)
        {

        }

        // 檢視套印樣板
        private void linklabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        // 變更套印樣板
        private void linklabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        // 檢視功能變數總表
        private void linklabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 新寫法，依照　所選取的ESL 樣板設定，產生出動態對應階層的功能變數總表！！
            CreateFieldTemplate();
            return;
        }


        // 穎驊搬過來的工具，可以一次大量建立有規則的功能變數，可以省下很多時間。
        private void CreateFieldTemplate()
        {
            List<Term> termList = new List<Term>();

            Aspose.Words.Document doc = new Aspose.Words.Document();
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);

            // Create a paragraph style and specify some formatting for it.
            Aspose.Words.Style style = builder.Document.Styles.Add(Aspose.Words.StyleType.Paragraph, "ESLNameStyle");

            style.Font.Size = 24;
            style.Font.Bold = true;
            style.ParagraphFormat.SpaceAfter = 12;

            #region 固定變數

            builder.ParagraphFormat.Style = builder.Document.Styles["ESLNameStyle"];
            // 固定變數，不分　期中、期末、學期  (使用大字粗體)
            builder.Writeln("固定變數");

            builder.ParagraphFormat.Style = builder.Document.Styles["Normal"];

            builder.StartTable();
            builder.InsertCell();
            builder.Write("項目");
            builder.InsertCell();
            builder.Write("變數");
            builder.EndRow();
            foreach (string key in new string[]{
                    "學年度",
                    "學期",
                    "學號",
                    "年級",
                    "原班級名稱",
                    "學生英文姓名",
                    "學生中文姓名",
                    "電子報表辨識編號"
                })
            {
                builder.InsertCell();
                builder.Write(key);
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + key + " \\* MERGEFORMAT ", "«" + key + "»");
                builder.EndRow();
            }

            builder.EndTable();

            builder.Writeln();
            #endregion


            #region  解讀　description　XML
            // 取得ESL 描述 in description
            DataTable dt;
            QueryHelper qh = new QueryHelper();

            // 抓所有目前系統 有設定的ESL 評分樣版
            string selQuery = "SELECT id,name,description FROM exam_template WHERE description IS NOT NULL ORDER BY name  ";
            dt = qh.Select(selQuery);

            foreach (DataRow dr in dt.Rows)
            {
                string xmlStr = "<root>" + dr["description"].ToString() + "</root>";

                string eslTemplateName = dr["name"].ToString();

                XElement elmRoot = XElement.Parse(xmlStr);

                termList = ElmRootToTermlist(elmRoot);

                MergeFieldGenerator(builder, eslTemplateName, termList);

                termList.Clear(); // 每一個樣板 清完後 再加

                _templateIDNameDict.Add(dr["id"].ToString(), dr["name"].ToString());
            }


            #endregion


            #region 儲存檔案
            string inputReportName = "合併欄位總表";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".docx");

            if (System.IO.File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!System.IO.File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {
                doc.Save(path, Aspose.Words.SaveFormat.Docx);
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".docx";
                sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        doc.Save(path, Aspose.Words.SaveFormat.Docx);
                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            #endregion
        }

        private List<Term> ElmRootToTermlist(XElement elmRoot)
        {
            List<Term> termList = new List<Term>();

            //解析讀下來的 descriptiony 資料，打包成物件群 
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

                                if (a.Type == "Comment") // 假如是 評語類別，多讀一項 輸入限制屬性
                                {
                                    a.InputLimit = ele_assessment.Attribute("InputLimit").Value;
                                }

                                a.IndicatorsList = new List<Indicators>();

                                if (ele_assessment.Element("Indicators") != null)
                                {

                                    foreach (XElement ele_Indicator in ele_assessment.Element("Indicators").Elements("Indicator"))
                                    {
                                        Indicators i = new Indicators();

                                        i.Name = ele_Indicator.Attribute("Name").Value;
                                        i.Description = ele_Indicator.Attribute("Description").Value;

                                        a.IndicatorsList.Add(i);
                                    }
                                }
                                s.AssessmentList.Add(a);
                            }
                            t.SubjectList.Add(s);
                        }

                        termList.Add(t); // 整理成大包的termList 後面用此來拚功能變數總表
                    }
                }
            }

            return termList;
        }

        private void MergeFieldGenerator(Aspose.Words.DocumentBuilder builder, string eslTemplateName,List<Term> termList)
        {
            #region 成績變數

            int termCounter = 1;

            // 2018/6/15 穎驊備註 以下整理 功能變數 最常使用的 string..Trim().Replace(' ', '_').Replace('"', '_') 
            // >> 其用意為避免Word 功能變數合併列印時 會有一些奇怪的BUG ，EX: row["Final-Term評量_Science科目_In-Class Score子項目_分數1"] = "YOYO!"; >> 有空格印不出來 


            // Apply the paragraph style to the current paragraph in the document and add some text.
            builder.ParagraphFormat.Style = builder.Document.Styles["ESLNameStyle"];
            // 每一個 ESL 樣板的名稱 放在最上面 (使用大字粗體)
            builder.Writeln("樣板名稱: " + eslTemplateName);

            // Change to a paragraph style that has no list formatting. (將字體還原)
            builder.ParagraphFormat.Style = builder.Document.Styles["Normal"];
            builder.Writeln("");

            foreach (Term term in termList)
            {
                builder.Writeln(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "評量");

                builder.StartTable();
                builder.InsertCell();
                builder.Write("評量名稱");
                builder.InsertCell();
                builder.Write("評量分數");
                builder.InsertCell();
                builder.Write("評量比重");
                builder.EndRow();

                builder.InsertCell();
                
                builder.Write(term.Name);

                builder.InsertCell();
                
                // 2018/10/29 穎驊註解，和恩正討論後，不同樣板之間的 Term 名稱 會分不清楚， 因此在前面加 評分樣板作區別
                builder.InsertField("MERGEFIELD " + eslTemplateName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數" + " \\* MERGEFORMAT ", "«TS»");

                builder.InsertCell();
                
                builder.InsertField("MERGEFIELD " + eslTemplateName.Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重" + " \\* MERGEFORMAT ", "«TW»");

                //termCounter++;

                builder.EndRow();
                builder.EndTable();

                builder.Writeln();

                int subjectCounter = 1;



                foreach (Subject subject in term.SubjectList)
                {
                    builder.Writeln(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "科目分數型成績");

                    builder.StartTable();
                    builder.InsertCell();
                    builder.Write("科目名稱");
                    builder.InsertCell();
                    builder.Write("科目分數");
                    builder.InsertCell();
                    builder.Write("科目比重");
                    builder.EndRow();


                    builder.InsertCell();
                    builder.Write(subject.Name);
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數" + " \\* MERGEFORMAT ", "«SS»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重" + " \\* MERGEFORMAT ", "«SW»");

                    //subjectCounter++;

                    builder.EndRow();
                    builder.EndTable();

                    builder.StartTable();
                    builder.InsertCell();
                    builder.Write("子項目名稱");
                    builder.InsertCell();
                    builder.Write("比重");
                    builder.InsertCell();
                    builder.Write("分數");
                    builder.InsertCell();
                    builder.Write("教師");

                    builder.EndRow();

                    int assessmentCounter = 1;

                    bool assessmentContainsIndicator = false;

                    bool assessmentContainsComment = false;

                    foreach (Assessment assessment in subject.AssessmentList)
                    {
                        if (assessment.Type == "Indicator") // 檢查看有沒有　　Indicator　，有的話，會另外再畫一張表專放Indicator                       
                        {
                            assessmentContainsIndicator = true;
                        }

                        if (assessment.Type == "Comment") // 檢查看有沒有　　Comment　，有的話，會另外再畫一張表專放Comment                        
                        {
                            assessmentContainsComment = true;
                        }

                        if (assessment.Type != "Score") //  非分數型成績 跳過 不寫入
                        {
                            continue;
                        }


                        builder.InsertCell();
                        builder.Write(assessment.Name);

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重" + " \\* MERGEFORMAT ", "«AW»");

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數" + " \\* MERGEFORMAT ", "«AS»");

                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師" + " \\* MERGEFORMAT ", "«AT»");

                        assessmentCounter++;

                        builder.EndRow();
                    }

                    builder.EndTable();
                    builder.Writeln();


                    // 處理Indicator
                    if (assessmentContainsIndicator)
                    {
                        builder.Writeln(term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "指標型成績");

                        builder.StartTable();
                        builder.InsertCell();
                        builder.Write("項目");
                        builder.InsertCell();
                        builder.Write("指標");
                        builder.InsertCell();
                        builder.Write("教師");
                        builder.EndRow();

                        assessmentCounter = 1;
                        foreach (Assessment assessment in subject.AssessmentList)
                        {
                            if (assessment.Type == "Indicator") // 檢查看有沒有　Indicator　，專為 Indicator 畫張表
                            {
                                builder.InsertCell();
                                builder.Write(assessment.Name);
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "指標" + " \\* MERGEFORMAT ", "«I»");
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師" + " \\* MERGEFORMAT ", "«T»");
                                builder.EndRow();
                                //assessmentCounter++;
                            }
                        }
                        builder.EndTable();
                        builder.Writeln();
                    }

                    // 處理Comment
                    if (assessmentContainsComment)
                    {
                        builder.Writeln(term.Name + "/" + subject.Name + "評語型成績");

                        builder.StartTable();
                        builder.InsertCell();
                        builder.Write("項目");
                        builder.InsertCell();
                        builder.Write("評語");
                        builder.InsertCell();
                        builder.Write("教師");
                        builder.EndRow();

                        assessmentCounter = 1;
                        foreach (Assessment assessment in subject.AssessmentList)
                        {
                            if (assessment.Type == "Comment") // 檢查看有沒有　Comment　，專為 Comment 畫張表
                            {
                                builder.InsertCell();
                                builder.Write(assessment.Name);
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "評語" + " \\* MERGEFORMAT ", "«C»");
                                builder.InsertCell();
                                builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師" + " \\* MERGEFORMAT ", "«T»");
                                builder.EndRow();
                                assessmentCounter++;
                            }
                        }
                        builder.EndTable();
                        builder.Writeln();
                    }

                }


                builder.Writeln();
                builder.Writeln();
            }

            #endregion


        }

        private void _bkWork_DoWork(object sender, DoWorkEventArgs e)
        {

            List<string> studentIDList = new List<string>();
            List<string> courseIDList = new List<string>();

            // 選擇的學生名單 
            studentIDList = K12.Presentation.NLDPanels.Student.SelectedSource;




            #region 取得本學期有設定ESL 評分樣版的課程清單
            QueryHelper qh = new QueryHelper();

            // 選出 本學期有設定ESL 評分樣版的清單
            string sqlCourseID = @"
    SELECT 
		id
		,course_name
		,school_year
		,semester
		,ref_exam_template_id 
	FROM course 
	WHERE 
	school_year =" + K12.Data.School.DefaultSchoolYear +
    @"AND semester = " + K12.Data.School.DefaultSemester +
    @"AND ref_exam_template_id IN(
		SELECT id 
		FROM exam_template  
		WHERE description IS NOT NULL)";

            DataTable dtCourseID = qh.Select(sqlCourseID);

            foreach (DataRow row in dtCourseID.Rows)
            {
                string id = "" + row["id"];

                courseIDList.Add(id);
            }

            if (courseIDList.Count == 0)
            {
                MsgBox.Show("本學期沒有任何設定ESL樣板的課程。", "錯誤!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            #endregion


            #region  解讀　description　XML

            // 取得ESL 描述 in description
            DataTable dtTemplateDescription;

            string selQuery = "SELECT id,name,description FROM exam_template WHERE description IS NOT NULL";

            dtTemplateDescription = qh.Select(selQuery);

            Dictionary<string, List<Term>> termListDict = new Dictionary<string, List<Term>>();

            List<List<Term>> termListCollection = new List<List<Term>>();

            //整理目前的ESL 課程資料
            if (dtTemplateDescription.Rows.Count > 0)
            {
                foreach (DataRow dr in dtTemplateDescription.Rows)
                {
                    List<Term> termList = new List<Term>();

                    string xmlStr = "<root>" + dr["description"].ToString() + "</root>";
                    XElement elmRoot = XElement.Parse(xmlStr);

                    termList = ElmRootToTermlist(elmRoot);

                    termListDict.Add(dr["name"].ToString(), termList);

                }

                _mergeDataTable = GetMergeField(termListDict);
            }


            #endregion

            // 取得課程基本資料 (教師)
            List<K12.Data.CourseRecord> courseList = K12.Data.Course.SelectByIDs(courseIDList);

            // 取得教師基本資料 
            List<K12.Data.TeacherRecord> teacherList = K12.Data.Teacher.SelectAll();

            //取得學生基本資料
            List<K12.Data.StudentRecord> studentList = K12.Data.Student.SelectByIDs(studentIDList);


            #region 取得ESL 課程成績
            _bw.ReportProgress(20, "取得ESL課程成績");


            int progress = 80;
            decimal per = (decimal)(100 - progress) / studentIDList.Count;
            int count = 0;

            string course_ids = string.Join("','", courseIDList);

            string student_ids = string.Join("','", studentIDList);

            // 建立成績結構
            foreach (string stuID in studentIDList)
            {
                _scoreDict.Add(stuID, new Dictionary<string, string>());
            }

            // 按照時間順序抓， 如果有相同的成績結構， 以後來新的 取代前的
            string sqlScore = "SELECT * FROM $esl.gradebook_assessment_score WHERE ref_course_id IN ('" + course_ids + "') AND ref_student_id IN ('" + student_ids + "') ORDER BY last_update,ref_student_id "; 

            DataTable dtScore = qh.Select(sqlScore);

            foreach (DataRow row in dtScore.Rows)
            {
                string termWord = "" + row["term"];
                string subjectWord = "" + row["subject"];
                string assessmentWord = "" + row["assessment"];

                string id = "" + row["ref_student_id"];

                // 有教師自訂的子項目成績就跳掉 不處理
                if ("" + row["custom_assessment"] != "")
                {
                    continue;
                }

                // 要設計一個模式 處理 三種成績
                CourseRecord courseRecord = courseList.Find(c => c.ID == "" + row["ref_course_id"]);

                // 項目都有，為assessment 成績
                if (termWord != "" && "" + subjectWord != "" && "" + assessmentWord != "")
                {
                    if (_scoreDict.ContainsKey(id))
                    {
                        // 指標型成績
                        if (_indicatorList.Contains("" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_')))
                        {
                            string scoreKey = "評量" + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_');

                            if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "指標"))
                            {
                                _scoreDict[id].Add(scoreKey + "_" + "指標", "" + row["value"]);
                            }
                            else
                            {
                                _scoreDict[id][scoreKey + "_" + "指標"] = "" + row["value"];
                            }

                            if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "教師"))
                            {
                                _scoreDict[id].Add(scoreKey + "_" + "教師", teacherList.Find(t => t.ID == "" + row["ref_teacher_id"]).Name); //教師名稱
                            }
                            else
                            {
                                _scoreDict[id][scoreKey + "_" + "教師"] = teacherList.Find(t => t.ID == "" + row["ref_teacher_id"]).Name; //教師名稱
                            }

                        }
                        // 評語型成績
                        else if (_commentList.Contains("" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "_" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_')))
                        {
                            string scoreKey = "評量" + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_');

                            if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "評語"))
                            {
                                _scoreDict[id].Add(scoreKey + "_" + "評語", "" + row["value"]);
                            }
                            else
                            {
                                _scoreDict[id][scoreKey + "_" + "評語"] = "" + row["value"];
                            }

                            if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "教師"))
                            {
                                _scoreDict[id].Add(scoreKey + "_" + "教師", teacherList.Find(t => t.ID == "" + row["ref_teacher_id"]).Name); //教師名稱
                            }
                            else
                            {
                                _scoreDict[id][scoreKey + "_" + "教師"] = teacherList.Find(t => t.ID == "" + row["ref_teacher_id"]).Name; //教師名稱
                            }

                        }
                        // 分數型成績
                        else
                        {
                            string scoreKey = "評量" + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessmentWord.Trim().Replace(' ', '_').Replace('"', '_');

                            if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "分數"))
                            {
                                _scoreDict[id].Add(scoreKey + "_" + "分數", "" + row["value"]);
                                _scoreDict[id].Add(scoreKey + "_" + "比重", _itemDict[courseRecord.AssessmentSetup.Name][scoreKey + "_" + "比重"]);
                            }
                            else
                            {
                                _scoreDict[id][scoreKey + "_" + "分數"] = "" + row["value"];
                            }

                            if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "教師"))
                            {
                                _scoreDict[id].Add(scoreKey + "_" + "教師", teacherList.Find(t => t.ID == "" + row["ref_teacher_id"]).Name); //教師名稱
                            }
                            else
                            {
                                _scoreDict[id][scoreKey + "_" + "教師"] = teacherList.Find(t => t.ID == "" + row["ref_teacher_id"]).Name; //教師名稱
                            }
                            
                        }
                    }
                }

                // 沒有assessment，為subject 成績
                if (termWord != "" && "" + subjectWord != "" && "" + assessmentWord == "")
                {
                    if (_scoreDict.ContainsKey(id))
                    {
                        string scoreKey = "評量" + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subjectWord.Trim().Replace(' ', '_').Replace('"', '_');
                        
                        if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "分數"))
                        {
                            _scoreDict[id].Add(scoreKey + "_" + "分數", "" + row["value"]);
                            _scoreDict[id].Add(scoreKey + "_" + "比重", _itemDict[courseRecord.AssessmentSetup.Name][scoreKey + "_" + "比重"]);
                        }
                        else
                        {
                            _scoreDict[id][scoreKey + "_" + "分數"] = "" + row["value"];
                        }

                    }
                }
                // 沒有assessment、subject，為term 成績
                if (termWord != "" && "" + subjectWord == "" && "" + assessmentWord == "")
                {
                    

                    
                    if (_scoreDict.ContainsKey(id))
                    {
                        // 2018/10/29 穎驊註解，和恩正討論後，不同樣板之間的 Term 名稱 會分不清楚， 因此在前面加 評分樣板作區別
                        string scoreKey = courseRecord.AssessmentSetup.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "評量" + "_" + termWord.Trim().Replace(' ', '_').Replace('"', '_');
                        
                        if (!_scoreDict[id].ContainsKey(scoreKey + "_" + "分數"))
                        {
                            _scoreDict[id].Add(scoreKey + "_" + "分數", "" + row["value"]);
                            _scoreDict[id].Add(scoreKey + "_" + "比重", _itemDict[courseRecord.AssessmentSetup.Name][scoreKey + "_" + "比重"]);
                        }
                        else
                        {
                            _scoreDict[id][scoreKey + "_" + "分數"] = "" + row["value"];
                        }
                    }
                }

            }
            #endregion




            foreach (StudentRecord stuRecord in studentList)
            {
                string id = stuRecord.ID;

                DataRow row = _mergeDataTable.NewRow();

                row["電子報表辨識編號"] = "系統編號{" + stuRecord.ID + "}"; // 學生系統編號

                row["學年度"] = K12.Data.School.DefaultSchoolYear;
                row["學期"] = K12.Data.School.DefaultSemester;
                row["學號"] = stuRecord.StudentNumber;
                row["年級"] = stuRecord.Class != null ? "" + stuRecord.Class.GradeYear : "";

                row["原班級名稱"] = stuRecord.Class != null ? "" + stuRecord.Class.Name : "";
                row["學生英文姓名"] = stuRecord.EnglishName;
                row["學生中文姓名"] = stuRecord.Name;

                
                //foreach (string mergeKey in _itemDict.Keys)
                //{
                //    if (row.Table.Columns.Contains(mergeKey))
                //    {                                               
                //        row[mergeKey] = _itemDict[mergeKey];
                //    }
                //}

                if (_scoreDict.ContainsKey(id))
                {
                    foreach (string mergeKey in _scoreDict[id].Keys)
                    {
                        if (row.Table.Columns.Contains(mergeKey))
                        {
                            row[mergeKey] = _scoreDict[id][mergeKey];
                        }
                    }
                }

                _mergeDataTable.Rows.Add(row);

                count++;
                progress += (int)(count * per);
                _bw.ReportProgress(progress);

            }


            try
            {
                // 載入使用者所選擇的 word 檔案
                _doc = new Document(_wordURL);
            }
            catch (Exception ex)
            {
                MsgBox.Show(ex.Message);
                e.Cancel = true;
                return;
            }


            _doc.MailMerge.Execute(_mergeDataTable);

            e.Result = _doc;
        }

        private void _worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                this.Close();
                return;
            }

            FISCA.Presentation.MotherForm.SetStatusBarMessage(" ESL學生成績單產生完成。");

            Document doc = (Document)e.Result;
            doc.MailMerge.DeleteFields();

            // 電子報表功能先暫時不製做
            #region 電子報表
            //// 檢查是否上傳電子報表
            //if (chkUploadEPaper.Checked)
            //{
            //    List<Document> docList = new List<Document>();
            //    foreach (Section ss in doc.Sections)
            //    {
            //        Document dc = new Document();
            //        dc.Sections.Clear();
            //        dc.Sections.Add(dc.ImportNode(ss, true));
            //        docList.Add(dc);
            //    }

            //    Update_ePaper up = new Update_ePaper(docList, "超額比序項目積分證明單", PrefixStudent.系統編號);
            //    if (up.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
            //    {
            //        MsgBox.Show("電子報表已上傳!!");
            //    }
            //    else
            //    {
            //        MsgBox.Show("已取消!!");
            //    }
            //} 
            #endregion

            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "ESL學生成績單.docx";
            sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    doc.Save(sd.FileName, Aspose.Words.SaveFormat.Docx);
                    System.Diagnostics.Process.Start(sd.FileName);
                }
                catch
                {
                    MessageBox.Show("檔案儲存失敗");
                }
            }

            this.Close();
        }

        private void _worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }
        
        private DataTable GetMergeField(Dictionary<string, List<Term>> termListDict)
        {
            // 計算權重用的字典(因為使用者在介面設定的權重數值 不一定就是想在報表上顯示的)
            // 目前康橋報表希望能夠將，每一個Subject、assessment 的比重 換算成為對於期中考的比例
            Dictionary<string, float> weightCalDict = new Dictionary<string, float>();

            DataTable dataTable = new DataTable();

            #region 固定變數
            // 固定變數
            // 基本資料
            dataTable.Columns.Add("學年度");
            dataTable.Columns.Add("學期");
            dataTable.Columns.Add("學號");
            dataTable.Columns.Add("年級");
            dataTable.Columns.Add("原班級名稱");
            dataTable.Columns.Add("學生英文姓名");
            dataTable.Columns.Add("學生中文姓名");        
            dataTable.Columns.Add("電子報表辨識編號");
            #endregion


            // 2018/6/15 穎驊備註 以下整理 功能變數 最常使用的 string.Trim().Replace(' ', '_').Replace('"', '_') 
            // >> 其用意為避免Word 功能變數合併列印時 會有一些奇怪的BUG ，EX: row["Final-Term評量_Science科目_In-Class Score子項目_分數1"] = "YOYO!"; >> 有空格印不出來 

            foreach (string templateName in termListDict.Keys)
            {
                //每一個 template 清空一次 weight 計算用 字典
                weightCalDict.Clear();

                _itemDict.Add(templateName,new Dictionary<string, string>());

                foreach (Term term in termListDict[templateName])
                {
                    // 2018/10/29 穎驊註解，和恩正討論後，不同樣板之間的 Term 名稱 會分不清楚， 因此在前面加 評分樣板作區別

                    if (!dataTable.Columns.Contains(templateName.Replace(' ', '_').Replace('"', '_') +  "_" +"評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重"))
                        dataTable.Columns.Add(templateName.Replace(' ', '_').Replace('"', '_') +  "_" + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重");

                    if (!dataTable.Columns.Contains(templateName.Replace(' ', '_').Replace('"', '_') +  "_" + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數"))
                        dataTable.Columns.Add(templateName.Replace(' ', '_').Replace('"', '_') +  "_" + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數"); // Term 分數本身 先暫時這樣處理之後要有類別整理

                    if (!_itemDict[templateName].ContainsKey(templateName.Replace(' ', '_').Replace('"', '_') +  "_" + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重"))
                    {                        
                        _itemDict[templateName].Add(templateName.Replace(' ', '_').Replace('"', '_') +  "_" + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重", term.Weight);
                    }


                    // 計算比重用，先整理 Term 、 Subject  的 總和
                    foreach (Subject subject in term.SubjectList)
                    {
                        // Term
                        if (!weightCalDict.ContainsKey(term.Name + "_SubjectTotal"))
                        {
                            if (float.TryParse(subject.Weight, out float f))
                            {
                                weightCalDict.Add(term.Name + "_SubjectTotal", f);
                            }
                        }
                        else
                        {
                            if (float.TryParse(subject.Weight, out float f))
                            {
                                weightCalDict[term.Name + "_SubjectTotal"] += f;
                            }
                        }

                        // Subject
                        if (!weightCalDict.ContainsKey(term.Name + "_" + subject.Name))
                        {
                            if (float.TryParse(subject.Weight, out float f))
                            {
                                weightCalDict.Add(term.Name + "_" + subject.Name, f);
                            }
                        }

                    }

                    foreach (Subject subject in term.SubjectList)
                    {
                        if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重"))
                            dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重");
                        if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數"))
                            dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數"); // Subject 分數本身 先暫時這樣處理之後要有類別整理


                        string subjectWieght = "" + Math.Round((float.Parse(subject.Weight) * 100) / (weightCalDict[term.Name + "_SubjectTotal"]), 2, MidpointRounding.ToEven);

                        _itemDict[templateName].Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重", subjectWieght); //subject比重 

                        // 計算比重用，先整理 Assessment  的 總和
                        foreach (Assessment assessment in subject.AssessmentList)
                        {
                            if (!weightCalDict.ContainsKey(term.Name + "_" + subject.Name + "_AssessmentTotal"))
                            {
                                if (float.TryParse(assessment.Weight, out float f))
                                {
                                    weightCalDict.Add(term.Name + "_" + subject.Name + "_AssessmentTotal", f);
                                }
                            }
                            else
                            {
                                if (float.TryParse(assessment.Weight, out float f))
                                {
                                    weightCalDict[term.Name + "_" + subject.Name + "_AssessmentTotal"] += f;
                                }
                            }
                        }



                        foreach (Assessment assessment in subject.AssessmentList)
                        {
                            if (assessment.Type == "Score") //分數型成績 才增加
                            {
                                if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重"))
                                    dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重");
                                if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數"))
                                    dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數");
                                if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師"))
                                    dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師");

                                string assessmentWieght = "" + Math.Round((weightCalDict[term.Name + "_" + subject.Name] * float.Parse(assessment.Weight) * 100) / (weightCalDict[term.Name + "_SubjectTotal"] * weightCalDict[term.Name + "_" + subject.Name + "_AssessmentTotal"]), 2, MidpointRounding.ToEven);

                                _itemDict[templateName].Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重", assessmentWieght); //assessment比重 
                                
                            }
                            if (assessment.Type == "Indicator") // 檢查看有沒有　　Indicator　，有的話另外存List 做對照
                            {
                                if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "指標"))
                                    dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "指標");
                                if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師"))
                                    dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師");

                                string key = term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_');

                                // 作為對照 key 為 courseID_termName_subjectName_assessment_Name
                                if (!_indicatorList.Contains(key))
                                {
                                    _indicatorList.Add(key);
                                }

                                
                            }

                            if (assessment.Type == "Comment") // 檢查看有沒有　　Comment　，有的話另外存List 做對照
                            {
                                if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "評語"))
                                    dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "評語");
                                if (!dataTable.Columns.Contains("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師"))
                                    dataTable.Columns.Add("評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "教師");
                                

                                string key = term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_');

                                // 作為對照 key 為 courseID_termName_subjectName_assessment_Name
                                if (!_commentList.Contains(key))
                                {
                                    _commentList.Add(key);
                                }
                            }

                        }
                    }

                }

            }
  
            return dataTable;
        }

    }
}
