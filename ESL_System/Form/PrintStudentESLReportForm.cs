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

        private List<Term> _termList = new List<Term>();


        public PrintStudentESLReportForm()
        {
            InitializeComponent();
        }

        // 列印
        private void btnPrint_Click(object sender, EventArgs e)
        {

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
                    "英文課程名稱",
                    "原班級名稱",
                    "學生英文姓名",
                    "學生中文姓名",
                    "教師一",
                    "教師二",
                    "教師三",
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
            string selQuery = "SELECT id,name,description FROM exam_template WHERE description IS NOT NULL ";
            dt = qh.Select(selQuery);



            foreach (DataRow dr in dt.Rows)
            {
                string xmlStr = "<root>" + dr["description"].ToString() + "</root>";

                string eslTemplateName = dr["name"].ToString();

                XElement elmRoot = XElement.Parse(xmlStr);

                ElmRootToTermlist(elmRoot);

                MergeFieldGenerator(builder, eslTemplateName);

                _termList.Clear(); // 每一個樣板 清完後 再加
            }




            #endregion






            #region 儲存檔案
            string inputReportName = "合併欄位總表";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

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
                doc.Save(path, Aspose.Words.SaveFormat.Doc);
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Excel檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        doc.Save(path, Aspose.Words.SaveFormat.Doc);
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



        private void ElmRootToTermlist(XElement elmRoot)
        {

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

                        _termList.Add(t); // 整理成大包的termList 後面用此來拚功能變數總表
                    }
                }
            }

        }


        private void MergeFieldGenerator(Aspose.Words.DocumentBuilder builder,string eslTemplateName)
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

            foreach (Term term in _termList)
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
                //builder.InsertField("MERGEFIELD " + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "名稱" + termCounter + " \\* MERGEFORMAT ", "«I" + termCounter + "»");
                builder.Write(term.Name);

                builder.InsertCell();
                //builder.InsertField("MERGEFIELD " + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "分數" + termCounter + " \\* MERGEFORMAT ", "«TS" + termCounter + "»");
                builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數" + " \\* MERGEFORMAT ", "«TS»");

                builder.InsertCell();
                //builder.InsertField("MERGEFIELD " + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "比重" + termCounter+ " \\* MERGEFORMAT ", "«TW" + termCounter + "»");
                builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數" + " \\* MERGEFORMAT ", "«TW»");

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
                        //builder.InsertField("MERGEFIELD " + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "名稱" + assessmentCounter + " \\* MERGEFORMAT ", "«I" + assessmentCounter + "»");
                        builder.Write(assessment.Name);

                        builder.InsertCell();
                        //builder.InsertField("MERGEFIELD " + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "比重" + assessmentCounter + " \\* MERGEFORMAT ", "«AW" + assessmentCounter + "»");
                        builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "比重" + " \\* MERGEFORMAT ", "«AW»");

                        builder.InsertCell();
                        //builder.InsertField("MERGEFIELD " + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "分數" + assessmentCounter + " \\* MERGEFORMAT ", "«S" + assessmentCounter + "»");
                        builder.InsertField("MERGEFIELD " + "評量" + "_" + term.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + subject.Name.Trim().Replace(' ', '_').Replace('"', '_') + "/" + assessment.Name.Trim().Replace(' ', '_').Replace('"', '_') + "_" + "分數" + " \\* MERGEFORMAT ", "«AS»");

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


    }
}
