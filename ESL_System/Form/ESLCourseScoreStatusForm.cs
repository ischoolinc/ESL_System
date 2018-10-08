using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using Aspose.Cells;

using K12.Data;
using System.Xml;
using System.Data;
using FISCA.Data;
using System.Xml.Linq;

namespace ESL_System.Form
{
    public partial class ESLCourseScoreStatusForm : FISCA.Presentation.Controls.BaseForm
    {

        private List<CourseListViewItem> _courseListViewItemList;
        private List<string> _CourseIDList;
        private List<ESLCourse> _ESLCourseList = new List<ESLCourse>();


        // 目標樣板ID
        private string _targetTemplateID;

        // 目標試別(Term)名稱
        private string _targetTermName;



        //  ESL 課程ID 與 評分樣版ID 的對照
        private Dictionary<string, string> _ESLCourseIDExamTemIDDict = new Dictionary<string, string>();

        //  評分樣版名稱 與 評分樣版ID 的對照
        private Dictionary<string, string> _ExamTemNameExamTemIDDict = new Dictionary<string, string>();

        //   <評分樣版ID,ESLTemplate>
        private Dictionary<string, ESLTemplate> _ESLTemplateDict = new Dictionary<string, ESLTemplate>();


        public ESLCourseScoreStatusForm(List<string> eslCouseList)
        {
            InitializeComponent();
            _CourseIDList = eslCouseList;

            GetESLTemplate();

            FillCboTemplate();



        }

        private void GetESLTemplate()
        {
            string courseIDs = string.Join(",", _CourseIDList);

            #region 取得ESL 課程資料
            // 2018/06/12 抓取課程且其有ESL 樣板設定規則的，才做後續整理，  在table exam_template 欄位 description 不為空代表其為ESL 的樣板
            string query = @"
                    SELECT 
                        course.id AS courseID
                        ,course.course_name
                        ,exam_template.description 
                        ,exam_template.id AS templateID
                        ,exam_template.name AS templateName
                    FROM course 
                    LEFT JOIN  exam_template ON course.ref_exam_template_id =exam_template.id  
                    WHERE course.id IN( " + courseIDs + ") AND  exam_template.description IS NOT NULL  ";

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(query);

            _CourseIDList.Clear(); // 清空

            //整理目前的ESL 課程資料
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    ESLCourse record = new ESLCourse();

                    _CourseIDList.Add("" + dr[0]); // 加入真正的 是ESL 課程ID

                    record.ID = "" + dr[0]; //課程ID
                    record.CourseName = "" + dr[1]; //課程名稱
                    record.Description = "" + dr[2]; // ESL 評分樣版設定

                    _ESLCourseList.Add(record);

                    if (!_ESLCourseIDExamTemIDDict.ContainsKey("" + dr["courseID"]))
                    {
                        _ESLCourseIDExamTemIDDict.Add("" + dr["courseID"], "" + dr["templateID"]);
                    }

                    if (!_ESLTemplateDict.ContainsKey("" + dr["templateID"]))
                    {
                        ESLTemplate template = new ESLTemplate();

                        template.ID = "" + dr["templateID"];
                        template.ESLTemplateName = "" + dr["templateName"];
                        template.Description = "" + dr["description"];

                        _ESLTemplateDict.Add("" + dr["templateID"], template);
                    }




                }
            }
            #endregion

            #region 解析ESL 課程 計算規則
            // 解析計算規則
            foreach (string templateID in _ESLTemplateDict.Keys)
            {
                string xmlStr = "<root>" + _ESLTemplateDict[templateID].Description + "</root>";
                XElement elmRoot = XElement.Parse(xmlStr);

                //解析讀下來的 descriptiony 資料
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
                            t.Ref_exam_id = ele_term.Attribute("Ref_exam_id").Value;

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

                            

                            _ESLTemplateDict[templateID].TermList.Add(t);
                        }
                    }
                }
            }
            #endregion

        }


        private void FillCboTemplate()
        {

            cboTemplate.Items.Clear();

            foreach (string templateID in _ESLTemplateDict.Keys)
            {
                cboTemplate.Items.Add(_ESLTemplateDict[templateID].ESLTemplateName);

                _ExamTemNameExamTemIDDict.Add(_ESLTemplateDict[templateID].ESLTemplateName, _ESLTemplateDict[templateID].ID);
            }
        }


        /// <summary>
        /// 當試別改變時觸發
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboExam_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshListView();
        }

        /// <summary>
        /// 更新ListView
        /// </summary>
        private void RefreshListView()
        {
            if (cboExam.SelectedItem == null) return;

            _targetTermName = "" + cboExam.SelectedItem;

            // 清內容
            listView.Items.Clear();

            // 清表頭
            listView.Columns.Clear();

            listView.Columns.Add("課程名稱",200);

            // 依目前 所選樣板、試別 動態產生 listView 表頭
            foreach (Term term in _ESLTemplateDict[_targetTemplateID].TermList)
            {
                if (term.Name == _targetTermName)
                {
                    foreach (Subject subject in term.SubjectList)
                    {
                        foreach (Assessment assessment in subject.AssessmentList)
                        {
                            listView.Columns.Add(assessment.Name, assessment.Name.Length * 9);
                        }
                    }
                }
            }


            //LoadCourses(exam_id);
            //SortItemList();
            //FillCourses(GetDisplayList());
        }


        /// <summary>
        /// 依試別取得所有關聯課程
        /// </summary>
        /// <param name="exam_id"></param>
        private void LoadCourses(string exam_id)
        {

        }

        /// <summary>
        /// 取得某試別的所有成績記錄
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, List<ESLScore>> GetCourseScores(string exam_id, List<string> CourseIDs)
        {
            #region 過濾非在校生

            #endregion

            #region 依 CourseID，將成績分開成不同 List

            Dictionary<string, List<ESLScore>> courseScoreDict = new Dictionary<string, List<ESLScore>>();



            #endregion

            return courseScoreDict;
        }

        /// <summary>
        /// 將課程填入ListView
        /// </summary>
        private void FillCourses(List<CourseListViewItem> list)
        {
            if (list.Count <= 0) return;

            listView.SuspendLayout();
            listView.Items.Clear();
            listView.Items.AddRange(list.ToArray());
            listView.ResumeLayout();
        }

        /// <summary>
        /// 將 CourseListViewItemList 排序
        /// </summary>
        private void SortItemList()
        {
            _courseListViewItemList.Sort(delegate (CourseListViewItem a, CourseListViewItem b) { return a.Text.CompareTo(b.Text); });
        }

        /// <summary>
        /// 按下「關閉」時觸發
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 改變「僅顯示未完成輸入之課程」時觸發
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkDisplayNotFinish_CheckedChanged(object sender, EventArgs e)
        {
            FillCourses(GetDisplayList());
        }

        /// <summary>
        /// 取得要顯示的 CourseListViewItemList
        /// </summary>
        /// <returns></returns>
        private List<CourseListViewItem> GetDisplayList()
        {
            if (chkDisplayNotFinish.Checked == true)
            {
                List<CourseListViewItem> list = new List<CourseListViewItem>();
                foreach (CourseListViewItem item in _courseListViewItemList)
                {
                    if (item.IsFinish) continue;
                    list.Add(item);
                }
                return list;
            }
            else
            {
                return _courseListViewItemList;
            }
        }

        ///// <summary>
        ///// 按下「匯出到 Excel」時觸發
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void btnExport_Click(object sender, EventArgs e)
        //{
        //    if (listView.Items.Count <= 0) return;

        //    saveFileDialog1.FileName = string.Format("{0}學年度{1}學期{2}課程成績輸入狀況", intSchoolYear.Value, (intSemester.Value == 1) ? "上" : "下", ((KeyValuePair<string, string>)cboExam.SelectedItem).Value);
        //    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
        //    {
        //        Workbook book = new Workbook();
        //        book.Worksheets.Clear();
        //        Worksheet ws = book.Worksheets[book.Worksheets.Add()];
        //        ws.Name = string.Format("{0}學年度 {1}學期 {2}", intSchoolYear.Value, (intSemester.Value == 1) ? "上" : "下", ((KeyValuePair<string, string>)cboExam.SelectedItem).Value);

        //        #region 加入 Header

        //        int row = 0;
        //        ws.Cells[row, chCourseName.Index].PutValue(chCourseName.Text);
        //        ws.Cells[row, chTeacher.Index].PutValue(chTeacher.Text);
        //        ws.Cells[row, chScore.Index].PutValue(chScore.Text);
        //        ws.Cells[row, chAssignment.Index].PutValue(chAssignment.Text);
        //        ws.Cells[row, chText.Index].PutValue(chText.Text);

        //        #endregion

        //        #region 加入每一筆課程輸入狀況

        //        listView.SuspendLayout();
        //        foreach (CourseListViewItem item in listView.Items)
        //        {
        //            row++;
        //            ws.Cells[row, chCourseName.Index].PutValue(item.Text);
        //            ws.Cells[row, chTeacher.Index].PutValue(item.SubItems[chTeacher.Index].Text);
        //            ws.Cells[row, chScore.Index].PutValue(item.SubItems[chScore.Index].Text);
        //            ws.Cells[row, chAssignment.Index].PutValue(item.SubItems[chAssignment.Index].Text);
        //            ws.Cells[row, chText.Index].PutValue(item.SubItems[chText.Index].Text);
        //        }
        //        listView.ResumeLayout();

        //        #endregion

        //        ws.AutoFitColumns();

        //        try
        //        {
        //            book.Save(saveFileDialog1.FileName, FileFormatType.Excel2003);
        //            Framework.MsgBox.Show("匯出完成。");
        //        }
        //        catch (Exception ex)
        //        {
        //            Framework.MsgBox.Show("匯出失敗。" + ex.Message);
        //        }
        //    }
        //}

        /// <summary>
        /// 按下「重新整理」時觸發
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRefresh_Click(object sender, EventArgs e)
        {

            RefreshListView();
        }

        /// <summary>
        /// 每一筆課程的評量狀況
        /// </summary>
        private class CourseListViewItem : ListViewItem
        {
            private const string Format = "{0}/{1}";
            private int _total;
            private int _scoreCount, _assignmentCount, _textCount;

            private string ScoreField { get { return string.Format(Format, _scoreCount, _total); } }
            private string AssignmentField { get { return string.Format(Format, _assignmentCount, _total); } }
            private string TextField { get { return string.Format(Format, _textCount, _total); } }

            private bool _is_finish;
            public bool IsFinish { get { return _is_finish; } }

            //public CourseListViewItem(JHSchool.CourseRecord course, HC.JHAEIncludeRecord aei, List<HC.JHSCETakeRecord> sceList)
            //{
            //    _is_finish = true;

            //    TeacherRecord teacher = course.GetFirstTeacher();

            //    //_total = course.GetAttends().Count;
            //    _total = SCAttend.GetCourseStudentCount(course.ID);
            //    Calculate(sceList);

            //    if (aei.UseScore) _is_finish &= (_scoreCount == _total);
            //    if (aei.UseAssignmentScore) _is_finish &= (_assignmentCount == _total);
            //    if (aei.UseText) _is_finish &= (_textCount == _total);

            //    this.Text = course.Name;
            //    this.SubItems.Add((teacher != null) ? teacher.Name : "");
            //    this.SubItems.Add(aei.UseScore ? ScoreField : "").ForeColor = (_scoreCount == _total) ? Color.Black : Color.Red;
            //    this.SubItems.Add(aei.UseAssignmentScore ? AssignmentField : "").ForeColor = (_assignmentCount == _total) ? Color.Black : Color.Red;
            //    this.SubItems.Add(aei.UseText ? TextField : "").ForeColor = (_textCount == _total) ? Color.Black : Color.Red;
            //}

            //private void Calculate(List<HC.JHSCETakeRecord> sceList)
            //{
            //    _scoreCount = _assignmentCount = _textCount = 0;

            //    foreach (var sce in sceList)
            //    {
            //        if (sce.Score.HasValue) _scoreCount++;
            //        if (sce.AssignmentScore.HasValue) _assignmentCount++;
            //        if (!string.IsNullOrEmpty(sce.Text)) _textCount++;
            //    }
            //}
        }

        private void ESLCourseScoreStatusForm_Load(object sender, EventArgs e)
        {
            
        }

        private void cboTemplate_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboExam.Enabled = true;

            _targetTemplateID = _ExamTemNameExamTemIDDict["" + cboTemplate.SelectedItem];

            FillcboExam();
        }

        private void FillcboExam()
        {
            cboExam.Items.Clear();

            foreach (Term term in (_ESLTemplateDict[_targetTemplateID].TermList))
            {
                cboExam.Items.Add(term.Name);
            }
        }


    }
}

