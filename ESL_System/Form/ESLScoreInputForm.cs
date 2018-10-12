﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;


namespace ESL_System.Form
{
    public partial class ESLScoreInputForm : FISCA.Presentation.Controls.BaseForm
    {

        private string _targetCourseName;
        private string _targetTermName;
        private string _targetSubjectName;
        private string _targetAssessmentName;

        // 目標分數種類(Score、Indicator、Comment)
        private string _targetScoreType;

        // 指標性評語 可以使用項目
        private List<Indicators> _targetIndicatorList = new List<Indicators>();

        // 評語型項目 字數上限
        private string _commentLimit;

        private ESLTemplate _eslTemplate;

        private Dictionary<string, List<ESLScore>> _scoreDict = new Dictionary<string, List<ESLScore>>();

        private List<K12.Data.SCAttendRecord> _scaList;

        private BackgroundWorker _worker;

       
        public ESLScoreInputForm(string targetCourseName, ESLTemplate eslTemplate, Dictionary<string, List<ESLScore>> scoreDict, string targetTermName, List<K12.Data.SCAttendRecord> scaList)
        {
            InitializeComponent();

            _targetCourseName = targetCourseName;
            _targetTermName = targetTermName;

            _eslTemplate = eslTemplate;

            

            _scoreDict = scoreDict;

            _scaList = scaList;

            labelX1.Text = _targetCourseName; // 課程名稱

            FillCboSubject(); // 填科目選項


        }

        public ESLScoreInputForm(string targetCourseName, ESLTemplate eslTemplate, Dictionary<string, List<ESLScore>> scoreDict, string targetTermName, string targetSubjectName, string targetAssessmentName, List<K12.Data.SCAttendRecord> scaList)
        {
            InitializeComponent();

            _targetCourseName = targetCourseName;
            _targetTermName = targetTermName;
            _targetSubjectName = targetSubjectName;
            _targetAssessmentName = targetAssessmentName;

            _eslTemplate = eslTemplate;

            _targetScoreType = _eslTemplate.TermList.Find(t => t.Name == _targetTermName).SubjectList.Find(s => s.Name == _targetSubjectName).AssessmentList.Find(a => a.Name == _targetAssessmentName).Type;

            _targetIndicatorList = _eslTemplate.TermList.Find(t => t.Name == _targetTermName).SubjectList.Find(s => s.Name == _targetSubjectName).AssessmentList.Find(a => a.Name == _targetAssessmentName).IndicatorsList;

            _commentLimit = _eslTemplate.TermList.Find(t => t.Name == _targetTermName).SubjectList.Find(s => s.Name == _targetSubjectName).AssessmentList.Find(a => a.Name == _targetAssessmentName).InputLimit;

            _scoreDict = scoreDict;

            _scaList = scaList;

            labelX1.Text = _targetCourseName + "/" + _targetSubjectName + "/" + _targetAssessmentName ; // 課程名稱 / 科目名稱 / 評量名稱

            cboSubject.Visible = false;
            cboAssessment.Visible = false;
            labelX2.Visible = false;
            labelX3.Visible = false;

            FillScore();
        }

        private void cboSubject_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboAssessment.Enabled = true;

            FillcboAssessment();

        }

        private void cboAssessment_SelectedIndexChanged(object sender, EventArgs e)
        {
            _targetAssessmentName = "" + cboAssessment.SelectedItem;
            FillScore();
        }

        private void FillcboAssessment()
        {
            cboAssessment.Items.Clear();

            _targetSubjectName = "" + cboSubject.SelectedItem;

            List<string> assessmentList = new List<string>();

            foreach (ESLScore scoreItem in _scoreDict[_targetSubjectName])
            {
                if (!assessmentList.Contains(scoreItem.Assessment))
                {
                    assessmentList.Add(scoreItem.Assessment);
                }
            }

            foreach (string assessmentName in assessmentList)
            {
                cboAssessment.Items.Add(assessmentName);
            }
        }

        private void FillCboSubject()
        {

            cboSubject.Items.Clear();

            foreach (string subjectName in _scoreDict.Keys)
            {
                cboSubject.Items.Add(subjectName);
            }
        }

        private void FillScore()
        {
            dataGridViewX1.Rows.Clear();

           
            List<ESLScore> eslScoreList = new List<ESLScore>();

            foreach (ESLScore scoreItem in _scoreDict[_targetSubjectName])
            {
                if (scoreItem.Assessment == _targetAssessmentName)
                {
                    DataGridViewRow row = new DataGridViewRow();

                    row.CreateCells(dataGridViewX1);

                    K12.Data.SCAttendRecord scar = _scaList.Find(sca => sca.RefStudentID == scoreItem.RefStudentID && sca.RefCourseID == scoreItem.RefCourseID);

                    row.Cells[0].Value = scar.Student.Class != null ? scar.Student.Class.Name : "";
                    row.Cells[1].Value = scar.Student != null ? "" + scar.Student.SeatNo : "";
                    row.Cells[2].Value = scar.Student != null ? "" + scar.Student.Name : "";
                    row.Cells[3].Value = scar.Student != null ? "" + scar.Student.StudentNumber : "";
                    row.Cells[4].Value = scoreItem.RefTeacherName;
                    row.Cells[5].Value = scoreItem.HasValue ? scoreItem.Value : "";

                    row.Tag =  scoreItem.RefStudentID;  // row tag 用studentID 就夠

                    dataGridViewX1.Rows.Add(row);

                }
            }

            // 依   學號 排序
            dataGridViewX1.Sort(ColStudentNumber, ListSortDirection.Ascending);

        }


        // 儲存
        private void buttonX1_Click(object sender, EventArgs e)
        {

            _worker = new BackgroundWorker();
            _worker.DoWork += new DoWorkEventHandler(Worker_DoWork);
            _worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Worker_RunWorkerCompleted);
            _worker.ProgressChanged += new ProgressChangedEventHandler(Worker_ProgressChanged);
            _worker.WorkerReportsProgress = true;

            // 暫停畫面控制項
            cboSubject.SuspendLayout();
            cboAssessment.SuspendLayout();
            dataGridViewX1.SuspendLayout();
            buttonX1.Enabled = false;
           
            _worker.RunWorkerAsync();

        }



        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {

            //拚SQL
            // 兜資料
            List<string> dataList = new List<string>();

            List<ESLScore> updateESLscoreList = new List<ESLScore>(); // 最後要update ESLscoreList
            List<ESLScore> insertESLscoreList = new List<ESLScore>(); // 最後要indert ESLscoreList


            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {               
                foreach (ESLScore scoreItem in _scoreDict[_targetSubjectName])
                {
                    // 評量相同 、 學生ID 相同 則為該成績
                    if (scoreItem.Assessment == _targetAssessmentName && scoreItem.RefStudentID == ""+ row.Tag)
                    {
                        // 本來有成績， 但值不同了 加入 更新名單
                        if (scoreItem.HasValue && scoreItem.Value != "" + row.Cells[5].Value)
                        {
                            scoreItem.Value = "" + row.Cells[5].Value; // 新分數

                            updateESLscoreList.Add(scoreItem);
                        }

                        // 本來沒成績， 本次新填的值 ，加入新增名單
                        if (!scoreItem.HasValue && "" + row.Cells[5].Value !="")
                        {
                            scoreItem.Value = "" + row.Cells[5].Value; // 新分數

                            scoreItem.HasValue = true; 

                            insertESLscoreList.Add(scoreItem);
                        }
                    }
                }
            }



            foreach (ESLScore score in updateESLscoreList)
            {
                string data = string.Format(@"
                SELECT
                    '{0}'::BIGINT AS ref_student_id
                    ,'{1}'::BIGINT AS ref_course_id
                    ,'{2}'::BIGINT AS ref_teacher_id
                    ,'{3}'::TEXT AS term
                    ,'{4}'::TEXT AS subject
                    ,'{5}'::TEXT AS assessment
                    ,'{6}'::TEXT AS custom_assessment
                    ,'{7}'::TEXT AS value
                    ,'{8}'::INTEGER AS uid
                    ,'UPDATE'::TEXT AS action
                ", score.RefStudentID, score.RefCourseID, score.RefTeacherID, score.Term, score.Subject,score.Assessment,"",score.Value, score.ID);

                dataList.Add(data);
            }

            foreach (ESLScore score in insertESLscoreList)
            {
                string data = string.Format(@"
                SELECT
                    '{0}'::BIGINT AS ref_student_id
                    ,'{1}'::BIGINT AS ref_course_id
                    ,'{2}'::BIGINT AS ref_teacher_id
                    ,'{3}'::TEXT AS term
                    ,'{4}'::TEXT AS subject
                    ,'{5}'::TEXT AS assessment
                    ,'{6}'::TEXT AS custom_assessment
                    ,'{7}'::TEXT AS value
                    ,'{8}'::INTEGER AS uid
                    ,'INSERT'::TEXT AS action
                ", score.RefStudentID, score.RefCourseID, score.RefTeacherID, score.Term, score.Subject,score.Assessment,"",score.Value, 0);  // insert 給 uid = 0

                dataList.Add(data);
            }

            string Data = string.Join(" UNION ALL", dataList);


            string sql = string.Format(@"
WITH score_data_row AS(			 
                {0}     
),update_score AS(	    
    Update $esl.gradebook_assessment_score
    SET
        ref_student_id = score_data_row.ref_student_id
        ,ref_course_id = score_data_row.ref_course_id
        ,ref_teacher_id = score_data_row.ref_teacher_id
        ,term = score_data_row.term
        ,subject = score_data_row.subject
        ,assessment = score_data_row.assessment
        ,custom_assessment = score_data_row.custom_assessment
        ,value = score_data_row.value
    FROM 
        score_data_row    
    WHERE $esl.gradebook_assessment_score.uid = score_data_row.uid  
        AND score_data_row.action ='UPDATE'
    RETURNING  $esl.gradebook_assessment_score.* 
)
INSERT INTO $esl.gradebook_assessment_score(
	ref_student_id	
	,ref_course_id
	,ref_teacher_id
	,term
	,subject
    ,assessment
    ,custom_assessment
	,value
)
SELECT 
	score_data_row.ref_student_id::BIGINT AS ref_student_id	
	,score_data_row.ref_course_id::BIGINT AS ref_course_id	
	,score_data_row.ref_teacher_id::BIGINT AS ref_teacher_id	
	,score_data_row.term::TEXT AS term	
	,score_data_row.subject::TEXT AS subject	
    ,score_data_row.assessment::TEXT AS assessment	
    ,score_data_row.custom_assessment::TEXT AS custom_assessment
	,score_data_row.value::TEXT AS value	
FROM
	score_data_row
WHERE action ='INSERT'", Data);



            UpdateHelper uh = new UpdateHelper();

            _worker.ReportProgress(90, "上傳成績...");

            //執行sql
            uh.Execute(sql);

            _worker.ReportProgress(100, "ESL 評量成績上傳完成。");



        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // 繼續 畫面控制項
            cboSubject.ResumeLayout();
            cboAssessment.ResumeLayout();
            dataGridViewX1.ResumeLayout();            
            buttonX1.Enabled = true;

            MsgBox.Show("上傳完成!");
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("", e.ProgressPercentage);
        }

        private void dataGridViewX1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {


            DataGridViewCell cell = dataGridViewX1.Rows[e.RowIndex].Cells[e.ColumnIndex];

            cell.ErrorText = String.Empty;

            if (_targetScoreType == "Score")
            {
                int i = 0;

                if (!int.TryParse("" + e.FormattedValue, out i))
                {
                    cell.ErrorText = "請輸入數值。";

                }
            }
            if (_targetScoreType == "Indicator")
            {
                if (_targetIndicatorList.Find(indicator => indicator.Name == "" + e.FormattedValue) ==null)
                {
                    List<string> indicatorList = new List<string>();

                    string indicators = "";

                    foreach (Indicators indicator in _targetIndicatorList)
                    {
                        indicatorList.Add(indicator.Name);
                    }

                    indicators = string.Join("、", indicatorList);

                    cell.ErrorText = "請輸入" + indicators + "之一的文字";
                }


            }
            if (_targetScoreType == "Comment")
            {
                if (e.FormattedValue.ToString().Length > int.Parse(_commentLimit))
                {
                    cell.ErrorText = "請輸入小於" + _commentLimit + "字數的評語";
                }
            }


        }

        private void dataGridViewX1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                //只檢查分數欄，有錯誤就不給存
                DataGridViewCell cell = dataGridViewX1.Rows[row.Index].Cells[5];

                if (cell.ErrorText != String.Empty)
                {
                    buttonX1.Enabled = false;
                }
                else
                {
                    buttonX1.Enabled = true;
                }

            }

        }
    }
}
