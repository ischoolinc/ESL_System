using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Data;
using DevComponents.DotNetBar;
using System.Xml.Linq;
using K12.Data;
using System.Xml;
using System.Data;
using FISCA.Presentation.Controls;
using K12.Data;

namespace ESL_System
{
    // 2018/6/12 穎驊新增 用來結算 ESL Term 、Subject 成績用
    class CalculateTermScore_Old
    {
        private string target_exam_id; //目標試別id

        private List<string> _courseIDList;
        private List<ESLCourse> _ESLCourseList = new List<ESLCourse>();

        private List<string> _TargetCourseExam = new List<string>(); //本次須計算的目標課程試別 <courseID + _ + ExamName>


        private Dictionary<string, List<ESLScore>> _scoreDict = new Dictionary<string, List<ESLScore>>(); // 取得原本成績階層結構(Assessment)
        private Dictionary<string, List<ESLScore>> _scoreUpsertDict = new Dictionary<string, List<ESLScore>>(); // 計算後要上傳、新增的成績放在這

        private Dictionary<string, List<Term>> _scoreTemplateDict = new Dictionary<string, List<Term>>(); // 各課程的分數計算規則 

        private Dictionary<string, decimal> _scoreRatioDict = new Dictionary<string, decimal>(); // 各課程的分數比例權重分子

        private Dictionary<string, decimal> _scoreRatioTotalDict = new Dictionary<string, decimal>(); // 各課程的分數比例權重分母

        private Dictionary<string, ESLScore> _subjectScoreDict = new Dictionary<string, ESLScore>(); // 計算用的科目成績字典
        private Dictionary<string, ESLScore> _termScoreDict = new Dictionary<string, ESLScore>(); // 計算用的評量成績字典

        private List<ESLScore> _eslscoreList = new List<ESLScore>(); // 先暫時這樣儲存 上傳用，之後會想改用scoreUpsertDict

        public CalculateTermScore_Old(List<string> courseIDList, string exam_id)
        {
            _courseIDList = courseIDList;
            target_exam_id = exam_id;
        }

        public void CalculateESLTermScore()
        {
            string courseIDs = string.Join(",", _courseIDList);


            // 2018/06/12 抓取課程且其有ESL 樣板設定規則的，才做後續整理，  在table exam_template 欄位 description 不為空代表其為ESL 的樣板
            string query = @"
SELECT 
    course.id
    ,course.course_name
    ,exam_template.description 
FROM course 
LEFT JOIN  exam_template ON course.ref_exam_template_id =exam_template.id  
WHERE course.id IN( " + courseIDs + ") AND  exam_template.description IS NOT NULL  ";



            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(query);

            _courseIDList.Clear(); // 清空

            //整理目前的ESL 課程資料
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    ESLCourse record = new ESLCourse();

                    _courseIDList.Add("" + dr[0]); // 加入真正的 是ESL 課程ID

                    record.ID = "" + dr[0]; //課程ID
                    record.CourseName = "" + dr[1]; //課程名稱
                    record.Description = "" + dr[2]; // ESL 評分樣版設定

                    _ESLCourseList.Add(record);
                }
            }

            // 解析計算規則
            foreach (ESLCourse course in _ESLCourseList)
            {
                string xmlStr = "<root>" + course.Description + "</root>";
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

                            if (target_exam_id == t.Ref_exam_id)
                            {
                                _TargetCourseExam.Add(course.ID + "_" + t.Name);
                            }

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

                            if (!_scoreTemplateDict.ContainsKey(course.ID))
                            {
                                _scoreTemplateDict.Add(course.ID, new List<Term>());

                                _scoreTemplateDict[course.ID].Add(t);
                            }
                            else
                            {
                                _scoreTemplateDict[course.ID].Add(t);
                            }
                        }
                    }
                }
            }



            string eslCourseIDs = string.Join(",", _courseIDList);// 把真正是ESL 課程的ID 列出來

            query = "SELECT* FROM $esl.gradebook_assessment_score WHERE ref_course_id in( " + eslCourseIDs + ")"; // 抓取這些ESL 的所有成績資料
            dt = qh.Select(query);

            // 整理 既有的成績資料 (Assessment)
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    // 算是整理成績資料最重要的一行， assessment 欄位 為空，還有濾掉 custom_assessment 成績 代表此成績 是 Subject 或是 Term 成績
                    if ("" + dr["assessment"] == "" || "" + dr["custom_assessment"] != "")
                    {
                        continue;
                    }

                    // 假如本成績資料 非在本次所選的評量，也將其拿掉
                    if (!_TargetCourseExam.Contains("" + dr["ref_course_id"] + "_" + "" + dr["term"]))
                    {
                        continue;
                    }

                    if (!_scoreDict.ContainsKey("" + dr["ref_course_id"]))
                    {
                        ESLScore score = new ESLScore();

                        score.ID = "" + dr["uid"];
                        score.RefCourseID = "" + dr["ref_course_id"];
                        score.RefStudentID = "" + dr["ref_student_id"];
                        score.RefTeacherID = "" + dr["ref_teacher_id"];
                        score.Term = "" + dr["term"];
                        score.Subject = "" + dr["subject"];
                        score.Assessment = "" + dr["assessment"];
                        score.Custom_Assessment = "" + dr["custom_assessment"];
                        score.Value = "" + dr["value"];

                        _scoreDict.Add("" + dr["ref_course_id"], new List<ESLScore>());

                        _scoreDict["" + dr["ref_course_id"]].Add(score);
                    }
                    else
                    {
                        ESLScore score = new ESLScore();

                        score.ID = "" + dr["uid"];
                        score.RefCourseID = "" + dr["ref_course_id"];
                        score.RefStudentID = "" + dr["ref_student_id"];
                        score.RefTeacherID = "" + dr["ref_teacher_id"];
                        score.Term = "" + dr["term"];
                        score.Subject = "" + dr["subject"];
                        score.Assessment = "" + dr["assessment"];
                        score.Custom_Assessment = "" + dr["custom_assessment"];
                        score.Value = "" + dr["value"];

                        _scoreDict["" + dr["ref_course_id"]].Add(score);
                    }
                }
            }



            //計算成績 比例加減
            foreach (KeyValuePair<string, List<Term>> p in _scoreTemplateDict)
            {
                foreach (Term t in p.Value)
                {
                    t.SubjectTotalWeight = 0;

                    foreach (Subject s in t.SubjectList)
                    {
                        //t.SubjectTotalWeight += decimal.Parse(s.Weight); // 加總比例 之後才可以換算分別比例

                        s.AssessmentTotalWeight = 0;

                        foreach (Assessment a in s.AssessmentList)
                        {
                            if (a.Type == "Score") // 只取分數型成績
                            {
                                //s.AssessmentTotalWeight += decimal.Parse(a.Weight); // 加總比例 之後才可以換算分別比例
                            }
                        }
                    }
                }
            }

            foreach (KeyValuePair<string, List<Term>> p in _scoreTemplateDict)
            {
                string key_subject = "";
                string key_assessment = "";

                foreach (Term t in p.Value)
                {
                    foreach (Subject s in t.SubjectList)
                    {
                        foreach (Assessment a in s.AssessmentList)
                        {
                            if (a.Type == "Score") // 只取分數型成績
                            {
                                key_assessment = p.Key + "_" + t.Name + "_" + s.Name + "_" + a.Name;

                                decimal ratio_assessment = decimal.Parse(a.Weight);

                                _scoreRatioDict.Add(key_assessment, ratio_assessment);
                            }
                        }
                        key_subject = p.Key + "_" + t.Name + "_" + s.Name;

                        decimal ratio = decimal.Parse(s.Weight);

                        _scoreRatioDict.Add(key_subject, ratio);
                    }
                }
            }

            //計算成績 依照比例 等量換算成分數 儲存 在 scoreUpsertDict (Subject 成績)
            foreach (KeyValuePair<string, List<ESLScore>> p in _scoreDict)
            {
                string key_score_assessment = "";

                string key_subject = "";
                string key_term = "";

                decimal subject_score_partial;

                foreach (ESLScore score in p.Value)
                {
                    key_score_assessment = p.Key + "_" + score.Term + "_" + score.Subject + "_" + score.Assessment; // 查分數比例的KEY 這些就夠

                    key_subject = p.Key + "_" + score.RefStudentID + "_" + score.Term + "_" + score.Subject; // 寫給學生的Subject 成績 還必須要有 studentID 才有獨立性

                    key_term = p.Key + "_" + score.RefStudentID + "_" + score.Term; // 寫給學生的Term 成績 還必須要有 studentID 才有獨立性

                    if (_scoreRatioDict.ContainsKey(key_score_assessment))
                    {
                        decimal assementScore;
                        if (decimal.TryParse(score.Value, out assementScore))
                        {
                            subject_score_partial = Math.Round(assementScore * _scoreRatioDict[key_score_assessment], 2, MidpointRounding.ToEven); // 四捨五入到第二位

                            // 處理 subject分母
                            if (!_scoreRatioTotalDict.ContainsKey(key_subject))
                            {
                                _scoreRatioTotalDict.Add(key_subject, _scoreRatioDict[key_score_assessment]);
                            }
                            else
                            {
                                _scoreRatioTotalDict[key_subject] += _scoreRatioDict[key_score_assessment];
                            }


                            if (!_subjectScoreDict.ContainsKey(key_subject))
                            {
                                ESLScore subjectScore = new ESLScore();

                                subjectScore.RefCourseID = score.RefCourseID;
                                subjectScore.RefStudentID = score.RefStudentID;
                                subjectScore.RefTeacherID = score.RefTeacherID;
                                subjectScore.Term = score.Term;
                                subjectScore.Subject = score.Subject;
                                subjectScore.Score = subject_score_partial;

                                _subjectScoreDict.Add(key_subject, subjectScore);
                            }
                            else
                            {
                                _subjectScoreDict[key_subject].Score += subject_score_partial;
                            }
                        }
                        else
                        {
                            //assementScore = 0; // 轉失敗(可能沒有輸入)，當0 分
                        }
                    }
                }
            }


            // 計算Subject成績後，現在將各自加權後的成績除以各自的的總權重
            foreach (KeyValuePair<string, ESLScore> score in _subjectScoreDict)
            {
                string ratioTotalKey = score.Value.RefCourseID + "_" + score.Value.RefStudentID + "_" + score.Value.Term + "_" + score.Value.Subject;

                _subjectScoreDict[score.Key].Score = Math.Round(_subjectScoreDict[score.Key].Score / _scoreRatioTotalDict[ratioTotalKey],2,MidpointRounding.ToEven);

            }



            //計算成績 依照比例 等量換算成分數 儲存 在 scoreUpsertDict (term 成績)
            foreach (ESLScore subjectScore in _subjectScoreDict.Values)
            {
                string key_score_subject = subjectScore.RefCourseID + "_" + subjectScore.Term + "_" + subjectScore.Subject; // 查分數比例的KEY 這些就夠

                string key_term = "";

                decimal term_score_partial;

                if (_scoreRatioDict.ContainsKey(key_score_subject))
                {
                    key_term = subjectScore.RefCourseID + "_" + subjectScore.RefStudentID + "_" + subjectScore.Term; // 寫給學生的Term 成績 還必須要有 studentID 才有獨立性

                    term_score_partial = Math.Round(subjectScore.Score * _scoreRatioDict[key_score_subject], 2,MidpointRounding.ToEven); // 四捨五入到第二位


                    // 處理 term 分母
                    if (!_scoreRatioTotalDict.ContainsKey(key_term))
                    {
                        _scoreRatioTotalDict.Add(key_term, _scoreRatioDict[key_score_subject]);
                    }
                    else
                    {
                        _scoreRatioTotalDict[key_term] += _scoreRatioDict[key_score_subject];
                    }



                    if (!_termScoreDict.ContainsKey(key_term))
                    {
                        ESLScore termScore = new ESLScore();

                        termScore.RefCourseID = subjectScore.RefCourseID;
                        termScore.RefStudentID = subjectScore.RefStudentID;
                        termScore.RefTeacherID = subjectScore.RefTeacherID;
                        termScore.Term = subjectScore.Term;
                        termScore.Score = term_score_partial;

                        _termScoreDict.Add(key_term, termScore);
                    }
                    else
                    {
                        _termScoreDict[key_term].Score += term_score_partial;
                    }
                }
            }

            // 計算Term成績後，現在將各自加權後的成績除以各自的的總權重
            foreach (KeyValuePair<string, ESLScore> score in _termScoreDict)
            {
                string ratioTotalKey = score.Value.RefCourseID + "_" + score.Value.RefStudentID + "_" + score.Value.Term;

                _termScoreDict[score.Key].Score = Math.Round(_termScoreDict[score.Key].Score / _scoreRatioTotalDict[ratioTotalKey],2,MidpointRounding.ToEven);

            }



            //2018/6/14 先暫時這樣整理，之後會想要用 scoreUpsertDict ，資料會比較齊
            foreach (ESLScore score in _subjectScoreDict.Values)
            {
                _eslscoreList.Add(score);
            }

            foreach (ESLScore score in _termScoreDict.Values)
            {
                _eslscoreList.Add(score);
            }


            // 取的學生修課紀錄
            List<K12.Data.SCAttendRecord> scaList = SCAttend.SelectByCourseIDs(_courseIDList);

            // 取得舊有評量成績
            List<SCETakeRecord> scetakeList_Old = SCETake.SelectByCourseAndExam(_courseIDList, target_exam_id);

            List<SCETakeRecord> updateList = new List<SCETakeRecord>();
            List<SCETakeRecord> insertList = new List<SCETakeRecord>();


            // 更新舊評量分數
            foreach (SCETakeRecord scet in scetakeList_Old)
            {
                foreach (ESLScore score in _termScoreDict.Values)
                {
                    if (scet.RefCourseID == score.RefCourseID && scet.RefStudentID == score.RefStudentID && GetScore(scet) != "" + score.Score)
                    {
                        SetScore(scet, "" + score.Score);

                        updateList.Add(scet);
                    }
                }
            }


            // 新增 評量分數
            foreach (SCAttendRecord sca in scaList)
            {
                // 沒在舊評量成績 就是本次要新增的項目
                if (!scetakeList_Old.Any(s => s.RefSCAttendID == sca.ID))
                {
                    foreach (ESLScore score in _termScoreDict.Values)
                    {
                        if (sca.RefCourseID == score.RefCourseID && sca.RefStudentID == score.RefStudentID)
                        {
                            SCETakeRecord sce = new SCETakeRecord();
                            sce.RefSCAttendID = sca.ID;
                            sce.RefExamID = target_exam_id;
                            sce.RefStudentID = sca.RefStudentID;
                            sce.RefCourseID = sca.RefCourseID;
                            SetScore(sce, "" + score.Score);
                            insertList.Add(sce);
                        }

                    }
                }

            }


            K12.Data.SCETake.Update(updateList);
            K12.Data.SCETake.Insert(insertList);




            //拚SQL
            // 兜資料
            List<string> dataList = new List<string>();


            // 沒有新增任何成績資料，代表所選ESL 課程都沒有成績，不需執行SQL
            if (_eslscoreList.Count == 0)
            {
                return;
            }


            foreach (ESLScore score in _eslscoreList)
            {
                string data = string.Format(@"
                SELECT
                    '{0}'::BIGINT AS ref_student_id
                    ,'{1}'::BIGINT AS ref_course_id
                    ,'{2}'::BIGINT AS ref_teacher_id
                    ,'{3}'::TEXT AS term
                    ,'{4}'::TEXT AS subject
                    ,'{5}'::TEXT AS value
                ", score.RefStudentID, score.RefCourseID, score.RefTeacherID, score.Term, score.Subject, score.Score);

                dataList.Add(data);
            }

            string Data = string.Join(" UNION ALL", dataList);


            string sql = string.Format(@"
WITH score_data_row AS(			 
                {0}     
),delete_score AS(
	DELETE
	FROM
		$esl.gradebook_assessment_score
	WHERE assessment IS NULL
)
INSERT INTO $esl.gradebook_assessment_score(
	ref_student_id	
	,ref_course_id
	,ref_teacher_id
	,term
	,subject
	,value
)
SELECT 
	score_data_row.ref_student_id::BIGINT AS ref_student_id	
	,score_data_row.ref_course_id::BIGINT AS ref_course_id	
	,score_data_row.ref_teacher_id::BIGINT AS ref_teacher_id	
	,score_data_row.term::TEXT AS term	
	,score_data_row.subject::TEXT AS subject	
	,score_data_row.value::TEXT AS value	
FROM
	score_data_row", Data);

            UpdateHelper uh = new UpdateHelper();

            //執行sql
            uh.Execute(sql);


            MsgBox.Show("計算完成!");




        }


        private string GetScore(SCETakeRecord sce)
        {
            XmlElement elem = sce.ToXML().SelectSingleNode("Extension/Extension/Score") as XmlElement;

            string score = elem == null ? string.Empty : elem.InnerText;

            return score;
        }


        private void SetScore(SCETakeRecord sce, string score)
        {
            XmlElement root = sce.ToXML();
            XmlElement elem = root.SelectSingleNode("Extension/Extension/Score") as XmlElement;

            decimal d;
            decimal.TryParse(score, out d);

            if (elem != null)
            {
                elem.InnerText = d + "";
            }

            sce.Load(root);
        }




    }



}
