using ESL_System.Model;
using ESL_System.UDT;
using FISCA.UDT;
using JHSchool.Data;
using K12.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ESL_System.Service
{
    /// <summary>
    ///  撈取資料 整理資料用
    /// </summary>
    class DataService
    {

        List<SubjectInfoForGAPLevel> SubGPAMapping;
        List<ScoreGPAMapping> ScoreGPAMapping;

       public DataService()
        {
            AccessHelper accessHelper = new AccessHelper();
            SubGPAMapping = accessHelper.Select<SubjectInfoForGAPLevel>();
            ScoreGPAMapping = accessHelper.Select<ScoreGPAMapping>();
        }

        /// <summary>
        /// 取得學生學期科目文字描述
        /// </summary>
        /// <param name="studentIDs"></param>
        /// <param name="schoolYear"></param>
        /// <param name="semester"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, string>> GetSemsSubjText(List<string> studentIDs, int schoolYear, int semester)
        {
            Dictionary<string, Dictionary<string, string>> dicSubjTextInfos = new Dictionary<string, Dictionary<string, string>>();

            List<JHSemesterScoreRecord> ListJHSemesterScoreRecord = JHSemesterScore.SelectByStudentIDs(studentIDs);


            foreach (JHSemesterScoreRecord scorerecord in ListJHSemesterScoreRecord)
            {
                Dictionary<string, SubjectScore> subjDic = scorerecord.Subjects;  // 取得科目資料s

                Dictionary<string, SubjectScore> subjDicf = subjDic.Where(x => x.Value.SchoolYear == schoolYear && x.Value.Semester == semester).ToDictionary(x => x.Key, x => x.Value); ;


                foreach (string subjName in subjDicf.Keys)
                {

                    SubjectScore subjectScore = subjDic[subjName];


                    if (!dicSubjTextInfos.ContainsKey(scorerecord.RefStudentID))
                    {
                        dicSubjTextInfos.Add(scorerecord.RefStudentID, new Dictionary<string, string>());

                    }

                    if (!dicSubjTextInfos[scorerecord.RefStudentID].ContainsKey(subjName))
                    {

                        dicSubjTextInfos[scorerecord.RefStudentID].Add(subjName, subjectScore.Text);
                    }
                }
            }
            return dicSubjTextInfos;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public decimal? GetFinalGPA(string subject, decimal score)
        {
            bool Iscontain = this.SubGPAMapping.Any(x => x.Subject == subject);
            if (Iscontain)
            {
                SubjectInfoForGAPLevel maping = this.SubGPAMapping.Where(x => x.Subject == subject).First();

                if (maping.IsAP)
                {
                    return GetAP(score);
                }
                else if (maping.IsHoner)
                {
                    return GetHoner(score);
                }
                else  // standandard
                {
                    return GetStandar(score);
                }
            }
            else
            {

                return GetStandar(score);

            }

        }



        public decimal? GetHoner(decimal score)
        {
            foreach (ScoreGPAMapping scoreGPAMapping in this.ScoreGPAMapping)
            {
                if (score >= scoreGPAMapping.MinScore && score < scoreGPAMapping.MaxScore)
                {

                    return scoreGPAMapping.Honers;

                }
            }
            return null;

        }



        public decimal? GetAP(decimal score)
        {
            foreach (ScoreGPAMapping scoreGPAMapping in this.ScoreGPAMapping)
            {
                if (score >= scoreGPAMapping.MinScore && score < scoreGPAMapping.MaxScore)
                {

                    return scoreGPAMapping.AP;

                }
            }
            return null;

        }


        public decimal? GetStandar(decimal score)
        {
            foreach (ScoreGPAMapping scoreGPAMapping in this.ScoreGPAMapping)
            {
                if (score >= scoreGPAMapping.MinScore && score < scoreGPAMapping.MaxScore)
                {
                    return scoreGPAMapping.GPA;

                }
            }
            return null;

        }


    }
}
