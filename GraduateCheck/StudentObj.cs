using K12.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GraduateCheck
{
    class StudentObj
    {
        public StudentRecord Student;
        public ClassRecord Class;
        public Dictionary<string, SemesterScoreRecord> GradeScores;
        public const decimal _Greater = 0.5m;
        public const decimal _Less = 0.25m;
        public const string _Elective = "Elective";
        public static List<string> _Domain = new List<string>() { "Chinese", "English", "Mathematics", "Science", "Social Studies", "Physical Education", "Elective" };

        public StudentObj(StudentRecord student, ClassRecord cr)
        {
            Student = student;
            Class = cr;

            GradeScores = new Dictionary<string, SemesterScoreRecord>() {
                { "9上", null }, { "9下", null } ,{ "10上", null }, { "10下", null } ,{ "11上", null }, { "11下", null } ,{ "12上", null }, { "12下", null } 
            };
        }

        public void LoadRecord(SemesterScoreRecord record, string grade)
        {
            if (!GradeScores.ContainsKey(grade))
                return;

            if (GradeScores[grade] == null)
                GradeScores[grade] = record;

            //如果有相同grade的record,以較新的為主
            if (record.SchoolYear > GradeScores[grade].SchoolYear)
                GradeScores[grade] = record;
            else if (record.SchoolYear == GradeScores[grade].SchoolYear)
            {
                if (record.Semester > GradeScores[grade].Semester)
                    GradeScores[grade] = record;
            }
        }

        public decimal GetDomainCredit(string domain)
        {
            decimal retVal = 0;

            List<string> already_learned = new List<string>();
            List<string> sibling_learned = new List<string>();

            foreach (string grade in GradeScores.Keys)
            {
                SemesterScoreRecord ssr = GradeScores[grade];
                if (ssr == null)
                    continue;

                decimal max = 0;

                foreach (SubjectScore subject in ssr.Subjects.Values)
                {
                    decimal temp = 0;

                    //判斷得分
                    if (subject.Domain == domain && subject.Score.HasValue && subject.Score.Value >= 60 && subject.Period.HasValue && subject.Period.Value >= 4)
                        temp = _Greater;
                    else if (subject.Domain == domain && subject.Score.HasValue && subject.Score.Value >= 60 && subject.Period.HasValue && subject.Period.Value >= 2)
                        temp = _Less;

                    //保留此學期該領域的最高分
                    if (temp > max)
                        max = temp;

                    //選修課程中節數未達4節,須修滿上下學期
                    if (domain == _Elective)
                    {
                        //已修習過的選修科目不再給分
                        if (temp == _Greater)
                        {
                            if (!already_learned.Contains(subject.Subject))
                                already_learned.Add(subject.Subject);
                            else
                                temp = 0;
                        }
                        else if (temp == _Less)
                        {
                            temp = 0;
                            string source_key = grade + "_" + subject.Subject;
                            string target_key = GetSiblingGrade(grade) + "_" + subject.Subject;

                            if (sibling_learned.Contains(target_key) && !sibling_learned.Contains(source_key))
                            {
                                if (!already_learned.Contains(subject.Subject))
                                {
                                    already_learned.Add(subject.Subject);
                                    temp = _Greater;
                                }
                            }
                                
                            if(!sibling_learned.Contains(source_key))
                                sibling_learned.Add(source_key);
                        }

                        retVal += temp;
                    }
                }

                //非選修科目就加最高分
                if (domain != _Elective)
                    retVal += max;
            }

            return retVal;
        }

        public string CheckSems()
        {
            List<string> sems = new List<string>();
            foreach (string grade in GradeScores.Keys)
            {
                if (GradeScores[grade] == null)
                    sems.Add(grade);
            }

            if (sems.Count > 0)
                return "學期成績缺漏:" + string.Join(",", sems);
            else
                return string.Empty;
        }

        public string CheckCredit()
        {
            string retVal = "";

            decimal Chinese = 0;
            decimal English = 0;
            decimal SS = 0;
            decimal Math = 0;
            decimal Science_ = 0;
            decimal PE = 0;
            decimal Elective = 0;

            Chinese = GetDomainCredit("Chinese");
            if (Chinese < 4)
                retVal += "Chinese不足4分(" + Chinese + ")";

            English = GetDomainCredit("English");
            if (English < 4)
                retVal += "English不足4分(" + English + ")";

            SS = GetDomainCredit("Social Studies");
            if (SS < 3)
                retVal += "Social Studies不足3分(" + SS + ")";

            Math = GetDomainCredit("Mathematics");
            if (Math < 3)
                retVal += "Mathematics不足3分(" + Math + ")";

            Science_ = GetDomainCredit("Science");
            if (Science_ < 3)
                retVal += "Science不足3分(" + Science_ + ")";

            PE = GetDomainCredit("Physical Education");
            if (PE < 2)
                retVal += "Physical Education不足2分(" + PE + ")";

            Elective = GetDomainCredit("Elective");
            if (Elective < 4)
                retVal += "Elective不足4分(" + Elective + ")";

            return retVal;
        }

        public string CheckArt(List<string> arts)
        {
            bool HasArtSubject = false;
            foreach (string grade in GradeScores.Keys)
            {
                if (HasArtSubject)
                    break;

                SemesterScoreRecord ssr = GradeScores[grade];
                if (ssr == null)
                    continue;

                foreach (SubjectScore subject in ssr.Subjects.Values)
                {
                    if (arts.Contains(subject.Subject) && subject.Score.HasValue && subject.Score.Value >= 60)
                    {
                        HasArtSubject = true;
                        break;
                    }
                }
            }

            if (!HasArtSubject)
                return "未修習過美術課程";
            else
                return string.Empty;
        }

        //取得相關年級判斷
        public static string GetSiblingGrade(string grade)
        {
            switch (grade)
            {
                case "9上":
                    return "9下";
                case "9下":
                    return "9上";
                case "10上":
                    return "10下";
                case "10下":
                    return "10上";
                case "11上":
                    return "11下";
                case "11下":
                    return "11上";
                case "12上":
                    return "12下";
                case "12下":
                    return "12上";
                default:
                    return string.Empty;
            }
        }
    }
}
