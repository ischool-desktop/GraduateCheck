using Aspose.Cells;
using FISCA.Data;
using FISCA.Presentation.Controls;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GraduateCheck
{
    public partial class Checker : BaseForm
    {
        private List<string> _ids;
        private List<string> _ArtSubjects;
        private BackgroundWorker _BW;
        private const string _ConfigCode = "GraduateCheckArts";
        Campus.Configuration.ConfigData _CD;

        public Checker()
        {
            InitializeComponent();

            _ids = K12.Presentation.NLDPanels.Student.SelectedSource;

            _ArtSubjects = new List<string>();

            _BW = new BackgroundWorker();
            _BW.DoWork += new DoWorkEventHandler(BW_DoWork);
            _BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BW_Completed);

            _CD = Campus.Configuration.Config.User[_ConfigCode];

            foreach(string art in _CD["美學清單"].Split(','))
            {
                if(!string.IsNullOrWhiteSpace(art))
                    dgvArts.Rows.Add(art);
            }
        }

        private void BW_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("畢業判斷完成");
            FormLock(true);
            Workbook wb = e.Result as Workbook;
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "另存新檔";
            save.FileName = "未達畢業標準名單.xls";
            save.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";

            if (save.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    wb.Save(save.FileName, Aspose.Cells.SaveFormat.Excel97To2003);
                    System.Diagnostics.Process.Start(save.FileName);
                }
                catch
                {
                    MessageBox.Show("檔案儲存失敗");
                }
            }
        }

        private void BW_DoWork(object sender, DoWorkEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("畢業判斷計算中...");
            Dictionary<string, StudentObj> student_obj_dic = new Dictionary<string, StudentObj>();

            //班級catch
            Dictionary<string, ClassRecord> class_catch = new Dictionary<string, ClassRecord>();
            foreach (ClassRecord cr in K12.Data.Class.SelectAll())
            {
                if (!class_catch.ContainsKey(cr.ID))
                    class_catch.Add(cr.ID, cr);
            }

            //取得學生資料
            foreach (StudentRecord student in K12.Data.Student.SelectByIDs(_ids))
            {
                ClassRecord cr = class_catch.ContainsKey(student.RefClassID) ? class_catch[student.RefClassID] : new ClassRecord();

                if (!student_obj_dic.ContainsKey(student.ID))
                    student_obj_dic.Add(student.ID, new StudentObj(student, cr));
            }

            //測試用(複製學期歷程)
            //QueryHelper q = new QueryHelper();
            //UpdateHelper u = new UpdateHelper();

            //DataTable dt = q.Select("select sems_history from student where id=3782");
            //string xml = dt.Rows[0]["sems_history"] + "";

            //string ids = string.Join(",",_ids);
            //u.Execute("update student set sems_history='" + xml + "' where id in (" + ids + ")");

            //學期歷程快照
            Dictionary<string, string> sems_history = new Dictionary<string, string>();
            foreach (SemesterHistoryRecord shr in K12.Data.SemesterHistory.SelectByStudentIDs(_ids))
            {
                foreach (SemesterHistoryItem item in shr.SemesterHistoryItems)
                {
                    string key = item.RefStudentID + "_" + item.SchoolYear + "_" + item.Semester;

                    if (!sems_history.ContainsKey(key))
                    {
                        if (item.Semester == 1)
                            sems_history.Add(key, item.GradeYear + "上");
                        else
                            sems_history.Add(key, item.GradeYear + "下");
                    }
                }
            }

            //整理符合範圍的學期成績
            foreach (SemesterScoreRecord r in K12.Data.SemesterScore.SelectByStudentIDs(_ids))
            {
                string key = r.RefStudentID + "_" + r.SchoolYear + "_" + r.Semester;

                if (sems_history.ContainsKey(key))
                {
                    string grade = sems_history[key];
                    if (grade == "9上" || grade == "9下" || grade == "10上" || grade == "10下" || grade == "11上" || grade == "11下" || grade == "12上" || grade == "12下")
                    {
                        student_obj_dic[r.RefStudentID].LoadRecord(r, grade);
                    }
                }
            }

            List<StudentObj> objs = student_obj_dic.Values.ToList();
            objs.Sort(delegate(StudentObj x, StudentObj y)
            {
                string xx = x.Class.Name.PadLeft(20, '0');
                xx += (x.Student.SeatNo + "").PadLeft(3, '0');

                string yy = y.Class.Name.PadLeft(20, '0');
                yy += (y.Student.SeatNo + "").PadLeft(3, '0');

                return xx.CompareTo(yy);
            });

            Workbook wb = new Workbook();
            wb.Worksheets[0].Name = "未達畢業標準名單";
            Cells cs = wb.Worksheets[0].Cells;

            //Print Title
            cs[0, 0].PutValue("班級");
            cs[0, 1].PutValue("座號");
            cs[0, 2].PutValue("姓名");
            cs[0, 3].PutValue("學號");
            cs[0, 4].PutValue("未達畢業標準原因");

            int index = 1;
            foreach (StudentObj obj in objs)
            {
                List<string> strs = new List<string>();

                string reason1 = obj.CheckSems();
                string reason2 = obj.CheckArt(_ArtSubjects);
                string reason3 = obj.CheckCredit();

                if (!string.IsNullOrWhiteSpace(reason1))
                    strs.Add(reason1);
                if (!string.IsNullOrWhiteSpace(reason2))
                    strs.Add(reason2);
                if (!string.IsNullOrWhiteSpace(reason3))
                    strs.Add(reason3);

                string all_reason = string.Join(";", strs);

                if (string.IsNullOrWhiteSpace(all_reason))
                    continue;

                cs[index, 0].PutValue(obj.Class.Name);
                cs[index, 1].PutValue(obj.Student.SeatNo + "");
                cs[index, 2].PutValue(obj.Student.Name + "" + obj.Student.EnglishName);
                cs[index, 3].PutValue(obj.Student.StudentNumber);
                cs[index, 4].PutValue(all_reason);
                index++;
            }

            wb.Worksheets[0].AutoFitColumns();

            e.Result = wb;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            CD_Save();
            FormLock(false);

            if (_BW.IsBusy)
                MessageBox.Show("系統忙碌中,請稍後再試...");
            else
                _BW.RunWorkerAsync();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FormLock(bool b)
        {
            dgvArts.Enabled = b;
        }

        private void CD_Save()
        {
            List<string> arts = new List<string>();
            foreach (DataGridViewRow row in dgvArts.Rows)
            {
                string subj = row.Cells[0].Value + "";
                //不給打逗號
                subj = subj.Replace(",", "");

                if (row.IsNewRow || string.IsNullOrWhiteSpace(subj))
                    continue;

                if (!arts.Contains(subj))
                    arts.Add(subj);

                if (!_ArtSubjects.Contains(subj))
                    _ArtSubjects.Add(subj);
            }

            string art = string.Join(",", arts);

            _CD["美學清單"] = art;
            _CD.Save();
        }
    }
}
