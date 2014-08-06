using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GraduateCheck
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void main()
        {
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["學生", "教務"];
            item1["畢業作業"]["畢業資格檢查"].Enable = false;
            item1["畢業作業"]["畢業資格檢查"].Click += delegate
            {
                new Checker().ShowDialog();
            };

            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                item1["畢業作業"]["畢業資格檢查"].Enable = K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0 && Permissions.GraduateCheck權限;
            };

            Catalog permission = RoleAclSource.Instance["學生"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.GraduateCheck, "畢業資格檢查"));
        }
    }
}
