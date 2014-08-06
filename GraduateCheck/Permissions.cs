using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GraduateCheck
{
    class Permissions
    {
        public static string GraduateCheck { get { return "GraduateCheck.D59FD290-075B-43ED-A126-D0439D5D94A7"; } }

        public static bool GraduateCheck權限
        {
            get { return FISCA.Permission.UserAcl.Current[GraduateCheck].Executable; }
        }
    }
}
