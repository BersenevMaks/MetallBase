using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MetalBase.HelpClasses
{
    class C_Lists
    {
        private List<string> get_TolshList()
        {
            List<string> tolshes = new List<string>() { "2","3","4","5","7","8","10","12","14","16","18","20","22","26","30","32","36","40","45","50","55","60","65","70","80","90","100","110","120","130","140","150","160","180" };
            return tolshes;
        }

        public List<string> GetTolshes()
        {
            return get_TolshList();
        }
    }
}
