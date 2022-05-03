using System;
using System.Collections.Generic;
using System.Text;

namespace MetalBase
{
    class InfoOrganization
    {
        internal string OrgName { get; set; }
        internal string OrgAdress { get; set; }
        internal List<string> SkladAdr {get; set;}
        /// <summary>
        /// массив 2х или 3х строк, первая - имя, вторая - телефон, третья - Email
        /// </summary>
        internal List<string[]> Manager { get; set; }
        internal string OrgTel { get; set; }
        internal string Email { get; set; }
        internal string Site { get; set; }
        internal string Inn_Kpp { get; set; }
        internal string r_s { get; set; }
        internal string k_s { get; set; }
        internal string BIK { get; set; }


    }
}
