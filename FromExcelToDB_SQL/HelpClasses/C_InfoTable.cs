using System;
using System.Collections.Generic;
using System.Text;

namespace MetalBase
{
    class C_InfoTable
    {
        internal int StartCol { get; set; }
        internal int StartRow { get; set; }
        internal List<int> ListExcelIndexTab { get; set; }
        internal List<int> ListdtProductIndexRow { get; set; }

        //public int EndExcelRowFact;
        internal string Name { get; set; }
        internal string Type { get; set; }
        internal string Standart { get; set; }
        internal string Mark { get; set; }
        internal int LastRowExcel { get; set; }

        //public List<int> listPrices;//если нужно хранить несколько номеров столбцов с ценами
        //public List<int> listMarks;//если нужно хранить несколько номеров столбцов с марками
    }
}
