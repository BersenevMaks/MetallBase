using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;

namespace MetalBase
{
    class Class_StroiTehCentr
    {
        private string filePath;

        public void Set(string Path)
        {
            filePath = Path;
        }

        public void GetTableFromExcel()
        {
            dtProduct.Columns.Add("Название");
            dtProduct.Columns.Add("Тип");
            dtProduct.Columns.Add("Диаметр (высота), мм");
            dtProduct.Columns.Add("Толщина (ширина), мм");
            dtProduct.Columns.Add("Метраж, м (длина, мм)");
            dtProduct.Columns.Add("Мерность (т, м, мм)");
            dtProduct.Columns.Add("Марка");
            dtProduct.Columns.Add("Стандарт");
            dtProduct.Columns.Add("Класс");
            dtProduct.Columns.Add("Цена");
            dtProduct.Columns.Add("Примечание");
            ReadExcel();
            //return dtProduct;
        }

        private Excel.Application excelapp;
        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;
        private Excel.Sheets excelsheets;
        bool isExcelOpen = false;
        DataTable dtProduct = new DataTable();

        string orgname = "";

        public string NameOrg() { return orgname; }

        private void ReadExcel()
        {
            InfoOrganization infoOrg = new InfoOrganization
            {
                SkladAdr = new List<string>(),
                Manager = new List<string[]>()
            };

            excelapp = new Excel.Application();
            try
            {
                orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?)").Match(Path.GetFileName(filePath)).Value;
                infoOrg.OrgName = orgname;
                excelappworkbooks = excelapp.Workbooks;
                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                    isExcelOpen = true;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла " + orgname + "\n\n" + ex.ToString()); isExcelOpen = false; }

                string temp = "", tmp = "", standart = "", mark = "", name = "", type = "", price = "", prim = "";
                string diam = "", tolsh = "", metraj = "", mera = "", skladPrim = "";
                var regexParam = new C_RegexParamProduct();
                double dDiam;
                double dTolsh;
                double dMetraj;

                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    //if (excelworksheet.Index != 9) continue;
                    //MessageBox.Show(excelsheets.Count.ToString());
                    var tab = new C_InfoTable();
                    var naaame = excelworksheet.Name;
                    List<C_InfoTable> tabs = new List<C_InfoTable>();
                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol <= 10) cCelCol = 10;
                    if (cCelCol > 10) cCelCol = 25;
                    //cCelCol = 10;

                    int ColName = 0, ColMera = 0, ColPrice = 0, ColGost = 0, ColSoder = 0, ColDiam = 0, ColTolsh = 0, ColMark = 0;

                    int Max = cCelCol * cCelRow;
                    SetMaxValProgressBar(Max);
                    //Поиск заголовков столбцов
                    int progress = 0;

                    for (int j = 4; j <= cCelRow; j++) //строки
                    {
                        int jj = j;
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange;
                            cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"^название", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    tab = new C_InfoTable
                                    {
                                        StartCol = i,
                                        StartRow = jj
                                    };
                                    tabs.Add(tab);
                                    j = cCelRow + 1;
                                    break;
                                }
                            }
                            if (progress < Max) ProcessChanged(progress++);
                            else ProcessChanged(Max);
                        }
                    }

                    ProcessChanged(0);
                    Max = tabs.Count;
                    SetMaxValProgressBar(Max);
                    progress = 0;
                    for (int k = 0; k < tabs.Count; k++)
                    {
                        ColPrice = 0; ColName = 0; ColMera = 0; ColSoder = 0; ColGost = 0;
                        name = ""; type = "";
                        Excel.Range cellRange;
                        tab = tabs[k];
                        int endRow = cCelRow;
                        if (k < tabs.Count - 1)   // определение последней строки в текущей минитаблице
                            if (tab.StartCol == tabs[k + 1].StartCol)
                                endRow = tabs[k + 1].StartRow - 1;
                            else if (k < tabs.Count - 2)
                                if (tab.StartCol == tabs[k + 2].StartCol)
                                    endRow = tabs[k + 2].StartRow - 1;
                                else if (k < tabs.Count - 3)
                                    if (tab.StartCol == tabs[k + 3].StartCol)
                                        endRow = tabs[k + 3].StartRow - 1;

                        // // // поиск имени продукции
                        ProcessChanged(0); //установить прогрессбар на 0
                        SetMaxValProgressBar(cCelCol); Max = cCelCol; //установить максимум для прогрессбара
                        progress = 0;
                        int jjj = tab.StartRow;
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            cellRange = (Excel.Range)excelworksheet.Cells[jjj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"\bназвание\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColName = i;
                                    continue;
                                }
                                if (new Regex(@"\bвес\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColMera = i;
                                    continue;
                                }
                                if (new Regex(@"\bцена\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColPrice = i;
                                    cellRange = (Excel.Range)excelworksheet.Cells[jjj + 1, ColName];
                                    if (cellRange.MergeArea.Columns.Count > 1)
                                    {
                                        for (int ii = cellRange.MergeArea.Column; i <= cellRange.MergeArea.Columns.Count - 1; ii++)
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[jjj, ii];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                if (new Regex(@"\bот\s*5.*до\s*10\b", RegexOptions.IgnoreCase).IsMatch(temp)) ColPrice = ii;
                                            }
                                        }
                                    }
                                }
                            }
                            if (progress < Max) ProcessChanged(progress++);
                            else ProcessChanged(Max);
                        }

                        ProcessChanged(0);//установить прогрессбар на 0
                        progress = 0;
                        SetMaxValProgressBar(endRow * (k + 1)); //установить максимум для прогрессбара
                        Max = endRow * (k + 1);
                        for (int j = tab.StartRow + 1; j <= endRow; j++)
                        {
                            diam = ""; tolsh = ""; metraj = ""; mera = ""; standart = ""; mark = ""; dTolsh = 0; dMetraj = 0;
                            if (ColName > 0)
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[j, ColName];
                                if (cellRange.MergeArea.Columns.Count > 1)
                                {
                                    cellRange = (Excel.Range)excelworksheet.Cells[j, cellRange.MergeArea.Column];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        name = "";
                                        type = "";
                                        if (regexParam.RegName.IsMatch(temp))
                                        {
                                            name = StringFirstUp(regexParam.RegName.Match(temp).Value);
                                            if (new Regex(@"оцин|эл.св|рыж|г/к|х/к|х/д|б/ш|рифл|проф", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                type = new Regex(@"оцин|эл.св|рыж|г/к|х/к|х/д|б/ш|рифл|проф", RegexOptions.IgnoreCase).Match(temp).Value.ToLower();
                                                switch (type)
                                                {
                                                    case "оцин":
                                                        type = "оцинкованная";
                                                        break;
                                                    case "эл.св":
                                                        type = "электросварная";
                                                        break;
                                                    case "рыж":
                                                        type = "рыжая";
                                                        break;
                                                    case "г/к":
                                                        type = "горячекатанная";
                                                        break;
                                                    case "х/к":
                                                        type = "холоднокатанная";
                                                        break;
                                                    case "х/д":
                                                        type = "холоднодеформированная";
                                                        break;
                                                    case "б/ш":
                                                        type = "бесшовная";
                                                        break;
                                                    case "рифл":
                                                        type = "рифленая";
                                                        break;
                                                    case "проф":
                                                        type = "профильная";
                                                        break;
                                                    default:
                                                        type = "";
                                                        break;
                                                }
                                                if (!string.IsNullOrEmpty(type) && name.ToLower().Contains("лист")) type = type.Replace("ая", "ый");
                                            }
                                            else if (string.IsNullOrEmpty(type))
                                                type = regexParam.RegType.Match(temp).Value;
                                            
                                        }
                                        else name = "Труба";
                                    }
                                }
                                else
                                {
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        temp = temp.Trim();
                                        prim = temp;

                                        //diamDxDxD
                                        if (new Regex(@"\w\s+\d+(?:[,.]\d+)?\s+\d+(?:[,.]\d+)?\s*[mм]\s+\d+(?:[,.]\d+)?[mм]", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            diam = new Regex(@"(?<=\w\s+)\d+(?:[,.]\d+)?(?=\s+\d+(?:[,.]\d+)?\s*[mм]\s+\d+(?:[,.]\d+)?[mм])", RegexOptions.IgnoreCase).Match(temp).Value.Replace(".", ",");
                                            tolsh = new Regex(@"(?<=\w\s+\d+(?:[,.]\d+)?\s+)\d+(?:[,.]\d+)?(?=\s*[mм]\s+\d+(?:[,.]\d+)?[mм])", RegexOptions.IgnoreCase).Match(temp).Value.Replace(".", ",");
                                            metraj = new Regex(@"(?<=\w\s+\d+(?:[,.]\d+)?\s+\d+(?:[,.]\d+)?\s*[mм]\s+)\d+(?:[,.]\d+)?(?=[mм])", RegexOptions.IgnoreCase).Match(temp).Value.Replace(".", ",");
                                            if (new Regex(@"\d\d?,?\d*", RegexOptions.IgnoreCase).IsMatch(tolsh)) double.TryParse(tolsh, out dTolsh);
                                            if (new Regex(@"\d\d?,?\d*", RegexOptions.IgnoreCase).IsMatch(metraj)) double.TryParse(metraj, out dMetraj);
                                            if (dTolsh != 0)
                                                tolsh = (dTolsh * 1000).ToString();
                                            if (dMetraj != 0)
                                                metraj = (dMetraj * 1000).ToString();
                                        }
                                        //diamDxD
                                        else if (new Regex(@"\w\s+\d+(?:[,.]\d+)?\s+\d+(?:[,.]\d+)?\s*[mм]", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            diam = new Regex(@"(?<=\w\s+)\d+(?:[,.]\d+)?(?=\s+\d+(?:[,.]\d+)?\s*[mм])", RegexOptions.IgnoreCase).Match(temp).Value.Replace(".", ",");
                                            metraj = new Regex(@"(?<=\w\s+\d+(?:[,.]\d+)?\s+)\d+(?:[,.]\d+)?(?=\s*[mм])", RegexOptions.IgnoreCase).Match(temp).Value.Replace(".", ",");
                                            if (new Regex(@"\d\d?(?:,\d*)?", RegexOptions.IgnoreCase).IsMatch(metraj)) double.TryParse(metraj, out dMetraj);
                                            if (dMetraj != 0)
                                                metraj = (dMetraj * 1000).ToString();
                                        }
                                        //diam D
                                        else if (new Regex(@"\w\s+\d+(?:[,.]\d+)?\s*", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            diam = new Regex(@"(?<=\w\s+)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                        }
                                        if (string.IsNullOrEmpty(metraj))
                                        {
                                            if (new Regex(@"(?<=\d\w\d\s)\d+(?:[,.]\d+)?(?=\s*[mм])", RegexOptions.IgnoreCase).IsMatch(temp))
                                                metraj = new Regex(@"(?<=\d\w\d\s)\d+(?:[,.]\d+)?(?=\s*[mм])", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if (new Regex(@"\d\d?(?:,\d*)?", RegexOptions.IgnoreCase).IsMatch(metraj)) double.TryParse(metraj, out dMetraj);
                                            if (dMetraj != 0)
                                                metraj = (dMetraj * 1000).ToString();
                                        }
                                        mark = regexParam.RegMark.Match(temp).Value;
                                        standart = new Regex(@"\s[\d\-]+(?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        if (name.ToLower().Contains("швеллер")) type = new Regex(@"(?<=\d)у|п", RegexOptions.IgnoreCase).Match(temp).Value;
                                        if (name.ToLower().Contains("электрод")) type = regexParam.RegType.Match(temp).Value;
                                        if (ColPrice > 0)
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[j, ColPrice];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                price = temp;
                                            }
                                        }
                                        
                                        if (ColMera > 0)
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[j, ColMera];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                mera = temp;
                                            }
                                        }

                                        if (!String.IsNullOrEmpty(diam))
                                        {
                                            DataRow row = dtProduct.NewRow();
                                            row["Название"] = name;
                                            if (string.IsNullOrEmpty(type))
                                                row["Тип"] = tab.Type;
                                            else row["Тип"] = type;
                                            if (string.IsNullOrEmpty(row["Тип"].ToString())) row["Тип"] = "тип не указан";
                                            else row["Тип"] = row["Тип"].ToString().ToLower();
                                            row["Диаметр (высота), мм"] = diam;
                                            row["Толщина (ширина), мм"] = tolsh;
                                            row["Метраж, м (длина, мм)"] = metraj;
                                            row["Мерность (т, м, мм)"] = mera;
                                            if (String.IsNullOrEmpty(mark))
                                                row["Марка"] = tab.Mark;
                                            else row["Марка"] = mark;
                                            if (String.IsNullOrEmpty(standart))
                                                row["Стандарт"] = tab.Standart;
                                            else row["Стандарт"] = standart;
                                            row["Класс"] = "";
                                            row["Цена"] = price;
                                            row["Примечание"] = prim;
                                            dtProduct.Rows.Add(row);
                                        }
                                    }
                                }
                            }
                            if (j * (k + 1) < Max) ProcessChanged(j * (k + 1));
                            else ProcessChanged(Max);
                        }
                    }

                    //поиск информации об организации ТОЛЬКО на ПЕРВОМ листе
                    if (tabs.Count > 0 && dtProduct.Rows.Count > 0 && excelworksheet.Index == 1)
                    {
                        Max = (tabs[0].StartRow - 1) * cCelCol;
                        SetMaxValProgressBar(Max);
                        Excel.Range cellRange;
                        for (int j = 1; j <= tabs[0].StartRow - 1; j++) //строки
                        {
                            for (int i = 1; i <= cCelCol; i++) //столбцы
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    FillInfoOrg(infoOrg, temp, regexParam, excelworksheet, j, i);
                                }

                                if (i * j < Max) ProcessChanged(i * j);
                                else ProcessChanged(Max);
                            }
                        }
                    }
                }

                if (isExcelOpen)
                {
                    excelappworkbook.Close();
                    excelapp.Quit();
                }

                SetInfoOrganization(infoOrg);
                WorkCompleted(dtProduct);

            }
            catch (Exception ex) { WorkCompleted(dtProduct); MessageBox.Show("Ошибка в функции ReedExcel() в " + this.ToString() + "\n\n" + ex.ToString()); }
        }

        private static string StringFirstUp(string StringIn)
        {
            string StringOut = "";
            if (!String.IsNullOrEmpty(StringIn))
            {
                if (StringIn.Length > 2)
                    StringOut = StringIn.Substring(0, 1).ToUpper() + StringIn.Substring(1, StringIn.Length - 1).ToLower();
                else StringOut = StringIn;
            }
            return StringOut;
        }

        private static void FillInfoOrg(InfoOrganization infoOrg, string temp, C_RegexParamProduct regexParam, Excel.Worksheet worksheet, int j, int i)
        {
            //if (regexParam.OrgAdresFull.IsMatch(temp))
            //{
            //    infoOrg.OrgAdress = regexParam.OrgAdresFull.Match(temp).Value;//new Regex(@"(?<=офис.*:\s*).*(?:\w|\)|\.)", RegexOptions.IgnoreCase).Match(temp).Value;
            //}
            //if (string.IsNullOrEmpty(infoOrg.OrgAdress))
            //    if (new Regex(@"г\.\s*(?:\w+\s?){1,2}", RegexOptions.IgnoreCase).IsMatch(temp))
            //        infoOrg.OrgAdress = new Regex(@"г\.\s*(?:\w+\s?){1,2}", RegexOptions.IgnoreCase).Match(temp).Value;
            if (new Regex(@"(?:\(\d+\)\s*)(?:\s*[\d\-+]+)+", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.OrgTel = new Regex(@"(?:\(\d+\)\s*)?(?:\s*[\d\-+]+)+", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            //if (regexParam.INN_KPP.IsMatch(temp))
            //{
            //    infoOrg.Inn_Kpp = regexParam.INN_KPP.Match(temp).Value;
            //}
            //if (regexParam.R_S.IsMatch(temp))
            //{
            //    infoOrg.r_s = regexParam.R_S.Match(temp).Value;
            //}
            //if (regexParam.K_S.IsMatch(temp))
            //{
            //    infoOrg.k_s = regexParam.K_S.Match(temp).Value;
            //}
            //if (regexParam.BIK.IsMatch(temp))
            //{
            //    infoOrg.BIK = regexParam.BIK.Match(temp).Value;
            //}

            //if (new Regex(@"[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})", RegexOptions.IgnoreCase).IsMatch(temp))
            //{
            //    foreach (Match m in new Regex(@"[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})", RegexOptions.IgnoreCase).Matches(temp))
            //        if (String.IsNullOrEmpty(infoOrg.Email)) infoOrg.Email = m.Value;
            //        else infoOrg.Email += "; " + m.Value;
            //}
            //if (new Regex(@"w+\.[\w\-]+(?:\.[\w\-]+)?", RegexOptions.IgnoreCase).IsMatch(temp))
            //{
            //    infoOrg.Site = new Regex(@"w+\.[\w\-]+(?:\.[\w\-]+)?", RegexOptions.IgnoreCase).Match(temp).Value;
            //}
            //if (new Regex(@"г\.\w+[,\w\.\s\d]+(?=\s*$)", RegexOptions.IgnoreCase).IsMatch(temp))
            //{
            //    infoOrg.SkladAdr.Add(new Regex(@"г\.\w+[,\w\.\s\d]+(?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value);
            //}
            //if (new Regex(@"(?:\w+\s*)+тел.*\.(?:ru|com|info|рф)", RegexOptions.IgnoreCase).IsMatch(temp))
            //{
            //    string tmp = new Regex(@"(?:\w+\s*)+тел.*\.(?:ru|com|info|рф)", RegexOptions.IgnoreCase).Match(temp).Value;
            //    string[] manager = new string[]
            //    {
            //        new Regex(@"(?:\w+\s*)+(?=тел.*\.(?:ru|com|info|рф))", RegexOptions.IgnoreCase).Match(temp).Value, //имя
            //        new Regex(@"(?<=(?:\w+\s*)+тел.*)[\d\-]+(?=\s*e-.*(?:ru|com|info|рф))", RegexOptions.IgnoreCase).Match(temp).Value, //телефон
            //        new Regex(@"(?<=(?:\w+\s*)+тел.*[\d\-]+\s*e-mail:).*(?:ru|com|info|рф)", RegexOptions.IgnoreCase).Match(temp).Value //email
            //    };
            //    infoOrg.Manager.Add(manager);
            //}
        }

        private List<double> GetIncrementingMassiv(string[] strParams)
        {
            List<double> Ddiam = new List<double>();
            List<double> ch = new List<double>();
            string str;
            foreach (string s in strParams)
            {
                str = s.Replace('.', ',');
                Ddiam.Add(Convert.ToDouble(s));
            }
            if (strParams.Length > 1)
            {
                double increment = 0;
                if (Ddiam[1] >= 1 && Ddiam[1] < 4) increment = 0.5;
                if (Ddiam[1] >= 4 && Ddiam[1] < 50) increment = 2;
                if (Ddiam[1] >= 50 && Ddiam[1] < 100) increment = 10;
                if (Ddiam[1] >= 100 && Ddiam[1] < 1000) increment = 100;
                if (Ddiam[1] >= 1000 && Ddiam[1] < 5000) increment = 500;
                if (Ddiam[1] >= 5000 && Ddiam[1] < 50000) increment = 1000;

                if (increment > 0)
                {
                    for (double d = Ddiam[0]; d <= Ddiam[1]; d += increment)
                    {
                        if (d != Ddiam[0] && d % 1 == 1)
                            d -= 0.1;
                        ch.Add(d);
                        if (d + increment > Ddiam[1] && d != Ddiam[1]) ch.Add(Ddiam[1]);
                    }
                    if (ch.Count > 0) Ddiam = ch;
                }
            }
            return Ddiam;
        }

        public event Action<int> ProcessChanged; //установить текущее значение прогрессбара

        public event Action<int> SetMaxValProgressBar; //установить максимальное значение для прогрессбара

        public event Action<InfoOrganization> SetInfoOrganization;

        public event Action<DataTable> WorkCompleted;
    }
}
