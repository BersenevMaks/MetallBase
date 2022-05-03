using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;

namespace MetalBase
{
    class Class_UPTK
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

                string temp = "", tmp = "", tempRaz = "", mark = "", standart = "", name = "", type = "", price = "", prim = "";
                string diam = "", tolsh = "", metraj = "", mera = "";
                var regexParam = new C_RegexParamProduct();

                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    //MessageBox.Show(excelsheets.Count.ToString());
                    var tab = new C_InfoTable();
                    List<C_InfoTable> tabs = new List<C_InfoTable>();
                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol <= 10) cCelCol = 10;
                    if (cCelCol > 10) cCelCol = 20;
                    cCelCol = 10;

                    int lastRow = 0;
                    int ColName = 0, ColRaz = 0, ColMark = 0, ColMera = 0, ColPrice = 0;

                    int Max = cCelCol * cCelRow;
                    SetMaxValProgressBar(Max);
                    //Поиск заголовков столбцов
                    int progress = 0;

                    for (int j = 6; j <= cCelRow; j++) //строки
                    {
                        int jj = j;
                        for (int i = 1; i <= 1; i++) //столбцы
                        {
                            Excel.Range cellRange;
                            cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"^наименование", RegexOptions.IgnoreCase).IsMatch(temp))
                                { tab = new C_InfoTable(); ColName = i; tab.StartRow = jj; tabs.Add(tab); }
                            }
                        }
                    }

                    if(tabs.Count > 0)
                    {
                        for (int i = 2; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange;
                            cellRange = (Excel.Range)excelworksheet.Cells[tabs[0].StartRow, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"размер", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColRaz = i; }
                                else if (new Regex(@"марка", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColMark = i; }
                                else if (new Regex(@"кол.*во", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColMera = i; }
                                else if (new Regex(@"цена", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColPrice = i; }

                                if (progress < Max) ProcessChanged(progress++);
                                else ProcessChanged(Max);
                            }
                        }
                    }
                    ProcessChanged(0);
                    Max = cCelRow;
                    SetMaxValProgressBar(Max);
                    bool NotAName = false;
                    if (ColRaz != 0)
                    {
                        for (int k = 0; k < tabs.Count; k++)
                        {
                            int endRow = cCelRow;
                            if (k < tabs.Count - 1)
                                endRow = tabs[k + 1].StartRow;
                            tab = tabs[k];
                            Excel.Range cellRange;
                            for (int jj = tab.StartRow + 1; jj <= endRow; jj++)
                            {
                                tab = new C_InfoTable();
                                cellRange = (Excel.Range)excelworksheet.Cells[jj, ColRaz];
                                if (cellRange.Value != null)
                                    tempRaz = cellRange.Value.ToString().Trim();
                                else tempRaz = "";
                                if (tempRaz != "")
                                {
                                    cellRange = (Excel.Range)excelworksheet.Cells[jj, ColName];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        if (regexParam.RegName.IsMatch(temp))
                                        {
                                            name = regexParam.RegName.Match(temp).Value;
                                            name = StringFirstUp(name);
                                            type = new Regex(@"[cc]-образный", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if (type == "") type = regexParam.RegType.Match(temp).Value;
                                            NotAName = false;
                                        }
                                        else if (new Regex(@"лист|круг", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            name = new Regex(@"лист|круг", RegexOptions.IgnoreCase).Match(temp).Value;
                                            name = StringFirstUp(name);
                                            type = new Regex(@"[cc]-образный", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if (type == "") type = regexParam.RegType.Match(temp).Value;
                                            NotAName = false;
                                        }
                                        else { NotAName = true; continue; }
                                    }
                                    else if (NotAName) continue;
                                    
                                    diam = ""; tolsh = ""; metraj = "";
                                    if (new Regex(@"(?<=^|\s)\d+(?:[,\.]\d+)?.{0,4}\s*.*[xх]\s*\d+(?:[,\.]\d+)?(?:\s*[xх]\s*\d+(?:[,\.]\d+)?)+(?:\s|\(|\?|м\.?|\s*$|-|н)", RegexOptions.IgnoreCase).IsMatch(tempRaz))
                                    {
                                        tolsh = new Regex(@"(?<=[xх]\s*)\d+(?:[,\.]\d+)?(?=\s*[xх]\s*\d+(?:[,\.]\d+)?\s*м?\s*$|-|н)", RegexOptions.IgnoreCase).Match(tempRaz).Value;
                                        diam = new Regex(@"(?<=^|\s)\d+(?:[,\.]\d+)?(?=.{0,4}\s*[xх])", RegexOptions.IgnoreCase).Match(tempRaz).Value;
                                        metraj = new Regex(@"(?<=[xх]\s*)\d+(?:[,\.]\d+)?(?=\s|\(|\?|м\.?|\s*$)", RegexOptions.IgnoreCase).Match(tempRaz).Value;
                                    }
                                    else if (new Regex(@"(?<=^|\s)\d+(?:[,\.]\d+)?\s*.*[xх]\s*\d+(?:[,\.]\d+)?(?:\s|\(|\?|м\.?|\s*$|-)", RegexOptions.IgnoreCase).IsMatch(tempRaz))
                                    {
                                        diam = new Regex(@"(?<=^|\s)\d+(?:[,\.]\d+)?(?=\s*[xх])", RegexOptions.IgnoreCase).Match(tempRaz).Value;
                                        tolsh = new Regex(@"(?<=[xх]\s*)\d+(?:[,\.]\d+)?(?=\s|\(|\?|м\.?|\s*$|-)", RegexOptions.IgnoreCase).Match(tempRaz).Value;
                                        if (name.ToLower() == "арматура" || name.ToLower() == "швеллер") { metraj = tolsh; tolsh = ""; }
                                    }
                                    else if (new Regex(@"(?<=^|\s)\d+(?:[,\.]\d+)?(?=\s|\(|\?|м\.?|\s*$|-)", RegexOptions.IgnoreCase).IsMatch(tempRaz))
                                    {
                                        diam = new Regex(@"(?<=^|\s)\d+(?:[,\.]\d+)?(?=\s|\(|\?|м\.?|\s*$|-)", RegexOptions.IgnoreCase).Match(tempRaz).Value;
                                    }
                                    prim = tempRaz;
                                    if (diam != "")
                                    {
                                        if (ColMark > 0)
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[jj, ColMark];
                                            if (cellRange.Value != null)
                                                tmp = cellRange.Value.ToString().Trim();
                                            else tmp = "";
                                            if (tmp != "")
                                            {
                                                mark = tmp;
                                            }
                                            else if (cellRange.MergeArea.Rows.Count == 1) mark = "";
                                        }
                                        if (ColMera > 0)
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[jj, ColMera];
                                            if (cellRange.Value != null)
                                                tmp = cellRange.Value.ToString().Trim();
                                            else tmp = "";
                                            if (tmp != "")
                                            {
                                                mera = tmp;
                                            }
                                        }
                                        if (ColPrice > 0)
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[jj, ColPrice];
                                            if (cellRange.Value != null)
                                                tmp = cellRange.Value.ToString().Trim();
                                            else tmp = "";
                                            if (tmp != "")
                                            {
                                                price = tmp;
                                            }
                                        }

                                        if (dtProduct.Rows.Count > 0)
                                        {
                                            lastRow = dtProduct.Rows.Count - 1;
                                            tab.LastRowExcel = jj;
                                        }
                                        else lastRow = 0;

                                        DataRow row = dtProduct.NewRow();

                                        row["Название"] = name;
                                        row["Тип"] = type;
                                        if (string.IsNullOrEmpty(row["Тип"].ToString())) row["Тип"] = "тип не указан";
                                        else row["Тип"] = row["Тип"].ToString().ToLower();
                                        row["Диаметр (высота), мм"] = diam;
                                        row["Толщина (ширина), мм"] = tolsh;
                                        row["Метраж, м (длина, мм)"] = metraj;
                                        row["Мерность (т, м, мм)"] = mera;
                                        row["Марка"] = mark;
                                        row["Стандарт"] = standart;
                                        row["Класс"] = "";
                                        row["Цена"] = price;
                                        row["Примечание"] = prim;
                                        dtProduct.Rows.Add(row);
                                    }
                                }
                                else
                                {
                                    cellRange = (Excel.Range)excelworksheet.Cells[jj, ColName];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        if (regexParam.RegName.IsMatch(temp))
                                        {
                                            name = regexParam.RegName.Match(temp).Value;
                                            name = StringFirstUp(name);
                                            type = new Regex(@"[cc]-образный", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if (type == "") type = regexParam.RegType.Match(temp).Value;
                                            NotAName = false;
                                        }
                                        else if (new Regex(@"лист|круг", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            name = new Regex(@"лист|круг", RegexOptions.IgnoreCase).Match(temp).Value;
                                            name = StringFirstUp(name);
                                            type = new Regex(@"[cc]-образный", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if(type=="") type = regexParam.RegType.Match(temp).Value;
                                            NotAName = false;
                                        }
                                        else { NotAName = true; continue; }
                                    }
                                    else if (NotAName) continue;

                                    if (progress < Max) ProcessChanged(jj);
                                    else ProcessChanged(Max);
                                }
                            }
                        }
                    }

                    //поиск информации об организации
                    if(tabs.Count>0)
                    if (tabs[0].StartRow - 1 > 0)
                    {
                        Max = (tabs[0].StartRow - 1) * cCelCol;
                        SetMaxValProgressBar(Max);

                        for (int j = 1; j < tabs[0].StartRow - 1; j++) //строки
                        {
                            for (int i = 1; i <= cCelCol; i++) //столбцы
                            {
                                Excel.Range cellRange;
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
                    //excelappworkbook.Close();
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
            if (new Regex(@"г\.\s*\w+,?\s* ул\.\s*\w+\s*\d+", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.OrgAdress = new Regex(@"г\.\s*\w+,?\s* ул\.\s*\w+\s*\d+", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (new Regex(@"(?<=Тел.*\s)(?:\(\d{3,5}\))?(?:\s\d+-\d+\s*,?)+", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.OrgTel = new Regex(@"(?<=Тел.*\s)(?:\(\d{3,5}\))?(?:\s\d+-\d+\s*,?)+", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (new Regex(@"[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.Email = new Regex(@"[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (new Regex(@"(?:www\.|http:)(?:[а-яёa-z0-9_-]{1,32}(?::[а-яёa-z0-9_-]{1,32})?@)?(?:(?:[а-яёa-z0-9-]{1,128}\.)+(?:ru|su|com|net|org|mil|edu|arpa|gov|biz|info|aero|inc|name|рф|[а-яёa-z]{2}))", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.Site = new Regex(@"(?:www\.|http:)(?:[а-яёa-z0-9_-]{1,32}(?::[а-яёa-z0-9_-]{1,32})?@)?(?:(?:[а-яёa-z0-9-]{1,128}\.)+(?:ru|su|com|net|org|mil|edu|arpa|gov|biz|info|aero|inc|name|рф|[а-яёa-z]{2}))", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (new Regex(@"адрес\s*базы", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.SkladAdr.Add(new Regex(@"(?<=дрес\s*базы\s*:?\s*)[\w\s\.\d]+(?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value);
            }
            if (new Regex(@"(?:\d\s*-\s*)?\d{3,4}\s*-\s*\d{3,4}\s*-\s*\d{2,4}\s*-\s*\d{2,4}", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                string telefon = new Regex(@"(?:\d\s*-\s*)?\d{3,4}\s*-\s*\d{3,4}\s*-\s*\d{2,4}\s*-\s*\d{2,4}", RegexOptions.IgnoreCase).Match(temp).Value;
                Excel.Range range = (Excel.Range)worksheet.Cells[j + 1, i];
                if (range.Value != null)
                {
                    temp = range.Value.ToString().Trim();
                }
                else temp = "";
                if (temp != "")
                {
                    if (new Regex(@"\w+(?:\s*\w+)*", RegexOptions.IgnoreCase).IsMatch(temp))
                    {
                        string[] manager = new string[] { new Regex(@"\w+(?:\s*\w+)*", RegexOptions.IgnoreCase).Match(temp).Value, telefon };
                        infoOrg.Manager.Add(manager);
                    }
                }
            }
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
                if (Ddiam[1] >= 50) increment = 10;
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

        public event Action<int> ProcessChanged;

        public event Action<int> SetMaxValProgressBar;

        public event Action<InfoOrganization> SetInfoOrganization;

        public event Action<DataTable> WorkCompleted;

    }
}
