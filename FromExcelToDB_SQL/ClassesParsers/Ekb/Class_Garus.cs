using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;

namespace MetalBase
{
    class Class_Garus
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
                orgname = "Гарус";
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

                string temp = "", tmp = "", price = "", prim = "";
                string diam = "", tolsh = "", metraj = "", mera = "";
                var regexParam = new C_RegexParamProduct();

                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    if (excelworksheet.Name.ToLower() == "кубы")
                    {
                        ReadKubSheet(infoOrg, ref temp, ref tmp, ref diam, ref tolsh, ref price, ref metraj, ref prim, ref mera, excelworksheet, regexParam);
                        continue;
                    }
                    var tab = new C_InfoTable();
                    List<C_InfoTable> tabs = new List<C_InfoTable>();
                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol <= 10) cCelCol = 10;
                    if (cCelCol > 10) cCelCol = 20;
                    cCelCol = 10;

                    int lastRow = 0;
                    int ColDiam = 0, ColRaz = 6, ColMark = 2, ColDlina = 0, ColMera = 0, ColPrice = 0;

                    int Max = cCelCol * cCelRow;
                    SetMaxValProgressBar(Max);
                    //Поиск заголовков столбцов
                    int progress = 0;
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        int jj = j;
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange;
                            cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim().Replace(" ","");
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"^диаметр", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColDiam = i; j = cCelRow+1; tab.StartRow = jj; }
                                else if (new Regex(@"размер", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColRaz = 6; }
                                else if (new Regex(@"марка", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColMark = 2; }
                                else if (new Regex(@"длина", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColDlina = i; }
                                else if (new Regex(@"\bвес\b", RegexOptions.IgnoreCase).IsMatch(temp))
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
                    if (ColDiam != 0)
                    {
                        Excel.Range cellRange;
                        for (int jj = tab.StartRow + 1; jj <= cCelRow; jj++) 
                        {
                            tab = new C_InfoTable();
                            cellRange = (Excel.Range)excelworksheet.Cells[jj, ColDiam];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp == "")
                            {
                                for (int i = 2; i < cCelCol; i++)
                                {
                                    cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim().Replace(" ", "");
                                    else temp = "";
                                    
                                    if (temp != "")
                                    {
                                        if (regexParam.RegName.IsMatch(temp))
                                        {
                                            tab.Name = regexParam.RegName.Match(temp).Value;
                                        }
                                        else if (new Regex(@"лист|круг", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            tab.Name = new Regex(@"лист|круг", RegexOptions.IgnoreCase).Match(temp).Value;
                                        }
                                        else tab.Name = "";
                                        if (tab.Name != "")
                                        {
                                            tab.Name = StringFirstUp(tab.Name);
                                            tab.Type = regexParam.RegType.Match(temp).Value;
                                            if (tab.Type == "") tab.Type = new Regex(@"на столбы|нержавеющ(?:ий|ая)", RegexOptions.IgnoreCase).Match(temp).Value;
                                            tab.StartRow = jj + 1;
                                            tabs.Add(tab);
                                            break;
                                        }
                                    }
                                }
                            }
                            if (progress < Max) ProcessChanged(progress++);
                            else ProcessChanged(Max);
                        }

                        ProcessChanged(0);
                        Max = cCelRow;
                        SetMaxValProgressBar(Max);

                        for (int k = 0; k < tabs.Count; k++)
                        {
                            tab = tabs[k];
                            int endRow = cCelRow;
                            if (k < tabs.Count - 1)
                                endRow = tabs[k + 1].StartRow - 2;

                            for (int jj = tab.StartRow + 1; jj <= endRow; jj++) //строки
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[jj, ColDiam];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    if (!String.IsNullOrEmpty(tab.Name))
                                    {
                                        temp = temp.Replace(" ", "");
                                        temp = temp.Replace('/', 'х');

                                        diam = ""; tolsh = ""; metraj = "";
                                        if (new Regex(@"(?<=^)\d+(?:[,\.]\d+)?\s*[xх]\s*\d+(?:[,\.]\d+)?\s*[xх]\s*\d+(?:[,\.]\d+)?(?:\s*$|\(|\?)", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            diam = new Regex(@"(?<=[xх]\s*)\d+(?:[,\.]\d+)?(?=\s*[xх])", RegexOptions.IgnoreCase).Match(temp).Value;
                                            tolsh = new Regex(@"(?<=^)\d+(?:[,\.]\d+)?(?=\s*[xх])", RegexOptions.IgnoreCase).Match(temp).Value;
                                            metraj = new Regex(@"(?<=[xх]\s*)\d+(?:[,\.]\d+)?(?=\s*$|\(|\?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        }
                                        else if (new Regex(@"(?<=^)\d+(?:[,\.]\d+)?\s*[xх]\s*\d+(?:[,\.]\d+)?(?:\s*$|\(|\?)", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            diam = new Regex(@"(?<=^)\d+(?:[,\.]\d+)?(?=\s*[xх])", RegexOptions.IgnoreCase).Match(temp).Value;
                                            tolsh = new Regex(@"(?<=[xх]\s*)\d+(?:[,\.]\d+)?(?=\s*$|\(|\?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        }
                                        else if (new Regex(@"(?<=^)\d+(?:[,\.]\d+)?(?=\s*$|\(|\?)", RegexOptions.IgnoreCase).IsMatch(temp))
                                            diam = new Regex(@"(?<=^)\d+(?:[,\.]\d+)?(?=\s*$|\(|\?)", RegexOptions.IgnoreCase).Match(temp).Value;

                                        if (!string.IsNullOrEmpty(diam))
                                        {
                                            prim = "диаметр: " + diam + "; размер: ";
                                            if (ColMark > 0)
                                            {
                                                cellRange = (Excel.Range)excelworksheet.Cells[jj, ColMark];
                                                if (cellRange.Value != null)
                                                    tmp = cellRange.Value.ToString().Trim();
                                                else tmp = "";
                                                if (tmp != "")
                                                {
                                                    tab.Mark = tmp;
                                                }
                                            }
                                            if (ColDlina > 0)
                                            {
                                                cellRange = (Excel.Range)excelworksheet.Cells[jj, ColDlina];
                                                if (cellRange.Value != null)
                                                    tmp = cellRange.Value.ToString().Trim();
                                                else tmp = "";
                                                if (tmp != "")
                                                {
                                                    metraj = tmp;
                                                }
                                            }
                                            if (ColRaz > 0)
                                            {
                                                cellRange = (Excel.Range)excelworksheet.Cells[jj, ColRaz];
                                                if (cellRange.Value != null)
                                                    tmp = cellRange.Value.ToString().Trim();
                                                else tmp = "";
                                                if (tmp != "")
                                                {
                                                    prim += tmp;
                                                }
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
                                            row["Название"] = tab.Name;
                                            row["Тип"] = tab.Type;
                                            if (string.IsNullOrEmpty(row["Тип"].ToString())) row["Тип"] = "тип не указан";
                                            else row["Тип"] = row["Тип"].ToString().ToLower();
                                            row["Диаметр (высота), мм"] = diam;
                                            row["Толщина (ширина), мм"] = tolsh;
                                            row["Метраж, м (длина, мм)"] = metraj;
                                            row["Мерность (т, м, мм)"] = mera;
                                            row["Марка"] = tab.Mark;
                                            row["Стандарт"] = tab.Standart;
                                            row["Класс"] = "";
                                            row["Цена"] = price;
                                            row["Примечание"] = prim;
                                            dtProduct.Rows.Add(row);
                                        }
                                    }
                                }

                                if (jj < Max) ProcessChanged(jj);
                                else ProcessChanged(Max);
                            }
                        }
                    }

                    if(tabs.Count>0)
                    if (tabs[0].StartRow-1 > 0)
                    {
                            Max = (tabs[0].StartRow - 1) * cCelCol;
                                SetMaxValProgressBar(Max);
                        for (int j = 1; j < tabs[0].StartRow-1; j++) //строки
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
                                    FillInfoOrg(infoOrg, temp, regexParam);
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

        private void ReadKubSheet(InfoOrganization infoOrg, ref string temp, ref string tmp, ref string diam, ref string tolsh, ref string price, ref string metraj, ref string prim, ref string mera, Excel.Worksheet excelworksheet, C_RegexParamProduct regexParam)
        {
            string mark = "";
            var tab = new C_InfoTable();
            List<C_InfoTable> tabs = new List<C_InfoTable>();
            int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
            if (cCelCol <= 10) cCelCol = 10;
            if (cCelCol > 10) cCelCol = 20;
            cCelCol = 10;

            int lastRow = 0;
            int ColTolsh = 0, ColShir = 0, ColMark = 0, ColDlina = 0, ColMera = 0, ColPrice = 0;

            int Max = cCelCol * cCelRow;
            SetMaxValProgressBar(Max);
            //Поиск заголовков столбцов
            int progress = 0;
            for (int j = 1; j <= cCelRow; j++) //строки
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
                        if (new Regex(@"длина", RegexOptions.IgnoreCase).IsMatch(temp))
                        { ColDlina = i; j = cCelRow + 1; tab.StartRow = jj; }
                        else if (new Regex(@"толщина", RegexOptions.IgnoreCase).IsMatch(temp))
                        { ColTolsh = i; }
                        else if (new Regex(@"марка", RegexOptions.IgnoreCase).IsMatch(temp))
                        { ColMark = i; }
                        else if (new Regex(@"ширина", RegexOptions.IgnoreCase).IsMatch(temp))
                        { ColShir = i; }
                        else if (new Regex(@"масса", RegexOptions.IgnoreCase).IsMatch(temp))
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
            if (ColDlina != 0)
            {
                Excel.Range cellRange;
                for (int jj = tab.StartRow + 1; jj <= cCelRow; jj++)
                {
                    tab = new C_InfoTable();
                    cellRange = (Excel.Range)excelworksheet.Cells[jj, ColDlina];
                    if (cellRange.Value != null)
                        temp = cellRange.Value.ToString().Trim();
                    else temp = "";
                    if (temp != "")
                    {
                        diam = new Regex(@"\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                        prim = "длина: " + tmp;
                        if (!String.IsNullOrEmpty(diam))
                        {
                            if (ColTolsh > 0)
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[jj, ColTolsh];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    tolsh = new Regex(@"\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                    prim += "; толщина: " + tmp;
                                }
                            }
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
                                }
                                if (ColShir > 0)
                                {
                                    cellRange = (Excel.Range)excelworksheet.Cells[jj, ColShir];
                                    if (cellRange.Value != null)
                                        tmp = cellRange.Value.ToString().Trim();
                                    else tmp = "";
                                    if (tmp != "")
                                    {
                                        metraj = tmp;
                                    prim += "; ширина: " + tmp;
                                    }
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
                                row["Название"] = "Куб";
                                row["Тип"] = "тип не указан";
                                row["Диаметр (высота), мм"] = diam;
                                row["Толщина (ширина), мм"] = tolsh;
                                row["Метраж, м (длина, мм)"] = metraj;
                                row["Мерность (т, м, мм)"] = mera;
                                row["Марка"] = mark;
                                row["Стандарт"] = "";
                                row["Класс"] = "";
                                row["Цена"] = price;
                                row["Примечание"] = prim;
                                dtProduct.Rows.Add(row);
                            }
                        }
                    
                    if (progress < Max) ProcessChanged(jj);
                    else ProcessChanged(Max);
                }
            }
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

        private static void FillInfoOrg(InfoOrganization infoOrg, string temp, C_RegexParamProduct regexParam)
        {
            if (new Regex(@"г\.\s*\w+,?\s* ул\.\s*\w+\s*\d+", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.OrgAdress = new Regex(@"г\.\s*\w+,?\s* ул\.\s*\w+\s*\d+", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (regexParam.OrgMobileTelefon.IsMatch(temp))
            {
                infoOrg.OrgTel = regexParam.OrgMobileTelefon.Match(temp).Value;
            }
            if (new Regex(@"\d\s*\(\d+\).+\s*\d+-\d+-\d+(?:\(\d+\))?", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.OrgTel = new Regex(@"\d\s*\(\d+\).+\s*\d+-\d+-\d+(?:\(\d+\))?", RegexOptions.IgnoreCase).Match(temp).Value;
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
