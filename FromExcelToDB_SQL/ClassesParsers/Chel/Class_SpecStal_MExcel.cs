using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;

namespace MetalBase
{
    class Class_SpecStal_M
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
                string diam = "", tolsh = "", metraj = "", mera = "";
                var regexParam = new C_RegexParamProduct();
                List<double> dDiam;
                List<double> dTolsh;
                List<double> dMetraj;

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
                    int ColDiam = 0, ColTolsh = 0, ColDlina = 0, ColMark = 0, ColMera = 0, ColPrice = 0, ColGost = 0;

                    int Max = cCelCol * cCelRow;
                    SetMaxValProgressBar(Max);
                    //Поиск заголовков столбцов
                    int progress = 0;

                    for (int j = 10; j <= cCelRow; j++) //строки
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
                                if (new Regex(@"^ширина", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColDiam = i; tab.StartRow = jj; j = cCelRow + 1;
                                }
                                if (new Regex(@"толщина", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColTolsh = i; }
                                if (new Regex(@"длина", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColDlina = i; }
                                else if (new Regex(@"^сталь", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColMark = i; }
                                else if (new Regex(@"^вес", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColMera = i; }
                                else if (new Regex(@"^н[\/]д", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColGost = i; }
                                else if (new Regex(@"^цена", RegexOptions.IgnoreCase).IsMatch(temp))
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
                        for (int jj = tab.StartRow - 6; jj <= cCelRow; jj++)
                        {
                            cellRange = (Excel.Range)excelworksheet.Cells[jj, ColDiam];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                dDiam = new List<double>();
                                dTolsh = new List<double>();
                                dMetraj = new List<double>();

                                if (temp.Split('-').Length == 2)
                                    dDiam = GetIncrementingMassiv(temp.Split('-'));
                                else
                                {
                                    double d;
                                    double.TryParse(temp, out d);
                                    if (d != 0)
                                    {
                                        dDiam.Add(d);
                                    }
                                }


                                if (dDiam.Count > 0)
                                {
                                    if (ColTolsh > 0)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, ColTolsh];
                                        if (cellRange.Value != null)
                                            tmp = cellRange.Value.ToString().Trim();
                                        else tmp = "";
                                        if (tmp != "")
                                        {
                                            if (tmp.Split('-').Length == 2)
                                                dTolsh = GetIncrementingMassiv(tmp.Split('-'));
                                            else dTolsh.Add(Convert.ToDouble(tmp));
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
                                            if (tmp.Split('-').Length == 2)
                                                dMetraj = GetIncrementingMassiv(tmp.Split('-'));
                                            else dMetraj.Add(Convert.ToDouble(tmp));
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
                                    if (ColGost > 0)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, ColGost];
                                        if (cellRange.Value != null)
                                            tmp = cellRange.Value.ToString().Trim();
                                        else tmp = "";
                                        if (tmp != "")
                                        {
                                            standart = tmp;
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
                                            tmp = tmp.Replace("р.", "");
                                            price = tmp;
                                        }
                                    }

                                    if (dtProduct.Rows.Count > 0)
                                    {
                                        lastRow = dtProduct.Rows.Count - 1;
                                        tab.LastRowExcel = jj;
                                    }
                                    else lastRow = 0;

                                    for (int d = 0; d < dDiam.Count; d++)
                                        for (int t = 0; t < dTolsh.Count; t++)
                                            for (int m = 0; m < dMetraj.Count; m++)
                                            {
                                                DataRow row = dtProduct.NewRow();
                                                row["Название"] = "Лист";
                                                row["Тип"] = type;
                                                if (string.IsNullOrEmpty(row["Тип"].ToString())) row["Тип"] = "тип не указан";
                                                else row["Тип"] = row["Тип"].ToString().ToLower();
                                                row["Диаметр (высота), мм"] = dDiam[d];
                                                row["Толщина (ширина), мм"] = dTolsh[t];
                                                row["Метраж, м (длина, мм)"] = dMetraj[m];
                                                row["Мерность (т, м, мм)"] = mera;
                                                row["Марка"] = mark;
                                                row["Стандарт"] = standart;
                                                row["Класс"] = "";
                                                row["Цена"] = price;
                                                row["Примечание"] = prim;
                                                dtProduct.Rows.Add(row);
                                            }
                                }
                            }
                            else
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[jj, 1];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    if (temp.Contains("наличии")) prim = "в наличии";
                                    else if (temp.Contains("заказ")) prim = "под заказ";
                                }
                            }
                            if (progress < Max) ProcessChanged(progress++);
                            else ProcessChanged(Max);
                        }
                    }

                    //поиск информации об организации

                    if (dtProduct.Rows.Count > 0)
                    {
                        Max = (tab.StartRow - 1) * cCelCol;
                        SetMaxValProgressBar(Max);
                        Excel.Range cellRange;
                        for (int j = 1; j <= tab.StartRow - 1; j++) //строки
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
            if (new Regex(@"(?<=адрес\s*:?\s*)(?:\d+\s*,?\s*)?(?:\s*(?:г|ул|д)\.\s*[\w\d]+\s*,?)+", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.OrgAdress = new Regex(@"(?<=адрес\s*:?\s*)(?:\d+\s*,?\s*)?(?:\s*(?:г|ул|д)\.\s*[\w\d]+\s*,?)+", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (new Regex(@"(?<=ел(?:ефон\s*)?[\/]?(?:\s*\(?факс\)?)\s*:\s*)(?:[\d\(\)\s-]+(?:доб\s*\.?\s*\d{1,5})?\s*;?)+", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.OrgTel = new Regex(@"(?<=ел(?:ефон\s*)?[\/]?(?:\s*\(?факс\)?)\s*:\s*)(?:[\d\(\)\s-]+(?:доб\s*\.?\s*\d{1,5})?\s*;?)+", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (regexParam.INN_KPP.IsMatch(temp))
            {
                infoOrg.Inn_Kpp = regexParam.INN_KPP.Match(temp).Value;
            }
            if (regexParam.R_S.IsMatch(temp))
            {
                infoOrg.r_s = regexParam.R_S.Match(temp).Value;
            }
            if (regexParam.K_S.IsMatch(temp))
            {
                infoOrg.k_s = regexParam.K_S.Match(temp).Value;
            }
            if (regexParam.BIK.IsMatch(temp))
            {
                infoOrg.BIK = regexParam.BIK.Match(temp).Value;
            }

            if (new Regex(@"[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                foreach(Match m in new Regex(@"[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})", RegexOptions.IgnoreCase).Matches(temp))
                    if(String.IsNullOrEmpty(infoOrg.Email)) infoOrg.Email = m.Value;
                    else infoOrg.Email += "; " + m.Value;
            }
            if (new Regex(@"(?:www\.|http:)(?:[а-яёa-z0-9_-]{1,32}(?::[а-яёa-z0-9_-]{1,32})?@)?(?:(?:[а-яёa-z0-9-]{1,128}\.)+(?:ru|su|com|net|org|mil|edu|arpa|gov|biz|info|aero|inc|name|рф|[а-яёa-z]{2}))", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.Site = new Regex(@"(?:www\.|http:)(?:[а-яёa-z0-9_-]{1,32}(?::[а-яёa-z0-9_-]{1,32})?@)?(?:(?:[а-яёa-z0-9-]{1,128}\.)+(?:ru|su|com|net|org|mil|edu|arpa|gov|biz|info|aero|inc|name|рф|[а-яёa-z]{2}))", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (new Regex(@"Склад\s*:\s*(?:\+7\s*|8\s*)(?:\(?[\d-]+\)?\s*)+\w+\s*", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.SkladAdr.Add(new Regex(@"Склад\s*:\s*(?:\+7\s*|8\s*)(?:\(?[\d-]+\)?\s*)+\w+\s*", RegexOptions.IgnoreCase).Match(temp).Value);
            }
            if (new Regex(@"(?<=обильные\s*:)(?:\s*(?:\+7\s*|8\s*)(?:\(?[\d-]+\)?\s*)\w+\s*;?)+", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                string[] telefons = new Regex(@"(?<=обильные\s*:)(?:\s*(?:\+7\s*|8\s*)(?:\(?[\d-]+\)?\s*)\w+\s*;?)+", RegexOptions.IgnoreCase).Match(temp).Value.Split(';');
                foreach (string tel in telefons)
                {
                    string telefon = new Regex(@"\s*(?:\+7\s*|8\s*)(?:\(?[\d-]+\)?)", RegexOptions.IgnoreCase).Match(tel).Value;
                    string[] manager = new string[] { new Regex(@"(?<=\d\s+)\w+(?=;|\s|\s*$)", RegexOptions.IgnoreCase).Match(tel).Value, telefon };
                    infoOrg.Manager.Add(manager);
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

        public event Action<int> ProcessChanged;

        public event Action<int> SetMaxValProgressBar;

        public event Action<InfoOrganization> SetInfoOrganization;

        public event Action<DataTable> WorkCompleted;
    }
}
