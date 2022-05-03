using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;

namespace MetalBase
{
    class Class_inkomMetal
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

                string temp = "", tmp = "", price = "", prim = "";
                string diam = "", tolsh = "", metraj = "", mera = "";
                var regexParam = new C_RegexParamProduct();

                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    var tab = new C_InfoTable();
                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;


                    int lastRow = 0;
                    int ColRaz = 0, ColMark = 0, ColDlina = 0, ColMera = 0, ColPrice = 0;

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
                                if (new Regex(@"^размер", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColRaz = i; j = jj; tab.StartRow = jj; }
                                else if (new Regex(@"Марка", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColMark = i; }
                                else if (new Regex(@"длина", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColDlina = i; }
                                else if (new Regex(@"\bвес\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColMera = i; }
                                else if (new Regex(@"Цена", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColPrice = i; }

                                if (progress < Max) ProcessChanged(progress++);
                                else ProcessChanged(Max);
                            }
                        }
                    }
                    ProcessChanged(0);
                    Max = cCelRow;
                    SetMaxValProgressBar(Max);
                    if (ColRaz != 0)
                    {
                        for (int jj = tab.StartRow + 1; jj <= cCelRow; jj++) //строки
                        {
                            Excel.Range cellRange;
                            cellRange = (Excel.Range)excelworksheet.Cells[jj, ColRaz];
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
                                    if (regexParam.DiamDxDxD.IsMatch(temp))
                                    {
                                        diam = new Regex(@"(?<=[xх]\s*)\d+(?:[,\.]\d+)?(?=\s*[xх])", RegexOptions.IgnoreCase).Match(temp).Value;
                                        tolsh = new Regex(@"(?<=^)\d+(?:[,\.]\d+)?(?=\s*[xх])", RegexOptions.IgnoreCase).Match(temp).Value;
                                        metraj = new Regex(@"(?<=[xх]\s*)\d+(?:[,\.]\d+)?(?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                    }
                                    else if (regexParam.DiamDxD.IsMatch(temp))
                                    {
                                        diam = new Regex(@"(?<=^)\d+(?:[,\.]\d+)?(?=\s*[xх])", RegexOptions.IgnoreCase).Match(temp).Value;
                                        tolsh = new Regex(@"(?<=[xх]\s*)\d+(?:[,\.]\d+)?(?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                    }
                                    else if (regexParam.DiamD.IsMatch(temp))
                                        diam = regexParam.DiamD.Match(temp).Value;

                                    if (!string.IsNullOrEmpty(diam))
                                    {
                                        prim = temp;
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
                            else
                            if (cellRange.MergeArea.Columns.Count > 2)
                            {
                                if (ColRaz > 1)
                                {
                                    cellRange = (Excel.Range)excelworksheet.Cells[jj, 1];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                }
                                if (temp != "")
                                {
                                    if (regexParam.RegName.IsMatch(temp)) tab.Name = regexParam.RegName.Match(temp).Value;
                                    else tab.Name = "";
                                }
                                if (tab.Name != "")
                                {
                                    tab.Name = StringFirstUp(tab.Name);
                                    tab.Type = regexParam.RegType.Match(temp).Value;
                                }
                            }
                            if (jj < Max) ProcessChanged(jj);
                            else ProcessChanged(Max);
                        }
                    }

                    if (tab.StartRow > 0)
                    {
                        for (int j = 1; j < tab.StartRow; j++) //строки
                        {
                            for (int i = 1; i <= cCelCol; i++) //столбцы
                            {
                                Max = tab.StartRow * cCelCol;
                                SetMaxValProgressBar(Max);

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
                        ProcessChanged(Max);
                        int currentProgress = 0;
                        for (int j = tab.LastRowExcel + 10; j > tab.LastRowExcel; j--) //строки
                        {
                            for (int i = 1; i <= cCelCol; i++) //столбцы
                            {
                                Max = cCelRow / 5;
                                SetMaxValProgressBar(Max);

                                Excel.Range cellRange;
                                cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    FillInfoOrg(infoOrg, temp, regexParam);
                                }

                                if (currentProgress < Max) ProcessChanged(currentProgress++);
                                else ProcessChanged(Max);
                            }
                        }
                        ProcessChanged(Max);
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
            catch (Exception ex) { MessageBox.Show("Ошибка в функции ReedExcel() в " + this.ToString() + "\n\n" + ex.ToString()); }
        }

        private static string StringFirstUp(string StringIn)
        {
            string StringOut = "";
            if (StringIn.Length > 2)
                StringOut = StringIn.Substring(0, 1).ToUpper() + StringIn.Substring(1, StringIn.Length - 1).ToLower();
            else StringOut = StringIn;

            return StringOut;
        }

        private static void FillInfoOrg(InfoOrganization infoOrg, string temp, C_RegexParamProduct regexParam)
        {
            if (new Regex(@"(?<=Адрес\s*:\s*)[\s\w\.,\d]+", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.OrgAdress = new Regex(@"(?<=Адрес\s*:\s*)[\s\w\.,\d]+", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (regexParam.OrgMobileTelefon.IsMatch(temp))
            {
                infoOrg.OrgTel = regexParam.OrgMobileTelefon.Match(temp).Value;
            }
            if (new Regex(@"(?<=тел\\факс\s).*", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.OrgTel = new Regex(@"(?<=тел\\факс\s).*", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (regexParam.EMail.IsMatch(temp))
            {
                infoOrg.Email = regexParam.EMail.Match(temp).Value;
            }
            if (regexParam.Site.IsMatch(temp))
            {
                infoOrg.Site = regexParam.Site.Match(temp).Value;
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
