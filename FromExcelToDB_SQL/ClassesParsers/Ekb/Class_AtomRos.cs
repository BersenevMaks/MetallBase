using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;

namespace MetalBase
{
    class ClassAtomRos
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
            ReedExcel();
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

        private void ReedExcel()
        {
            InfoOrganization infoOrg = new InfoOrganization();
            infoOrg.SkladAdr = new List<string>();
            infoOrg.Manager = new List<string[]>();

            excelapp = new Excel.Application();
            try
            {
                orgname = "Атом-Рос";
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


                    int lastRow;
                    int ColName = 0, ColMark = 0, ColTU = 0, ColMera = 0, ColPrice = 0;

                    int Max = cCelCol * cCelRow;
                    SetMaxValProgressBar(Max);
                    //Поиск заголовков столбцов
                    int progress = 0;
                    for (int j = 4; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            int jj = j;
                            Excel.Range cellRange;
                            cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"^наименование", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColName = i; j = jj; tab.StartRow = jj; }
                                else if (new Regex(@"Марка", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColMark = i;  }
                                else if (new Regex(@"норм.*документ", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColTU = i; }
                                else if (new Regex(@"кол.*во", RegexOptions.IgnoreCase).IsMatch(temp))
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
                        if (ColName != 0)
                        {
                            for (int jj = tab.StartRow; jj <= cCelRow; jj++) //строки
                            {
                                Excel.Range cellRange;
                                cellRange = (Excel.Range)excelworksheet.Cells[jj, ColName];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    if (cellRange.MergeArea.Columns.Count > 2)
                                    {
                                        if (regexParam.RegName.IsMatch(temp)) tab.Name = regexParam.RegName.Match(temp).Value;
                                        else tab.Name = "";
                                    }
                                    if (!String.IsNullOrEmpty(tab.Name))
                                    {
                                    diam = ""; tolsh = ""; metraj = "";
                                    if (regexParam.DiamDxDxD.IsMatch(temp))
                                    {
                                        diam = new Regex(@"(?<=[xх]\s*)\d+(?:[,\.]\d+)?(?=\s*[xх])", RegexOptions.IgnoreCase).Match(temp).Value;
                                        tolsh = new Regex(@"(?<=^)\d+(?:[,\.]\d+)?(?=\s*[xх])", RegexOptions.IgnoreCase).Match(temp).Value;
                                        metraj = new Regex(@"(?<=[xх]\s*)\d+(?:[,\.]\d+)?(?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        if (tab.Name.ToLower() == "лист")
                                            try
                                            {
                                                if (double.Parse(diam) > double.Parse(tolsh))
                                                {
                                                    string str = diam;
                                                    diam = tolsh;
                                                    tolsh = str;
                                                }

                                            }
                                            catch { }
                                    }
                                    else if (regexParam.DiamDxD.IsMatch(temp))
                                    {
                                        diam = new Regex(@"(?<=^)\d+(?:[,\.]\d+)?(?=\s*[xх])", RegexOptions.IgnoreCase).Match(temp).Value;
                                        tolsh = new Regex(@"(?<=[xх]\s*)\d+(?:[,\.]\d+)?(?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        if (tab.Name.ToLower() == "лист")
                                            try
                                            {
                                                if (double.Parse(diam) > double.Parse(tolsh))
                                                {
                                                    string str = diam;
                                                    diam = tolsh;
                                                    tolsh = str;
                                                }

                                            }
                                            catch { }
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
                                            if (ColTU > 0)
                                            {
                                                cellRange = (Excel.Range)excelworksheet.Cells[jj, ColTU];
                                                if (cellRange.Value != null)
                                                    tmp = cellRange.Value.ToString().Trim();
                                                else tmp = "";
                                                if (tmp != "")
                                                {
                                                    tab.Standart = tmp;
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
                                                lastRow = dtProduct.Rows.Count - 1;
                                            else lastRow = 0;
                                            DataRow row = dtProduct.NewRow();
                                            row["Название"] = tab.Name;
                                            row["Тип"] = tab.Type;
                                            if (string.IsNullOrEmpty(row["Тип"].ToString())) row["Тип"] = "не указан";
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
                    
                    if (tab.StartRow > 0)
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
                                    if (new Regex(@"(?<=Адрес\s*:\s*)[\s\w\.,\d]+", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        infoOrg.OrgAdress = new Regex(@"(?<=Адрес\s*:\s*)[\s\w\.,\d]+", RegexOptions.IgnoreCase).Match(temp).Value;
                                    }
                                    if (new Regex(@"(?<=Склад\s*:\s*).+", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        infoOrg.SkladAdr.Add(new Regex(@"(?<=Склад\s*:\s*).+", RegexOptions.IgnoreCase).Match(temp).Value);
                                    }
                                    if (regexParam.OrgMobileTelefon.IsMatch(temp))
                                    {
                                        infoOrg.OrgTel = regexParam.OrgMobileTelefon.Match(temp).Value;
                                    }
                                    if (regexParam.EMail.IsMatch(temp))
                                    {
                                        infoOrg.Email = regexParam.EMail.Match(temp).Value;
                                    }
                                    if (new Regex(@"(?:т/ф\.)?\(\d+\)\s*(?:-?\d+)+", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        string[] strManager = new string[2];
                                        strManager[0] = new Regex(@"(?:т/ф\.)?\s*\(\d+\)\s*(?:-?\d+)+[,\s*\w+\.\d+]+\s*(?=\s\w+\s*$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        strManager[1] = new Regex(@"(?<=\d+\s)\w+(?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        infoOrg.Manager.Add(strManager);
                                    }
                                }

                                if (i * j < Max) ProcessChanged(i * j);
                                else ProcessChanged(Max);
                            }
                        }
                    ProcessChanged(Max);
                }

                if (isExcelOpen)
                {
                    excelappworkbook.Close();
                    excelapp.Quit();
                }

                SetInfoOrganization(infoOrg);
                workCompleted(dtProduct);


            }
            catch (Exception ex) { MessageBox.Show("Ошибка в функции ReedExcel() в " + this.ToString() + "\n\n" + ex.ToString()); }
        }

        private List<double> getIncrementingMassiv(string[] strParams)
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

        public event Action<DataTable> workCompleted;

    }
}
