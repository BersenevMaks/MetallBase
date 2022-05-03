using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.IO;

namespace MetalBase
{
    class Class_A_Group_Trub
    {
        private string filePath;

        public void Set(string Path)
        { filePath = Path; }

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
                var regexParam = new C_RegexParamProduct();

                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    var tab = new C_InfoTable();
                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;

                    string[] diam, tolsh, metraj;
                    diam = new string[] { "" };
                    tolsh = new string[] { "" };
                    metraj = new string[] { "" };
                    string[] diamTolsh;
                    List<double> Ddiam = new List<double>(), Dtolsh = new List<double>(), Dmetraj = new List<double>();
                    List<double> ch = new List<double>();
                    int lastRow;
                    int firstRow = 0;

                    int Max = cCelCol * cCelRow;
                    SetMaxValProgressBar(Max);

                    for (int i = 1; i <= cCelCol; i++) //столбцы
                    {
                        for (int j = 1; j <= cCelRow; j++) //строки
                        {
                            int jj = j; //запоминаем строку, чтобы потом вернуться если че
                            Excel.Range cellRange;
                            cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (regexParam.RegName.IsMatch(temp))
                                {
                                    tab = new C_InfoTable();
                                    tab.Name = regexParam.RegName.Match(temp).Value;
                                    tab.Standart = regexParam.RegTU.Match(temp).Value;
                                    jj++;
                                    //ищем гост
                                    cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        if (regexParam.RegTU.IsMatch(temp))
                                        {
                                            foreach (Match m in regexParam.RegTU.Matches(temp))
                                            {
                                                if (tab.Standart != "") tab.Standart += "; " + regexParam.RegTU.Match(temp).Value;
                                                else tab.Standart = regexParam.RegTU.Match(temp).Value;
                                            }
                                        }
                                    }

                                    for (jj++; jj < cCelRow; jj++)
                                    {
                                        Ddiam = new List<double>();
                                        Dtolsh = new List<double>();
                                        Dmetraj = new List<double>();
                                        Dmetraj.Add(0.0);
                                        price = "";
                                        prim = "";
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                                        if (cellRange.Value != null)
                                            tmp = cellRange.Value.ToString().Trim();
                                        else tmp = "";
                                        if (tmp != "")
                                        {
                                            prim = tmp;
                                            if (new Regex(@"(?:\d+(?:[,\.]\d+)?)?\s*[xх]\s*\d+(?:[,\.]\d+)?\s*[,;]\s*(?:\d+(?:[,\.]\d+)?)?\s*[xх]\s*\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).IsMatch(tmp))
                                            {
                                                diamTolsh = tmp.Split(';', ',');
                                            }
                                            else if (new Regex(@"\d+(?:[,\.]\d+)?\s*[xх]\s*\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).IsMatch(tmp))
                                            {
                                                diamTolsh = new string[] { tmp };
                                            }
                                            else break;
                                        }
                                        else diamTolsh = new string[] { "" };

                                        if (diamTolsh.Length > 0)
                                            if (diamTolsh[0] != "0")
                                            {
                                                for (int s = 0; s < diamTolsh.Length; s++)
                                                {
                                                    string[] d_t = new string[]{
                                                    new Regex(@".*(?=[xх])", RegexOptions.IgnoreCase).Match(diamTolsh[s]).Value.Trim(),
                                                    new Regex(@"(?<=[xх])\s*\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).Match(diamTolsh[s]).Value.Trim()
                                                    };
                                                    if (d_t.Length > 1)
                                                    {
                                                        if (d_t[0].Length > 0)
                                                        {
                                                            string[] dia = d_t[0].Split('-');
                                                            List<double> tlist = getIncrementingMassiv(dia);
                                                            foreach (double d in tlist)
                                                                Ddiam.Add(d);
                                                        }
                                                        if (d_t[1].Length > 0)
                                                        {
                                                            d_t[1] = d_t[1].Replace('.', ',');
                                                            Dtolsh.Add(Convert.ToDouble(d_t[1]));
                                                        }
                                                    }
                                                    cellRange = (Excel.Range)excelworksheet.Cells[jj, i + 1];
                                                    if (cellRange.Value != null)
                                                        tmp = cellRange.Value.ToString().Trim();
                                                    else tmp = "";
                                                    if (tmp != "")
                                                    {
                                                        price = new Regex(@"[\d\s]+(?:\s*[,.]\s*[\d\s]+)?", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    }
                                                }

                                                if (Ddiam.Count > 0)
                                                    for (int d = 0; d < Ddiam.Count; d++)
                                                        for (int t = 0; t < Dtolsh.Count; t++)
                                                            for (int m = 0; m < Dmetraj.Count; m++)
                                                            {
                                                                if (d == 0 && t == 0 && m == 0)
                                                                {
                                                                    if (firstRow == 0) firstRow = jj - 2;
                                                                    dtProduct.Rows.Add();
                                                                    lastRow = dtProduct.Rows.Count - 1;
                                                                    dtProduct.Rows[lastRow]["Название"] = tab.Name;
                                                                    dtProduct.Rows[lastRow]["Тип"] = tab.Type;
                                                                    dtProduct.Rows[lastRow]["Стандарт"] = tab.Standart;
                                                                    dtProduct.Rows[lastRow]["Марка"] = tab.Mark;
                                                                    dtProduct.Rows[lastRow]["Цена"] = price;
                                                                    dtProduct.Rows[lastRow]["Примечание"] = prim;
                                                                    if (Ddiam[0] != 0) dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = Ddiam[0];
                                                                    if (Dtolsh[0] != 0) dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = Dtolsh[0];
                                                                    if (Dmetraj[0] != 0) dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = Dmetraj[0];
                                                                }
                                                                else
                                                                {
                                                                    if (dtProduct.Rows.Count > 0)
                                                                        lastRow = dtProduct.Rows.Count - 1;
                                                                    else lastRow = 0;
                                                                    DataRow row = dtProduct.NewRow();
                                                                    row["Название"] = dtProduct.Rows[lastRow]["Название"];
                                                                    row["Тип"] = dtProduct.Rows[lastRow]["Тип"];
                                                                    if (Ddiam[d] != 0) row["Диаметр (высота), мм"] = Ddiam[d];
                                                                    if (Dtolsh[t] != 0) row["Толщина (ширина), мм"] = Dtolsh[t];
                                                                    if (Dmetraj[m] != 0) row["Метраж, м (длина, мм)"] = Dmetraj[m];
                                                                    row["Мерность (т, м, мм)"] = dtProduct.Rows[lastRow]["Мерность (т, м, мм)"];
                                                                    row["Марка"] = dtProduct.Rows[lastRow]["Марка"];
                                                                    row["Стандарт"] = dtProduct.Rows[lastRow]["Стандарт"];
                                                                    row["Класс"] = dtProduct.Rows[lastRow]["Класс"];
                                                                    row["Цена"] = dtProduct.Rows[lastRow]["Цена"];
                                                                    row["Примечание"] = dtProduct.Rows[lastRow]["Примечание"];
                                                                    dtProduct.Rows.Add(row);
                                                                }
                                                            }
                                            }
                                        if (i * jj < Max) ProcessChanged(i * jj);
                                        else ProcessChanged(Max);
                                    }
                                }
                                else continue;
                                j = jj - 1;
                                if (i * j < Max) ProcessChanged(i * j);
                                else ProcessChanged(Max);
                            }
                            else continue;
                            if (i * j < Max) ProcessChanged(i * j);
                            else ProcessChanged(Max);
                        }
                    }
                    ProcessChanged(Max);
                    if (firstRow > 0)
                        for (int j = 1; j < firstRow; j++) //строки
                        {
                            for (int i = 1; i <= cCelCol; i++) //столбцы
                            {
                                Max = firstRow * cCelCol;
                                SetMaxValProgressBar(Max);

                                Excel.Range cellRange;
                                cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    if (new Regex(@"(?<=Офис\s*:\s*)[\s\w\.,\d]+(?=,\sтел)", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        infoOrg.OrgAdress = new Regex(@"(?<=Офис\s*:\s*)[\s\w\.,\d]+(?=,\sтел)", RegexOptions.IgnoreCase).Match(temp).Value;
                                    }
                                    if (new Regex(@"(?<=Склад\s*:\s*).+", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        infoOrg.SkladAdr.Add(new Regex(@"(?<=Склад\s*:\s*).+", RegexOptions.IgnoreCase).Match(temp).Value);
                                    }
                                    if (new Regex(@"(?<=\bтел.*\s)[\(\d\s\)-]+", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        infoOrg.OrgTel = new Regex(@"(?<=\bтел.*\s)[\(\d\s\)-]+", RegexOptions.IgnoreCase).Match(temp).Value;
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
            catch (Exception ex) { MessageBox.Show("Ошибка в функции ReedExcel() в "+this.ToString()+"\n\n" + ex.ToString()); }
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
