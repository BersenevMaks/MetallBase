using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace MetalBase.ClassesParsers.Chel
{
    class Class_MetallInvestExcel
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
                orgname = "МеталлИнвест";
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
                List<double> dDiam;
                List<string> dTolsh;
                List<double> dMetraj;
                List<string> razmers;
                List<string> marks;
                List<string> prices;

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

                    int ColDiam = 0, ColDlina = 0, ColMera = 0, ColGost = 0, ColTolsh = 0, ColMark = 0, ColPrice = 0;
                    var dtm = new HelpClasses.Class_DTM();
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
                                if (new Regex(@"^наименование", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    tab = new C_InfoTable
                                    {
                                        StartCol = i,
                                        StartRow = jj
                                    };
                                    tabs.Add(tab);
                                    //j = cCelRow;
                                    //break;
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
                        ColGost = 0; ColDiam = 0; ColMera = 0; ColTolsh = 0; ColPrice = 0;
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
                        marks = new List<string>();
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            cellRange = (Excel.Range)excelworksheet.Cells[jjj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"\bнаименование\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColDiam = i;
                                    continue;
                                }
                                if (new Regex(@"\bстенка\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColTolsh = i;
                                    continue;
                                }
                                //if (new Regex(@"\bдлина\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                //{
                                //    ColDlina = i;
                                //    continue;
                                //}
                                //if (new Regex(@"\bкол.*во\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                //{
                                //    ColMera = i;
                                //    continue;
                                //}
                                //if (new Regex(@"\bмарка\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                //{
                                //    ColMark = i;
                                //    continue;
                                //}
                                //if (new Regex(@"\bту\b|\bгост\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                //{
                                //    ColGost = i;
                                //    continue;
                                //}
                                if (new Regex(@"\bцена\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    
                                    ColPrice = i;
                                    cellRange = (Excel.Range)excelworksheet.Cells[jjj+1, i];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        tab.Mark = regexParam.RegMark.Match(temp).Value;
                                        if (!string.IsNullOrEmpty(tab.Mark))
                                        {
                                            marks.Add(tab.Mark);
                                            cellRange = (Excel.Range)excelworksheet.Cells[jjj + 1, i + 1];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                if (regexParam.RegMark.IsMatch(temp))
                                                {
                                                    tab.Mark = regexParam.RegMark.Match(temp).Value;
                                                    marks.Add(tab.Mark);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[jjj + 2, i];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                tab.Mark = regexParam.RegMark.Match(temp).Value;
                                                if (!string.IsNullOrEmpty(tab.Mark))
                                                {
                                                    marks.Add(tab.Mark);
                                                    cellRange = (Excel.Range)excelworksheet.Cells[jjj + 2, i + 1];
                                                    if (cellRange.Value != null)
                                                        temp = cellRange.Value.ToString().Trim();
                                                    else temp = "";
                                                    if (temp != "")
                                                    {
                                                        if (regexParam.RegMark.IsMatch(temp))
                                                        {
                                                            tab.Mark = regexParam.RegMark.Match(temp).Value;
                                                            marks.Add(tab.Mark);
                                                        }
                                                    }
                                                }
                                            }
                                    
                                        }
                                    }
                                    continue;
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
                            diam = ""; tolsh = ""; metraj = ""; mera = ""; price = ""; name = ""; prim = ""; standart = ""; mark = "";
                            razmers = new List<string>();
                            dTolsh = new List<string>();
                            prices = new List<string>();
                            if (ColDiam > 0)
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[j, ColDiam];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    prim = temp;
                                    if (cellRange.MergeArea.Columns.Count > 2)
                                    {
                                        tab.Name = regexParam.RegName.Match(temp).Value;
                                        if (!string.IsNullOrEmpty(tab.Name)) temp = temp.Replace(tab.Name, "");
                                        foreach (Match m in regexParam.RegTU.Matches(temp))
                                        {
                                            if (string.IsNullOrEmpty(tab.Standart))
                                                tab.Standart = m.Value;
                                            else tab.Standart += ", " + m.Value;
                                            temp = temp.Replace(m.Value, "");
                                        }

                                        foreach (Match m in regexParam.RegType.Matches(temp))
                                        {
                                            if (string.IsNullOrEmpty(tab.Type))
                                                tab.Type = m.Value;
                                            else tab.Type += " " + m.Value;
                                        }
                                        tab.Standart = "";
                                        tab.Type = "";
                                    }
                                    else
                                    {
                                        name = regexParam.RegName.Match(temp).Value;
                                        temp = temp.Replace('*', 'x');
                                        if (string.IsNullOrEmpty(name)) name = tab.Name;
                                        if (regexParam.DiamDxDxD_Only.IsMatch(temp))
                                            foreach (Match m in regexParam.DiamDxDxD_Only.Matches(temp))
                                                razmers.Add(m.Value);
                                        else if (regexParam.DiamDxD_Only.IsMatch(temp))
                                            foreach (Match m in regexParam.DiamDxD_Only.Matches(temp))
                                                razmers.Add(m.Value);
                                        else if (new Regex(@"(?<=\d),\s+(?=\d|\s*$)", RegexOptions.IgnoreCase).IsMatch(temp))
                                            foreach (string m in new Regex(@"(?<=\d),\s+(?=\d|\s*$)", RegexOptions.IgnoreCase).Split(temp))
                                                razmers.Add(m);
                                        else if (temp.Contains("-") && (name.ToLower().Contains("лист") || name.ToLower().Contains("арматура")))
                                            foreach (Match m in new Regex(@"\d+(?:[\.,]\d)?\s*\-\s*\d+(?:[\.,]\d)?", RegexOptions.IgnoreCase).Matches(temp))
                                                foreach (double d in GetIncrementingMassiv(m.Value.Split('-')))
                                                    razmers.Add(d.ToString());
                                        else
                                            razmers.Add(regexParam.DiamD_Only.Match(temp).Value);
                                        type = regexParam.RegType.Match(temp).Value;
                                        //if (!string.IsNullOrEmpty(type)) type = regexParam.SetRightEnding(type, name);
                                        if (regexParam.RegTypeShveller.IsMatch(temp))
                                            type = regexParam.RegTypeShveller.Match(temp).Value;
                                    }
                                }
                            }
                            if (ColTolsh > 0)
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[j, ColTolsh];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    prim += "," + temp;
                                    string[] vs = temp.Split('-');
                                    if(vs.Length>1)
                                    {
                                        foreach(string s in vs)
                                        {
                                            dTolsh.Add(s);
                                        }
                                    }
                                    else
                                    {
                                        vs = temp.Split(';');
                                        if (vs.Length > 1)
                                        {
                                            foreach (string s in vs)
                                            {
                                                dTolsh.Add(s);
                                            }
                                        }
                                        else
                                        {
                                            vs = temp.Split(',');
                                            if (vs.Length>2)
                                            {
                                                foreach(string s in vs)
                                                {
                                                    dTolsh.Add(s);
                                                }
                                            }
                                            else
                                            {
                                                dTolsh.Add(temp);
                                            }
                                        }
                                    }
                                }
                            }
                            
                            if (ColPrice > 0)
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[j, ColPrice];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    prices.Add(new Regex(@"\d+(?:[\.,]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value);
                                    prim += ", " + temp;
                                    if (marks.Count == 2)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[j, ColPrice + 1];
                                        if (cellRange.MergeArea.Rows.Count>1)
                                            cellRange = (Excel.Range)excelworksheet.Cells[cellRange.MergeArea.Row, ColPrice + 1];
                                        //if (cellRange.MergeArea.Columns.Count>1)
                                        //    cellRange = (Excel.Range)excelworksheet.Cells[j, ColPrice];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        prices.Add(new Regex(@"\d+(?:[\.,]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value);
                                        prim += ", " + temp;
                                    }
                                }
                            }

                            if (razmers.Count>0 &&(dTolsh.Count>0 || prices.Count>0) && (!string.IsNullOrEmpty(tab.Name) || !string.IsNullOrEmpty(name)))
                            {
                                foreach (string ts in razmers)
                                {
                                    dtm.CalcDTM(ts, name: name);
                                    diam = dtm.D();
                                    if (!string.IsNullOrEmpty(diam))
                                    {
                                        if (dTolsh.Count > 0)
                                            metraj = dtm.T();
                                        else
                                        {
                                            tolsh = dtm.T();
                                            metraj = dtm.M();
                                        }
                                        if (dTolsh.Count > 0)
                                            foreach (string sTolsh in dTolsh)
                                            {
                                                for (int i = 0; i < prices.Count; i++)
                                                {
                                                    if (i < marks.Count)
                                                        mark = marks[i];
                                                    else if (marks.Count > 0)
                                                        mark = marks[marks.Count - 1];
                                                    else mark = "";
                                                    DataRow row = dtProduct.NewRow();
                                                    if (!string.IsNullOrEmpty(name))
                                                        row["Название"] = StringFirstUp(name);
                                                    else row["Название"] = StringFirstUp(tab.Name);
                                                    if (string.IsNullOrEmpty(type))
                                                        row["Тип"] = tab.Type;
                                                    else row["Тип"] = type;
                                                    if (string.IsNullOrEmpty(row["Тип"].ToString())) row["Тип"] = "тип не указан";
                                                    else row["Тип"] = row["Тип"].ToString().ToLower();
                                                    row["Диаметр (высота), мм"] = diam;
                                                    row["Толщина (ширина), мм"] = sTolsh;
                                                    row["Метраж, м (длина, мм)"] = metraj;
                                                    row["Мерность (т, м, мм)"] = mera;
                                                    if (String.IsNullOrEmpty(mark))
                                                        row["Марка"] = tab.Mark;
                                                    else row["Марка"] = mark;
                                                    if (String.IsNullOrEmpty(tab.Standart))
                                                        row["Стандарт"] = standart;
                                                    else row["Стандарт"] = tab.Standart;
                                                    row["Класс"] = "";
                                                    row["Цена"] = prices[i];
                                                    row["Примечание"] = prim;
                                                    dtProduct.Rows.Add(row);
                                                }
                                            }
                                        else
                                            for (int i = 0; i < prices.Count; i++)
                                            {
                                                if (i < marks.Count)
                                                    mark = marks[i];
                                                else if (marks.Count > 0)
                                                    mark = marks[marks.Count - 1];
                                                else mark = "";
                                                DataRow row = dtProduct.NewRow();
                                                if (!string.IsNullOrEmpty(name))
                                                    row["Название"] = name;
                                                else row["Название"] = tab.Name;
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
                                                if (String.IsNullOrEmpty(tab.Standart))
                                                    row["Стандарт"] = standart;
                                                else row["Стандарт"] = tab.Standart;
                                                row["Класс"] = "";
                                                row["Цена"] = prices[i];
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
            if (regexParam.OrgAdresFully.IsMatch(temp))
            {
                infoOrg.OrgAdress = regexParam.OrgAdresFully.Match(temp).Value;
            }
            if (regexParam.OrgMobileTelefon_tel_X_XXX_XXX_XX_XX.IsMatch(temp))
            {
                foreach (Match m in regexParam.OrgMobileTelefon_tel_X_XXX_XXX_XX_XX.Matches(temp))
                {
                    if (string.IsNullOrEmpty(infoOrg.OrgTel)) infoOrg.OrgTel = m.Value;
                    else infoOrg.OrgTel += ", " + m.Value;
                }
            }
            if (new Regex(@"\d{3}\-\d{2}\-\d{2}", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                foreach (Match m in new Regex(@"\d{3}\-\d{2}\-\d{2}", RegexOptions.IgnoreCase).Matches(temp))
                {
                    if (string.IsNullOrEmpty(infoOrg.OrgTel)) infoOrg.OrgTel = m.Value;
                    else infoOrg.OrgTel += ", " + m.Value;
                }
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

            if (new Regex(@"[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(?>(.[a-z0-9-]+))*(.[a-z]{2,4})", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                foreach (Match m in new Regex(@"[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(?>(.[a-z0-9-]+))*(.[a-z]{2,4})", RegexOptions.IgnoreCase).Matches(temp))
                    if (String.IsNullOrEmpty(infoOrg.Email)) infoOrg.Email = m.Value;
                    else infoOrg.Email += "; " + m.Value;
            }
            if (new Regex(@"w+\.[\w\-]+(?:\.[\w\-]+)?", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.Site = new Regex(@"w+\.[\w\-]+(?:\.[\w\-]+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                if(string.IsNullOrEmpty(infoOrg.Site)) 
                    infoOrg.Site = new Regex(@"w+\.[\w\-]+(?:\.[\w\-]+)?", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (string.IsNullOrEmpty(infoOrg.Site))
                if (new Regex(@"(?<=https?(?::?//)\s*).*\.(?:ru|com|info|org)", RegexOptions.IgnoreCase).IsMatch(temp))
                {

                    infoOrg.Site = new Regex(@"(?<=https?(?::?//)\s*).*\.(?:ru|com|info|org)", RegexOptions.IgnoreCase).Match(temp).Value;
                }
            //if (new Regex(@"г\.\w+[,\w\.\s\d]+(?=\s*$)", RegexOptions.IgnoreCase).IsMatch(temp))
            //{
            //    infoOrg.SkladAdr.Add(new Regex(@"г\.\w+[,\w\.\s\d]+(?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value);
            //}
            //if (new Regex(@"(?<=воните\s*(?:[\+\d]{0,2}[\s\(]{0,2}[\)\d\s,-]{3,})+\s)\w{4,}(?:\s*\w+)", RegexOptions.IgnoreCase).IsMatch(temp))
            //{
            //    string[] manager = new string[]
            //    {
            //        new Regex(@"(?<=\+?\d\s*\(\d+\)[\s\d\-]+\s)\w+(?:\s*\w+)?", RegexOptions.IgnoreCase).Match(temp).Value, //имя
            //        new Regex(@"(?:[\+\d]{0,2}[\s\(]{0,2}[\)\d\s,-]{3,})+", RegexOptions.IgnoreCase).Match(temp).Value, //телефон
            //        "" //email
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

        private List<double> GetIncrementingTolshMass(string[] strParams)
        {
            List<double> MassTolsh = new List<double>();
            List<double> tolshSTD = new List<double> { 2,3,4,5,7,8,10,12,14,16,18,20,22,26,30,32,36,40,
                45,50,55,60,65,70,80,90,100,110,120,130,140,150,160,180 };
            return MassTolsh;
        }

        public event Action<int> ProcessChanged; //установить текущее значение прогрессбара

        public event Action<int> SetMaxValProgressBar; //установить максимальное значение для прогрессбара

        public event Action<InfoOrganization> SetInfoOrganization;

        public event Action<DataTable> WorkCompleted;
    }
}
