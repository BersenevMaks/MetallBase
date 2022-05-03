using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;
using MetalBase.HelpClasses;

namespace MetalBase.ClassesParsers.Chel
{
    class Class_MetallTorgExcel
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
                infoOrg.OrgName = StringFirstUp(orgname);
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
                var dtm = new Class_DTM();
                int indexOfLastRow = 1;
                int workListIndex = 0;

                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    if (excelworksheet.Name.ToString().ToLower().Contains("формул") ||
                        excelworksheet.Name.ToString().ToLower().Contains("сплав")) continue;
                    var listName = excelworksheet.Name.ToString();
                    
                    var tab = new C_InfoTable();
                    var naaame = excelworksheet.Name;
                    List<C_InfoTable> tabs = new List<C_InfoTable>();
                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol <= 10) cCelCol = 10;
                    if (cCelCol > 10) cCelCol = 25;
                    //cCelCol = 10;

                    int ColDiam = 0, ColMera = 0, ColGost = 0, ColDlina = 0, ColMark = 0, ColPrice = 0, ColName = 0;
                    int ColTolsh = 0;

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
                                //если на листе Трубы
                                if (listName.ToLower().Contains("труб"))
                                {
                                    if (new Regex(@"наименование", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        tab = new C_InfoTable
                                        {
                                            StartCol = i,
                                            StartRow = jj
                                        };
                                        tabs.Add(tab);
                                        j = cCelRow;
                                        break;
                                    }
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
                        if (workListIndex < 1) workListIndex = excelworksheet.Index;
                        ColGost = 0; ColDiam = 0; ColMera = 0; ColGost = 0; ColMark = 0;
                        name = ""; type = "";
                        Excel.Range cellRange;
                        tab = tabs[k];
                        int endRow = cCelRow;
                        if (k < tabs.Count - 1)   // определение последней строки в текущей минитаблице
                        {
                            if (tab.StartCol == tabs[k + 1].StartCol)
                                endRow = tabs[k + 1].StartRow - 1;
                        }
                        else if (k < tabs.Count - 2)
                        {
                            if (tab.StartCol == tabs[k + 2].StartCol)
                                endRow = tabs[k + 2].StartRow - 1;
                        }
                        else if (k < tabs.Count - 3)
                        {
                            if (tab.StartCol == tabs[k + 3].StartCol)
                                endRow = tabs[k + 3].StartRow - 1;
                        }

                        // поиск имени продукции
                        ProcessChanged(0); //установить прогрессбар на 0
                        SetMaxValProgressBar(cCelCol); Max = cCelCol; //установить максимум для прогрессбара
                        progress = 0;
                        int jjj = tab.StartRow;
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            //если на листе Трубы
                            if (listName.ToLower().Contains("труб"))
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[jjj, i];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    if (new Regex(@"мерность", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        ColMera = i;
                                        continue;
                                    }
                                    if (new Regex(@"цена.*тн)", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        ColPrice = i;
                                        continue;
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
                            diam = ""; tolsh = ""; metraj = ""; mera = ""; mark = ""; price = ""; name = ""; prim = ""; standart = "";
                            if (ColMark > 0)
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[j, ColMark];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    if (regexParam.RegName.IsMatch(temp))
                                    {
                                        name = StringFirstUp(regexParam.RegName.Match(temp).Value).Trim();
                                    }
                                    else if (new Regex(@"\bзакладная\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        name = StringFirstUp(new Regex(@"\bзакладная\b", RegexOptions.IgnoreCase).Match(temp).Value).Trim();
                                    }
                                    if (new Regex(@"\bкосынк\w\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        name = "Косынка";
                                    }
                                    if (new Regex(@"\bкруг\w\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        name = "Круг";
                                        tab.Name = "Круг";
                                    }
                                    if (new Regex(@"\bзаглушк\w\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        name = "Заглушка";
                                    }
                                    prim = temp;

                                    if (!string.IsNullOrEmpty(name))
                                        temp = temp.Replace(name, string.Empty).Trim();
                                    if (!string.IsNullOrEmpty(type))
                                        temp = temp.Replace(type, string.Empty).Trim();

                                    standart = regexParam.RegTU.Match(temp).Value;
                                    if (!string.IsNullOrEmpty(standart)) temp = temp.Replace(standart, string.Empty).Trim();
                                    //mark = new Regex(@"", RegexOptions.IgnoreCase).Match(temp).Value;
                                    mark = regexParam.RegMark.Match(temp).Value;
                                    if (!string.IsNullOrEmpty(mark)) temp = temp.Replace(mark, string.Empty).Trim();

                                    if (string.IsNullOrEmpty(mark))
                                        mark = new Regex(@"(?<=\s)\D(?:\d+?\w+?|\w+?\d+?|\-)+(?:[\w\d\-]+)?", RegexOptions.IgnoreCase).Match(temp).Value;


                                    dtm.CalcDTM(temp, type);
                                    diam = dtm.D();
                                    tolsh = dtm.T();
                                    metraj = dtm.M();

                                }
                            }
                            if (!string.IsNullOrEmpty(diam))
                            {
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

                                if (!String.IsNullOrEmpty(diam))
                                {
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
                                    row["Цена"] = price;
                                    row["Примечание"] = prim;
                                    dtProduct.Rows.Add(row);
                                    indexOfLastRow = j;
                                }
                            }
                            if (progress < Max) ProcessChanged(progress++);
                            else ProcessChanged(Max);
                        }
                    }

                    //поиск информации об организации ТОЛЬКО на ПЕРВОМ листе
                    if (tabs.Count > 0 && dtProduct.Rows.Count > 0 && excelworksheet.Index == workListIndex)
                    {
                        Max = (tabs[0].StartRow - 1) * cCelCol;
                        SetMaxValProgressBar(Max);
                        progress = 0;
                        Excel.Range cellRange;
                        for (int j = 1; j <= tabs[0].StartRow; j++) //строки
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
                                if (progress < Max) ProcessChanged(progress++);
                                else ProcessChanged(Max);
                            }
                        }
                        for (int j = cCelRow; j >= cCelRow - (cCelRow - indexOfLastRow); j--) //строки
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
                            }
                        }
                    }
                }

                if (isExcelOpen && excelappworkbook != null && excelapp != null)
                {
                    excelapp.DisplayAlerts = false;
                    excelappworkbook.Close();
                    excelapp.Quit();
                }

                SetInfoOrganization(infoOrg);
                WorkCompleted(dtProduct);

            }
            catch (Exception ex)
            {
                WorkCompleted(dtProduct);
                if (new Regex(@"защищ\w+\s*лист", RegexOptions.IgnoreCase).IsMatch(ex.ToString()))
                {
                    MessageBox.Show("Файл под паролем!\nЛист защищен!\nСнимите защиту листа!");
                }
                else
                    MessageBox.Show("Ошибка в функции ReedExcel() в " + this.ToString() + "\n\n" + ex.ToString());
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

        private static void FillInfoOrg(InfoOrganization infoOrg, string temp, C_RegexParamProduct regexParam, Excel.Worksheet worksheet, int j, int i)
        {
            temp = temp.Replace("\r\a", " ").Replace("\r", " ").Replace("\n", " ");
            if (new Regex(@"\d{6}(?:,?\s*\w+(?:\s*\w+)?)?,?\s*г.\s*\w+(?:\s*\w+)*,?\s*ул.\s*\w+(?:\s*\w+)*,?\s*(?:д.)?\s*[\d\w]+(?:,?\s*оф.?\s*[\d\-]+)?", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.OrgAdress = new Regex(@"\d{6}(?:,?\s*\w+(?:\s*\w+)?)?,?\s*г.\s*\w+(?:\s*\w+)*,?\s*ул.\s*\w+(?:\s*\w+)*,?\s*(?:д.)?\s*[\d\w]+(?:,?\s*оф.?\s*[\d\-]+)?", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (new Regex(@"(?<=тел.*факс\s)[\(\)\d\s\-]*", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.OrgTel = new Regex(@"(?<=тел.*факс\s)[\(\)\d\s\-]*", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (regexParam.EMail.IsMatch(temp))
            {
                infoOrg.Email = regexParam.EMail.Match(temp).Value;
            }
            if (string.IsNullOrEmpty(infoOrg.Email)) infoOrg.Email = new Regex(@"(?<=mail.*:).*@.*\.\w{2,4}(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
            //if (regexParam.Site.IsMatch(temp))
            //{
            //    infoOrg.Site = regexParam.Site.Match(temp).Value;
            //}
            else if (new Regex(@"(?:https?:?//|(?:https?:?//)?www\.)?(?:[а-яёa-z0-9_-]{1,32}(?::[а-яёa-z0-9_-]{1,32})?)?(?:(?:[а-яёa-z0-9-]{1,128}\.)+(?:ru|su|com|net|org|mil|edu|arpa|gov|biz|info|aero|inc|name|рф))", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.Site = new Regex(@"(?:https?:?//|(?:https?:?//)?www\.)?(?:[а-яёa-z0-9_-]{1,32}(?::[а-яёa-z0-9_-]{1,32})?)?(?:(?:[а-яёa-z0-9-]{1,128}\.)+(?:ru|su|com|net|org|mil|edu|arpa|gov|biz|info|aero|inc|name|рф))", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (new Regex(@"(?<=ИНН(?:/КПП)?\s*:?\s*|инн\s*|кпп\s*)\d{9,15}(?:\s*/\s*\d{9,15})?", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                if (string.IsNullOrEmpty(infoOrg.Inn_Kpp))
                    infoOrg.Inn_Kpp = new Regex(@"(?<=ИНН(?:/КПП)?\s*:?\s*|инн\s*|кпп\s*)\d{9,15}(?:\s*/\s*\d{9,15})?", RegexOptions.IgnoreCase).Match(temp).Value;
                else infoOrg.Inn_Kpp += "/" + new Regex(@"(?<=ИНН(?:/КПП)?\s*:?\s*|инн\s*|кпп\s*)\d{9,15}(?:\s*/\s*\d{9,15})?", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (new Regex(@"(?<=Р.\s*сч\s*|рас.\s*сч|р[\\/]с\s*)[\d\s]+(?=\s[\w]|\s\s|$|,)", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.r_s = new Regex(@"(?<=Р.\s*сч\s*|рас.\s*сч|р[\\/]с\s*)[\d\s]+(?=\s[\w]|\s\s|$|,)", RegexOptions.IgnoreCase).Match(temp).Value.Trim().Replace(" ", "");
            }
            if (new Regex(@"(?<=к(?:ор)?\s*.сч\s*)\d+|(?<=к(?:ор)?[\\/]с\s*)[\d\s]+(?=\s[\w]|\s\s|$|,)", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.k_s = new Regex(@"\s+", RegexOptions.IgnoreCase).Replace(
                    new Regex(@"(?<=к(?:ор)?\s*.сч\s*)\d+|(?<=к(?:ор)?[\\/]с\s*)[\d\s]+(?=\s[\w]|\s\s|$|,)", RegexOptions.IgnoreCase).Match(temp).Value, "");
            }
            if (new Regex(@"(?<=\bбик\b\s*)[\d\s]+(?=\s[\w]|\s\s|$|,)", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.BIK = new Regex(@"\s+", RegexOptions.IgnoreCase).Replace(
                    new Regex(@"(?<=\bбик\b\s*)[\d\s]+(?=\s[\w]|\s\s|$|,)", RegexOptions.IgnoreCase).Match(temp).Value.Trim().Replace(" ", ""), "");
            }
            if (new Regex(@"(?<=адрес\s*склада\s*:\s+).*(?=\s*$)", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.SkladAdr.Add(new Regex(@"(?<=адрес\s*склада\s*:\s+).*(?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value);
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

        public event Action<int> ProcessChanged; //установить текущее значение прогрессбара

        public event Action<int> SetMaxValProgressBar; //установить максимальное значение для прогрессбара

        public event Action<InfoOrganization> SetInfoOrganization;

        public event Action<DataTable> WorkCompleted;
    }
}
