using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.IO;
using MetalBase.HelpClasses;

namespace MetalBase.ClassesParsers.Chel
{
    class Class_MetallGradExcel
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
                infoOrg.OrgName = StringFirstUp(orgname) + "_excel";
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
                int indexOfLastRow = 1;

                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    if (excelworksheet.Name.ToString().ToLower().Contains("формул") ||
                        excelworksheet.Name.ToString().ToLower().Contains("сплав")) continue;
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
                        ColGost = 0; ColDiam = 0; ColMera = 0; ColGost = 0;
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
                                if (new Regex(@"наименование", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColName = i;
                                    continue;
                                }
                                if (new Regex(@"мерность", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColMera = i;
                                    continue;
                                }
                                if (new Regex(@"цена", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColPrice = i;
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
                            diam = ""; tolsh = ""; metraj = ""; mera = ""; mark = ""; price = ""; name = ""; prim = ""; standart = "";
                            if (ColName > 0)
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[j, ColName];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    if (regexParam.RegName.IsMatch(temp) || new Regex(@"круги", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        name = StringFirstUp(regexParam.RegName.Match(temp).Value).Trim();
                                        tab.Name = StringFirstUp(regexParam.RegName.Match(temp).Value).Trim();

                                        type = regexParam.RegType.Match(temp).Value;
                                        tab.Type = regexParam.RegType.Match(temp).Value;
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

                                    //diamDxDxD
                                    if (new Regex(@"\d+(?:[,.]\d+)?\s*[xх\*×]\s*\d+(?:[,.]\d+)?\s*[xх\*×]\s*\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        diam = new Regex(@"\d+(?:[,.]\d+)?(?=\s*[xх\*×]\s*\d+(?:[,.]\d+)?\s*[xх\*×]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        tolsh = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*×]\s*)\d+(?:[,.]\d+)?(?=\s*[xх\*×]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        metraj = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*×]\s*\d+(?:[,.]\d+)?\s*[xх\*×]\s*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                        if (double.Parse(tolsh) > double.Parse(diam)) { var max = diam; diam = tolsh; tolsh = max; }
                                        if (new Regex(@"профиль", RegexOptions.IgnoreCase).IsMatch(type))
                                            if (double.Parse(tolsh) > double.Parse(metraj)) { var max = tolsh; tolsh = metraj; metraj = max; }

                                    }
                                    //diamDxD
                                    else if (new Regex(@"\d+(?:[,.]\d+)?\s*[xх\*×]\s*\d+(?:[,.]\d+)?\s*", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        diam = new Regex(@"\d+(?:[,.]\d+)?(?=\s*[xх\*×]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        tolsh = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*×]\s*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                        try
                                        {
                                            if (double.Parse(tolsh) > double.Parse(diam)) { var max = diam; diam = tolsh; tolsh = max; }
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("Ошибка преобразования\n diam = " + diam + ", tolsh = " + tolsh + "\n" + ex.ToString());
                                        }
                                    }
                                    else diam = new Regex(@"\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                    if (string.IsNullOrEmpty(metraj) && !string.IsNullOrEmpty(tolsh))
                                    {
                                        metraj = new Regex(@"\d+(?:[,.]\d+)?(?=\s*мм)", RegexOptions.IgnoreCase).Match(temp).Value;
                                    }

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
                    if (tabs.Count > 0 && dtProduct.Rows.Count > 0 && excelworksheet.Index == 1)
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
            temp = temp.Replace("\r\a", " ").Replace("\r", " ").Replace("\n", " ");
            if (new Regex(@"\d{6}(?:,?\s*\w+(?:\s*\w+)?)?,?\s*г.\s*\w+(?:\s*\w+)*,?\s*ул.\s*\w+(?:\s*\w+)*,?\s*(?:д.)?\s*[\d\w]+(?:,?\s*оф.?\s*[\d\-]+)?", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.OrgAdress = new Regex(@"\d{6}(?:,?\s*\w+(?:\s*\w+)?)?,?\s*г.\s*\w+(?:\s*\w+)*,?\s*ул.\s*\w+(?:\s*\w+)*,?\s*(?:д.)?\s*[\d\w]+(?:,?\s*оф.?\s*[\d\-]+)?", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (new Regex(@"(?<=тел\s*:\s*)(?:[\d+-]+(?:,\s)?)+(?=\s)", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.OrgTel = new Regex(@"(?<=тел\s*:\s*)(?:[\d+-]+(?:,\s)?)+(?=\s)", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (regexParam.EMail.IsMatch(temp))
            {
                infoOrg.Email = regexParam.EMail.Match(temp).Value;
            }
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
    class Class_MetallGradWord
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
            ReadWord();
            //return dtProduct;
        }

        Word._Application application;
        Word._Document document;
        Object missingObj = System.Reflection.Missing.Value;
        Object trueObj = true;
        Object falseObj = false;
        bool isOpenWord = false;

        DataTable dtProduct = new DataTable();

        string orgname = "";


        public string NameOrg() { return orgname; }

        private void ReadWord()
        {
            InfoOrganization infoOrg = new InfoOrganization
            {
                SkladAdr = new List<string>(),
                Manager = new List<string[]>()
            };

            //создаем обьект приложения word
            application = new Word.Application();
            // создаем путь к файлу
            Object templatePathObj = filePath;
            Word.Tables tables;
            try
            {
                orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?)").Match(Path.GetFileName(filePath)).Value;
                infoOrg.OrgName = StringFirstUp(orgname)+"_word";

                document = application.Documents.Open(ref templatePathObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj);

                tables = document.Tables;
                isOpenWord = true;

                int lastRow = 0;
                int ColDiam = 0, ColTolsh = 0, ColMera = 0, ColMark = 0, ColPrim = 0, ColTU = 0, ColOpis = 0, ColPrice = 0, ColName = 0;
                C_InfoTable tab;

                string temp = "", tmp = "", price = "", prim = "", name = "", type = "";
                string diam = "", tolsh = "", metraj = "", mera = "", standart = "", mark = "";
                var regexParam = new C_RegexParamProduct();

                int progress = 0; //для прогрессбара счетчик

                for (int t = 1; t <= tables.Count; t++)
                {
                    tab = new C_InfoTable();
                    Word.Table wTab = tables[t];
                    int cCelCol = wTab.Columns.Count;
                    int cCelRow = wTab.Rows.Count;
                    //if (cCelCol < 10) cCelCol = 10;
                    //if (cCelCol > 20) cCelCol = 20;

                    int Max = cCelCol * cCelRow;
                    int dif = 0;
                    SetMaxValProgressBar(Max);
                    //Поиск заголовков столбцов
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        int jj = j;
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            if (i <= wTab.Rows[jj].Cells.Count)
                            {
                                temp = wTab.Cell(jj, i).Range.Text;
                                temp = temp.Replace("\r\a", string.Empty).Trim();
                                if (new Regex(@"^наименование", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColName = i;
                                    if (ColName > 1)
                                    {
                                        if (wTab.Rows[jj].Cells.Count < wTab.Rows[jj + 1].Cells.Count)
                                        {
                                            dif = wTab.Rows[jj + 1].Cells.Count - wTab.Rows[jj].Cells.Count;
                                        }
                                    }
                                    j = cCelRow; tab.StartRow = jj; continue;
                                }
                                else if (new Regex(@"\bмерность\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColMera = i; continue; }
                                else if (new Regex(@"цена", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColPrice = i; continue; }

                                if (progress < Max) ProcessChanged(progress++);
                                else ProcessChanged(Max);
                            }
                        }
                    }
                    if (dif > 0) { ColName += dif; ColMera += dif; ColPrice += dif; }
                    Max = cCelRow;
                    SetMaxValProgressBar(Max);
                    progress = 0;
                    name = "Труба";
                    if (ColName != 0)
                    {
                        for (int jj = tab.StartRow + 1; jj <= cCelRow; jj++) //строки
                        {
                            temp = wTab.Cell(jj, ColName).Range.Text;
                            temp = temp.Replace("\r\a", string.Empty).Trim();
                            temp = temp.Replace(".", ",");
                            if (temp != "")
                            {
                                diam = ""; tolsh = ""; metraj = ""; mera = ""; standart = ""; mark = "";

                                prim = temp;
                                name = StringFirstUp(regexParam.RegName.Match(temp).Value);
                                if (!string.IsNullOrEmpty(name))
                                {
                                    temp = temp.Replace(name, "").Trim();
                                    type = regexParam.GetTypeLongFromShort(regexParam.RegType.Match(temp).Value, name);
                                    if (string.IsNullOrEmpty(standart))
                                    {
                                        standart = regexParam.RegTU.Match(temp).Value;
                                        if (!string.IsNullOrEmpty(standart)) temp = temp.Replace(standart, "");
                                    }
                                    if (string.IsNullOrEmpty(mark))
                                    {
                                        mark = regexParam.RegMark.Match(temp).Value;
                                        if (!string.IsNullOrEmpty(mark)) temp = temp.Replace(mark, "");
                                    }
                                }
                                var dtm = new Class_DTM(temp, type);
                                diam = dtm.D();
                                tolsh = dtm.T();
                                metraj = dtm.M();
                                ////diamDxDxD
                                //if (new Regex(@"\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                //{
                                //    diam = new Regex(@"\d+(?:[,.]\d+)?(?=\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                //    tolsh = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*]\s*)\d+(?:[,.]\d+)?(?=\s*[xх\*]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                //    metraj = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*[xх\*]\s*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                //    if (double.Parse(tolsh) > double.Parse(diam)) { var max = diam; diam = tolsh; tolsh = max; }
                                //}
                                ////diamDxD
                                //else if (new Regex(@"\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*", RegexOptions.IgnoreCase).IsMatch(temp))
                                //{
                                //    diam = new Regex(@"\d+(?:[,.]\d+)?(?=\s*[xх\*]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                //    tolsh = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*]\s*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                //    try
                                //    {
                                //        if (double.Parse(tolsh) > double.Parse(diam)) { var max = diam; diam = tolsh; tolsh = max; }
                                //    }
                                //    catch (Exception ex)
                                //    {
                                //        MessageBox.Show("Ошибка преобразования\n diam = " + diam + ", tolsh = " + tolsh + "\n" + ex.ToString());
                                //    }

                                //}
                                ////diam D
                                //else if (new Regex(@"\d+(?:[,.]\d+)?\s*", RegexOptions.IgnoreCase).IsMatch(temp))
                                //{
                                //    diam = new Regex(@"\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                //}
                                if (!string.IsNullOrEmpty(name) && string.IsNullOrEmpty(diam))
                                {
                                    tab.Name = StringFirstUp(name);
                                    tab.Type = regexParam.RegType.Match(temp).Value;
                                }


                                if (!string.IsNullOrEmpty(diam))
                                {
                                    if (ColMera > 0)
                                    {
                                        tmp = "";
                                        tmp = wTab.Cell(jj, ColMera).Range.Text;
                                        tmp = tmp.Replace("\r\a", string.Empty).Trim();
                                        tmp = tmp.Replace("м", string.Empty).Trim();
                                        if (tmp != "")
                                        {
                                            mera = tmp;
                                        }
                                    }
                                    if (ColPrice > 0)
                                    {
                                        tmp = "";
                                        tmp = wTab.Cell(jj, ColPrice).Range.Text;
                                        tmp = tmp.Replace("\r\a", string.Empty).Trim();
                                        if (tmp != "")
                                        {
                                            price = new Regex(@"\d+(?:[,.]\d+)?",RegexOptions.IgnoreCase).Match(tmp).Value;
                                        }
                                    }

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
                                }
                            }

                            if (progress < Max) ProcessChanged(progress++);
                            else ProcessChanged(Max);
                        }
                    }
                }

                SetMaxValProgressBar(10);
                ProcessChanged(0);
                progress = 0;
                //поиск информации об организации в первых 10 параграфах
                for (int i = 1; i <= 10; i++)
                {
                    temp = document.Paragraphs[i].Range.Text;
                    temp = temp.Replace("\r\a", string.Empty).Trim();
                    FillInfoOrg(infoOrg, temp, regexParam);
                    ProcessChanged(progress++);
                }
                //SetMaxValProgressBar(document.Paragraphs.Count);
                //ProcessChanged(0);
                //progress = 0;
                //for (int i = 1; i <= document.Paragraphs.Count; i++)
                //{
                //    temp = document.Paragraphs[i].Range.Text;
                //    temp = temp.Replace("\r\a", string.Empty).Trim();
                //    FillInfoOrg(infoOrg, temp, regexParam);
                //    ProcessChanged(progress++);
                //}

                SetInfoOrganization(infoOrg);
                WorkCompleted(dtProduct);
            }
            catch (Exception ex) { MessageBox.Show("Ошибка при обработке файла " + Path.GetFileName(filePath) + "\n\n" + ex.ToString()); }
            if (isOpenWord)
            {
                document.Close();
                application.Quit(missingObj, missingObj, missingObj);
            }
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
            temp = temp.Replace("\r\a", " ").Replace("\r", " ").Replace("\n", " ");
            if (new Regex(@"\d{6}(?:,?\s*\w+(?:\s*\w+)?)?,?\s*г.\s*\w+(?:\s*\w+)*,?\s*ул.\s*\w+(?:\s*\w+)*,?\s*(?:д.)?\s*[\d\w]+(?:,?\s*оф.?\s*[\d\-]+)?", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.OrgAdress = new Regex(@"\d{6}(?:,?\s*\w+(?:\s*\w+)?)?,?\s*г.\s*\w+(?:\s*\w+)*,?\s*ул.\s*\w+(?:\s*\w+)*,?\s*(?:д.)?\s*[\d\w]+(?:,?\s*оф.?\s*[\d\-]+)?", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (new Regex(@"(?<=тел\s*:\s*)(?:[\d+-]+(?:,\s)?)+(?=\s|\s*$)", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                infoOrg.OrgTel = new Regex(@"(?<=тел\s*:\s*)(?:[\d+-]+(?:,\s)?)+(?=\s|\s*$)", RegexOptions.IgnoreCase).Match(temp).Value;
            }
            if (regexParam.EMail.IsMatch(temp))
            {
                infoOrg.Email = regexParam.EMail.Match(temp).Value;
            }
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


        public event Action<int> ProcessChanged;

        public event Action<int> SetMaxValProgressBar;

        public event Action<InfoOrganization> SetInfoOrganization;

        public event Action<DataTable> WorkCompleted;
    }
}
