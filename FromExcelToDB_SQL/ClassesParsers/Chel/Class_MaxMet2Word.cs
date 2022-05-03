using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;
using MetalBase.HelpClasses;

namespace MetalBase.ClassesParsers.Chel
{
    class Class_MaxMet2Word
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
        private Excel.Application excelapp;
        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;
        private Excel.Sheets excelsheets;

        Word._Application application;
        Word._Document document;
        Object missingObj = System.Reflection.Missing.Value;
        Object trueObj = true;
        Object falseObj = false;
        bool isOpenWord = false;
        bool isExcelOpen = false;

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
                infoOrg.OrgName = "МаксМет2";

                document = application.Documents.Open(ref templatePathObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj);

                tables = document.Tables;
                isOpenWord = true;

                int progress = 0; //для прогрессбара счетчик

                for (int t = 1; t <= tables.Count; t++)
                {
                    Word.Table wTab = tables[t];

                    wTab.Range.Copy();

                    excelapp = new Excel.Application();
                    excelappworkbooks = excelapp.Workbooks;
                    excelappworkbook = excelapp.Workbooks.Add(1);
                    excelsheets = excelappworkbook.Worksheets;
                    isExcelOpen = true;
                    Excel.Worksheet excelworksheet = (Excel.Worksheet)excelsheets[1];

                    excelworksheet.Paste();
                    //excelapp.Visible = true;

                    var tab = new C_InfoTable();
                    var naaame = excelworksheet.Name;
                    List<C_InfoTable> tabs = new List<C_InfoTable>();
                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol <= 10) cCelCol = 10;
                    if (cCelCol > 10) cCelCol = 25;

                    int Max = cCelCol * cCelRow;
                    SetMaxValProgressBar(Max);

                    int ColName = 0, ColDiam = 0, ColTolsh = 0, ColMera = 0, ColMark = 0, ColPrim = 0, ColGost = 0, ColPrice = 0;

                    string temp = "", price = "", prim = "", name = "", type = "";
                    string diam = "", tolsh = "", metraj = "", mera = "", standart = "", mark = "";
                    var regexParam = new C_RegexParamProduct();
                    var dtm = new Class_DTM();

                    //Поиск заголовков столбцов
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        int jj = j;
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange;
                            cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                            if (cellRange.Value != null)
                                temp = regexParam.DelSpacesInWords(cellRange.Value.ToString().Trim());
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"^диаметр", RegexOptions.IgnoreCase).IsMatch(temp))
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
                                if (new Regex(@"^диаметр", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColDiam = i;
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
                        int maxCountCols = 1; //максимальное количество столбцов в таблице
                        int lastDiamCol = 1; //текущий номер столбца ColMera
                        for (int j = 1; j <= endRow; j++)
                        {
                            diam = ""; tolsh = ""; metraj = ""; mera = ""; name = ""; prim = ""; standart = ""; 
                            
                            if (ColDiam > 0)
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[j, ColDiam];
                                if (cellRange.MergeArea.Columns.Count >= maxCountCols)
                                {
                                    if (cellRange.MergeArea.Columns.Count > maxCountCols)
                                        maxCountCols = cellRange.MergeArea.Columns.Count;
                                    cellRange = (Excel.Range)excelworksheet.Cells[j, cellRange.MergeArea.Column];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        temp = regexParam.DelSpacesInWords(temp);
                                        if (regexParam.RegName.IsMatch(temp))
                                            tab.Name = regexParam.RegName.Match(temp).Value;
                                        type = regexParam.RegType.Match(temp).Value.Trim();

                                        mark = "";
                                        mark = regexParam.RegMark.Match(temp).Value;
                                        if (!string.IsNullOrEmpty(mark))
                                            temp = temp.Replace(mark, string.Empty);
                                        price = "";
                                        price = new Regex(@"(?<=цена\s*)\d+(?:[\s\.,]+\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                    }
                                }
                                else
                                {
                                    if (ColDiam + cellRange.MergeArea.Columns.Count - 1 > lastDiamCol)
                                        lastDiamCol = ColDiam + cellRange.MergeArea.Columns.Count - 1;

                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        dtm.CalcDTM(temp);
                                        diam = dtm.D();
                                        tolsh = dtm.T();
                                        metraj = dtm.M();

                                        prim = tab.Name+", " + temp;

                                    }

                                    //чтение столбца ColMera
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[j, lastDiamCol+1];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            temp = temp.Replace(".", ",");
                                            foreach (string numb in temp.Split('+'))
                                            {
                                                mera = new Regex(@"\d+(?:[\s\.,]+\d+)?", RegexOptions.IgnoreCase).Match(numb).Value;
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
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            if (j * (k + 1) < Max) ProcessChanged(j * (k + 1));
                            else ProcessChanged(Max);

                        }

                        if (isExcelOpen)
                        {
                            excelappworkbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                            excelapp.Quit();
                        }
                    }
                }
                SetMaxValProgressBar(document.Sections.Count);
                ProcessChanged(0);
                progress = 0;
                //// поиск информации об организации в первых 10 параграфах
                //for (int i = 1; i <= document.Sections.Count; i++)
                //{
                //    temp = document.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text;
                //    temp = temp.Replace("\r\a", string.Empty).Trim();
                //    FillInfoOrg(infoOrg, temp, regexParam);
                //    ProcessChanged(progress++);
                //}

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
            if (regexParam.OrgAdresFully.IsMatch(temp) || regexParam.OrgAdresFull.IsMatch(temp) || regexParam.OrgAdres.IsMatch(temp))
            {
                infoOrg.OrgAdress = regexParam.OrgAdresFully.Match(temp).Value;
                if (string.IsNullOrEmpty(infoOrg.OrgAdress)) infoOrg.OrgAdress = regexParam.OrgAdresFull.Match(temp).Value;
                if (string.IsNullOrEmpty(infoOrg.OrgAdress)) infoOrg.OrgAdress = regexParam.OrgAdres.Match(temp).Value;
            }
            if (regexParam.OrgMobileTelefon.IsMatch(temp))
            {
                infoOrg.OrgTel = regexParam.OrgMobileTelefon.Match(temp).Value;
            }
            if (regexParam.EMail.IsMatch(temp))
            {
                infoOrg.Email = regexParam.EMail.Match(temp).Value;
            }
            if (regexParam.Site.IsMatch(temp))
            {
                infoOrg.Site = regexParam.Site.Match(temp).Value;
            }
            if (regexParam.INN.IsMatch(temp))
            {
                infoOrg.Inn_Kpp = regexParam.INN.Match(temp).Value;
            }
            if (regexParam.BIK.IsMatch(temp))
            {
                infoOrg.BIK = regexParam.BIK.Match(temp).Value;
            }
            if (regexParam.R_S.IsMatch(temp))
            {
                infoOrg.r_s = regexParam.R_S.Match(temp).Value;
            }
            if (regexParam.K_S.IsMatch(temp))
            {
                infoOrg.k_s = regexParam.K_S.Match(temp).Value;
            }
            if (new Regex(@"(?<=тел.?\s*).*\s\b\w+\b", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                string[] manager = new string[]
                {
                    new Regex(@"(?<=ел.?[\d\-\s]+\s+)\w+(?:\s*\w+)?", RegexOptions.IgnoreCase).Match(temp).Value, //имя
                    new Regex(@"(?:[\+\d]{0,2}[\s\(]{0,2}[\)\d\s,-]{3,})+", RegexOptions.IgnoreCase).Match(temp).Value, //телефон
                    "" //email
                };
                infoOrg.Manager.Add(manager);
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
