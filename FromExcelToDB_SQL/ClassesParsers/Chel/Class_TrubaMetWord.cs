using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.IO;

namespace MetalBase.ClassesParsers.Chel
{
    class Class_TrubaMetWord
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
                infoOrg.OrgName = orgname;

                document = application.Documents.Open(ref templatePathObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj);

                tables = document.Tables;
                isOpenWord = true;

                int lastRow = 0;
                int ColDiam = 0, ColTolsh = 0, ColMera = 0, ColMark = 0, ColPrim = 0, ColTU = 0, ColOpis = 0, ColPrice = 0;
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
                                if (new Regex(@"^труба", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColDiam = i; j = cCelRow; tab.StartRow = jj; continue; }
                                else if (new Regex(@"описание", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColOpis = i; continue; }
                                else if (new Regex(@"\bтн\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColMera = i; continue; }
                                else if (new Regex(@"цена", RegexOptions.IgnoreCase).IsMatch(temp))
                                { ColPrice = i; continue; }

                                if (progress < Max) ProcessChanged(progress++);
                                else ProcessChanged(Max);
                            }
                        }
                    }
                    Max = cCelRow;
                    SetMaxValProgressBar(Max);
                    progress = 0;
                    name = "Труба";
                    if (ColDiam != 0)
                    {
                        for (int jj = tab.StartRow + 1; jj <= cCelRow; jj++) //строки
                        {
                            temp = wTab.Cell(jj, ColDiam).Range.Text;
                            temp = temp.Replace("\r\a", string.Empty).Trim();
                            temp = temp.Replace(".", ",");
                            if (temp != "")
                            {
                                diam = ""; tolsh = ""; metraj = ""; mera = ""; standart = ""; mark = "";
                                if (new Regex(@"\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    prim = "D=" + temp;
                                    //diamDxDxD
                                    if (new Regex(@"\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        diam = new Regex(@"\d+(?:[,.]\d+)?(?=\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        tolsh = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*]\s*)\d+(?:[,.]\d+)?(?=\s*[xх\*]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        metraj = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*[xх\*]\s*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                        if (double.Parse(tolsh) > double.Parse(diam)) { var max = diam; diam = tolsh; tolsh = max; }
                                    }
                                    //diamDxD
                                    else if (new Regex(@"\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        diam = new Regex(@"\d+(?:[,.]\d+)?(?=\s*[xх\*]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        tolsh = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*]\s*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                        try
                                        {
                                            if (double.Parse(tolsh) > double.Parse(diam)) { var max = diam; diam = tolsh; tolsh = max; }
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("Ошибка преобразования\n diam = " + diam + ", tolsh = " + tolsh + "\n" + ex.ToString());
                                        }

                                    }
                                    //diam D
                                    else if (new Regex(@"\d+(?:[,.]\d+)?\s*", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        diam = new Regex(@"\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                    }
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
                                            price = tmp;
                                        }
                                    }
                                    if (ColOpis > 0)
                                    {
                                        tmp = "";
                                        tmp = wTab.Cell(jj, ColOpis).Range.Text;
                                        tmp = tmp.Replace("\r\a", string.Empty).Trim();
                                        if (tmp != "")
                                        {
                                            standart = regexParam.RegTU.Match(tmp).Value;
                                            if(!string.IsNullOrEmpty(standart)) tmp = tmp.Replace(standart, "");
                                            mark = regexParam.RegMark.Match(tmp).Value;
                                            if(!string.IsNullOrEmpty(mark)) tmp = tmp.Replace(mark, "");
                                            tmp = tmp.Trim();
                                            tmp = tmp.Replace(" ", "");
                                            if (!string.IsNullOrEmpty(tmp))
                                                prim += "  Примеч: " + tmp;
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

                SetMaxValProgressBar(document.Sections.Count);
                ProcessChanged(0);
                progress = 0;
                // поиск информации об организации в первых 10 параграфах
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    temp = document.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text;
                    temp = temp.Replace("\r\a", string.Empty).Trim();
                    FillInfoOrg(infoOrg, temp, regexParam);
                    ProcessChanged(progress++);
                }
                SetMaxValProgressBar(document.Paragraphs.Count);
                ProcessChanged(0);
                progress = 0;
                for (int i = 1; i <= document.Paragraphs.Count; i++)
                {
                    temp = document.Paragraphs[i].Range.Text;
                    temp = temp.Replace("\r\a", string.Empty).Trim();
                    FillInfoOrg(infoOrg, temp, regexParam);
                    ProcessChanged(progress++);
                }

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
                if(string.IsNullOrEmpty(infoOrg.OrgAdress)) infoOrg.OrgAdress = regexParam.OrgAdresFull.Match(temp).Value;
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
