using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.IO;

namespace MetalBase.ClassesParsers.Chel
{
    class Class_UralTeploEnergoServiceWord
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

                int lastRow = 0, progress = 0;
                int ColName = 0, ColDlina = 0, ColMera = 0, ColMark = 0, ColPrice = 0;
                C_InfoTable tab;

                string temp = "", tmp = "", price = "", prim = "";
                string diam = "", tolsh = "", metraj = "", mera = "", mark = "", type = "", name = "", standart = "", dlina = "";
                var regexParam = new C_RegexParamProduct();

                int colP = 10, dpc = document.Paragraphs.Count;
                if (dpc > 0)
                    colP = document.Paragraphs.Count;
                tab = new C_InfoTable();

                SetMaxValProgressBar(colP);
                progress = 0;
                ProcessChanged(progress);

                for (int i = 1; i < colP; i++)
                {
                    diam = ""; tolsh = ""; metraj = ""; mera = ""; mark = ""; type = ""; name = ""; standart = ""; dlina = ""; price = ""; prim = "";
                    //temp = document.Paragraphs[i].Range.Text;
                    temp = document.Paragraphs[i].Range.Text;
                    temp = temp.Replace("\r\a", string.Empty).Trim();
                    temp = new Regex(@"\s\s\s\s", RegexOptions.IgnoreCase).Replace(temp, "~").Trim();
                    string[] tempMass = temp.Split('~');
                    if (tempMass.Length > 0)
                        for (int j = 0; j < tempMass.Length; j++)
                        {
                            diam = ""; tolsh = ""; metraj = ""; mera = ""; mark = ""; type = ""; name = ""; standart = ""; dlina = ""; price = ""; prim = "";
                            if (tempMass[j] != "")
                            {
                                tempMass[j] = tempMass[j].Trim();
                                name = regexParam.RegName.Match(tempMass[j]).Value;
                                //diamDxDxD
                                if (new Regex(@"\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).IsMatch(tempMass[j]))
                                {
                                    diam = new Regex(@"\d+(?:[,.]\d+)?(?=\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(tempMass[j]).Value;
                                    tolsh = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*]\s*)\d+(?:[,.]\d+)?(?=\s*[xх\*]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(tempMass[j]).Value;
                                    metraj = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*[xх\*]\s*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(tempMass[j]).Value;
                                    if (double.Parse(tolsh) > double.Parse(diam)) { var max = diam; diam = tolsh; tolsh = max; }
                                }
                                //diamDxD
                                else if (new Regex(@"\d+(?:[,.]\d+)?\s*[xх\*]\s*\d+(?:[,.]\d+)?\s*", RegexOptions.IgnoreCase).IsMatch(tempMass[j]))
                                {
                                    diam = new Regex(@"\d+(?:[,.]\d+)?(?=\s*[xх\*]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(tempMass[j]).Value;
                                    tolsh = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*]\s*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(tempMass[j]).Value;
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
                                if (!string.IsNullOrEmpty(name) && (new Regex(@"\d+(?:[,.]\d+)?\s*", RegexOptions.IgnoreCase).IsMatch(tempMass[j])))
                                {
                                    diam = new Regex(@"\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(tempMass[j]).Value;
                                }
                                if (!string.IsNullOrEmpty(diam))
                                {
                                    tab.Name = "Труба";
                                    prim = tempMass[j];
                                    name = regexParam.RegName.Match(tempMass[j]).Value;
                                    if (!string.IsNullOrEmpty(name)) tempMass[j] = tempMass[j].Replace(name, "").Trim();
                                    standart = regexParam.RegTU.Match(tempMass[j]).Value;
                                    if (!string.IsNullOrEmpty(standart)) tempMass[j] = tempMass[j].Replace(standart, "").Trim();
                                    mark = regexParam.RegMark.Match(tempMass[j]).Value;
                                    if (!string.IsNullOrEmpty(mark)) tempMass[j] = tempMass[j].Replace(mark, "").Trim();
                                    price = new Regex(@"(?<=цена\s*)\d+(?:\s?[\.,]\s?\d+)?", RegexOptions.IgnoreCase).Match(tempMass[j]).Value;
                                    tempMass[j] = new Regex(@"цена\s*\d+(?:\s?[\.,]\s?\d+)?", RegexOptions.IgnoreCase).Replace(tempMass[j], "");
                                    mera = new Regex(@"\d+(?:\s?[\.,]\s?\d+)?(?=\s*тн\.?)", RegexOptions.IgnoreCase).Match(tempMass[j]).Value;
                                    tempMass[j] = new Regex(@"\d+(?:\s?[\.,]\s?\d+)?\s*тн\.?", RegexOptions.IgnoreCase).Replace(tempMass[j], "");
                                    metraj = new Regex(@"\d+(?:\s?[\.,]\s?\d+)?(?=\s*м\.?)", RegexOptions.IgnoreCase).Match(tempMass[j]).Value;

                                    if (!String.IsNullOrEmpty(price) || !string.IsNullOrEmpty(standart) || !string.IsNullOrEmpty(name))
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

                                    if (progress < colP) ProcessChanged(progress++);
                                    else ProcessChanged(colP);

                                    continue;
                                }

                                if (new Regex(@"\d{6}\s*г.\w+,?\s*ул.(?:\s*\w+\s*){1,2},?\s*д\.\d+(?:[-\\/][\d\w]+)?(?:,?\s*[\w-]+[\d\w])?", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    infoOrg.OrgAdress = new Regex(@"\d{6}\s*г.\w+,?\s*ул.(?:\s*\w+\s*){1,2},?\s*д\.\d+(?:[-\\/][\d\w]+)?(?:,?\s*[\w-]+[\d\w])?", RegexOptions.IgnoreCase).Match(temp).Value;
                                }
                                if (new Regex(@"\+79\d{9}(?=\s*$|\s)", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    infoOrg.OrgTel = new Regex(@"\+79\d{9}(?=\s*$|\s)", RegexOptions.IgnoreCase).Match(temp).Value;
                                }
                                if (regexParam.EMail.IsMatch(temp))
                                {
                                    infoOrg.Email = regexParam.EMail.Match(temp).Value;
                                }
                                if (new Regex(@"(?<=\s|^)(?:www\.)?(?:[а-яёa-z0-9_-]{1,32}(?::[а-яёa-z0-9_-]{1,32})?)?(?:(?:[а-яёa-z0-9-]{1,128}\.)+(?:ru|su|com|net|org|mil|edu|arpa|gov|biz|info|aero|inc|name|рф))", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    infoOrg.Site = new Regex(@"(?<=\s|^)(?:www\.)?(?:[а-яёa-z0-9_-]{1,32}(?::[а-яёa-z0-9_-]{1,32})?)?(?:(?:[а-яёa-z0-9-]{1,128}\.)+(?:ru|su|com|net|org|mil|edu|arpa|gov|biz|info|aero|inc|name|рф))", RegexOptions.IgnoreCase).Match(temp).Value;
                                }
                                if (new Regex(@"адрес\s*базы", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    infoOrg.SkladAdr.Add(new Regex(@"(?<=дрес\s*базы\s*:?\s*)[\w\s\.\d]+(?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value);
                                }
                                if (new Regex(@"(?<=ИНН/КПП\s*:?\s*)\d{9,15}\s*/\s*\d{9,15}", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    infoOrg.Inn_Kpp = new Regex(@"(?<=ИНН/КПП\s*:?\s*)\d{9,15}\s*/\s*\d{9,15}", RegexOptions.IgnoreCase).Match(temp).Value;
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
                                if (progress < colP) ProcessChanged(progress++);
                                else ProcessChanged(colP);
                            }
                        }
                }

                ProcessChanged(colP);
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
