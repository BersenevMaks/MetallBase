using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.IO;

namespace MetalBase
{
    class Class_MetallPromSnab
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
                int ColName = 0, ColDlina = 0, ColMera = 0, ColMark = 0, ColPrice = 0;
                C_InfoTable tab;

                string temp = "", tmp = "", price = "", prim = "";
                string diam = "", tolsh = "", metraj = "", mera = "", name = "", standart = "", dlina = "";
                var regexParam = new C_RegexParamProduct();

                for (int t = 1; t <= 1; t++)
                {
                    tab = new C_InfoTable();

                    Word.Table wTab = tables[t];
                    int cCelCol = wTab.Columns.Count;
                    int cCelRow = wTab.Rows.Count;
                    //if (cCelCol < 10) cCelCol = 10;
                    //if (cCelCol > 20) cCelCol = 20;
                    int Max = cCelRow;
                    SetMaxValProgressBar(Max);
                    for (int jj = tab.StartRow + 1; jj <= cCelRow; jj++) //строки
                    {
                        diam = ""; tolsh = ""; metraj = ""; price = ""; prim = ""; mera = ""; standart = "";
                        temp = wTab.Cell(jj, 1).Range.Text;
                        temp = temp.Replace("\r\a", string.Empty).Trim();
                        if (temp != "")
                        {
                            if (regexParam.RegName.IsMatch(temp))
                            {
                                tab.Standart = "";
                                tab.Name = StringFirstUp(regexParam.RegName.Match(temp).Value);
                                tab.Type = regexParam.RegType.Match(temp).Value;
                                tab.Standart = regexParam.RegTU.Match(temp).Value;
                            }
                            if (regexParam.RegType.IsMatch(temp))
                                tab.Type = regexParam.RegType.Match(temp).Value;

                            if (regexParam.RegTU.IsMatch(temp))
                                tab.Standart = regexParam.RegTU.Match(temp).Value;

                            temp = temp.Replace(" ", "");
                            temp = temp.Replace('/', 'х');


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
                            if (string.IsNullOrEmpty(diam)) continue;

                            prim = temp;

                            tmp = "";
                            tmp = wTab.Cell(jj, 2).Range.Text;
                            tmp = tmp.Replace("\r\a", string.Empty).Trim();
                            if (tmp != "")
                            {
                                tab.Mark = regexParam.RegMark.Match(tmp).Value;
                                metraj = new Regex(@"(?<=l\s*=\s*)(\d+(?:[,.]\d+)?\s*-?\s*,?\s*)+", RegexOptions.IgnoreCase).Match(tmp).Value;
                                if(string.IsNullOrEmpty(metraj))
                                    metraj = new Regex(@"^(?:\d+(?:[,.]\d+)?,?\s*)?", RegexOptions.IgnoreCase).Match(tmp).Value;
                                if (regexParam.RegType.IsMatch(tmp))
                                    tab.Type = regexParam.RegType.Match(tmp).Value;
                                if (regexParam.RegTU.IsMatch(tmp))
                                {
                                    standart = regexParam.RegTU.Match(tmp).Value;
                                    if (standart.Contains(diam)) diam = "";
                                }
                            }

                            tmp = "";
                            tmp = wTab.Cell(jj, 3).Range.Text;
                            tmp = tmp.Replace("\r\a", string.Empty).Trim();
                            if (tmp != "")
                            {
                                mera = tmp;
                            }

                            tmp = "";
                            tmp = wTab.Cell(jj, 4).Range.Text;
                            tmp = tmp.Replace("\r\a", string.Empty).Trim();
                            if (tmp != "")
                            {
                                price = tmp;
                            }

                            if (!string.IsNullOrEmpty(tab.Standart) && !string.IsNullOrEmpty(diam))
                            {
                                if (tab.Standart.Contains(diam)) diam = "";
                            }

                            if (dtProduct.Rows.Count > 0)
                            {
                                lastRow = dtProduct.Rows.Count - 1;
                                tab.LastRowExcel = jj;
                            }
                            else lastRow = 0;
                            if (!String.IsNullOrEmpty(diam))
                            {
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
                                row["Стандарт"] = standart;
                                if(string.IsNullOrEmpty(row["Стандарт"].ToString()))
                                    row["Стандарт"] = tab.Standart;
                                row["Класс"] = "";
                                row["Цена"] = price;
                                row["Примечание"] = prim;
                                dtProduct.Rows.Add(row);
                            }

                            if (jj < Max) ProcessChanged(jj);
                            else ProcessChanged(Max);

                        }
                    }
                }

                // поиск информации об организации в первых 10 параграфах
                int colP = 10, dsc = document.Shapes.Count;
                if (dsc > 0)
                    colP = document.Shapes.Count;
                for (int i = 1; i < colP; i++)
                {
                    //temp = document.Paragraphs[i].Range.Text;
                    temp = document.Shapes[1].TextFrame.TextRange.Text;

                    temp = temp.Replace("\r\a", string.Empty).Trim();
                    if (temp != "")
                    {
                        if (new Regex(@"\d+\s*,\s*(?:г\.)?\s*\w+\s*,\s*(?:\w+\.?)?\s*\w+\s*,\s*д\.?\s*6(?:\s*,\s*оф\.?\s*\d+\w+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                        {
                            infoOrg.OrgAdress = new Regex(@"\d+\s*,\s*(?:г\.)?\s*\w+\s*,\s*(?:\w+\.?)?\s*\w+\s*,\s*д\.?\s*6(?:\s*,\s*оф\.?\s*\d+\w+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                        }
                        if (new Regex(@"(?<=Тел/\s*факс:\s*)[\(\)\d\s,\-]+", RegexOptions.IgnoreCase).IsMatch(temp))
                        {
                            infoOrg.OrgTel = new Regex(@"(?<=Тел/\s*факс:\s*)[\(\)\d\s,\-]+", RegexOptions.IgnoreCase).Match(temp).Value;
                        }
                        if (new Regex(@"(?<=тел\\факс\s).*", RegexOptions.IgnoreCase).IsMatch(temp))
                        {
                            infoOrg.OrgTel = new Regex(@"(?<=тел\\факс\s).*", RegexOptions.IgnoreCase).Match(temp).Value;
                        }
                        if (regexParam.EMail.IsMatch(temp))
                        {
                            infoOrg.Email = regexParam.EMail.Match(temp).Value;
                        }
                        if (new Regex(@"(?:http://)?\w+\.(?:ru|com|info|рф)", RegexOptions.IgnoreCase).IsMatch(temp))
                        {
                            infoOrg.Site = new Regex(@"(?:http://)?\w+\.(?:ru|com|info|рф)", RegexOptions.IgnoreCase).Match(temp).Value;
                        }
                        else if (new Regex(@"(?<=\s|^)(?:www\.)?(?:[а-яёa-z0-9_-]{1,32}(?::[а-яёa-z0-9_-]{1,32})?)?(?:(?:[а-яёa-z0-9-]{1,128}\.)+(?:ru|su|com|net|org|mil|edu|arpa|gov|biz|info|aero|inc|name|рф))", RegexOptions.IgnoreCase).IsMatch(temp))
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
                    }
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
