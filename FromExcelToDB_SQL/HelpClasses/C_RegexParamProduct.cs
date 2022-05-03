using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace MetalBase
{
    class C_RegexParamProduct
    {
        private Regex regName = new Regex(@"лента|(?<=\s|^)лист|арматура|\bдвутавр\b|\bмуфта\b|полоса|угол|швеллер|труб[аы]|балк\w|\bрулон\b|\bкруг\b|шестигранник|шгр|квадрат|полоса|катанка|быстрорез|проволока|профиль|поковка|отвод\w?|тройник\w?\b|металлолом|штанга|плит\w|колес\w|шифер|сетка|поковк\w|флан\w?ц\w?|заглушк\w|тройник|переход|задвижка|клапан(?:\s|\s*$)|блок(?:\s|\s*$)|электрод\w?|\bстружка\b|\bподкат\b|\bобрезь\b|\bотходы\b|\bсутунка\b|\bлом\b|\bзаготовка\b|\bметиз\w?|\bболт\w?\b|\bгайк[аи]\b|\bшайб[аы]\b|шпала|фторопласт(?:\s+втулка)?|\bвтулка\b|\bпруток\b|\bстолбики\b|\bдуги\b|\bлопости\b|\bкосынк\w\b|\bзажим\b|шарниры?|шпильк[аи]|\bшплинт\b|\bгвозд\w?\b|\bдюбель-гвозд\w|затвор|кран|входн\w\w\s+групп\w|перчатк\w|\bрельс\w?|\bдробь\b|\bсва[ия]\b|\bстойк\w\b|шпунт\w?|\bНЛЗ\b|\bштрипс\b", RegexOptions.IgnoreCase);
        private Regex regName2 = new Regex(@"сталь", RegexOptions.IgnoreCase);
        private Regex regType = new Regex(@"\w+ое(?=\b|\s|\d)|\w+ые(?=\b|\s|\d)|\w+ие(?=\b|\s|\d)|\w+ый(?=\b|\s|\d)|\w+ая(?=\b|\s|\d)|\w+ой(?:\s*проч)?|\w+ий|г[\/]к|ВГП|проф|ЭС|некондиция|нлг|[мрmp]+\-3|ОЗС\-4|ОК\-46|УОНИ\-13/55|II\s*сорт|д/забора|для\s*сваи|\dзакладная\d|\bв\s+изоляции\b", RegexOptions.IgnoreCase);
        private Regex regTypeShveller = new Regex(@"(?<=\d)(?:б|м|ш|kк|[уy]|п)", RegexOptions.IgnoreCase);
        private Regex regTypeShveller2 = new Regex(@"(?<=[\s\d])(?:б|м|ш|kк|[уy]|п)(?=\d\s|$)", RegexOptions.IgnoreCase);
        private Regex regDiam = new Regex(@"\d+(?:[,.]\d+)?(?:[xх/-]\d+(?:[,.]\d+)?)?(?:(?:\s*[x*х]\s*\d+(?:[,.]\d+)?(?:[/-]\d+(?:[,.]\d+)?)?){2,}(?:[/-]\d+(?:[,.]\d+)?)?)?(?=\s|;|$|-\d\s)|\d+(?:[,.]\d+)?\s?\*\s?\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase);
        private Regex regDiam2 = new Regex(@"(?<=\s)\d+(?:[,.]\d+)?(?:[xх/-]\d+(?:[,.]\d+)?)?(?:\s*[x*х]\s*\d+(?:[,.]\d+)?(?:[/-]\d+(?:[,.]\d+)?)?)(?=\s|$)", RegexOptions.IgnoreCase);
        private Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.]*)*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*|\bост\b\s*?\d\.\d+(?:\-\d+)?|\bТУ\b\s+\d\d\.\d\d\.\d\d-\d+\-\d+\-\d+", RegexOptions.IgnoreCase);
        private Regex regMark = new Regex(@"(?:\d{1,3}[ШСТУХ]+\d{0,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+(?:\d{0,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+)*\d{0,3})(?=\s+|$)|(?:\b(?:Ст.|ст.)(?:\s*\d{1,2})\b)(?=\s+|$)|(?:\d{,3}[ХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)(?=\s+|$)|(?:[АA]-?\d)(?=\s+|$)|(?:[Сс][Тт]\.?\s?)?\d{1,2}[гГ]\d{1,2}[cCсС]|(?<=ст\.)\d{1,2}[хфа]+(?=\s|$)|(?<=\s)(?:ст|сталь)\.?\s?\d{1,2}[_\w]+|(?<=\s)(?:ст|сталь)\.?\s?[_\w]+\d{1,2}|(?<=\s)[0-4]\d[хx]\d\d[ТН]*(?=\s)|AISI\s*\d+\w*?\b|ст12Х1МФ|ст08пс-5|ст08пс|ст10|ст\d{1,2}\b|ст\d{1,2}пс\d{0,2}|ст\d{1,2}сп\b|ст\d{1,2}сп\d{0,2}(?:/сп\d{0,2})?\b|ст3сп5\b|\b[cс]\d{3}\b|s355(17Г1С)|s355|стУ8А|\b[cс]255|\b[cс]345|[аa]500[сc]|35гс|din\s\d+|амг\d+\w*|(?:св)?ам[цг][\w\d]*|д\d+а\w*|в\d+а\w*|а\d[мн]|ад\d+\w*|в\d+\w*\d*|1561\w*\d*|а[вдк]\w*\d+|в95.*\b|д(?:16|20)\w+|д1т|д\d\d?(?=\s+\d+(?:\s|$|[хx]))|э[пи][\d\w\-]+(?=\s|$)|\bав\w?\b|\d\d?[xх]\d.*[нмш][\d\w]*(?=\s)|[рp]\d\d?[kкфм][\d\w]*|[рp]\d+(?=\s|$)|внл\-3|ди\d\d?(?=\s|$)|НМцАК\s*2-2-1|НХ\s9,5|\b\d+юч\b|\b[xх]\d+мф?\b|\b\d+[сc]\d*?\b|\bСТ\s*0\b|хн\d+[мвтудюцгб\d-]+?(?=\s|$)|\d+нхм(?:\-\w+)(?=\s|$)|(?<=\s|^)40х(?=\s|$)|нихром\s\d{2}(?=\s|$)|\d+-\d+[xх]\d+н[\w\b]+\b|1D|зеркало\sв\sбум|зеркало\sв\sпленке|шлиф.*?\s*в\sпленке|\d+г\d+[фбгю]+|\d+фа\b|(?<=^|\s)к\d+(?=\s|$)|Алюмель|Хромель|фехраль|Никель|ОТ-\d-\d|Латунь|60\d\dТ\d|\d\d?\(\d\d?\)Х\d\d?Н\d\d?Т|\b[cс]\d+-\d", RegexOptions.IgnoreCase);
        private Regex regMark2 = new Regex(@"(?:\d{1,3}[ШСТУХ]+\d{0,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+(?:\d{0,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+)*\d{0,3})(?=\s+|$)|(?:\b(?:Ст.|ст.)(?:\s*\d{1,2})\b)(?=\s+|$)|(?:\d{,3}[ХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)(?=\s+|$)|(?:[АA]-?\d)(?=\s+|$)|(?:[Сс][Тт]\.?\s?)?\d{1,2}[гГ]\d{1,2}[cCсС]|(?<=ст\.)\d{1,2}[хфа]+(?=\s|$)|(?<=\s)(?:ст|сталь)\.?\s?\d{1,2}[_\w]+|(?<=\s)(?:ст|сталь)\.?\s?[_\w]+\d{1,2}|(?<=\s)[0-4]\d[хx]\d\d[ТН]*(?=\s)|AISI\s*\d+\w*?\b|ст12Х1МФ|ст08пс-5|ст08пс|ст10|ст3|ст2пс|ст3сп|ст3сп5|ст3сп5|s355(17Г1С)|s355|стУ8А|\b[cс]255|\b[cс]345|[аa]500[сc]|35гс|din\s\d+|амг\d+\w*|(?:св)?ам[цг][\w\d]*", RegexOptions.IgnoreCase);
        private Regex diamD = new Regex(@"(?<=^)\d+(?:[,\.]\d+)?(?=(\s*[xх+]\s*\D*)?\s*$)", RegexOptions.IgnoreCase);
        private Regex diamD_Only = new Regex(@"\d+(?:[,\.]\d+)?(?=(\s*[xх+]\s*\D*)?\s*)", RegexOptions.IgnoreCase);
        private Regex diamDxD = new Regex(@"^\s*\d+(?:[,.]\d+)?\s*[xх]\s*\d+(?:[,.]\d+)?(?=(\s*[xх]?\s*\D*)?\s*$)", RegexOptions.IgnoreCase);
        private Regex diamDxD_Only = new Regex(@"\s*\d+(?:[,.]\d+)?\s*[xх]\s*\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase);
        private Regex diamDxDxD = new Regex(@"^\s*\d+(?:[,.]\d+)?\s*[xх]\s*\d+(?:[,.]\d+)?\s*[xх]\s*\d+(?:[,.]\d+)?\s*\+?\s*$", RegexOptions.IgnoreCase);
        private Regex diamDxDxD_Only = new Regex(@"\s*\d+(?:[,.]\d+)?\s*[xх]\s*\d+(?:[,.]\d+)?\s*[xх]\s*\d+(?:[,.]\d+)?\s*\+?\s*", RegexOptions.IgnoreCase);
        private Regex orgMobileTelefon = new Regex(@"\d\s*-\s*\d{3,4}\s*-\s*\d{3,4}\s*-\s*\d{2,4}\s*-\s*\d{2,4}|(?<=тел/факс\s*)(?:\(\d{3,5}\))?(?:\s*\d+\s*\d+\s*\d+,?)+|(?<=тел\.?(?:ефон)?\s*)(?:(?:\(\d+\)\s*)?(?:[\d-]+),?\s*)+", RegexOptions.IgnoreCase);
        private Regex orgMobileTelefon_tel_XXXXXXXXXX = new Regex(@"(?<=тел\s*:?\s*)\+?\s*\d{10,11}", RegexOptions.IgnoreCase);
        private Regex orgMobileTelefon_tel_XXX_XXX_XX_XX = new Regex(@"(?<=тел\s*:?\s*)(?:\+?\d{1,2}\-?)?\s*\d{2,3}\s*\-\s*\d{2,3}\s*\-\s*\d{2,3}\s*\-\s*\d{2,4}", RegexOptions.IgnoreCase);
        private Regex orgTelefon_XXX_XX_XX = new Regex(@"[\d\(\)]+\s*\-\s*\d{2,3}\s*\-\s*\d{2,3}", RegexOptions.IgnoreCase);
        private Regex eMail = new Regex(@"\s[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})");
        private Regex site = new Regex("(?:www\\.)(?:[а-яёa-z0-9_-]{1,32}(?::[а-яёa-z0-9_-]{1,32})?@)?(?:(?:[а-яёa-z0-9-]{1,128}\\.)+(?:ru|su|com|net|org|mil|edu|arpa|gov|biz|info|aero|inc|name|рф|[а-яёa-z]{2}))", RegexOptions.IgnoreCase);
        private Regex iNN_KPP = new Regex(@"(?<=ИНН\s*[\/]\s*КПП\s*:?\s*)\d+\s*[\/]\s*\d+(?=(?:\s*ОКПО\s*\d+)?)", RegexOptions.IgnoreCase);
        private Regex iNN = new Regex(@"(?<=ИНН\s*:?\s*)\d+", RegexOptions.IgnoreCase);
        private Regex oGRN = new Regex(@"(?<=огрн\s*:?\s*)\d+", RegexOptions.IgnoreCase);
        private Regex oKPO = new Regex(@"(?<=окпо\s*:?\s*)\d+", RegexOptions.IgnoreCase);
        private Regex r_S = new Regex(@"(?<=р\s*[\/]\s*[cс]\s*:?\s*)\d+", RegexOptions.IgnoreCase);
        private Regex k_S = new Regex(@"(?<=[кk]\s*[\/]\s*[cс]\s*:?\s*)\d+", RegexOptions.IgnoreCase);
        private Regex bik = new Regex(@"(?<=БИК\s*:?\s*)\d+", RegexOptions.IgnoreCase);
        private Regex orgAdres = new Regex(@"(?:\d{6}\s*)?(г\.\s*\w+,?\s*)?ул\.\s*\w+(?:\s*\w+)?,?\s*[\w\d]+", RegexOptions.IgnoreCase);
        private Regex orgAdresFull = new Regex(@"(?:\d{6}\s*)?(г\.\s*\w+,?)?(?:\s*(?:ул|пр(?:-к?т)?|д|оф)\.\s*\w+(?:\s*\w+)?,?)+", RegexOptions.IgnoreCase);
        private Regex orgAdresFully = new Regex(@"(?<=^\s*|\s)(?:\d{6}\s*,?\s*)?(?:\s*\D+\s*,\s*)?(?:\w+\s*\w+\s*\.?,?\s*)?(?:\s*г\.\s*\w+,?)\s*(?:(?:ул\.?|пр(?:-кт)?|пер)\.?\s*(?:\d*\s*)?\w+(?:\s\w+)?,?)?\s*д?(?:ом)?\.?\s*\d+(?:\w)?,?\s*(?:оф(?:ис)?\.?\s*\d+(?:\s*\w)?)?", RegexOptions.IgnoreCase);
        
        private DateTime getDateFromName(string filePath)
        {
            DateTime dateTime = DateTime.Now;
            if (new Regex(@"\.{2,}(?=xlsx?|docx?)", RegexOptions.IgnoreCase).IsMatch(filePath))
            {
                filePath = new Regex(@"\.{2,}(?=xlsx?|docx?)", RegexOptions.IgnoreCase).Replace(filePath, @".");
            }
            if (new Regex(@"(?<=.+[\s_\.])\d+[\._]\d+[\._]\d+(?=(?:г\.?)?\.[\w\d]{3,4}$)").IsMatch(Path.GetFileName(filePath)))
            {
                string[] calendar = new Regex(@"(?<=.+[\s_\.])\d+[\._]\d+[\._]\d+(?=(?:г\.?)?\.[\w\d]{3,4}$)").Match(Path.GetFileName(filePath)).Value.Split('.', '_');
                if (calendar[2].Length == 2) calendar[2] = "20" + calendar[2];
                dateTime = new DateTime(Convert.ToInt32(calendar[2]), Convert.ToInt32(calendar[1]), Convert.ToInt32(calendar[0]));
            }
            return dateTime;
        }

        private string regTypeShort(string str, string Name)
        {
            string temp = "";
            string result = "";
            if (new Regex(@"ГН|гк|рифл|проф|хк|ОЦ|эс|эсв", RegexOptions.IgnoreCase).IsMatch(str))
            {
                temp = new Regex(@"ГН|гк|рифл|проф|хк|ОЦ|эс|эсв", RegexOptions.IgnoreCase).Match(str).Value.ToLower().Trim();
                switch (temp)
                {
                    case "нлг":
                        result = "НЛГ";
                        break;
                    case "гн":
                        result = "гнут";
                        break;
                    case "гк":
                        result = "горячекатан";
                        break;
                    case "рифл":
                        result = "рифлен";
                        break;
                    case "хк":
                        result = "холоднокатан";
                        break;
                    case "оц":
                        result = "оцинкованн";
                        break;
                    case "эс":
                        result = "электросварн";
                        break;
                    case "эсв":
                        result = "электросварн";
                        break;
                    case "проф":
                        result = "профильн";
                        break;
                }
                if (!string.IsNullOrEmpty(result) && !result.Contains("НЛГ"))
                    if (!new Regex(@"\b[э]\w+\b", RegexOptions.IgnoreCase).IsMatch(Name))
                    {
                        if (new Regex(@"\b\w+[аеёиуыэь]\b", RegexOptions.IgnoreCase).IsMatch(Name))
                        {
                            result += "ая";
                        }
                        else if (new Regex(@"\b\w+[ояю]\b", RegexOptions.IgnoreCase).IsMatch(Name))
                        {
                            result += "ое";
                        }
                        else if (new Regex(@"\b\w+[бвгджзклмнпрстфхцчшщъ]\b", RegexOptions.IgnoreCase).IsMatch(Name))
                        {
                            result += "ый";
                        }
                    }
                    else result += "ой";
            }
            else if (new Regex(@"НЛГ|пвл|вгп|\dвр\d", RegexOptions.IgnoreCase).IsMatch(str))
            {
                result = new Regex(@"НЛГ|пвл|вгп|\dвр\d", RegexOptions.IgnoreCase).Match(str).Value.ToLower().Trim();
            }
            return result;
        }

        public string RegTypeShort(string str, string Name)
        {
            return regTypeShort(str, Name);
        }

        public string GetTypeLongFromShort(string stringType, string Name)
        {
            if (new Regex(@"оцин|эл.св|рыж|г/к|х/к|х/д|б/ш|рифл|проф|э/св|оц|эл/св", RegexOptions.IgnoreCase).IsMatch(stringType))
            {
                stringType = new Regex(@"оцин|эл.св|э/св|рыж|г/к|х/к|х/д|б/ш|рифл|проф|оц|эл/св", RegexOptions.IgnoreCase).Match(stringType).Value.ToLower();
                switch (stringType)
                {
                    case "эл/св":
                        stringType = "электросварн";
                        break;
                    case "оцин":
                        stringType = "оцинкованн";
                        break;
                    case "оц":
                        stringType = "оцинкованн";
                        break;
                    case "эл.св":
                        stringType = "электросварн";
                        break;
                    case "э/св":
                        stringType = "электросварн";
                        break;
                    case "рыж":
                        stringType = "рыж";
                        break;
                    case "г/к":
                        stringType = "горячекатан";
                        break;
                    case "х/к":
                        stringType = "холоднокатан";
                        break;
                    case "х/д":
                        stringType = "холоднодеформированн";
                        break;
                    case "б/ш":
                        stringType = "бесшовн";
                        break;
                    case "рифл":
                        stringType = "рифлен";
                        break;
                    case "проф":
                        stringType = "профильн";
                        break;
                    default:
                        stringType = "";
                        break;
                }

                if (!string.IsNullOrEmpty(stringType))
                    stringType = RightEnding(stringType, Name);
            }
            return stringType;
        }

        public string GetTypeLongFromShortEndNull(string stringType, string Name)
        {
            stringType = stringType.Replace("x", "х");
            stringType = stringType.Replace("k", "к");
            if (new Regex(@"оцин|эл.св|рыж|г/к|х/к|х/д|б/ш|рифл|проф|э/св|оц", RegexOptions.IgnoreCase).IsMatch(stringType))
            {
                stringType = new Regex(@"оцин|эл.св|э/св|рыж|г/к|х/к|х/д|б/ш|рифл|проф|оц", RegexOptions.IgnoreCase).Match(stringType).Value.ToLower();
                switch (stringType)
                {
                    case "оцин":
                        stringType = "оцинкованн";
                        break;
                    case "оц":
                        stringType = "оцинкованн";
                        break;
                    case "эл.св":
                        stringType = "электросварн";
                        break;
                    case "э/св":
                        stringType = "электросварн";
                        break;
                    case "рыж":
                        stringType = "рыж";
                        break;
                    case "г/к":
                        stringType = "горячекатан";
                        break;
                    case "х/к":
                        stringType = "холоднокатан";
                        break;
                    case "х/д":
                        stringType = "холоднодеформированн";
                        break;
                    case "б/ш":
                        stringType = "бесшовн";
                        break;
                    case "рифл":
                        stringType = "рифлен";
                        break;
                    case "проф":
                        stringType = "профильн";
                        break;
                    default:
                        stringType = "";
                        break;
                }

                if (!string.IsNullOrEmpty(stringType))
                    stringType = RightEnding(stringType, Name);
            }
            else stringType = "";
            return stringType;
        }

        public string SetRightEnding(string type, string name)
        {
            if (type.Length > 4 && !type.ToLower().Contains("в изоляц"))
            {
                type = type.Remove(type.Length - 2, 2);
                if (!string.IsNullOrEmpty(type))
                    type = RightEnding(type, name);
            }
            else
                type = GetTypeLongFromShortEndNull(type, name);
            return type;
        }

        private string RightEnding(string type, string name)
        {
            if (new Regex(@"\b[э]\w+\b", RegexOptions.IgnoreCase).IsMatch(type) && !type.ToLower().Contains("электросвар"))
            { type += "ой"; }
            else
            {
                if (new Regex(@"\b\w+[аеёиуыэюь]\b", RegexOptions.IgnoreCase).IsMatch(name))
                {
                    type += "ая";
                }
                else if (new Regex(@"\b\w+[оя]\b", RegexOptions.IgnoreCase).IsMatch(name))
                {
                    type += "ое";
                }
                else if (new Regex(@"\b\w+[бвгджзклмнпрстфхцчшщъ]\b", RegexOptions.IgnoreCase).IsMatch(name))
                {
                    type += "ый";
                }
            }
            return type;
        }


        //изменить тип на нержавеющий, если марка подходящая
        public string GetTypeIfMarkNerj(string Name)
        {
            string type = "";
            if (new Regex(@"\b\w+[аеёиуыэюь]\b", RegexOptions.IgnoreCase).IsMatch(Name))
            {
                type += "нержавеющая";
            }
            else if (new Regex(@"\b\w+[оя]\b", RegexOptions.IgnoreCase).IsMatch(Name))
            {
                type += "нержавеющее";
            }
            else if (new Regex(@"\b\w+[бвгджзклмнпрстфхцчшщъ]\b", RegexOptions.IgnoreCase).IsMatch(Name))
            {
                type += "нержавеющий";
            }
            return type;
        }

        public Regex OrgMobileTelefon { get => orgMobileTelefon; set => orgMobileTelefon = value; }
        public Regex OrgMobileTelefon_tel_XXXXXXXXXXX { get => orgMobileTelefon_tel_XXXXXXXXXX; set => orgMobileTelefon_tel_XXXXXXXXXX = value; }
        public Regex OrgMobileTelefon_tel_X_XXX_XXX_XX_XX { get => orgMobileTelefon_tel_XXX_XXX_XX_XX; set => orgMobileTelefon_tel_XXX_XXX_XX_XX = value; }
        public Regex OrgTelefon_XXX_XX_XX { get => orgTelefon_XXX_XX_XX; set => orgTelefon_XXX_XX_XX = value; }
        public Regex DiamDxDxD { get => diamDxDxD; set => diamDxDxD = value; }
        public Regex DiamDxD { get => diamDxD; set => diamDxD = value; }
        public Regex DiamDxDxD_Only { get => diamDxDxD_Only; set => diamDxDxD_Only = value; }
        public Regex DiamDxD_Only { get => diamDxD_Only; set => diamDxD_Only = value; }
        public Regex DiamD { get => diamD; set => diamD = value; }
        public Regex DiamD_Only { get => diamD_Only; set => diamD_Only = value; }
        public Regex EMail { get => eMail; set => eMail = value; }
        public Regex Site { get => site; set => site = value; }
        public Regex RegMark { get => regMark; set => regMark = value; }
        public Regex RegMark2 { get => regMark2; set => regMark2 = value; }
        public Regex RegTU { get => regTU; set => regTU = value; }
        public Regex RegDiam2 { get => regDiam2; set => regDiam2 = value; }
        public Regex RegDiam { get => regDiam; set => regDiam = value; }
        public Regex RegTypeShveller { get => regTypeShveller; set => regTypeShveller = value; }
        /// <summary>
        /// Тип швеллера М,Ш,К,У,П с цифрой на конце
        /// </summary>
        public Regex RegTypeShveller2 { get => regTypeShveller2; set => regTypeShveller2 = value; }
        public Regex RegName2 { get => regName2; set => regName2 = value; }
        public Regex RegName { get => regName; set => regName = value; }
        public Regex RegType { get => regType; set => regType = value; }
        public Regex INN_KPP { get => iNN_KPP; set => iNN_KPP = value; }
        public Regex INN { get => iNN; set => iNN = value; }
        public Regex R_S { get => r_S; set => r_S = value; }
        public Regex K_S { get => k_S; set => k_S = value; }
        public Regex BIK { get => bik; set => bik = value; }
        public Regex OrgAdres { get => orgAdres; set => orgAdres = value; }
        public Regex OrgAdresFull { get => orgAdresFull; set => orgAdresFull = value; }
        public Regex OrgAdresFully { get => orgAdresFully; set => orgAdresFully = value; }
        public DateTime GetDateTimeFromName(string filepath)
        {
            return getDateFromName(filepath);
        }

        public string DelSpacesInWords(string inString)
        {
            string outString = "";
            outString = new Regex(@"(?<=\b\w\b)\s(?=\b\w\b)", RegexOptions.IgnoreCase).Replace(inString, "");
            return outString;
        }
    }
}
