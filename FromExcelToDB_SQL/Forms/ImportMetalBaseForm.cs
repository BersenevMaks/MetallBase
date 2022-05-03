using MetalBase.ClassesParsers.Chel;
using MetalBase.ClassesParsers.Ekb;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


namespace MetalBase
{
    public partial class ImportMetalBase : Form
    {
        string sqlConnString = @"Server=maks-pc\SQLEXPRESS;Database=MetalBase;User ID=metuser;Password=metuser";
        public ImportMetalBase(string connectionString)
        {
            InitializeComponent();
            if (connectionString != "")
                sqlConnString = connectionString;
        }

        bool isOpenSqlConnection = false;
        DataTable dtProduct = new DataTable();

        SqlConnection conn;
        List<int> listIndexOfNotEmptyName = new List<int>();
        List<int> listIndexOfEmptyName = new List<int>();
        List<int> listShiftIndex = new List<int>();
        C_RegexParamProduct regexParam = new C_RegexParamProduct();
        int countEmpty = 0;
        int colRow = 0;
        int countRowsForShift = 0;

        string nameProd = "";
        string orgname = "";

        private void btn_Connect_Click(object sender, EventArgs e)
        {

            if (!isOpenSqlConnection)
            {
                try
                {
                    conn = new SqlConnection(sqlConnString);
                    conn.Open();
                    isOpenSqlConnection = true;
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            }
            else
            {
                conn.Close();
                isOpenSqlConnection = false;
            }
        }

        private string filePath;
        private Excel.Application excelapp;
        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;
        private Excel.Sheets excelsheets;
        bool isExcelOpen = false;
        int colForName = 0; // номер столбца когда первый раз встретилось имя
        bool isGost = false; // маркер для определения откуда взят гост, из имени или из отдельного столбца, true - из имени
        bool isMark = false; // маркер для определения откуда взята марка, из имени или из отдельного столбца, true - из имени
        bool isTelefon = false; // маркер для опеределения был ли уже взят телефон, для исключения повторов
        int filesCountInDirectory = 0; //количество файлов для прогресс бара
        private DateTime datetimeNow = DateTime.Now;

        delegate void dataSourceDelegate();

        string ManualStringNameProd = "";

        void getNameProduct(string param)
        {
            nameProd = param;
        }

        private void btn_ReadExcel_Click(object sender, EventArgs e)
        {

            dataGridView1.DataSource = null;
            OpenFileDialog ofd = new OpenFileDialog();
            string md = Environment.GetFolderPath(Environment.SpecialFolder.Personal);//путь к Документам
            ofd.InitialDirectory = md;//System.IO.Directory.GetCurrentDirectory();
            ofd.Filter = "All|*.xls;*.xlsx|Excel|*.xls|Excel 2010|*.xlsx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                ChoosingShemaFile(ofd.FileName);
            }
        }

        ///<summary> 
        ///<remarks> Выбор шаблона для работы с файлом </remarks>
        ///<param name="FileName" >путь к файлу</param>
        ///</summary>
        private void ChoosingShemaFile(string FileName)
        {
            textBoxBIK.Text = "";
            textBoxOrgAdress.Text = "";
            textBoxOrgEmail.Text = "";
            textBoxOrgINN.Text = "";
            textBoxOrgKS.Text = "";
            textBoxOrgName.Text = "";
            textBoxOrgRS.Text = "";
            textBoxOrgSite.Text = "";
            textBoxOrgTelefon.Text = "";
            listViewAdrSklad.Items.Clear();
            listViewManager.Items.Clear();
            string name = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+\s*г?\.?\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?|\.docx?)").Match(Path.GetFileName(FileName)).Value.ToLower();
            name = GetFileNameForShemaFile(name, regexParam.GetDateTimeFromName(FileName));

            tsLabelClearingTable.Text = "выполняется обработка...";
            if (name != "")
                switch (name.Trim())
                {
                    case "ростехком160819":
                        Rostehcom(FileName); //специальный шаблон для А-групп от 12.07.19
                        break;
                    case "агрупп120719":
                        A_Grup120719(FileName); //специальный шаблон для А-групп от 12.07.19
                        break;
                    case "утксталь":
                        UTK_Stal_EKB(FileName); //специальный шаблон для Утк-Сталь
                        break;
                    case "континетчел":
                        KontinetChel(FileName); //специальный шаблон для Континет 
                        break;
                    case "челябинскпрофит29072019":
                        ChelyabinskProfit290719(FileName); //специальный шаблон для челябинск профит от 29.07.19
                        break;
                    case "стальтранзитзлат":
                        StalTranzitZlat(FileName); //специальный шаблон для СтальТранзит
                        break;
                    case "металлснабурал220719":
                        MetallSnabUral220719(FileName); //специальный шаблон для МеталлСнабУрал от 22.07.19
                        break;
                    case "трубанаскладе":
                        TrubaNaSklade(FileName); //специальный шаблон для Труба на складе
                        break;
                    case "металлбаза3":
                        MetallBasa3(FileName); //специальный шаблон для МеталлБаза 3 ЕКБ
                        break;
                    case "стальнойпрофиль":
                        StalnoyProfil(FileName); //специальный шаблон для Стальной профиль
                        break;
                    case "первнержком":
                        PervNerjKom(FileName); //специальный шаблон для Первая нержавеющая компания
                        break;
                    case "промметмск":
                        Prommet_MSK(FileName); //специальный шаблон для ПромМет_МСК
                        break;
                    case "амет":
                        Amet(FileName); //специальный шаблон для Амет
                        break;
                    case "алмас":
                        Almas(FileName); //специальный шаблон для Алмас
                        break;
                    case "металбрендурал":
                        MetalBrendUral(FileName); //специальный шаблон для МеталБрендУрал
                        break;
                    case "стальмашурал":
                        StalMashUral(FileName); //специальный шаблон для СтальМашУрал
                        break;
                    case "максмет2":
                        MaxMet2(FileName);   //специальный шаблон для МаксМет2
                        break;
                    case "максмет":
                        MaxMet(FileName);   //специальный шаблон для МаксМет
                        break;
                    case "атлантик":
                        Atlantic(FileName);   //специальный шаблон для Атлантик
                        break;
                    case "бинг":
                        Bing(FileName);   //специальный шаблон для Бинг
                        break;
                    case "стальком":
                        Stalcom(FileName);   //специальный шаблон для Стальком
                        break;
                    case "ростехком":
                        Rostehcom(FileName);   //специальный шаблон для Ростехком
                        break;
                    case "дакор":
                        Dakor(FileName);   //специальный шаблон для Дакор
                        break;
                    case "тдметиз":
                        TD_Metiz(FileName);   //специальный шаблон для ТД Метиз
                        break;
                    case "регионметпром":
                        RegionMetProm(FileName);   //специальный шаблон для РегионМетПром
                        break;
                    case "металлград":
                        MetallGrad(FileName);   //специальный шаблон для МеталлГрад
                        break;
                    case "уралтеплоэнергосервис":
                        UralTeploEnergoService(FileName);   //специальный шаблон для УралТеплоЭнергоСервис
                        break;
                    case "уралцентрсталь":
                        UralCentrStal(FileName);   //специальный шаблон для УралЦентрСталь
                        break;
                    case "стройтехнология":
                        StroiTehnolog(FileName);   //специальный шаблон для СтройТехнология
                        break;
                    case "снабметалсервис":
                        SnabMetalServis(FileName);   //специальный шаблон для СнабМеталСервис
                        break;
                    case "спецметкомплект":
                        SpecMetKomplekt(FileName);   //специальный шаблон для СпецМетКомплект
                        break;
                    case "трубамет":
                        TrubaMet(FileName);   //специальный шаблон для ТрубаМет
                        break;
                    case "спецтруба":
                        SpecTruba(FileName);   //специальный шаблон для СпецТруба
                        break;
                    case "трубмет": //специальный шаблон для трубмет
                        TrubMet(FileName);
                        break;
                    case "металл-база":    //специальный шаблон для Металл-База
                        Metall_Baza(FileName);
                        break;
                    case "стальной профиль":   //специальный шаблон для Стальной профиль
                        StalProfOld(FileName);
                        break;
                    case "инмет":       //специальный шаблон для Инмет
                        Inmet(FileName);
                        break;
                    case "инрост":      //специальный шаблон для Инрост
                        Inrost(FileName);
                        break;
                    case "трубный дом":      //специальный шаблон для Трубный дом
                        TrubDom(FileName);
                        break;
                    case "уралтрубосталь":      //специальный шаблон для УралТрубосталь
                        UralTrubomet(FileName);
                        break;
                    case "спк110517":         //специальный шаблон для СПК до 11.05.17
                        SPK110517(FileName);
                        break;
                    case "спк080719":         //специальный шаблон для СПК от 08.07.19
                        SPK080719(FileName);
                        break;
                    case "металлург":         //специальный шаблон для Металлург до 22.07.19
                        Metallurg(FileName);
                        break;
                    case "металлург220719":         //специальный шаблон для Металлург от 22.07.19
                        Metallurg220719(FileName);
                        break;
                    case "чзпт":            //специальный шаблон для ЧЗПТ
                        CHZPT(FileName);
                        break;
                    case "ммк":            //специальный шаблон для ММК
                        MMK(FileName);
                        break;
                    case "металлинвест":            //специальный шаблон для МаталлИнвест
                        MetallInvest(FileName);
                        break;
                    case "стихеев":            //специальный шаблон для Стихеев
                        Stiheev(FileName);
                        break;
                    case "евразметалл":            //специальный шаблон для ЕвразМеталл
                        Evraz(FileName);
                        break;
                    case "металлкомплект":            //специальный шаблон для МеталлКомплект
                        MettallKomplekt(FileName);
                        break;
                    case "устч":            //специальный шаблон для МеталлКомплект
                        USTCH(FileName);
                        break;
                    case "интермет":            //специальный шаблон для Интермет
                        Intermet(FileName);
                        break;
                    case "алисмет":            //специальный шаблон для Алисмет
                        AlisMet(FileName);
                        break;
                    case "металлсервисцентр":            //специальный шаблон для МеталлСервисЦентр
                        MetallServisCentr(FileName);
                        break;
                    case "промснабжение":            //специальный шаблон для ПромСнабжение
                        PromSnab(FileName);
                        break;
                    case "челябинскпрофит":            //специальный шаблон для ЧелябинскПрофит
                        ChelyabinskProfit(FileName);
                        break;
                    case "металлком":            //специальный шаблон для Металлком
                        MetallKom(FileName);
                        break;
                    case "проммет":            //специальный шаблон для Проммет
                        PromMet(FileName);
                        break;
                    case "ПромГруппа":            //специальный шаблон для ПромГруппа
                        PromMet(FileName);
                        break;
                    case "ПромГруппа_Круг":            //специальный шаблон для Промгруппа_Круг
                        PromMet(FileName);
                        break;
                    case "апогей":            //специальный шаблон для Апогей
                        Apogei(FileName);
                        break;
                    case "илеко":            //специальный шаблон для Илеко
                        Ileko(FileName);
                        break;
                    case "демидов":            //специальный шаблон для Демидов
                        Demidov(FileName);
                        break;
                    case "демидов_труба":            //специальный шаблон для Демидов Труба профильная
                        DemidovTruba(FileName);
                        break;
                    case "стальмаксимум":            //специальный шаблон для СтальМаксимум
                        StalMaksimum(FileName);
                        break;
                    case "вика":            //специальный шаблон для СтальМаксимум
                        Vika(FileName);
                        break;
                    case "пермь":            //специальный шаблон для Пермь
                        Perm(FileName);
                        break;
                    case "а_групп_труб":
                        AGroupTrub(FileName);   //специальный шаблон для AGroup трубы
                        break;
                    case "а_групп_труб_проф":
                        AGroupTrubProf(FileName);   //специальный шаблон для AGroup трубы профильные
                        break;
                    case "атомРос":
                        AtomRos(FileName);   //специальный шаблон для Атом-Рос
                        break;
                    case "инкомМеталл":
                        InkomMetal(FileName);   //специальный шаблон для инком Металл
                        break;
                    case "ксм":
                        KSM(FileName);   //специальный шаблон для ксм
                        break;
                    case "гранд_универсал":
                        GrandUniversal(FileName);   //специальный шаблон для гранд_универсал
                        break;
                    case "гарус":
                        Garus(FileName);   //специальный шаблон для гарус
                        break;
                    case "уптк":
                        UPTK(FileName);   //специальный шаблон для уптк
                        break;
                    case "атомпромкомплекс":
                        AtomPromKomp(FileName);   //специальный шаблон для атом пром комплекс
                        break;
                    case "уралметстрой":
                        UralMetStroi(FileName);   //специальный шаблон для УралМетСтрой
                        break;
                    case "спецстальм":
                        SpecStal(FileName);   //специальный шаблон для СпецСталь-М
                        break;
                    case "роспромцентр":
                        RosPromCentr(FileName);   //специальный шаблон для РосПромЦентр
                        break;
                    case "кузнецов":
                        Kuznetsov(FileName);   //специальный шаблон для Кузнецов
                        break;
                    case "теплообменныетрубы":
                        TeploobmenTrub(FileName);   //специальный шаблон для Теплообменные Трубы
                        break;
                    case "золотойвек":
                        ZolotoyVek(FileName);   //специальный шаблон для Золотой век до 26.11.17
                        break;
                    case "золотойвек240719":
                        ZolotoyVek240719(FileName);   //специальный шаблон для Золотой век от 24.07.19
                        break;
                    case "метчив":
                        Metchiv(FileName);   //специальный шаблон для Метчив
                        break;
                    case "меднаягора":
                        MedGora(FileName);   //специальный шаблон для Медная гора
                        break;
                    case "промметал":
                        Prommet(FileName);   //специальный шаблон для ПромМеталл
                        break;
                    case "сибметалл":
                        SibMetal(FileName);   //специальный шаблон для СибМеталл
                        break;
                    case "скат":
                        Skat(FileName);   //специальный шаблон для Скат
                        break;
                    case "металлпромснаб":
                        MetallPromSnab(FileName);   //специальный шаблон для МеталлПромСнаб
                        break;
                    case "стальмаркет":
                        StalMarket(FileName);   //специальный шаблон для СтальМаркет
                        break;
                    case "стройтехцентр":
                        StroiTehCentr(FileName);   //специальный шаблон для СтройТехЦентр
                        break;
                    case "тэл":
                        TEL(FileName);   //специальный шаблон для ТЭЛ
                        break;
                    case "умпц":
                        UMPC(FileName);   //специальный шаблон для УМПЦ
                        break;
                    case "уралпромметалл":
                        UralPromMetal(FileName);   //специальный шаблон для УралПромМеталл
                        break;
                    case "уралпромметалл120719":
                        UralPromMetal120719(FileName);   //специальный шаблон для УралПромМеталл
                        break;
                    case "уралчерметДо19.08.19":
                        UralCherMet_do19082019(FileName);   //специальный шаблон для УралЧерМет до 19.08.19
                        break;
                    case "уралчермет19.08.19":
                        UralcherMet190819(FileName);   //специальный шаблон для УралЧерМет от 19.08.19
                        break;
                    case "техномет":
                        TehnoMet(FileName);   //специальный шаблон для ТехноМет
                        break;
                    case "инплано":
                        Inplano(FileName);   //специальный шаблон для Инплано
                        break;
                    case "эгидапром":
                        EgidaProm(FileName);   //специальный шаблон для Эгидапром
                        break;
                    case "профмет<=25.01.18":
                        ProfMet250118(FileName);   //специальный шаблон для Профмет до 25.01.18
                        break;
                    case "профмет<=14.08.19":
                        ProfMet140819(FileName);   //специальный шаблон для Профмет от 25.01.18 до 14.08.19
                        break;
                    case "альфаметалл":
                        AlfaMetall(FileName);   //специальный шаблон для Альфаметал
                        break;
                    case "уральскаяметаллобаза":
                        UralskayaMetallobaza(FileName);   //специальный шаблон для уральская Металлобаза
                        break;
                    case "энергоальянс":
                        EnergoAlyans(FileName);   //специальный шаблон для Энергоальянс
                        break;

                    default: //для первых шаблонов или в качестве стандартного, если нет специального шаблона
                        OpenAndgetExcel(FileName);
                        break;
                }
            else
            {
                MessageBox.Show("Не удалось получить имя файла.\nПроверьте название файла:\n1.Имя должно быть отделено пробелом от даты;\n2.Дата должна быть формата дд.мм.гг или дд.мм.гггг;\n3.Проверьте работает ли файл если его просто открыть на компьютере.\n\nЕсли файл по прежнему не обрабатывается - обратитесь к разработчику");
            }
        }

        ///<summary> 
        ///<remarks> Выбор шаблона для работы с файлом </remarks>
        ///<param name="FileName" >путь к файлу</param>
        ///</summary>
        private string GetFileNameForShemaFile(string name, DateTime dateInFile)
        {
            string NameTemp = "";

            NameTemp = new Regex(@"\bростехком\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "ростехком"; }
            NameTemp = new Regex(@"\bутк\s*?[_\s-]\s*?сталь\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "утксталь"; }
            NameTemp = new Regex(@"\bконтинет(?:.*?чел)?\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "континетчел"; }
            NameTemp = new Regex(@"^металлург\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            {
                DateTime dt = new DateTime(2019, 07, 22);
                if (dateInFile >= dt)
                    name = "металлург220719";
                else name = "металлург";
            }
            NameTemp = new Regex(@"золотой\s*век", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            {
                DateTime dt = new DateTime(2019, 07, 24);
                if (dateInFile >= dt)
                    name = "золотойвек240719";
                else name = "золотойвек";
            }
            NameTemp = new Regex(@"\bчелябинск\s*?профит\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            {
                DateTime dt = new DateTime(2019, 7, 29);
                if (dateInFile >= dt)
                    name = "челябинскпрофит29072019";
                else name = "челябинскпрофит";
            }
            NameTemp = new Regex(@"\bсталь\s*транзит\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "стальтранзитзлат"; }
            NameTemp = new Regex(@"\bметалл?.*снаб.*урал\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            {
                DateTime dt = new DateTime(2019, 7, 22);
                if (dateInFile >= dt)
                    name = "металлснабурал220719";
            }
            NameTemp = new Regex(@"\bспк\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            {
                DateTime dt = new DateTime(2017, 5, 11);
                if (dateInFile <= dt)
                    name = "спк110517";
                else
                    name = "спк080719";
            }
            NameTemp = new Regex(@"\bтруба\s*на\s*складе\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "трубанаскладе"; }
            NameTemp = new Regex(@"\bметалл?\s*база[\-\s]*3", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "металлбаза3"; }
            NameTemp = new Regex(@"\bстальной\s*профиль\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "стальнойпрофиль"; }
            NameTemp = new Regex(@"\bпервая\s*нержавеющая\s*компания\s*", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "первнержком"; }
            NameTemp = new Regex(@"\bпро.*мет(?:\b\s*мск|\b.*мск)", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "промметмск"; }
            NameTemp = new Regex(@"\bамет\b||\bamet\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "амет"; }
            NameTemp = new Regex(@"\bалмас\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "алмас"; }
            NameTemp = new Regex(@"\bметал.*бр[еэ]нд.*урал\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "металбрендурал"; }
            NameTemp = new Regex(@"\bстальмашурал\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "стальмашурал"; }
            NameTemp = new Regex(@"\bмакс(?:[-\s]*)?мет\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            {
                if (new Regex(@"\bмакс(?:[-\s]*)?мет\b[-\s_]*2(?=(?:\d+\.\d+.\d+)|$)", RegexOptions.IgnoreCase).IsMatch(name))
                    name = "максмет2";
                else name = "максмет";
            }
            NameTemp = new Regex(@"\bатлантик\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "атлантик"; }
            NameTemp = new Regex(@"\bбинг\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "бинг"; }
            NameTemp = new Regex(@"стальком", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "стальком"; }
            NameTemp = new Regex(@"ростехком", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "ростехком"; }
            NameTemp = new Regex(@"\bдакор\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            {
                name = "дакор";
            }
            NameTemp = new Regex(@"тд\s*метиз", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "тдметиз"; }
            NameTemp = new Regex(@"\bрегионметпром\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "регионметпром"; }
            NameTemp = new Regex(@"\bМеталл?\s*град\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "металлград"; }
            NameTemp = new Regex(@"\bурал\s*тепло\s*энерго\s*сервис\b|\bутэс\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "уралтеплоэнергосервис"; }
            NameTemp = new Regex(@"\bурал\s*центр\s*сталь\b|\bуцс\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "уралцентрсталь"; }
            NameTemp = new Regex(@"\bстрой.*технология\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "стройтехнология"; }
            NameTemp = new Regex(@"\bснабметалл?сервис\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "снабметалсервис"; }
            NameTemp = new Regex(@"\bспецметкомплект\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "спецметкомплект"; }
            NameTemp = new Regex(@"\bтрубамет\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "трубамет"; }
            NameTemp = new Regex(@"\bспецтруба\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "спецтруба"; }
            //NameTemp = new Regex(@".*стальной.*профиль.*", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            //if (NameTemp != "")
            //{ name = "стальной профиль"; }
            NameTemp = new Regex(@"метал.*инвест", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "металлинвест"; }
            NameTemp = new Regex(@"уст\s*-?\s*ч", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "устч"; }
            NameTemp = new Regex(@"интер.*мет", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "интермет"; }
            NameTemp = new Regex(@"алис.*мет", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "алисмет"; }
            NameTemp = new Regex(@"МСЦ|Метал.*серв.*цент.*", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "металлсервисцентр"; }
            NameTemp = new Regex(@"металл?комм?\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "металлком"; }
            NameTemp = new Regex(@"металл?о?[\s\-]*комплект(?:[\s-]*M)?", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "металлкомплект"; }
            NameTemp = new Regex(@"пром.*мет\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "проммет"; }
            NameTemp = new Regex(@"апогей", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "апогей"; }
            NameTemp = new Regex(@"пром.*груп", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            {
                if (new Regex(@"пром.*груп.*круг", RegexOptions.IgnoreCase).IsMatch(name))
                { name = "ПромГруппа_Круг"; }
                else
                { name = "ПромГруппа"; }
            }
            NameTemp = new Regex(@"демидов", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            {
                if (new Regex(@"труба", RegexOptions.IgnoreCase).IsMatch(name))
                    name = "демидов_труба";
                else name = "демидов";
            }
            NameTemp = new Regex(@"сталь.*макс\w*\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "стальмаксимум"; }

            NameTemp = new Regex(@"вика", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "вика"; }

            NameTemp = new Regex(@"пермь", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "пермь"; }

            NameTemp = new Regex(@"\bа.?груп.*труб", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            {
                DateTime dt = new DateTime(2019, 7, 12);
                if (dateInFile < dt)
                {
                    if (!new Regex(@"а.?груп.*труб.*проф", RegexOptions.IgnoreCase).IsMatch(name))
                        name = "а_групп_труб";
                    else name = "а_групп_труб_проф";
                }
            }

            NameTemp = new Regex(@"\bа\s*?-\s*?групп?\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            {
                DateTime dt = new DateTime(2019, 7, 12);
                if (dateInFile >= dt)
                {
                    name = "агрупп120719";
                }
            }

            NameTemp = new Regex(@"атом.*рос", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "атомРос"; }

            NameTemp = new Regex(@"инком", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "инкомМеталл"; }
            NameTemp = new Regex(@"\bксм\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "ксм"; }
            NameTemp = new Regex(@"\bгранд.*универсал\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "гранд_универсал"; }
            NameTemp = new Regex(@"\bуптк\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "уптк"; }
            NameTemp = new Regex(@"атом\s*пром\s*комплекс", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "атомпромкомплекс"; }
            NameTemp = new Regex(@"урал\s*мет\s*строй", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "уралметстрой"; }
            NameTemp = new Regex(@"спец\s*сталь", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "спецстальм"; }
            NameTemp = new Regex(@"роспромцентр|уралметалл?ургкомплект", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "роспромцентр"; }
            NameTemp = new Regex(@"кузнецов", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "кузнецов"; }
            NameTemp = new Regex(@"теплоо\w+\s*труб", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "теплообменныетрубы"; }
            NameTemp = new Regex(@"метчив", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "метчив"; }
            NameTemp = new Regex(@"медная\s*гора", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "меднаягора"; }
            NameTemp = new Regex(@"\bпромметалл?\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "промметал"; }
            NameTemp = new Regex(@"сиб.*метал", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "сибметалл"; }
            NameTemp = new Regex(@"\bскат\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "скат"; }
            NameTemp = new Regex(@"\bметаллпромснаб\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "металлпромснаб"; }
            NameTemp = new Regex(@"\bстальмаркет\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "стальмаркет"; }
            NameTemp = new Regex(@"\bстройтехцентр\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "стройтехцентр"; }
            NameTemp = new Regex(@"\bтэл\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "тэл"; }
            NameTemp = new Regex(@"\bумпц\b|урал.*\bметалопром.*\bцентр\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "умпц"; }
            NameTemp = new Regex(@"\bуралпромметалл?\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            {
                DateTime dt = new DateTime(2019, 7, 12);
                if (dateInFile >= dt)
                    name = "уралпромметалл120719";
                else { name = "уралпромметалл"; }
            }
            NameTemp = new Regex(@"\bуралчермет\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            {
                DateTime dt = new DateTime(2019, 7, 19);
                if (dateInFile >= dt)
                    name = "уралчермет19.08.19";
                else { name = "уралчерметДо19.08.19"; }
            }
            NameTemp = new Regex(@"\bтехномет\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "техномет"; }
            NameTemp = new Regex(@"\bинплано\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "инплано"; }
            NameTemp = new Regex(@"\bэгидапром\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "эгидапром"; }
            NameTemp = new Regex(@"\bпрофмет\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            {
                DateTime dt = new DateTime(2018, 1, 25);
                if (dateInFile <= dt)
                    name = "профмет<=25.01.18";
                else
                    name = "профмет<=14.08.19";
            }
            NameTemp = new Regex(@"\bальфа-метал\w?\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "альфаметалл"; }
            NameTemp = new Regex(@"\bуральская\b\s*\bметал\w?обаза\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "уральскаяметаллобаза"; }
            NameTemp = new Regex(@"\bэнергоальянс\b", RegexOptions.IgnoreCase).Match(name).Value.ToString();
            if (NameTemp != "")
            { name = "энергоальянс"; }

            return name;
        }

        ///<summary> 
        ///<remarks> Открытие и чтение вордовского файла Стихеев </remarks>
        ///<param name="path" >путь к файлу</param>
        ///</summary>
        private void Stiheev(string path)
        {
            bool isOpenWord = false;
            textBoxPath.Text = path;
            filePath = path;

            SetNameFromName(filePath);
            SetDateFromName(filePath);

            Word._Application application;
            Word._Document document;
            Object missingObj = System.Reflection.Missing.Value;
            Object trueObj = true;
            Object falseObj = false;
            //создаем обьект приложения word
            application = new Word.Application();
            // создаем путь к файлу
            Object templatePathObj = path;
            Word.Tables tables;
            try
            {
                document = application.Documents.Open(ref templatePathObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj);

                tables = document.Tables;
                isOpenWord = true;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                tsLabeltotalSheets.Text = "1";

                Regex regName = new Regex(@"(?!\w+ая|\w+ые|\w+ое|\w+ой|\w+ый|\w\d\w?|\d+(?:[,.]\d+)?(?:\s*[xх*]\s*\d+(?:[,.]\d+)?)+|\bгост\b)\b\w{3,}\b(?=\s+\d|\s+\w+|\s*$)", RegexOptions.IgnoreCase);
                Regex regType = new Regex(@"\w+ая|\w+ые|\w+ое|\w+ой|\w+ый|эл[\/]св\w*\b|оц\w*|ВГП", RegexOptions.IgnoreCase);
                Regex regDiam = new Regex(@"\d+(?:[,.]\d+)?(?:\s*[xх*]\s*\d+(?:[,.]\d+)?)+", RegexOptions.IgnoreCase);

                string name = "";

                for (int t = 1; t <= tables.Count; t++)
                {
                    string temp = "";
                    Word.Table tab = tables[t];
                    int cColTab = tab.Columns.Count;
                    int cRowTab = tab.Rows.Count;
                    tsLabelcurrSheet.Text = "1";
                    tsPb1.Maximum = cColTab * cRowTab;
                    bool isAdded = false;

                    tsLabelClearingTable.Text = "Поиск наименований";
                    tsPb1.Value = 0;
                    for (int i = 1; i < cRowTab; i++)
                    {
                        for (int j = 1; j < cColTab; j++)
                        {
                            temp = tab.Cell(i, j).Range.Text;
                            temp = temp.Replace("\r\a", string.Empty);
                            if (temp != "")
                            {
                                if (new Regex(@"наимен", RegexOptions.IgnoreCase).IsMatch(temp))
                                {

                                    tsLabelClearingTable.Text = "Заполнение";
                                    tsPb1.Value = 0;
                                    tsPb1.Maximum = cRowTab;
                                    for (int ii = i + 1; ii <= cRowTab; ii++)
                                    {
                                        name = "";
                                        temp = tab.Cell(ii, j).Range.Text;
                                        temp = temp.Replace("\r\a", string.Empty);
                                        if (temp != "")
                                        {
                                            if (name != "" || regName.IsMatch(temp))
                                            {
                                                if (regDiam.IsMatch(temp))
                                                {
                                                    dtProduct.Rows.Add();
                                                    isAdded = true;
                                                    int lastRow = dtProduct.Rows.Count - 1;
                                                    dtProduct.Rows[lastRow]["Название"] =
                                                        regName.Match(temp).Value;
                                                    if (dtProduct.Rows[lastRow]["Название"].ToString().Trim() == "")
                                                        dtProduct.Rows[lastRow]["Название"] = name;
                                                    dtProduct.Rows[lastRow]["Тип"] = regType.Match(temp).Value;
                                                    if (new Regex(@"\d+(?:[.,]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).IsMatch(temp))
                                                    {
                                                        dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"\d+(?:[.,]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(temp).Value;
                                                        dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\d+(?:[.,]\d+)?\s*[xх*]\s*)\d+(?:[.,]\d+)?(?=\s|$|\\|/|\()").Match(temp).Value;
                                                        dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=\d+(?:[.,]\d+)?\s*[xх*]\s*)\d+(?:[.,]\d+)?(?=\s*[xх*])").Match(temp).Value;
                                                    }
                                                }
                                                else if (new Regex(@"(?<=\s)\d+(?:[,.]\d+)?(?=[\s\/$,])", RegexOptions.IgnoreCase).IsMatch(temp))
                                                {
                                                    dtProduct.Rows.Add();
                                                    isAdded = true;
                                                    int lastRow = dtProduct.Rows.Count - 1;
                                                    dtProduct.Rows[lastRow]["Название"] =
                                                        regName.Match(temp).Value;
                                                    if (dtProduct.Rows[lastRow]["Название"].ToString().Trim() == "")
                                                        dtProduct.Rows[lastRow]["Название"] = name;
                                                    dtProduct.Rows[lastRow]["Тип"] = regType.Match(temp).Value;
                                                    dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=\s)\d+(?:[,.]\d+)?(?=[\s\/$])", RegexOptions.IgnoreCase).Match(temp).Value;
                                                }
                                                if (dtProduct.Rows.Count > 0 && isAdded)
                                                {
                                                    int lastRow = dtProduct.Rows.Count - 1;
                                                    dtProduct.Rows[lastRow]["Примечание"] = temp;
                                                    temp = tab.Cell(ii, j + 2).Range.Text;
                                                    temp = temp.Replace("\r\a", string.Empty);
                                                    if (temp != "")
                                                    {
                                                        dtProduct.Rows[lastRow]["Цена"] = new Regex(@"\d+(?:\s*[,.]\s*\d+)?", RegexOptions.IgnoreCase).Match(temp.Trim()).Value;
                                                    }
                                                    temp = tab.Cell(ii, j + 1).Range.Text;
                                                    temp = temp.Replace("\r\a", string.Empty);
                                                    if (temp != "")
                                                    {
                                                        dtProduct.Rows[lastRow]["Мерность (т, м, мм)"] = new Regex(@"\d+(?:\s*[,.]\s*\d+)?", RegexOptions.IgnoreCase).Match(temp.Trim()).Value;
                                                    }
                                                    if (new Regex(@"эл[\/]св\w*", RegexOptions.IgnoreCase).IsMatch(dtProduct.Rows[lastRow]["Тип"].ToString().Trim()))
                                                        dtProduct.Rows[lastRow]["Тип"] = "Электросварная";
                                                    if (new Regex(@"оц\w*", RegexOptions.IgnoreCase).IsMatch(dtProduct.Rows[lastRow]["Тип"].ToString().Trim()))
                                                        dtProduct.Rows[lastRow]["Тип"] = "Оцинкованный";
                                                    if (dtProduct.Rows[lastRow]["Тип"].ToString().Trim() == "") dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                                    isAdded = false;
                                                }
                                            }
                                            else if (!new Regex(@"\d", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                if (regName.IsMatch(temp))
                                                {
                                                    name = regName.Match(temp).Value;
                                                }
                                            }
                                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                                            else tsPb1.Value = tsPb1.Maximum;
                                        }
                                    }
                                    i = cRowTab;
                                    break;
                                }
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        }
                    }
                    clearingTable();
                    dataGridView1.DataSource = dtProduct;
                }

            }
            catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла " + Path.GetFileName(path) + "\n\n" + ex.ToString()); }
            if (isOpenWord)
            {
                application.Quit(missingObj, missingObj, missingObj);

            }
        }

        ///<summary> 
        ///<remarks> Открытие и чтение вордовского файла Интермет </remarks>
        ///<param name="path" >путь к файлу</param>
        ///</summary>
        private void Intermet(string path)
        {
            bool isOpenWord = false;
            textBoxPath.Text = path;
            filePath = path;

            SetNameFromName(filePath);
            SetDateFromName(filePath);

            Word._Application application;
            Word._Document document;
            Object missingObj = System.Reflection.Missing.Value;
            Object trueObj = true;
            Object falseObj = false;
            //создаем обьект приложения word
            application = new Word.Application();
            // создаем путь к файлу
            Object templatePathObj = path;
            Word.Tables tables;
            try
            {
                document = application.Documents.Open(ref templatePathObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj);

                tables = document.Tables;
                isOpenWord = true;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                tsLabeltotalSheets.Text = "1";

                Regex regName = new Regex(@"лента|лист|арматура|полоса|угол|швеллер|труб|пвл", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ый\s*с\s*чеч\w*\.?\s*риф\w+\.?|\w+ая|\w+ой(?:\s*проч)|\w+ий", RegexOptions.IgnoreCase);
                Regex regSklad = new Regex(@"(?<=\w*\s*склад.*\s*:\s*)\w+\.\s*\w+,\s*\w+(?:\s*\w+)?,\s*\d+\.\s*.Металл-база.", RegexOptions.IgnoreCase);
                Regex regDiam = new Regex(@"(?<=\s)\d+(?:[,.]\d+)?(?=\s|[бшкм])", RegexOptions.IgnoreCase);
                Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*|ТУ\s*[\d+\.?]+[-\d+]+", RegexOptions.IgnoreCase);
                Regex regMark = new Regex(@"\bст[\s\.]*[\dгсп]+(?:\s*-\s*мд)?|[aа]\s*-\s*i{1,3}(?:\s*-\s*мд)?|[aа]т?\d+[cс]?(?:\s*-\s*мд)?|(?<=\d)[бмшк]\d?", RegexOptions.IgnoreCase);

                for (int t = 1; t <= tables.Count; t++)
                {
                    string tmp = "";
                    int lastRow = 0;
                    structTab tab = new structTab
                    {
                        listExcelIndexTab = new List<int>(),
                        listdtProductIndexRow = new List<int>()
                    };

                    string diam = "", tolsh = "", metraj = "", name = "";

                    Word.Table wTab = tables[t];
                    int cCelCol = wTab.Columns.Count;
                    int cCelRow = wTab.Rows.Count;
                    tsLabelcurrSheet.Text = "1";

                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow;

                    for (int j = 1; j < cCelRow; j++)
                    {
                        int i = 2;
                        tab.StartRow = j;

                        if (wTab.Rows[j].Cells.Count < 2) i = wTab.Rows[j].Cells.Count;
                        if (wTab.Rows[j].Cells.Count == 2) i = 1;
                        tmp = wTab.Cell(j, i).Range.Text;
                        tmp = tmp.Replace("\r\a", string.Empty).Trim();
                        if (regName.IsMatch(tmp) || new Regex(@"\d+(?:[,.]\d+)?(?:\s*[хx\*]\s*\d+(?:[,.]\d+)?)+", RegexOptions.IgnoreCase).IsMatch(tmp) || regDiam.IsMatch(tmp))
                        {
                            dtProduct.Rows.Add();
                            lastRow = dtProduct.Rows.Count - 1;
                            tab.listExcelIndexTab.Add(j);
                            tab.listdtProductIndexRow.Add(lastRow);
                            dtProduct.Rows[lastRow]["Название"] = regName.Match(tmp).Value;
                            if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "труб") dtProduct.Rows[lastRow]["Название"] = "Труба";
                            if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "угол") dtProduct.Rows[lastRow]["Название"] = "Уголок";
                            if (dtProduct.Rows[lastRow]["Название"].ToString().Length > 1)
                                dtProduct.Rows[lastRow]["Название"] = dtProduct.Rows[lastRow]["Название"].ToString().Substring(0, 1).ToUpper() + dtProduct.Rows[lastRow]["Название"].ToString().Substring(1, dtProduct.Rows[lastRow]["Название"].ToString().Length - 1);

                            if (dtProduct.Rows[lastRow]["Название"].ToString() == "")
                            {
                                if (dtProduct.Rows.Count > 1)
                                    if (dtProduct.Rows[lastRow - 1]["Название"].ToString() != "")
                                        dtProduct.Rows[lastRow]["Название"] = dtProduct.Rows[lastRow - 1]["Название"];
                            }

                            dtProduct.Rows[lastRow]["Примечание"] = tmp;
                            dtProduct.Rows[lastRow]["Тип"] = regType.Match(tmp).Value;
                            if (new Regex(@"\w+ый\s*с\s*чеч\w*\.?\s*риф\w+\.?", RegexOptions.IgnoreCase).IsMatch(tmp))
                                dtProduct.Rows[lastRow]["Тип"] = "горячекатаный с чечевичным рифлением";
                            if (dtProduct.Rows[lastRow]["Тип"].ToString() == "") dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                            name = dtProduct.Rows[lastRow]["Название"].ToString();
                            if (new Regex(@"\d+(?:[,.]\d+)?(?:\s*[хx\*]\s*\d+(?:[,.]\d+)?)+", RegexOptions.IgnoreCase).IsMatch(tmp))
                            {
                                if (name.ToLower() == "лист" || name.ToLower() == "полоса")
                                {
                                    diam = new Regex(@"(?<=^|\s)\d+(?:[,.]\d+)?(?=\s*[хx\*]\s*\d+)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                    if (name.ToLower() == "уголок")
                                        tolsh = new Regex(@"(?<=\s*[хx\*]\s*\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s*[хx\*])", RegexOptions.IgnoreCase).Match(tmp).Value;
                                    else
                                        tolsh = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s*[хx\*]\s*\d+(?:[,.]\d+)?|\s)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                    metraj = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                }
                                else
                                {
                                    diam = new Regex(@"(?<=^|\s)\d+(?:[,.]\d+)?(?=\s*[хx\*]\s*\d+)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                    tolsh = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s*[хx\*]|\s)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                    dtProduct.Rows[lastRow]["Мерность (т, м, мм)"] = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                    metraj = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s*[хx\*]\d+(?:[,.]\d+)?\s*[хx\*])", RegexOptions.IgnoreCase).Match(tmp).Value;
                                    if (tolsh == "")
                                        tolsh = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s*[хx\*]|\s)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                }
                                if (!string.IsNullOrEmpty(tolsh))
                                    if (tolsh.Length > 4)
                                        tolsh = tolsh.Replace(",", "");
                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = tolsh;
                                if (!string.IsNullOrEmpty(diam))
                                    dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = diam;
                                if (!string.IsNullOrEmpty(metraj))
                                    if (metraj.Length > 4)
                                        metraj = metraj.Replace(",", "");
                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = metraj;
                            }
                            dtProduct.Rows[lastRow]["Марка"] = regMark.Match(tmp).Value;

                            foreach (Match m in regTU.Matches(tmp))
                            {
                                if (dtProduct.Rows[lastRow]["Стандарт"].ToString() == "") dtProduct.Rows[lastRow]["Стандарт"] = m.Value;
                                else dtProduct.Rows[lastRow]["Стандарт"] += "; " + m.Value;
                            }

                            if (wTab.Rows[j].Cells.Count > 2)
                            {
                                tmp = wTab.Cell(j, 3).Range.Text;
                                tmp = tmp.Replace("\r\a", string.Empty).Trim();
                                if (tmp != "")
                                {
                                    dtProduct.Rows[lastRow]["Марка"] = regMark.Match(tmp).Value;
                                }
                            }

                            if (wTab.Rows[j].Cells.Count > 3)
                            {
                                tmp = wTab.Cell(j, 4).Range.Text;
                                tmp = tmp.Replace("\r\a", string.Empty).Trim();
                                if (tmp != "")
                                {
                                    dtProduct.Rows[lastRow]["Цена"] = tmp;
                                }
                            }
                        }
                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        else tsPb1.Value = tsPb1.Maximum;
                    }
                }

                for (int i = 1; i <= 10; i++)
                {
                    string temp = document.Paragraphs[i].Range.Text;
                    temp = temp.Replace("\r\a", string.Empty).Trim();
                    if (temp != "")
                    {
                        Regex regOrgAdr = new Regex(@"(?<=адрес\s*:\s*)(?:\w+,)?(?:\s*\d+\s*,)?(?:\s*[\w+\s*]+,)?\s*[-\w+\s*]+(?:,\s*[\w+\s*]+)*(?:,\s*\d+)*(?:,\s*офис\s*[\d\w]+)?", RegexOptions.IgnoreCase);
                        if (regOrgAdr.IsMatch(temp))
                        {
                            textBoxOrgAdress.Text = regOrgAdr.Match(temp).Value;

                            Regex regOrgTel = new Regex(@"(?<=(?:Телефоны\s*:\s*)*)(?:\d*\s*\(\d+\)\s*)?(\s*\d+(?:-\d+)+,?\s*)+", RegexOptions.IgnoreCase);
                            if (regOrgTel.IsMatch(temp))
                                textBoxOrgTelefon.Text = regOrgTel.Match(temp).Value;
                        }
                        Regex regManager = new Regex(@"менеджер", RegexOptions.IgnoreCase);
                        if (regManager.IsMatch(temp))
                        {
                            ListViewItem lvi = new ListViewItem(new Regex(@"(?<=менеджер\s*:?\s*)(?:\w+\s+){1,3}(?=\s\s|\d)", RegexOptions.IgnoreCase).Match(temp).Value); //имя менеджера
                            lvi.SubItems.Add(new Regex(@"(?<=менеджер\s*:?\s*(?:\w+\s+){1,3}\s*)[\d-]+(?:\s*\d+[(\s]?\d+[)\s]?\s?[\d-]+)?", RegexOptions.IgnoreCase).Match(temp).Value);           //телефон менеджера
                            if (lvi.SubItems[0].Text != "icq") listViewManager.Items.Add(lvi);
                        }

                        Regex regMail = new Regex(@"(?<=(?:эл.*почта\s*:\s*)?)[\w\d\.-]+@[\w\d\.-]+.(?:ru|com|рф|info)", RegexOptions.IgnoreCase);
                        if (regMail.IsMatch(temp))
                            textBoxOrgEmail.Text = regMail.Match(temp).Value;
                    }
                }

                clearingTable();
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка при обработке файла " + Path.GetFileName(path) + "\n\n" + ex.ToString()); }
            if (isOpenWord)
            {
                application.Quit(missingObj, missingObj, missingObj);

            }
        }

        ///<summary> 
        ///<remarks> Открытие и чтение вордовского файла АлисМет </remarks>
        ///<param name="path" >путь к файлу</param>
        ///</summary>
        private void AlisMet(string path)
        {
            bool isOpenWord = false;
            textBoxPath.Text = path;
            filePath = path;

            SetNameFromName(filePath);
            SetDateFromName(filePath);

            Word._Application application;
            Word._Document document;
            Object missingObj = System.Reflection.Missing.Value;
            Object trueObj = true;
            Object falseObj = false;
            //создаем обьект приложения word
            application = new Word.Application();
            // создаем путь к файлу
            Object templatePathObj = path;
            Word.Tables tables;
            try
            {
                document = application.Documents.Open(ref templatePathObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj);

                tables = document.Tables;
                isOpenWord = true;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                tsLabeltotalSheets.Text = "1";

                Regex regName = new Regex(@"лента|лист|арматура|полоса|угол|швеллер|труб|круг|шестигранник|шгр|квадрат|полоса|сталь|катанка|быстрорез|прокат", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ый\s*с\s*чеч\w*\.?\s*риф\w+\.?|\w+ая|\w+ой(?:\s*проч)|\w+ий", RegexOptions.IgnoreCase);
                Regex regSklad = new Regex(@"(?<=\w*\s*склад.*\s*:\s*)\w+\.\s*\w+,\s*\w+(?:\s*\w+)?,\s*\d+\.\s*.Металл-база.", RegexOptions.IgnoreCase);
                Regex regDiam = new Regex(@"^\d+(?:[,\.]\d+)?$", RegexOptions.IgnoreCase);
                Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*|ТУ\s*[\d+\.?]+[-\d+]+", RegexOptions.IgnoreCase);
                Regex regMark = new Regex(@"\bст[\s\.]*[\dгсп]+(?:\s*-\s*мд)?|[aа]\s*-\s*i{1,3}(?:\s*-\s*мд)?|[aа]т?\d+[cс]?(?:\s*-\s*мд)?|(?<=\d)[бмшк]\d?", RegexOptions.IgnoreCase);
                string tmp = "";
                string temp = "";
                int lastRow = 0;
                structTab tab;
                int jj = 0;

                int ColDiam = 0, ColTolsh = 0, ColMark = 0, ColKolVo = 0, ColPrice = 0;

                for (int t = 1; t <= tables.Count; t++)
                {
                    tab = new structTab
                    {
                        listExcelIndexTab = new List<int>(),
                        listdtProductIndexRow = new List<int>()
                    };

                    Word.Table wTab = tables[t];
                    int cCelCol = wTab.Columns.Count;
                    int cCelRow = wTab.Rows.Count;
                    tsLabelcurrSheet.Text = "1";

                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;

                    for (int j = 1; j <= cCelRow; j++)
                    {
                        jj = j;
                        for (int i = 1; i <= cCelCol; i++)
                        {
                            if (i <= wTab.Rows[jj].Cells.Count)
                            {
                                temp = wTab.Cell(jj, i).Range.Text;
                                temp = temp.Replace("\r\a", string.Empty).Trim();
                                if (new Regex(@"диаметр", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColDiam = i;
                                    tab.StartRow = jj;
                                    j = cCelRow;
                                }
                                if (new Regex(@"стенка", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColTolsh = i;
                                    tab.StartRow = jj;
                                    j = cCelRow;
                                }
                                if (new Regex(@"марка", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColMark = i;
                                    tab.StartRow = jj;
                                    j = cCelRow;
                                }
                                if (new Regex(@"кол.*во", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColKolVo = i;
                                    tab.StartRow = jj;
                                    j = cCelRow;
                                }
                                if (new Regex(@"цена", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColPrice = i;
                                    tab.StartRow = jj;
                                    j = cCelRow;
                                }
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }

                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow;

                    nameProd = "";

                    for (int j = 1; j <= cCelRow; j++)
                    {

                        temp = wTab.Cell(j, ColDiam).Range.Text;
                        temp = temp.Replace("\r\a", string.Empty).Trim();
                        if (regDiam.IsMatch(temp) || regName.IsMatch(temp))
                        {
                            if (regName.IsMatch(temp))
                            {
                                tmp = wTab.Cell(j, ColDiam).Range.Text;
                                tmp = tmp.Replace("\r\a", string.Empty).Trim();
                                if (tmp != "")
                                {
                                    tab.Standart = "";
                                    nameProd = "Труба";
                                    if (regTU.IsMatch(tmp))
                                        foreach (Match m in regTU.Matches(tmp))
                                        {
                                            if (tab.Standart == "" || tab.Standart == null) tab.Standart = m.Value;
                                            else tab.Standart += "; " + m.Value;
                                        }
                                }
                            }

                            if (nameProd != "" && wTab.Rows[j].Cells.Count > 4)
                            {
                                dtProduct.Rows.Add();
                                lastRow = dtProduct.Rows.Count - 1;
                                tab.listExcelIndexTab.Add(j);
                                tab.listdtProductIndexRow.Add(lastRow);
                                dtProduct.Rows[lastRow]["Название"] = nameProd;
                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "труб") dtProduct.Rows[lastRow]["Название"] = "Труба";
                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "угол") dtProduct.Rows[lastRow]["Название"] = "Уголок";
                                if (dtProduct.Rows[lastRow]["Название"].ToString().Length > 1)
                                    dtProduct.Rows[lastRow]["Название"] = dtProduct.Rows[lastRow]["Название"].ToString().Substring(0, 1).ToUpper() + dtProduct.Rows[lastRow]["Название"].ToString().Substring(1, dtProduct.Rows[lastRow]["Название"].ToString().Length - 1);

                                dtProduct.Rows[lastRow]["Примечание"] = temp;
                                dtProduct.Rows[lastRow]["Тип"] = regType.Match(temp).Value;
                                if (dtProduct.Rows[lastRow]["Тип"].ToString() == "") dtProduct.Rows[lastRow]["Тип"] = "тип не указан";

                                dtProduct.Rows[lastRow]["Стандарт"] = tab.Standart;

                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDiam.Match(temp).Value;
                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = "";

                                tmp = wTab.Cell(j, ColTolsh).Range.Text;
                                tmp = tmp.Replace("\r\a", string.Empty).Trim();
                                if (tmp != "")
                                    dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = tmp;

                                tmp = wTab.Cell(j, ColMark).Range.Text;
                                tmp = tmp.Replace("\r\a", string.Empty).Trim();
                                if (tmp != "")
                                    dtProduct.Rows[lastRow]["Марка"] = tmp;

                                tmp = wTab.Cell(j, ColKolVo).Range.Text;
                                tmp = tmp.Replace("\r\a", string.Empty).Trim();
                                if (tmp != "")
                                    dtProduct.Rows[lastRow]["Мерность (т, м, мм)"] = tmp;

                                tmp = wTab.Cell(j, ColPrice).Range.Text;
                                tmp = tmp.Replace("\r\a", string.Empty).Trim();
                                if (tmp != "")
                                    dtProduct.Rows[lastRow]["Цена"] = new Regex(@"\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).Match(tmp).Value;
                            }
                        }
                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        else tsPb1.Value = tsPb1.Maximum;
                    }
                }

                temp = "";
                if (temp != "")
                {
                    Regex regOrgAdr = new Regex(@"(?<=адрес\s*:\s*)(?:\w+,)?(?:\s*\d+\s*,)?(?:\s*[\w+\s*]+,)?\s*[-\w+\s*]+(?:,\s*[\w+\s*]+)*(?:,\s*\d+)*(?:,\s*офис\s*[\d\w]+)?", RegexOptions.IgnoreCase);
                    if (regOrgAdr.IsMatch(temp))
                    {
                        textBoxOrgAdress.Text = regOrgAdr.Match(temp).Value;

                        Regex regOrgTel = new Regex(@"(?<=(?:Телефоны\s*:\s*)*)(?:\d*\s*\(\d+\)\s*)?(\s*\d+(?:-\d+)+,?\s*)+", RegexOptions.IgnoreCase);
                        if (regOrgTel.IsMatch(temp))
                            textBoxOrgTelefon.Text = regOrgTel.Match(temp).Value;
                    }
                    Regex regManager = new Regex(@"менеджер", RegexOptions.IgnoreCase);
                    if (regManager.IsMatch(temp))
                    {
                        ListViewItem lvi = new ListViewItem(new Regex(@"(?<=менеджер\s*:?\s*)(?:\w+\s+){1,3}(?=\s\s|\d)", RegexOptions.IgnoreCase).Match(temp).Value); //имя менеджера
                        lvi.SubItems.Add(new Regex(@"(?<=менеджер\s*:?\s*(?:\w+\s+){1,3}\s*)[\d-]+(?:\s*\d+[(\s]?\d+[)\s]?\s?[\d-]+)?", RegexOptions.IgnoreCase).Match(temp).Value);           //телефон менеджера
                        if (lvi.SubItems[0].Text != "icq") listViewManager.Items.Add(lvi);
                    }

                    Regex regMail = new Regex(@"(?<=(?:эл.*почта\s*:\s*)?)[\w\d\.-]+@[\w\d\.-]+.(?:ru|com|рф|info)", RegexOptions.IgnoreCase);
                    if (regMail.IsMatch(temp))
                        textBoxOrgEmail.Text = regMail.Match(temp).Value;
                }


                //clearingTable();
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка при обработке файла " + Path.GetFileName(path) + "\n\n" + ex.ToString()); }
            if (isOpenWord)
            {
                application.Quit(missingObj, missingObj, missingObj);

            }
        }

        ///<summary> 
        ///<remarks> Открытие и чтение вордовского файла Вика </remarks>
        ///<param name="path" >путь к файлу</param>
        ///</summary>
        private void Vika(string path)
        {
            bool isOpenWord = false;
            textBoxPath.Text = path;
            filePath = path;

            SetNameFromName(filePath);
            SetDateFromName(filePath);

            Word._Application application;
            Word._Document document;
            Object missingObj = System.Reflection.Missing.Value;
            Object trueObj = true;
            Object falseObj = false;
            //создаем обьект приложения word
            application = new Word.Application();
            // создаем путь к файлу
            Object templatePathObj = path;
            Word.Tables tables;
            try
            {
                document = application.Documents.Open(ref templatePathObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                    ref missingObj, ref missingObj);

                tables = document.Tables;
                isOpenWord = true;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                tsLabeltotalSheets.Text = "1";

                Regex regName = new Regex(@"лента|лист|арматура|полоса|угол|швеллер|труб|круг|шестигранник|шгр|квадрат|полоса|сталь|катанка|быстрорез|прокат", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ый\s*с\s*чеч\w*\.?\s*риф\w+\.?|\w+ая|\w+ой(?:\s*проч)|\w+ий", RegexOptions.IgnoreCase);
                Regex regSklad = new Regex(@"(?<=\w*\s*склад.*\s*:\s*)\w+\.\s*\w+,\s*\w+(?:\s*\w+)?,\s*\d+\.\s*.Металл-база.", RegexOptions.IgnoreCase);
                Regex regDiam = new Regex(@"(?<=^(?:\w+)?\s*)\d+(?:[,\.]\d+)?$", RegexOptions.IgnoreCase);
                Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*|ТУ\s*[\d+\.?]+[-\d+]+", RegexOptions.IgnoreCase);
                Regex regMark = new Regex(@"[\d\w]+(?=\s+L)", RegexOptions.IgnoreCase);

                string tmp = "";
                string temp = "";
                int lastRow = 0;
                structTab tab;
                int jj = 0;

                int ColRaz = 0, ColMark = 0, ColKolVo = 0, ColPrice = 0;
                List<int> ColumnsRazmer = new List<int>();
                List<int> ColumnsMark = new List<int>();
                List<int> ColumnsKolvo = new List<int>();
                List<int> ColumnsPrice = new List<int>();

                for (int t = 1; t <= tables.Count; t++)
                {
                    tab = new structTab
                    {
                        listExcelIndexTab = new List<int>(),
                        listdtProductIndexRow = new List<int>()
                    };

                    Word.Table wTab = tables[t];
                    int cCelCol = wTab.Columns.Count;
                    int cCelRow = wTab.Rows.Count;
                    tsLabelcurrSheet.Text = "1";

                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;

                    for (int j = 1; j <= cCelRow; j++)
                    {
                        jj = j;
                        for (int i = 1; i <= cCelCol; i++)
                        {
                            if (i <= wTab.Rows[jj].Cells.Count)
                            {
                                temp = wTab.Cell(jj, i).Range.Text;
                                temp = temp.Replace("\r\a", string.Empty).Trim();
                                if (new Regex(@"Размер", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColRaz = i;
                                    ColumnsRazmer.Add(i);
                                    tab.StartRow = jj;
                                    j = cCelRow;
                                }

                                if (new Regex(@"марка", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColMark = i;
                                    ColumnsMark.Add(i);
                                    tab.StartRow = jj;
                                    //j = cCelRow;
                                }
                                if (new Regex(@"кол.*во", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColKolVo = i;
                                    ColumnsKolvo.Add(i);
                                    tab.StartRow = jj;
                                    //j = cCelRow;
                                }
                                if (new Regex(@"цена", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColPrice = i;
                                    ColumnsPrice.Add(i);
                                    tab.StartRow = jj;
                                    //j = cCelRow;
                                }
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }

                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow;

                    nameProd = "";

                    for (int i = 0; i < ColumnsRazmer.Count; i++)
                        for (int j = tab.StartRow + 1; j <= cCelRow; j++)
                        {

                            temp = wTab.Cell(j, ColumnsRazmer[i]).Range.Text;
                            temp = temp.Replace("\r\a", string.Empty).Trim();
                            if (temp != "")
                            {
                                nameProd = regName.Match(temp).Value;
                                if (nameProd == "")
                                {
                                    nameProd = "Круг";
                                }
                            }
                            if (regDiam.IsMatch(temp) || regName.IsMatch(temp))
                            {

                                dtProduct.Rows.Add();
                                lastRow = dtProduct.Rows.Count - 1;
                                tab.listExcelIndexTab.Add(j);
                                tab.listdtProductIndexRow.Add(lastRow);
                                dtProduct.Rows[lastRow]["Название"] = nameProd;
                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "труб") dtProduct.Rows[lastRow]["Название"] = "Труба";
                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "угол") dtProduct.Rows[lastRow]["Название"] = "Уголок";
                                if (dtProduct.Rows[lastRow]["Название"].ToString().Length > 1)
                                    dtProduct.Rows[lastRow]["Название"] = dtProduct.Rows[lastRow]["Название"].ToString().Substring(0, 1).ToUpper() + dtProduct.Rows[lastRow]["Название"].ToString().Substring(1, dtProduct.Rows[lastRow]["Название"].ToString().Length - 1);

                                dtProduct.Rows[lastRow]["Примечание"] = temp;
                                dtProduct.Rows[lastRow]["Тип"] = regType.Match(temp).Value;
                                if (dtProduct.Rows[lastRow]["Тип"].ToString() == "") dtProduct.Rows[lastRow]["Тип"] = "тип не указан";

                                dtProduct.Rows[lastRow]["Стандарт"] = "";//tab.Standart;

                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDiam.Match(temp).Value;

                                tmp = wTab.Cell(j, ColumnsMark[i]).Range.Text;
                                tmp = tmp.Replace("\r\a", string.Empty).Trim();
                                if (tmp != "")
                                    dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=L\s?=\s?)\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).Match(tmp).Value;

                                tmp = wTab.Cell(j, ColumnsMark[i]).Range.Text;
                                tmp = tmp.Replace("\r\a", string.Empty).Trim();
                                if (tmp != "")
                                    dtProduct.Rows[lastRow]["Марка"] = regMark.Match(tmp).Value;

                                tmp = wTab.Cell(j, ColKolVo).Range.Text;
                                tmp = tmp.Replace("\r\a", string.Empty).Trim();
                                if (tmp != "")
                                    dtProduct.Rows[lastRow]["Мерность (т, м, мм)"] = tmp;

                                tmp = wTab.Cell(j, ColPrice).Range.Text;
                                tmp = tmp.Replace("\r\a", string.Empty).Trim();
                                if (tmp != "")
                                    dtProduct.Rows[lastRow]["Цена"] = new Regex(@"\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).Match(tmp).Value;

                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                }

                for (int i = 1; i <= 10; i++)
                {
                    temp = document.Paragraphs[i].Range.Text;
                    temp = temp.Replace("\r\a", string.Empty).Trim();
                    if (temp != "")
                    {
                        Regex regOrgAdr = new Regex(@"(?<=адрес\s*:\s*)(?:\w+,)?(?:\s*\d+\s*,)?(?:\s*[\w+\s*]+,)?\s*[-\w+\s*]+(?:,\s*[\w+\s*]+)*(?:,\s*\d+)*(?:,\s*офис\s*[\d\w]+)?", RegexOptions.IgnoreCase);
                        if (regOrgAdr.IsMatch(temp))
                        {
                            textBoxOrgAdress.Text = regOrgAdr.Match(temp).Value;

                            Regex regOrgTel = new Regex(@"(?<=(?:Телефоны\s*:\s*)*)(?:\d*\s*\(\d+\)\s*)?(\s*\d+(?:-\d+)+,?\s*)+", RegexOptions.IgnoreCase);
                            if (regOrgTel.IsMatch(temp))
                                textBoxOrgTelefon.Text = regOrgTel.Match(temp).Value;
                        }
                        Regex regManager = new Regex(@"менеджер", RegexOptions.IgnoreCase);
                        if (regManager.IsMatch(temp))
                        {
                            ListViewItem lvi = new ListViewItem(new Regex(@"(?<=менеджер\s*:?\s*)(?:\w+\s+){1,3}(?=\s\s|\d)", RegexOptions.IgnoreCase).Match(temp).Value); //имя менеджера
                            lvi.SubItems.Add(new Regex(@"(?<=менеджер\s*:?\s*(?:\w+\s+){1,3}\s*)[\d-]+(?:\s*\d+[(\s]?\d+[)\s]?\s?[\d-]+)?", RegexOptions.IgnoreCase).Match(temp).Value);           //телефон менеджера
                            if (lvi.SubItems[0].Text != "icq") listViewManager.Items.Add(lvi);
                        }

                        Regex regMail = new Regex(@"(?<=(?:эл.*почта\s*:\s*)?)[\w\d\.-]+@[\w\d\.-]+.(?:ru|com|рф|info)", RegexOptions.IgnoreCase);
                        if (regMail.IsMatch(temp))
                            textBoxOrgEmail.Text = regMail.Match(temp).Value;

                        Regex regTelefon = new Regex(@"(?<=тел.*\s)(?:\s*\d+\(?\d*\)?\d*-?(?:-?\d+)+\s?;?)+(?=\s|\s*$)", RegexOptions.IgnoreCase);
                        if (regTelefon.IsMatch(temp))
                        {
                            textBoxOrgTelefon.Text = regTelefon.Match(temp).Value;
                        }
                    }
                }

                //clearingTable();
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка при обработке файла " + Path.GetFileName(path) + "\n\n" + ex.ToString()); }
            if (isOpenWord)
            {
                application.Quit(missingObj, missingObj, missingObj);

            }
        }

        ///<summary> 
        ///<remarks> Открытие и чтение экселевского файла </remarks>
        ///<param name="path" >путь к файлу</param>
        ///</summary>
        //Открытие и чтение экселевского файла
        private void OpenAndgetExcel(string path)
        {
            if (excelapp != null || excelappworkbook != null)
            {
                System.Threading.Thread.Sleep(100);
            }

            textBoxPath.Text = path;
            filePath = path;

            SetNameFromName(filePath);
            SetDateFromName(filePath);

            excelapp = new Excel.Application();
            //excelapp.Visible = true;

            isExcelOpen = true;
            excelappworkbooks = excelapp.Workbooks;
            try
            {
                excelappworkbook = excelapp.Workbooks.Open(filePath,
        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing);

                excelsheets = excelappworkbook.Worksheets;

                string temp = "";
                this.Focus();
                int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                //int hhh = 0; //посмотреть кол-во используемых строк на странице
                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                    if (!new Regex("в производстве|на порезку|на ул.", RegexOptions.IgnoreCase).IsMatch(excelworksheet.Name.ToString()))
                    {
                        tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                        int currentProgress = 0;
                        //tsLabelCurFile.Text = currentFile.ToString();
                        int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                        int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                        tsPb1.Maximum = cCelCol * cCelRow;
                        //if (hhh == 0) { MessageBox.Show(cCelCol.ToString()+" x "+cCelRow.ToString()); hhh = 1; } //посмотреть кол-во используемых строк на странице
                        countEmpty = 0;
                        listIndexOfNotEmptyName = new List<int>();
                        listIndexOfEmptyName = new List<int>();
                        listShiftIndex = new List<int>();
                        countRowsForShift = 0;

                        int colRowRas = 0;
                        int colRowPrice = 0;
                        int colRowKolvo = 0;
                        int colRowDiam = 0;
                        int colRowTolsh = 0;
                        int colRowMark = 0;
                        int colRowGost = 0;
                        colForName = 0;
                        isGost = false; // маркер для определения откуда взят гост, из имени или из отдельного столбца, true - из имени
                        isMark = false;

                        tsLabelClearingTable.Text = "Поиск наименований";
                        tsPb1.Value = 0;

                        for (int j = 1; j <= cCelRow; j++) //строки
                        {
                            for (int i = 1; i <= cCelCol; i++) //столбцы
                            {
                                #region Поиск наименования, названия

                                Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString();
                                else temp = "";
                                if (temp != "")
                                {
                                    #region ([Нн]аименование)|(?:[Нн]азвание)
                                    Regex reg = new Regex(@"(?:[Нн]аименование)|(?:[Нн]азвание)", RegexOptions.IgnoreCase);
                                    if (reg.IsMatch(temp))
                                    {
                                        string regstr = reg.Match(temp).Value;
                                        if (i > colForName)
                                        {
                                            for (int k = j + 1; k <= cCelRow; k++)
                                            {
                                                cellRange = (Excel.Range)excelworksheet.Cells[k, i];
                                                if (cellRange.Value != null)
                                                    temp = cellRange.Value.ToString();
                                                else temp = "";
                                                //bool cel = (bool)cellRange.MergeCells;
                                                if (temp != "" && temp != regstr)//&& !cel 
                                                {
                                                    colForName = i;
                                                    GetRegexNameFromString(temp, k);
                                                }
                                                else
                                                {
                                                    listIndexOfEmptyName.Add(k);
                                                    if (listIndexOfNotEmptyName.Count > 0)
                                                        countEmpty++;
                                                }

                                            }
                                        }
                                    }
                                    #endregion

                                    #region (?:.*номенклатура.*)
                                    reg = new Regex(@"(?:.*номенклатура.*)", RegexOptions.IgnoreCase);
                                    if (reg.IsMatch(temp))
                                    {
                                        string regstr = reg.Match(temp).Value;
                                        if (i > colForName)
                                        {
                                            for (int k = j + 1; k <= cCelRow; k++)
                                            {
                                                cellRange = (Excel.Range)excelworksheet.Cells[k, i];
                                                if (cellRange.Value != null)
                                                    temp = cellRange.Value.ToString().Trim();
                                                else temp = "";
                                                //bool cel = (bool)cellRange.MergeCells;
                                                if (temp != "" && temp != regstr) //&& !cel 
                                                {
                                                    colForName = i;
                                                    GetRegexNameNomeklFromString(temp, k, true);
                                                }
                                                else
                                                {
                                                    listIndexOfEmptyName.Add(k);
                                                    if (listIndexOfNotEmptyName.Count > 0)
                                                        countEmpty++;
                                                }

                                            }
                                        }
                                    }
                                    #endregion
                                }
                                tsPb1.Value++;
                                #endregion
                            }
                        }

                        tsLabelClearingTable.Text = "выполняется обработка листа";
                        tsPb1.Value = 0;
                        for (int j = 1; j <= cCelRow; j++) //строки
                        {
                            for (int i = 1; i <= cCelCol; i++) //столбцы
                            {
                                Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString();
                                else temp = "";
                                if (temp != "")
                                {
                                    if (new Regex(@"[Дд]иаметр|Размер.+\(Диам.+$", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        #region диаметр
                                        if (listIndexOfNotEmptyName.Count < 1) manualNameProd(excelworksheet, j, cCelRow, i, cCelCol, temp);
                                        int colRow = 0;
                                        int iDiam;
                                        string regstr = new Regex(@"[Дд]иаметр[\w\s,\.]+", RegexOptions.IgnoreCase).Match(temp).Value; // строка с заголовком
                                        if (listIndexOfNotEmptyName.Count > 0) //если найдено хоть одно имя
                                            for (int ii = colRowDiam; ii < listIndexOfNotEmptyName.Count; ii++)
                                            {
                                                iDiam = listIndexOfNotEmptyName[ii];
                                                if (colRow < countRowsForShift || countRowsForShift == 0)
                                                {
                                                    cellRange = (Excel.Range)excelworksheet.Cells[iDiam, i];
                                                    if (cellRange.Value != null)
                                                        temp = cellRange.Value.ToString();
                                                    else temp = "";
                                                    if (temp != "" && temp != regstr) //если строка не пустая и не равна строке заголовку
                                                    {
                                                        if (new Regex(@"\d\s?[xх]\s?\d", RegexOptions.IgnoreCase).IsMatch(temp))
                                                        {
                                                            if (colRowDiam > 0)
                                                            {
                                                                if (colRow < colRowDiam)
                                                                {
                                                                    try
                                                                    {
                                                                        GetRegexDiamFromString(temp, iDiam - listShiftIndex[colRowDiam] + countRowsIndt + countRowsForShift);
                                                                        GetRegexTolshFromString(temp, iDiam - listShiftIndex[colRowDiam] + countRowsIndt + countRowsForShift);
                                                                        GetRegexMarkFromString(temp, iDiam - listShiftIndex[colRowDiam] + countRowsIndt + countRowsForShift);
                                                                        GetRegexTUFromString(temp, iDiam - listShiftIndex[colRowDiam] + countRowsIndt + countRowsForShift);
                                                                        dtProduct.Rows[iDiam - listShiftIndex[colRowDiam] + countRowsIndt + countRowsForShift]["Примечание"] = temp;
                                                                        //dtProduct.Rows[iDiam - listShiftIndex[colRowDiam] + countRowsIndt + countRowsForShift]["Диаметр (высота), мм"] = strTmp;
                                                                    }
                                                                    catch (Exception ex) { MessageBox.Show("Ошибка при разборе диаметра строки диаметра, \nесли помимо диаметра там еще есть всякое\nпервый блок\n" + ex.ToString()); }
                                                                }
                                                                else
                                                                {
                                                                    try
                                                                    {
                                                                        GetRegexDiamFromString(temp, iDiam - listShiftIndex[colRowDiam] + countRowsIndt);
                                                                        GetRegexTolshFromString(temp, iDiam - listShiftIndex[colRowDiam] + countRowsIndt);
                                                                        GetRegexMarkFromString(temp, iDiam - listShiftIndex[colRowDiam] + countRowsIndt);
                                                                        GetRegexTUFromString(temp, iDiam - listShiftIndex[colRowDiam] + countRowsIndt);
                                                                        dtProduct.Rows[iDiam - listShiftIndex[colRowDiam] + countRowsIndt]["Примечание"] = temp;
                                                                        //dtProduct.Rows[iDiam - listShiftIndex[colRowDiam] + countRowsIndt]["Диаметр (высота), мм"] = strTmp;
                                                                    }
                                                                    catch (Exception ex) { MessageBox.Show("Ошибка при разборе диаметра строки диаметра, \nесли помимо диаметра там еще есть всякое\nвторой блок\n" + ex.ToString()); }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                try
                                                                {
                                                                    GetRegexDiamFromString(temp, iDiam - listShiftIndex[colRowDiam] + countRowsIndt);
                                                                    GetRegexTolshFromString(temp, iDiam - listShiftIndex[colRowDiam] + countRowsIndt);
                                                                    GetRegexMarkFromString(temp, iDiam - listShiftIndex[colRowDiam] + countRowsIndt);
                                                                    GetRegexTUFromString(temp, iDiam - listShiftIndex[colRowDiam] + countRowsIndt);
                                                                    dtProduct.Rows[iDiam - listShiftIndex[colRowDiam] + countRowsIndt]["Примечание"] = temp;
                                                                    //dtProduct.Rows[iDiam - listShiftIndex[colRowDiam] + countRowsIndt]["Диаметр (высота), мм"] = strTmp;
                                                                }
                                                                catch (Exception ex) { MessageBox.Show("Ошибка при разборе диаметра строки диаметра, \nесли помимо диаметра там еще есть всякое\nтретий блок\n" + ex.ToString()); }
                                                            }

                                                        }
                                                        else
                                                            try
                                                            {
                                                                if (colRowDiam > 0)
                                                                {
                                                                    if (colRow < colRowDiam)
                                                                    {
                                                                        dtProduct.Rows[iDiam - listShiftIndex[colRowDiam] + countRowsIndt + countRowsForShift]["Диаметр (высота), мм"] = temp;
                                                                    }
                                                                    else dtProduct.Rows[iDiam - listShiftIndex[colRowDiam] + countRowsIndt]["Диаметр (высота), мм"] = temp;
                                                                }
                                                                else dtProduct.Rows[iDiam - listShiftIndex[colRowDiam] + countRowsIndt]["Диаметр (высота), мм"] = temp;
                                                            }
                                                            catch (Exception ex) { MessageBox.Show("Ошибка №1000\n" + ex.ToString()); }

                                                    }
                                                    colRow++;
                                                    colRowDiam++;
                                                }
                                                else break;

                                            }
                                        #endregion
                                    }
                                    else if (new Regex(@"^[Рр][Аа][Зз][Мм][Ее][Рр]").IsMatch(temp))
                                    {
                                        #region Размер
                                        if (listIndexOfNotEmptyName.Count < 1) manualNameProd(excelworksheet, j, cCelRow, i, cCelCol, temp);
                                        int colRow = 0;
                                        int iDiam;
                                        string str = "";
                                        string regstr = new Regex(@"[Рр][Аа][Зз][Мм][Ее][Рр][\w\s,\.]+").Match(temp).Value;
                                        if (listIndexOfNotEmptyName.Count > 0) //если найдено хоть одно имя
                                            for (int ii = colRowDiam; ii < listIndexOfNotEmptyName.Count; ii++)
                                            {
                                                iDiam = listIndexOfNotEmptyName[ii];
                                                if (colRow < countRowsForShift || countRowsForShift == 0)
                                                {
                                                    cellRange = (Excel.Range)excelworksheet.Cells[iDiam, i];
                                                    if (cellRange.Value != null)
                                                        temp = cellRange.Value.ToString();
                                                    else temp = "";
                                                    if (temp != "" && temp != regstr)
                                                    {
                                                        Regex tmp = new Regex(@"\d+(?:,\d+)?(?=[xXхХ])");
                                                        if (tmp.IsMatch(temp))
                                                        {
                                                            str = tmp.Match(temp).Value;
                                                            try
                                                            {
                                                                if (colRowDiam > 0)
                                                                {
                                                                    if (colRow < colRowDiam)
                                                                    {
                                                                        dtProduct.Rows[iDiam - listShiftIndex[colRowDiam] + countRowsIndt + countRowsForShift]["Диаметр (высота), мм"] = str;
                                                                    }
                                                                    else dtProduct.Rows[iDiam - listShiftIndex[colRowDiam] + countRowsIndt]["Диаметр (высота), мм"] = str;
                                                                }
                                                                else dtProduct.Rows[iDiam - listShiftIndex[colRowDiam] + countRowsIndt]["Диаметр (высота), мм"] = str;
                                                            }
                                                            catch (Exception ex) { MessageBox.Show("Ошибка №1001\n" + ex.ToString()); }
                                                        }
                                                        tmp = new Regex(@"(?<=[xXхХ])\d+(?:,\d+)?");
                                                        if (tmp.IsMatch(temp))
                                                        {
                                                            str = tmp.Match(temp).Value;
                                                            if (colRowDiam > 0)
                                                            {
                                                                if (colRow < colRowDiam)
                                                                {
                                                                    dtProduct.Rows[iDiam - listShiftIndex[colRowDiam] + countRowsIndt + countRowsForShift]["Толщина (ширина), мм"] = str;
                                                                }
                                                                else dtProduct.Rows[iDiam - listShiftIndex[colRowDiam] + countRowsIndt]["Толщина (ширина), мм"] = str;
                                                            }
                                                            else dtProduct.Rows[iDiam - listShiftIndex[colRowDiam] + countRowsIndt]["Толщина (ширина), мм"] = str;
                                                        }
                                                    }
                                                    colRow++;
                                                    colRowDiam++;
                                                }
                                                else break;

                                            }
                                        #endregion
                                    }

                                    if (new Regex(@"^толщина|^стенка", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        try
                                        {
                                            #region толщина стенки
                                            if (listIndexOfNotEmptyName.Count < 1) manualNameProd(excelworksheet, j, cCelRow, i, cCelCol, temp);
                                            int colRow = 0;
                                            int iTolsh;
                                            string regstr = new Regex(@"[Тт]олщина[\w\s\.,]+").Match(temp).Value;
                                            if (listIndexOfNotEmptyName.Count > 0) //если найдено хоть одно имя
                                                for (int ii = colRowTolsh; ii < listIndexOfNotEmptyName.Count; ii++)
                                                {
                                                    iTolsh = listIndexOfNotEmptyName[ii];
                                                    if (colRow < countRowsForShift || countRowsForShift == 0)
                                                    {
                                                        cellRange = (Excel.Range)excelworksheet.Cells[iTolsh, i];
                                                        if (cellRange.Value != null)
                                                            temp = cellRange.Value.ToString();
                                                        else temp = "";
                                                        if (temp != "")
                                                        {
                                                            if (colRowTolsh > 0)
                                                            {
                                                                if (colRow < colRowTolsh)
                                                                {
                                                                    dtProduct.Rows[iTolsh - listShiftIndex[colRowTolsh] + countRowsIndt + countRowsForShift]["Толщина (ширина), мм"] = temp;
                                                                }
                                                                else dtProduct.Rows[iTolsh - listShiftIndex[colRowTolsh] + countRowsIndt]["Толщина (ширина), мм"] = temp;

                                                            }
                                                            else dtProduct.Rows[iTolsh - listShiftIndex[colRowTolsh] + countRowsIndt]["Толщина (ширина), мм"] = temp;
                                                        }
                                                        colRow++;
                                                        colRowTolsh++;
                                                    }
                                                    else break;

                                                }
                                            #endregion
                                        }
                                        catch (Exception ex) { MessageBox.Show("Ошибка №1002\n" + ex.ToString()); }
                                    }

                                    if (new Regex(@"Способ\s*производства\s*$", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        try
                                        {
                                            #region способ произодства, Тип
                                            int colRow = 0;
                                            int iType;
                                            string str = "";
                                            string regstr = new Regex(@"Способ\s*производства\s*$", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if (listIndexOfNotEmptyName.Count > 0) //если найдено хоть одно имя
                                                for (int ii = colRowDiam; ii < listIndexOfNotEmptyName.Count; ii++)
                                                {
                                                    iType = listIndexOfNotEmptyName[ii];
                                                    if (colRow < countRowsForShift || countRowsForShift == 0)
                                                    {
                                                        cellRange = (Excel.Range)excelworksheet.Cells[iType, i];
                                                        if (cellRange.Value != null)
                                                            temp = cellRange.Value.ToString();
                                                        else temp = "";
                                                        if (temp != "" && temp != regstr)
                                                        {
                                                            Regex tmp = new Regex(@"\d+(?:,\d+)?(?=[xXхХ])");
                                                            if (tmp.IsMatch(temp))
                                                            {
                                                                str = tmp.Match(temp).Value;
                                                                try
                                                                {
                                                                    if (colRowDiam > 0)
                                                                    {
                                                                        if (colRow < colRowDiam)
                                                                        {
                                                                            dtProduct.Rows[iType - listShiftIndex[colRowDiam] + countRowsIndt + countRowsForShift]["Тип"] = str;
                                                                        }
                                                                        else dtProduct.Rows[iType - listShiftIndex[colRowDiam] + countRowsIndt]["Тип"] = str;
                                                                    }
                                                                    else dtProduct.Rows[iType - listShiftIndex[colRowDiam] + countRowsIndt]["Тип"] = str;
                                                                }
                                                                catch (Exception ex) { MessageBox.Show("Ошибка №1032\n" + ex.ToString()); }
                                                            }
                                                        }
                                                        colRow++;
                                                        colRowDiam++;
                                                    }
                                                    else break;

                                                }
                                            #endregion
                                        }
                                        catch (Exception ex) { MessageBox.Show("Ошибка №1031\n" + ex.ToString()); }
                                    }

                                    if (new Regex(@"[Мм]арка|[Сс][Тт][Аа][Лл][Ьь](?:$|\s+)").IsMatch(temp) && !isMark)
                                    {
                                        try
                                        {
                                            #region Марка
                                            // если список индексов непустых строк имен не содержит ни одного имени
                                            // то выполнить функцию задания имени вручную
                                            if (listIndexOfNotEmptyName.Count < 1) manualNameProd(excelworksheet, j, cCelRow, i, cCelCol, temp);
                                            int colRow = 0;
                                            int iMark;
                                            if (listIndexOfNotEmptyName.Count > 0) //если найдено хоть одно имя
                                                for (int ii = colRowMark; ii < listIndexOfNotEmptyName.Count; ii++)
                                                {
                                                    iMark = listIndexOfNotEmptyName[ii];
                                                    if (colRow < countRowsForShift || countRowsForShift == 0)
                                                    {
                                                        cellRange = (Excel.Range)excelworksheet.Cells[iMark, i];
                                                        if (cellRange.Value != null)
                                                            temp = cellRange.Value.ToString();
                                                        else temp = "";
                                                        if (temp != "")
                                                        {
                                                            if (colRowMark > 0)
                                                            {
                                                                if (colRow < colRowMark)
                                                                {
                                                                    dtProduct.Rows[iMark - listShiftIndex[colRowMark] + countRowsIndt + countRowsForShift]["Марка"] = temp;
                                                                }
                                                                else dtProduct.Rows[iMark - listShiftIndex[colRowMark] + countRowsIndt]["Марка"] = temp;

                                                            }
                                                            else dtProduct.Rows[iMark - listShiftIndex[colRowMark] + countRowsIndt]["Марка"] = temp;
                                                        }
                                                        colRow++;
                                                        colRowMark++;
                                                    }
                                                    else break;

                                                }
                                            #endregion
                                        }
                                        catch (Exception ex) { MessageBox.Show("Ошибка №1003\n" + ex.ToString()); }
                                    }

                                    if (new Regex(@"[Гг][оО][сС][тТ]").IsMatch(temp) && !isGost)
                                    {
                                        try
                                        {
                                            #region Гост
                                            if (listIndexOfNotEmptyName.Count < 1) manualNameProd(excelworksheet, j, cCelRow, i, cCelCol, temp);
                                            int colRow = 0;
                                            int iGost;
                                            if (listIndexOfNotEmptyName.Count > 0) //если найдено хоть одно имя
                                                for (int ii = colRowGost; ii < listIndexOfNotEmptyName.Count; ii++)
                                                {
                                                    iGost = listIndexOfNotEmptyName[ii];
                                                    if (colRow < countRowsForShift || countRowsForShift == 0)
                                                    {
                                                        cellRange = (Excel.Range)excelworksheet.Cells[iGost, i];
                                                        if (cellRange.Value != null)
                                                            temp = cellRange.Value.ToString();
                                                        else temp = "";
                                                        if (temp != "")
                                                        {
                                                            if (colRowGost > 0)
                                                            {
                                                                if (colRow < colRowGost)
                                                                {
                                                                    dtProduct.Rows[iGost - listShiftIndex[colRowGost] + countRowsIndt + countRowsForShift]["Стандарт"] = temp;
                                                                }
                                                                else dtProduct.Rows[iGost - listShiftIndex[colRowGost] + countRowsIndt]["Стандарт"] = temp;

                                                            }
                                                            else dtProduct.Rows[iGost - listShiftIndex[colRowGost] + countRowsIndt]["Стандарт"] = temp;
                                                        }
                                                        colRow++;
                                                        colRowGost++;
                                                    }
                                                    else break;

                                                }
                                            #endregion
                                        }
                                        catch (Exception ex) { MessageBox.Show("Ошибка №1004\n" + ex.ToString()); }
                                    }

                                    if (new Regex(@"наличие\s*на\s*складе\s*$|наличие,\s*(?:тн|тонн)", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        #region наличие на складе
                                        if (listIndexOfNotEmptyName.Count < 1) manualNameProd(excelworksheet, j, cCelRow, i, cCelCol, temp);
                                        if (listIndexOfNotEmptyName.Count > 0) //если найдено хоть одно имя
                                        {
                                            for (int k = j + 1; k <= cCelRow; k++) //идем вниз от "наличие на складе"
                                            {
                                                cellRange = (Excel.Range)excelworksheet.Cells[k, i];
                                                if (cellRange.Value != null)
                                                    temp = cellRange.Value.ToString();
                                                else temp = "";
                                                bool stop = false; /*нужен чтобы остановить поиск вниз, если там уже прочитаны 
                                                                            * данные*/
                                                if (temp != "")
                                                {

                                                    Regex regTemp = new Regex(@"(?:[Оо]бщий)*\s*[Вв]ес");

                                                    if (regTemp.IsMatch(temp))
                                                    {
                                                        try
                                                        {
                                                            #region общий Мерность (т, м, мм)
                                                            colRow = 0;
                                                            foreach (int iVes in listIndexOfNotEmptyName)
                                                            {
                                                                cellRange = (Excel.Range)excelworksheet.Cells[iVes, i];
                                                                if (cellRange.Value != null)
                                                                    temp = cellRange.Value.ToString();
                                                                else temp = "";
                                                                if (temp != "")
                                                                {

                                                                    GetRegexVesFromString(temp, iVes - listShiftIndex[colRow] + countRowsIndt);
                                                                }
                                                                colRow++;
                                                            }
                                                            stop = true;
                                                            #endregion
                                                        }
                                                        catch (Exception ex) { MessageBox.Show("Ошибка №1005\n" + ex.ToString()); }
                                                    }

                                                    regTemp = new Regex(@"(?:[Мм]етраж)|(?:[Дд]лина)");
                                                    cellRange = (Excel.Range)excelworksheet.Cells[k, i + 1];
                                                    if (cellRange.Value != null)
                                                        temp = cellRange.Value.ToString();
                                                    else temp = "";
                                                    if (regTemp.IsMatch(temp))
                                                    {
                                                        try
                                                        {
                                                            #region Метраж, м (длина, мм)
                                                            colRow = 0;
                                                            foreach (int iDlina in listIndexOfNotEmptyName)
                                                            {
                                                                cellRange = (Excel.Range)excelworksheet.Cells[iDlina, i + 1];
                                                                if (cellRange.Value != null)
                                                                    temp = cellRange.Value.ToString();
                                                                else temp = "";
                                                                if (temp != "")
                                                                {

                                                                    GetRegexDlinaFromString(temp, iDlina - listShiftIndex[colRow] + countRowsIndt);
                                                                }
                                                                colRow++;
                                                            }
                                                            i++;
                                                            stop = true;
                                                            #endregion
                                                        }
                                                        catch (Exception ex) { MessageBox.Show("Ошибка №1006\n" + ex.ToString()); }
                                                    }

                                                    cellRange = (Excel.Range)excelworksheet.Cells[k, i + 2];
                                                    if (cellRange.Value != null)
                                                        temp = cellRange.Value.ToString();
                                                    else temp = "";
                                                    if (regTemp.IsMatch(temp))
                                                    {
                                                        try
                                                        {
                                                            #region Метраж, м (длина, мм)
                                                            colRow = 0;
                                                            foreach (int iDlina in listIndexOfNotEmptyName)
                                                            {
                                                                cellRange = (Excel.Range)excelworksheet.Cells[iDlina, i + 2];
                                                                if (cellRange.Value != null)
                                                                    temp = cellRange.Value.ToString();
                                                                else temp = "";
                                                                if (temp != "")
                                                                {

                                                                    GetRegexDlinaFromString(temp, iDlina - listShiftIndex[colRow] + countRowsIndt);
                                                                }
                                                                colRow++;
                                                            }
                                                            i++;
                                                            stop = true;
                                                            #endregion
                                                        }
                                                        catch (Exception ex) { MessageBox.Show("Ошибка №1007\n" + ex.ToString()); }
                                                    }
                                                    if (stop) break;
                                                }
                                            }
                                        }
                                        #endregion
                                    }
                                    else if (temp == "Раскрой" || temp == "раскрой")
                                    {
                                        try
                                        {
                                            #region Раскрой
                                            if (listIndexOfNotEmptyName.Count < 1) manualNameProd(excelworksheet, j, cCelRow, i, cCelCol, temp);
                                            int colRow = 0;
                                            int iRas;
                                            if (listIndexOfNotEmptyName.Count > 0) //если найдено хоть одно имя
                                                for (int ii = colRowRas; ii < listIndexOfNotEmptyName.Count; ii++)
                                                {
                                                    iRas = listIndexOfNotEmptyName[ii];
                                                    if (colRow < countRowsForShift || countRowsForShift == 0)
                                                    {
                                                        cellRange = (Excel.Range)excelworksheet.Cells[iRas, i];
                                                        if (cellRange.Value != null)
                                                            temp = cellRange.Value.ToString();
                                                        else temp = "";
                                                        if (temp != "")
                                                        {
                                                            if (colRowRas > 0)
                                                            {
                                                                if (colRow < colRowRas)
                                                                {
                                                                    dtProduct.Rows[iRas - listShiftIndex[colRowRas] + countRowsIndt + countRowsForShift]["Метраж, м (длина, мм)"] = temp;
                                                                }
                                                                else dtProduct.Rows[iRas - listShiftIndex[colRowRas] + countRowsIndt]["Метраж, м (длина, мм)"] = temp;
                                                            }
                                                            else dtProduct.Rows[iRas - listShiftIndex[colRowRas] + countRowsIndt]["Метраж, м (длина, мм)"] = temp;
                                                        }
                                                        colRow++;
                                                        colRowRas++;
                                                    }
                                                    else break;

                                                }
                                            #endregion
                                        }
                                        catch (Exception ex) { MessageBox.Show("Ошибка №1008\n" + ex.ToString()); }
                                    }
                                    else if (temp.ToLower() == "мерность")
                                    {
                                        try
                                        {
                                            #region мерность
                                            if (listIndexOfNotEmptyName.Count < 1) manualNameProd(excelworksheet, j, cCelRow, i, cCelCol, temp);
                                            int colRow = 0;
                                            int iRas;
                                            if (listIndexOfNotEmptyName.Count > 0) //если найдено хоть одно имя
                                                for (int ii = colRowRas; ii < listIndexOfNotEmptyName.Count; ii++)
                                                {
                                                    iRas = listIndexOfNotEmptyName[ii];
                                                    if (colRow < countRowsForShift || countRowsForShift == 0)
                                                    {
                                                        cellRange = (Excel.Range)excelworksheet.Cells[iRas, i];
                                                        if (cellRange.Value != null)
                                                            temp = cellRange.Value.ToString();
                                                        else temp = "";
                                                        if (temp != "")
                                                        {
                                                            if (colRowRas > 0)
                                                            {
                                                                if (colRow < colRowRas)
                                                                {
                                                                    dtProduct.Rows[iRas - listShiftIndex[colRowRas] + countRowsIndt + countRowsForShift]["Метраж, м (длина, мм)"] = temp;
                                                                }
                                                                else dtProduct.Rows[iRas - listShiftIndex[colRowRas] + countRowsIndt]["Метраж, м (длина, мм)"] = temp;
                                                            }
                                                            else dtProduct.Rows[iRas - listShiftIndex[colRowRas] + countRowsIndt]["Метраж, м (длина, мм)"] = temp;
                                                        }
                                                        colRow++;
                                                        colRowRas++;
                                                    }
                                                    else break;
                                                }
                                            #endregion
                                        }
                                        catch (Exception ex) { MessageBox.Show("Ошибка №1009\n" + ex.ToString()); }
                                    }

                                    else if (new Regex(@"[Кк]ол[\w-]+во|наличие,\s*(?:тн|тонн)|остаток|ост\.\s+\(т\)|^Вес\b", RegexOptions.IgnoreCase).IsMatch(temp))
                                        if (!new Regex(@"штук", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            #region количество
                                            if (listIndexOfNotEmptyName.Count < 1) manualNameProd(excelworksheet, j, cCelRow, i, cCelCol, temp);
                                            int colRow = 0;
                                            int iKol;
                                            if (listIndexOfNotEmptyName.Count > 0) //если найдено хоть одно имя
                                                for (int ii = colRowKolvo; ii < listIndexOfNotEmptyName.Count; ii++)
                                                {
                                                    iKol = listIndexOfNotEmptyName[ii];
                                                    if (colRow < countRowsForShift || countRowsForShift == 0)
                                                    {
                                                        cellRange = (Excel.Range)excelworksheet.Cells[iKol, i];
                                                        if (cellRange.Value != null)
                                                            temp = cellRange.Value.ToString();
                                                        else temp = "";
                                                        if (temp != "")
                                                        {
                                                            try
                                                            {
                                                                if (colRowKolvo > 0)
                                                                {
                                                                    if (colRow < colRowKolvo)
                                                                    {
                                                                        dtProduct.Rows[iKol - listShiftIndex[colRowKolvo] + countRowsIndt + countRowsForShift]["Мерность (т, м, мм)"] = temp;
                                                                    }
                                                                    else dtProduct.Rows[iKol - listShiftIndex[colRowKolvo] + countRowsIndt]["Мерность (т, м, мм)"] = temp;
                                                                }
                                                                else dtProduct.Rows[iKol - listShiftIndex[colRowKolvo] + countRowsIndt]["Мерность (т, м, мм)"] = temp;
                                                            }
                                                            catch (Exception ex) { MessageBox.Show("Ошибка №1010\n" + ex.ToString()); }
                                                        }
                                                        colRow++;
                                                        colRowKolvo++;
                                                    }
                                                    else break;

                                                }
                                            #endregion
                                        }

                                    if (new Regex(@"[Цц]ена").IsMatch(temp))
                                    {
                                        try
                                        {
                                            #region цена
                                            if (listIndexOfNotEmptyName.Count < 1) manualNameProd(excelworksheet, j, cCelRow, i, cCelCol, temp);
                                            int colRow = 0;
                                            int iPrice;
                                            if (listIndexOfNotEmptyName.Count > 0)
                                            {
                                                for (int ii = colRowPrice; ii < listIndexOfNotEmptyName.Count; ii++)
                                                {
                                                    iPrice = listIndexOfNotEmptyName[ii];
                                                    if (colRow < countRowsForShift || countRowsForShift == 0)
                                                    {
                                                        cellRange = (Excel.Range)excelworksheet.Cells[iPrice, i];
                                                        if (cellRange.Value != null)
                                                            temp = cellRange.Value.ToString();
                                                        else temp = "";
                                                        if (temp != "")
                                                        {
                                                            if (temp.IndexOf('\'') > 0)
                                                                temp = temp.Remove(temp.IndexOf('\''), 1);
                                                            if (colRowPrice > 0)
                                                            {
                                                                if (colRow < colRowPrice)
                                                                {
                                                                    GetRegexPriceFromString(temp, iPrice - listShiftIndex[colRowPrice] + countRowsIndt + countRowsForShift);
                                                                }
                                                                else GetRegexPriceFromString(temp, iPrice - listShiftIndex[colRowPrice] + countRowsIndt);
                                                            }
                                                            else GetRegexPriceFromString(temp, iPrice - listShiftIndex[colRowPrice] + countRowsIndt);
                                                        }
                                                        colRow++;
                                                        colRowPrice++;
                                                    }
                                                    else break;
                                                }
                                            }
                                            #endregion
                                        }
                                        catch (Exception ex) { MessageBox.Show("Ошибка №1011\n" + ex.ToString()); }
                                    }

                                    #region склад
                                    if (new Regex(@"[Аа]дрес(?=\s*(?:склад))|^[\.\s]+склад|^[\w\s\d,\.-]*(?=\(?\s*Склад\b\s*)", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        if (new Regex(@"адрес склада:?\s*(?:№\s*\d+)?$", RegexOptions.IgnoreCase).IsMatch(temp)) //для стальной мир
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[j, i + 1];
                                            if (cellRange.Value != null)
                                            {
                                                temp = cellRange.Value.ToString();
                                                ListViewItem lvi = new ListViewItem(temp);
                                                bool isIn = false;
                                                if (listViewAdrSklad.Items.Count < 1) listViewAdrSklad.Items.Add(lvi);
                                                for (int ii = 0; ii < listViewAdrSklad.Items.Count; ii++)
                                                {
                                                    if (listViewAdrSklad.Items[ii].SubItems[0].Text == temp) isIn = true;
                                                }
                                                if (!isIn) listViewAdrSklad.Items.Add(lvi);
                                            }
                                        }
                                        else if (new Regex(@"^Склад\s+-\s+", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            while (temp.IndexOf('"') > -1)
                                                temp = temp.Remove(temp.IndexOf('"'), 1);
                                            string tmp = new Regex(@"(?<=^склад\s+-\s+).+,\s\w+\.\w+[\s\w\d]+(?=,|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                            ListViewItem lvi = new ListViewItem(tmp);
                                            bool isIn = false;
                                            if (listViewAdrSklad.Items.Count < 1) listViewAdrSklad.Items.Add(lvi);
                                            for (int ii = 0; ii < listViewAdrSklad.Items.Count; ii++)
                                            {
                                                if (listViewAdrSklad.Items[ii].SubItems[0].Text == tmp) isIn = true;
                                            }
                                            if (!isIn) listViewAdrSklad.Items.Add(lvi);
                                        }
                                        else if (new Regex(@"^[\w\s\d,\.-]*(?=\(?\s*Склад\b\s*)", RegexOptions.IgnoreCase).IsMatch(temp)) //для стальной мир
                                        {
                                            string tmp = new Regex(@"^[\w\s\d,\.-]*(?=\(?\s*Склад\s*)", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if (tmp != "")
                                            {
                                                ListViewItem lvi = new ListViewItem(tmp);
                                                bool isIn = false;
                                                if (listViewAdrSklad.Items.Count < 1) listViewAdrSklad.Items.Add(lvi);
                                                for (int ii = 0; ii < listViewAdrSklad.Items.Count; ii++)
                                                {
                                                    if (listViewAdrSklad.Items[ii].SubItems[0].Text == tmp) isIn = true;
                                                }
                                                if (!isIn) listViewAdrSklad.Items.Add(lvi);
                                            }
                                        }

                                        else if (new Regex(@"^Склад\s+-\s+", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            string tmp = new Regex(@"(?<=^Склад\s+-\s+).+,\s\w+\.\w+[\s\w\d]+(?=,|$)").Match(temp).Value;
                                            ListViewItem lvi = new ListViewItem(tmp);
                                            bool isIn = false;
                                            if (listViewAdrSklad.Items.Count < 1) listViewAdrSklad.Items.Add(lvi);
                                            for (int ii = 0; ii < listViewAdrSklad.Items.Count; ii++)
                                            {
                                                if (listViewAdrSklad.Items[ii].SubItems[0].Text == tmp) isIn = true;
                                            }
                                            if (!isIn) listViewAdrSklad.Items.Add(lvi);
                                        }
                                        else
                                        {
                                            string tmp = new Regex(@"(?<=[Сс]клада*?\s)[№\w\s\d,.]+").Match(temp).Value;
                                            ListViewItem lvi = new ListViewItem(tmp);
                                            bool isIn = false;
                                            if (listViewAdrSklad.Items.Count < 1) listViewAdrSklad.Items.Add(lvi);
                                            for (int ii = 0; ii < listViewAdrSklad.Items.Count; ii++)
                                            {
                                                if (listViewAdrSklad.Items[ii].SubItems[0].Text == tmp) isIn = true;
                                            }
                                            if (!isIn) listViewAdrSklad.Items.Add(lvi);
                                        }

                                    }
                                    else if (new Regex(@"(?<=склад.*)г.*\sсклад\)", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        string tmp = new Regex(@"(?<=склад.*)г.*\sсклад\)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        ListViewItem lvi = new ListViewItem(tmp);
                                        bool isIn = false;
                                        if (listViewAdrSklad.Items.Count < 1) listViewAdrSklad.Items.Add(lvi);
                                        for (int ii = 0; ii < listViewAdrSklad.Items.Count; ii++)
                                        {
                                            if (listViewAdrSklad.Items[ii].SubItems[0].Text == tmp) isIn = true;
                                        }
                                        if (!isIn) listViewAdrSklad.Items.Add(lvi);
                                    }
                                    #endregion

                                    #region адрес организации
                                    if (!new Regex(@"склад", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        if (new Regex(@"(?:[Аа]дрес(?!\s*(?:склад)))|(?:\d+[\s,]\s?[Гг][\.\s]+\w+[,\s]+(?:(?:ул\.\s?)|(?:улица[\s:]+))\w+[,\s]+[\w\d.\\/]+\b(?:,?\s?(?:оф|офис)\.?\s?\d+(?:\w+)?)?)").IsMatch(temp))
                                        {
                                            if (textBoxOrgAdress.Text == "")
                                                textBoxOrgAdress.Text = new Regex(@"(?:(?<=[Аа]дрес\sофиса\s)[\w+\d+]+[\s*\w\d,.]*)|(?:\d+[\s,]\s?[Гг][\.\s]+\w+[,\s]+(?:(?:ул\.\s?)|(?:улица[\s:]+))\w+[,\s]+[\w\d.\s\\/]+(?:,?\s?(?:оф|офис)\.?\s?\d+(?:\w+)?)?)").Match(temp).Value;
                                            else
                                                textBoxOrgAdress.Text += ";" + new Regex(@"(?:(?<=[Аа]дрес\sофиса\s)[\w+\d+]+[\s*\w\d,.]*)|(?:\d+[\s,]\s?[Гг][\.\s]+\w+[,\s]+(?:(?:ул\.\s?)|(?:улица[\s:]+))\w+[,\s]+[\w\d.\s\\/]+(?:,?\s?(?:оф|офис)\.?\s?\d+(?:\w+)?)?)").Match(temp).Value;
                                        }
                                        else if (new Regex(@"адрес\s*офис\w+\s*$", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[j, i + 1];
                                            if (cellRange.Value != null)
                                            {
                                                temp = cellRange.Value.ToString();
                                                if (textBoxOrgAdress.Text == "")
                                                    textBoxOrgAdress.Text = temp;
                                                else
                                                    textBoxOrgAdress.Text += ";" + temp;
                                            }
                                        }
                                        else if (new Regex(@"г\.[\s+\w+,\.]+(?=,\s?тел)", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            if (textBoxOrgAdress.Text == "")
                                                textBoxOrgAdress.Text = new Regex(@"г\.[\s+\w+,\.]+(?=,\s?тел)", RegexOptions.IgnoreCase).Match(temp).Value;
                                            else
                                                textBoxOrgAdress.Text += new Regex(@"г\.[\s+\w+,\.]+(?=,\s?тел)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        }
                                    }
                                    else if (new Regex(@"(?<=офис.*)г.*\sоф.\d{1,3}", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        if (textBoxOrgAdress.Text == "")
                                            textBoxOrgAdress.Text = new Regex(@"(?<=офис.*)г.*\sоф.\d{1,3}", RegexOptions.IgnoreCase).Match(temp).Value;
                                        else
                                            textBoxOrgAdress.Text += new Regex(@"(?<=офис.*)г.*\sоф.\d{1,3}", RegexOptions.IgnoreCase).Match(temp).Value;
                                    }
                                    #endregion

                                    InfoOrganization(temp);
                                }
                                //progr.incrementProgress((i * j / totalProgress) * 100, 0);
                                //progr.Invoke(new Action<int>((p) => progr.incrementProgress(p, 0)), (i * j / totalProgress) * 100);
                                currentProgress++;
                                tsPb1.Value = currentProgress;
                            }
                        }
                        isTelefon = true;

                        countRowsIndt = dtProduct.Rows.Count;
                    }
                clearingTable();
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в импорте" + ex.ToString()); }
        }

        ///<summary> 
        ///<remarks> Открытие и чтение экселевского файла ТрубМет </remarks>
        ///<param name="path" >путь к файлу</param>
        ///</summary>
        private void TrubMet(string path)
        {
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                SetNameFromName(filePath);

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;
                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла TrubMet\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                    tsPb1.Maximum = cCelCol * cCelRow;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    tsLabelClearingTable.Text = "Поиск наименований";
                    tsPb1.Value = 0;
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString();
                            else temp = "";
                            if (temp != "")
                            {
                                Regex regDiam = new Regex(@"Размер.+\(Диам.+$", RegexOptions.IgnoreCase);
                                Regex regType = new Regex(@"способ\s*пр", RegexOptions.IgnoreCase);
                                Regex regMark = new Regex(@"^Сталь$", RegexOptions.IgnoreCase);
                                Regex regVes = new Regex(@"^В\sналичи.*тн$", RegexOptions.IgnoreCase);
                                Regex regPrice = new Regex(@"^Цена.\s*руб.*н$", RegexOptions.IgnoreCase);
                                Regex regSklad = new Regex(@"на\s*склад\s*по\s*адресу", RegexOptions.IgnoreCase);
                                Regex regManager = new Regex(@"Менеджер\s*:\s*", RegexOptions.IgnoreCase);

                                if (regDiam.IsMatch(temp))
                                {
                                    #region получение списка непустых строк
                                    nameProd = "Труба";
                                    for (int jj = j + 1; jj <= cCelRow; jj++) //строки
                                    {

                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            string tmp = new Regex(@"(?<=^\s*|ду\s*)\d+\s?[xх]\s?\d+", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if (tmp != "")
                                            {
                                                dtProduct.Rows.Add();   // добавить строку в результирующую таблицу
                                                int lastRow = dtProduct.Rows.Count - 1; //запомнить индекс последней строки
                                                dtProduct.Rows[lastRow]["Название"] = nameProd; /*записать наименование вручную указанного названия 
                                                                         * в ячейку названия продукции в результирующей таблице*/
                                                listIndexOfNotEmptyName.Add(jj);     //добавить в список индексов непустых значений индекс текущей строки
                                                int c = listIndexOfNotEmptyName[0] + countEmpty;    //запомнить в переменную индекс первого значения плюс количество пустых ячеек
                                                //если список непустых значений не пустой
                                                if (listIndexOfNotEmptyName.Count > 0)
                                                {
                                                    //то если список сдвигов индексов пустой
                                                    if (listShiftIndex.Count < 1)
                                                        // то занести в список сдвигов индексов индекс певого значения плюс количество пустых ячеек
                                                        listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                                                    /*если список сдвигов содержит больше 2х записей и индекс текущей строки меньше чем 
                                                      индекс строки в списке непустых значений в предпоследней записи */
                                                    else if (listIndexOfNotEmptyName.Count > 2 && jj < listIndexOfNotEmptyName[listIndexOfNotEmptyName.Count - 2])
                                                    {
                                                        countEmpty = 0; //сброс счета количества пустых ячеек
                                                        //занести в список сдвигов индексов индекс певого значения плюс количество пустых ячеек
                                                        listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                                                        //количество строк для сдвига равно текущему количеству строк в результирующей таблице
                                                        countRowsForShift = dtProduct.Rows.Count - 1;
                                                    }
                                                    else listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                                                }
                                            }
                                            else if (listIndexOfNotEmptyName.Count > 0) countEmpty++;
                                        }
                                        else if (listIndexOfNotEmptyName.Count > 0) countEmpty++;
                                    }
                                    ManualStringNameProd = nameProd;

                                    #endregion

                                    #region диаметр, ширина стенки, марка, стандарт

                                    string regstr = regDiam.Match(temp).Value; // строка с заголовком
                                    int iDiam = 0;

                                    for (int jj = 0; jj < listIndexOfNotEmptyName.Count; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[jj], i];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString();
                                        else temp = "";
                                        if (temp != "" && temp != regstr) //если строка не пустая и не равна строке заголовку
                                        {
                                            try
                                            {
                                                GetRegexDiamFromString(temp, iDiam); //поиск подстроки для диаметров
                                                GetRegexTolshFromStrWithSpases(temp, iDiam); //поиск подстроки с возможными пробелами перед толщиной
                                                GetRegexTUFromString(temp, iDiam);  //поиск подстроки для стандарта
                                                dtProduct.Rows[iDiam]["Примечание"] = temp;
                                                //dtProduct.Rows[iDiam - listShiftIndex[colRowDiam] + countRowsIndt + countRowsForShift]["Диаметр (высота), мм"] = strTmp;
                                            }
                                            catch (Exception ex) { MessageBox.Show("Ошибка при разборе диаметра строки диаметра, \nесли помимо диаметра там еще есть всякое\nпервый блок\n" + ex.ToString()); }
                                        }
                                        iDiam++;
                                    }
                                    #endregion
                                }

                                #region выделение подстроки типа
                                if (regType.IsMatch(temp)) // если есть совпадение по заголовку типа
                                {
                                    for (int jj = 0; jj < listIndexOfNotEmptyName.Count; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[jj], i];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            dtProduct.Rows[jj]["Тип"] = temp;
                                        }
                                    }
                                }
                                #endregion

                                #region выделение подстроки Стандарта

                                if (regMark.IsMatch(temp))
                                {
                                    for (int jj = 0; jj < listIndexOfNotEmptyName.Count; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[jj], i];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            dtProduct.Rows[jj]["Марка"] = temp;
                                        }
                                    }
                                }

                                #endregion

                                #region Вес, в наличии, тн

                                if (regVes.IsMatch(temp))
                                {
                                    for (int jj = 0; jj < listIndexOfNotEmptyName.Count; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[jj], i];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            dtProduct.Rows[jj]["Мерность (т, м, мм)"] = temp;
                                        }
                                    }
                                }

                                #endregion

                                #region Цена

                                if (regPrice.IsMatch(temp))
                                {
                                    for (int jj = 0; jj < listIndexOfNotEmptyName.Count; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[jj], i];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            dtProduct.Rows[jj]["Цена"] = temp;
                                        }
                                    }
                                }

                                #endregion

                                #region Склад

                                if (regSklad.IsMatch(temp))
                                {
                                    Regex sklad = new Regex(@"(?<=на\s*склад\s*по\s*адресу\.*\s*)г.*$", RegexOptions.IgnoreCase);
                                    string tmp = sklad.Match(temp).Value;
                                    ListViewItem lvi = new ListViewItem(tmp);
                                    bool isIn = false;
                                    if (listViewAdrSklad.Items.Count < 1) listViewAdrSklad.Items.Add(lvi);
                                    for (int ii = 0; ii < listViewAdrSklad.Items.Count; ii++)
                                    {
                                        if (listViewAdrSklad.Items[ii].SubItems[0].Text == tmp) isIn = true;
                                    }
                                    if (!isIn) listViewAdrSklad.Items.Add(lvi);
                                }

                                #endregion

                                #region Менеджер

                                if (regManager.IsMatch(temp))
                                {
                                    Regex manager = new Regex(@"(?<=Менеджер\s*:\s*).*$", RegexOptions.IgnoreCase);
                                    string name = manager.Match(temp).Value;
                                    string tmp = "";
                                    string str = "";
                                    ListViewItem lvi = new ListViewItem(name);   //имя менеджера
                                    cellRange = (Excel.Range)excelworksheet.Cells[j + 1, i];
                                    if (cellRange.Value != null)
                                        tmp = cellRange.Value.ToString();
                                    else tmp = "";
                                    if (tmp != "")
                                    {
                                        if (new Regex(@"(?<=тел.\s*).*(?=,)|(?<=тел.\s.*,\s*)\d[\d-]*", RegexOptions.IgnoreCase).IsMatch(tmp))
                                        {
                                            foreach (Match m in new Regex(@"(?<=тел.\s*).*(?=,)|(?<=тел.\s.*,\s*)\d[\d-]*", RegexOptions.IgnoreCase).Matches(tmp))
                                            {
                                                if (str == "") str = m.Value;
                                                else str += "; " + m.Value;
                                            }
                                        }
                                        lvi.SubItems.Add(str);
                                    }
                                    listViewManager.Items.Add(lvi);

                                    cellRange = (Excel.Range)excelworksheet.Cells[j + 2, i];
                                    lvi = new ListViewItem(name);
                                    if (cellRange.Value != null)
                                        tmp = cellRange.Value.ToString();
                                    else tmp = "";
                                    if (tmp != "")
                                    {
                                        lvi.SubItems.Add(new Regex(@"icq\s*\d{9}", RegexOptions.IgnoreCase).Match(tmp).Value);
                                    }
                                    listViewManager.Items.Add(lvi);
                                }

                                #endregion

                                InfoOrganization(temp);

                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        }
                    }
                }
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции TrubMet\n\n" + ex.ToString()); }
        }

        ///<summary> 
        ///<remarks> Открытие и чтение старого экселевского файла Стальной профиль </remarks>
        ///<param name="path" >путь к файлу</param>
        ///</summary>
        private void StalProfOld(string path)
        {
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                SetNameFromName(filePath);

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;
                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла TrubMet\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;


                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                    tsPb1.Maximum = cCelCol * cCelRow;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    tsLabelClearingTable.Text = "Поиск наименований";
                    tsPb1.Value = 0;

                    Regex regName = new Regex(@"^\w+(?=\s)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"(?<=^\w+\s)\w+", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"^\d+(?:,\d+)?(?=[xх*])", RegexOptions.IgnoreCase);
                    Regex regShirTolsh = new Regex(@"(?<=\d+(?:,\d+)?[xх*])\d+(?:,\d+)?(?=\s|$|\\|/|\()", RegexOptions.IgnoreCase);
                    Regex regDlina = new Regex(@"(?<=^\d+(?:,\d+)?[xх*])\d+(?:,\d+)?(?=[xх*])", RegexOptions.IgnoreCase);
                    int lastRow = 0;
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"типоразмер", RegexOptions.IgnoreCase).IsMatch(temp))
                                    for (int jj = j + 1; jj <= cCelRow; jj++)
                                    {
                                        string tmp = "";
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                                        if (cellRange.Value != null)
                                            tmp = cellRange.Value.ToString();
                                        else tmp = "";
                                        if (tmp != "")
                                        {
                                            dtProduct.Rows.Add();   // добавить строку в результирующую таблицу
                                            lastRow = dtProduct.Rows.Count - 1; //запомнить индекс последней строки
                                            dtProduct.Rows[lastRow]["Примечание"] = tmp;
                                            if (regName.IsMatch(tmp))
                                            {
                                                dtProduct.Rows[lastRow]["Название"] = regName.Match(tmp).Value;
                                            }
                                            if (regType.IsMatch(tmp))
                                            {
                                                dtProduct.Rows[lastRow]["Тип"] = regType.Match(tmp).Value;
                                            }
                                            else dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                            if (regDiam.IsMatch(tmp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDiam.Match(tmp).Value;
                                            }
                                            if (regShirTolsh.IsMatch(tmp))
                                            {
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regShirTolsh.Match(tmp).Value;
                                            }
                                            if (regDlina.IsMatch(tmp))
                                            {
                                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = regDlina.Match(tmp).Value;
                                            }

                                            GetRegexTUFromString(tmp, lastRow);

                                            listIndexOfNotEmptyName.Add(jj);     //добавить в список индексов непустых значений индекс текущей строки
                                            int c = listIndexOfNotEmptyName[0] + countEmpty;    //запомнить в переменную индекс первого значения плюс количество пустых ячеек
                                            //если список непустых значений не пустой
                                            if (listIndexOfNotEmptyName.Count > 0)
                                            {
                                                //то если список сдвигов индексов пустой
                                                if (listShiftIndex.Count < 1)
                                                    // то занести в список сдвигов индексов индекс певого значения плюс количество пустых ячеек
                                                    listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                                                /*если список сдвигов содержит больше 2х записей и индекс текущей строки меньше чем 
                                                  индекс строки в списке непустых значений в предпоследней записи */
                                                else if (listIndexOfNotEmptyName.Count > 2 && jj < listIndexOfNotEmptyName[listIndexOfNotEmptyName.Count - 2])
                                                {
                                                    countEmpty = 0; //сброс счета количества пустых ячеек
                                                    //занести в список сдвигов индексов индекс певого значения плюс количество пустых ячеек
                                                    listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                                                    //количество строк для сдвига равно текущему количеству строк в результирующей таблице
                                                    countRowsForShift = dtProduct.Rows.Count - 1;
                                                }
                                                else listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                                            }
                                        }
                                    }
                                if (new Regex(@"Наличие", RegexOptions.IgnoreCase).IsMatch(temp) && dtProduct.Rows.Count > 0)
                                {
                                    for (int jj = 0; jj < listIndexOfNotEmptyName.Count; jj++)
                                    {
                                        string tmp = "";
                                        cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[jj], i];
                                        if (cellRange.Value != null)
                                            tmp = cellRange.Value.ToString();
                                        else tmp = "";
                                        if (tmp != "")
                                        {
                                            dtProduct.Rows[listIndexOfNotEmptyName[jj] - listShiftIndex[jj] + countRowsForShift]["Мерность (т, м, мм)"] = tmp;
                                        }
                                    }
                                }

                                if (new Regex(@"Цена", RegexOptions.IgnoreCase).IsMatch(temp) && dtProduct.Rows.Count > 0)
                                {
                                    for (int jj = 0; jj < listIndexOfNotEmptyName.Count; jj++)
                                    {
                                        string tmp = "";
                                        cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[jj], i];
                                        if (cellRange.Value != null)
                                            tmp = cellRange.Value.ToString();
                                        else tmp = "";
                                        if (tmp != "")
                                        {
                                            dtProduct.Rows[listIndexOfNotEmptyName[jj] - listShiftIndex[jj] + countRowsForShift]["Цена"] = tmp;
                                        }
                                    }
                                }

                                #region Адрес
                                Regex regAddr = new Regex(@"^\s*\d{6}.*,\w+\s+\d+", RegexOptions.IgnoreCase);
                                if (regAddr.IsMatch(temp))
                                {
                                    Regex sklad = new Regex(@"^\s*\d{6}.*,\w+\s+\d+", RegexOptions.IgnoreCase);
                                    string tmp = sklad.Match(temp).Value;
                                    textBoxOrgAdress.Text = tmp;
                                }

                                #endregion

                                #region Склад
                                Regex regSklad = new Regex(@"(?<=Склад\s*-\s).*\w(?=\s*;)", RegexOptions.IgnoreCase);
                                if (regSklad.IsMatch(temp))
                                {
                                    Regex sklad = new Regex(@"(?<=Склад\s*-\s).*\w(?=\s*;)", RegexOptions.IgnoreCase);
                                    string tmp = sklad.Match(temp).Value;
                                    ListViewItem lvi = new ListViewItem(tmp);
                                    bool isIn = false;
                                    if (listViewAdrSklad.Items.Count < 1) listViewAdrSklad.Items.Add(lvi);
                                    for (int ii = 0; ii < listViewAdrSklad.Items.Count; ii++)
                                    {
                                        if (listViewAdrSklad.Items[ii].SubItems[0].Text == tmp) isIn = true;
                                    }
                                    if (!isIn) listViewAdrSklad.Items.Add(lvi);
                                }

                                #endregion

                                InfoOrganization(temp);
                            }

                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        }
                    }

                }
                clearingTable();
                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции StalProf\n\n" + ex.ToString()); }
        }

        ///<summary> 
        ///<remarks> Открытие и чтение экселевского файла Инрост </remarks>
        ///<param name="path" >путь к файлу</param>
        ///</summary>
        private void Inrost(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                SetNameFromName(filePath);

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;



                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла Inmet\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;



                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    List<structTab> tabs = new List<structTab>();
                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    tsPb1.Maximum = cCelCol * cCelRow;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    tsLabelClearingTable.Text = "Поиск наименований";
                    tsPb1.Value = 0;

                    int addedRow = 0;

                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"диаметр", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    for (int jj = j + 1; jj < cCelRow; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                                        if (cellRange.Value != null)
                                        {
                                            temp = cellRange.Value.ToString().Trim();
                                            if (cellRange.MergeArea.Count < 2)
                                            {
                                                dtProduct.Rows.Add();
                                                listIndexOfNotEmptyName.Add(jj);
                                                dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"] = "Труба";
                                                dtProduct.Rows[dtProduct.Rows.Count - 1]["Тип"] = "тип не указан";
                                                dtProduct.Rows[dtProduct.Rows.Count - 1]["Диаметр (высота), мм"] = temp;
                                            }
                                            else
                                            {
                                                for (int k = 0; k < cellRange.MergeArea.Count; k++)
                                                {
                                                    dtProduct.Rows.Add();
                                                    listIndexOfNotEmptyName.Add(jj);
                                                    dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"] = "Труба";
                                                    dtProduct.Rows[dtProduct.Rows.Count - 1]["Тип"] = "тип не указан";
                                                    dtProduct.Rows[dtProduct.Rows.Count - 1]["Диаметр (высота), мм"] = temp;
                                                    if (k < cellRange.MergeArea.Count - 1) jj++;
                                                }
                                            }
                                        }
                                    }
                                }

                                if (new Regex(@"марка", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    for (int jj = 0; jj < listIndexOfNotEmptyName.Count; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[jj], i];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0]]["Марка"] = temp;
                                        }
                                    }
                                }

                                if (new Regex(@"характеристик", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    for (int jj = 0; jj < listIndexOfNotEmptyName.Count; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[jj], i];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0]]["Примечание"] = temp;
                                            GetRegexTUFromString(temp, listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0]);
                                        }
                                    }
                                }

                                if (new Regex(@"кол.*во", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    for (int jj = 0; jj < listIndexOfNotEmptyName.Count; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[jj], i];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0]]["Мерность (т, м, мм)"] = temp;
                                        }
                                    }
                                }

                                if (new Regex(@"цена", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    for (int jj = 0; jj < listIndexOfNotEmptyName.Count; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[jj], i];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0]]["Цена"] = temp;
                                        }
                                    }
                                }

                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }
                    tsPb1.Value = 0;
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            if (i < 15)
                            {
                                InfoOrganization(temp);

                                Regex regAdr = new Regex(@"(?<=офис.*:\s+)\w.*", RegexOptions.IgnoreCase);
                                Regex regSklad = new Regex(@"(?<=склад.*\:\s?)\w.*", RegexOptions.IgnoreCase);

                                if (regAdr.IsMatch(temp))
                                {
                                    textBoxOrgAdress.Text = regAdr.Match(temp).Value;
                                }

                                if (regSklad.IsMatch(temp))
                                {
                                    string tmp = regSklad.Match(temp).Value;
                                    ListViewItem lvi = new ListViewItem(tmp);
                                    bool isIn = false;
                                    if (listViewAdrSklad.Items.Count < 1) listViewAdrSklad.Items.Add(lvi);
                                    for (int ii = 0; ii < listViewAdrSklad.Items.Count; ii++)
                                    {
                                        if (listViewAdrSklad.Items[ii].SubItems[0].Text == tmp) isIn = true;
                                    }
                                    if (!isIn) listViewAdrSklad.Items.Add(lvi);
                                }

                            }
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"стенка", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    for (int jj = 0; jj < listIndexOfNotEmptyName.Count; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[jj], i];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            Regex regSplit = new Regex(@"(?<=^|,|\s)\d+", RegexOptions.IgnoreCase);
                                            if (regSplit.IsMatch(temp))
                                            {
                                                int matchcount = regSplit.Matches(temp).Count;
                                                if (matchcount == 2)
                                                {
                                                    if (Convert.ToInt32(regSplit.Matches(temp)[0].Value.ToString()) > Convert.ToInt32(regSplit.Matches(temp)[1].Value.ToString()))
                                                    {
                                                        dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Толщина (ширина), мм"] = temp;
                                                    }
                                                    else
                                                    {
                                                        dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Толщина (ширина), мм"] = regSplit.Matches(temp)[0].Value.ToString();
                                                        DataRow row = dtProduct.NewRow();
                                                        row["Название"] = dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Название"];
                                                        row["Тип"] = dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Тип"];
                                                        row["Диаметр (высота), мм"] = dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Диаметр (высота), мм"];
                                                        row["Толщина (ширина), мм"] = regSplit.Matches(temp)[1].Value.ToString();
                                                        row["Марка"] = dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Марка"];
                                                        row["Стандарт"] = dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Стандарт"];
                                                        row["Примечание"] = dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Примечание"];
                                                        row["Мерность (т, м, мм)"] = dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Мерность (т, м, мм)"];
                                                        row["Цена"] = dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Цена"];
                                                        dtProduct.Rows.InsertAt(row, listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow + 1);
                                                        addedRow++;
                                                    }
                                                }
                                                else if (matchcount == 1)
                                                {
                                                    dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Толщина (ширина), мм"] = temp;
                                                }
                                                else if (matchcount > 2)
                                                {
                                                    dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Толщина (ширина), мм"] = regSplit.Matches(temp)[0].Value.ToString();
                                                    for (int m = 1; m < regSplit.Matches(temp).Count; m++)
                                                    {
                                                        DataRow row = dtProduct.NewRow();
                                                        row["Название"] = dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Название"];
                                                        row["Тип"] = dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Тип"];
                                                        row["Диаметр (высота), мм"] = dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Диаметр (высота), мм"];
                                                        row["Толщина (ширина), мм"] = regSplit.Matches(temp)[m].Value.ToString();
                                                        row["Марка"] = dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Марка"];
                                                        row["Стандарт"] = dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Стандарт"];
                                                        row["Примечание"] = dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Примечание"];
                                                        row["Мерность (т, м, мм)"] = dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Мерность (т, м, мм)"];
                                                        row["Цена"] = dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow]["Цена"];
                                                        dtProduct.Rows.InsertAt(row, listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0] + addedRow + 1);
                                                        addedRow++;
                                                    }
                                                }

                                            }
                                        }
                                    }
                                }
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }
                }
                clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции Inrost\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        ///<summary> 
        ///<remarks> Открытие и чтение экселевского файла ТрубДом </remarks>
        ///<param name="path" >путь к файлу</param>
        ///</summary>
        private void TrubDom(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                SetNameFromName(filePath);

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;



                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла TrubDom\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;



                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    List<structTab> tabs = new List<structTab>();
                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    tsPb1.Maximum = cCelCol * cCelRow;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    tsLabelClearingTable.Text = "Поиск наименований";
                    tsPb1.Value = 0;


                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"диаметр", RegexOptions.IgnoreCase).IsMatch(temp)) //находим диаметр и сдвигаемся от него влево, т.к. там нет заголовка столбца
                                {

                                    for (int jj = j + 1; jj <= cCelRow; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, i - 1];   // вот здесь сдвигаемся
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            dtProduct.Rows.Add();
                                            listIndexOfNotEmptyName.Add(jj);
                                            dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"] = new Regex(@"^\w+(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if (new Regex(@"(?<=^\w+\s+)\w+", RegexOptions.IgnoreCase).IsMatch(temp))
                                                dtProduct.Rows[dtProduct.Rows.Count - 1]["Тип"] = new Regex(@"(?<=^\w+\s+)\w+", RegexOptions.IgnoreCase).Match(temp).Value;
                                        }
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, i];  //а тут возвращаемся
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            dtProduct.Rows[dtProduct.Rows.Count - 1]["Диаметр (высота), мм"] = temp;
                                        }
                                    }
                                    continue;
                                }
                                if (new Regex(@"стенка", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    for (int jj = 0; jj < listIndexOfNotEmptyName.Count; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[jj], i];   // вот здесь сдвигаемся
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0]]["Толщина (ширина), мм"] = temp;
                                        }
                                    }
                                }

                                if (new Regex(@"цена", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    for (int jj = 0; jj < listIndexOfNotEmptyName.Count; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[jj], i];   // вот здесь сдвигаемся
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            dtProduct.Rows[listIndexOfNotEmptyName[jj] - listIndexOfNotEmptyName[0]]["Цена"] = temp;
                                        }
                                    }
                                }
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }
                }
                clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции TrubDom\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Структура для хранения минитабличек, на которые дробится основной файл
        /// </summary>
        struct structTab //Inmet, SPK, UralTrubomet
        {
            public int StartCol;
            public int StartRow;
            public int Razriv;

            public List<int> listExcelIndexTab;
            public List<int> listdtProductIndexRow;

            //public int EndExcelRowFact;
            public string Name;
            public string Type;
            public string Standart;

            //public List<int> listPrices;//если нужно хранить несколько номеров столбцов с ценами
            //public List<int> listMarks;//если нужно хранить несколько номеров столбцов с марками
        }

        ///<summary> 
        ///<remarks> Открытие и чтение экселевского файла Инмет </remarks>
        ///<param name="path" >путь к файлу</param>
        ///</summary>
        private void Inmet(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                SetNameFromName(filePath);

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;



                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла Inmet\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;



                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    List<structTab> tabs = new List<structTab>();
                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                    tsPb1.Maximum = cCelCol * cCelRow;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    tsLabelClearingTable.Text = "Поиск наименований";
                    tsPb1.Value = 0;

                    Regex regName = new Regex(@"^\w+(?=\s)", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"(?<=^|\s|ф|=|^ков.*\()\d+(?:,\d+)?(?=[xх*-]|\w\w+|\s|$|\(|\)|;)", RegexOptions.IgnoreCase);
                    Regex regShirTolsh = new Regex(@"(?<=\d+(?:,\d+)?[xх*]\s?ф?|\d+[xх]ф\()\d+(?:,\d+)?(?=\s|$|\\|/|\(|\+|рж|-|ко|обт)", RegexOptions.IgnoreCase);
                    Regex regDlina = new Regex(@"(?<=^\d+(?:,\d+)?[xх*]\s?ф?/?)\d+(?:,\d+)?(?=[xх*])", RegexOptions.IgnoreCase);


                    structTab tab = new structTab();
                    tsPb1.Maximum = cCelRow * cCelCol;
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                Regex regRazmer = new Regex(@"размер", RegexOptions.IgnoreCase);
                                Regex regMetalloprokat = new Regex(@"металлопрокат", RegexOptions.IgnoreCase);

                                if (regRazmer.IsMatch(temp))
                                {
                                    tab.StartCol = i;
                                    tab.StartRow = j;
                                    tab.listExcelIndexTab = new List<int>();
                                    tab.listdtProductIndexRow = new List<int>();
                                    tabs.Add(tab);
                                }

                                if (regMetalloprokat.IsMatch(temp))
                                {
                                    tab.StartCol = i;
                                    tab.StartRow = j;
                                    tab.listExcelIndexTab = new List<int>();
                                    tab.listdtProductIndexRow = new List<int>();
                                    tabs.Add(tab);
                                }

                                InfoOrganization(temp);

                                #region Адрес
                                Regex regAddr = new Regex(@"^\s*г\..*оф\.\d+\s", RegexOptions.IgnoreCase);
                                if (regAddr.IsMatch(temp))
                                {
                                    Regex sklad = new Regex(@"^\s*г\..*оф\.\d+\s", RegexOptions.IgnoreCase);
                                    string tmp = sklad.Match(temp).Value;
                                    textBoxOrgAdress.Text = tmp;
                                }

                                #endregion

                                #region Склад
                                Regex regSklad = new Regex(@"(?<=Склада\s*:\s).*д\.\d+\s(?=\s;?)", RegexOptions.IgnoreCase);
                                if (regSklad.IsMatch(temp))
                                {
                                    Regex sklad = new Regex(@"(?<=Склада\s*:\s).*д\.\d+\s(?=\s;?)", RegexOptions.IgnoreCase);
                                    string tmp = sklad.Match(temp).Value;
                                    ListViewItem lvi = new ListViewItem(tmp);
                                    bool isIn = false;
                                    if (listViewAdrSklad.Items.Count < 1) listViewAdrSklad.Items.Add(lvi);
                                    for (int ii = 0; ii < listViewAdrSklad.Items.Count; ii++)
                                    {
                                        if (listViewAdrSklad.Items[ii].SubItems[0].Text == tmp) isIn = true;
                                    }
                                    if (!isIn) listViewAdrSklad.Items.Add(lvi);
                                }

                                #endregion

                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }
                    int endIndexRow = 0;
                    tsPb1.Value = 0;
                    tsPb1.Maximum = tabs.Count;
                    for (int k = 0; k < tabs.Count; k++)
                    {

                        if (k < tabs.Count - 1)
                        {
                            if (tabs[k + 1].StartRow > tabs[k].StartRow)
                                endIndexRow = tabs[k + 1].StartRow;
                            else endIndexRow = tabs[k + 2].StartRow;
                        }
                        else endIndexRow = cCelRow;

                        bool stop = false;
                        int firstCol = tabs[k].StartCol;
                        int curCol = tabs[k].StartCol;
                        while (!stop)
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].StartRow, curCol];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"размер", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    for (int i = 1; i < cCelCol; i++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].StartRow - 1, i];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim().ToLower();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            dtProduct.Rows.Add();
                                            tabs[k].listExcelIndexTab.Add(tabs[k].StartRow - 1);
                                            tabs[k].listdtProductIndexRow.Add(dtProduct.Rows.Count - 1);
                                            dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"] = new Regex(@"(?!\w+ые|\w+ие|\w+ая|\d)(?<=^|\s)\w+(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if (dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"].ToString().ToLower() == "стали") dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"] = "Cталь";
                                            dtProduct.Rows[dtProduct.Rows.Count - 1]["Тип"] = new Regex(@"(?<=^|\s)\w+(?:ые|ие|ая)(?=,|\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        }
                                    }
                                    for (int i = tabs[k].StartRow + 1; i < endIndexRow; i++) //строки
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[i, curCol];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            dtProduct.Rows.Add();
                                            tabs[k].listExcelIndexTab.Add(i);

                                            int lastRow = dtProduct.Rows.Count - 1;
                                            tabs[k].listdtProductIndexRow.Add(lastRow);

                                            dtProduct.Rows[lastRow]["Примечание"] = temp;
                                            if (regDiam.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDiam.Match(temp).Value;
                                            }
                                            if (regShirTolsh.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regShirTolsh.Match(temp).Value;
                                            }
                                            if (regDlina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = regDlina.Match(temp).Value;
                                            }

                                            dtProduct.Rows[dtProduct.Rows.Count - 1]["Тип"] = "тип не указан";
                                        }
                                    }
                                }
                                if (tabs[k].listExcelIndexTab.Count > 0) // сюда обработку других столбцов
                                {
                                    if (new Regex(@"марка", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        for (int i = 0; i < tabs[k].listExcelIndexTab.Count; i++) //строки
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[i], curCol];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Марка"] = temp;
                                            }
                                        }
                                    }

                                    if (new Regex(@"цена", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        for (int i = 0; i < tabs[k].listExcelIndexTab.Count; i++) //строки
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[i], curCol];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Цена"] = temp;
                                            }
                                        }
                                    }

                                    if (new Regex(@"ост-к|вес", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        for (int i = 0; i < tabs[k].listExcelIndexTab.Count; i++) //строки
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[i], curCol];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Мерность (т, м, мм)"] = temp;
                                            }
                                        }
                                    }
                                }
                            }
                            countIteration++;
                            if (curCol == firstCol && curCol > 1) curCol--;
                            else if (curCol < firstCol) curCol += 2;
                            else if (curCol > firstCol && curCol - firstCol < 3) curCol++;
                            else if (curCol == 1 && firstCol == 1) curCol++;
                            else stop = true;
                        }
                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        else tsPb1.Value = tsPb1.Maximum;
                    }
                }
                clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции Inmet\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        ///<summary> 
        ///<remarks> Открытие и чтение экселевского файла УралТрубосталь </remarks>
        ///<param name="path" >путь к файлу</param>
        ///</summary>
        private void UralTrubomet(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                SetNameFromName(filePath);

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла UralTrubomet\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;



                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    List<structTab> tabs = new List<structTab>();
                    List<double> DiamStd = new List<double>();
                    List<double> TolshStd = new List<double>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    tsPb1.Maximum = cCelCol * cCelRow;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    tsLabelClearingTable.Text = "Поиск наименований";
                    tsPb1.Value = 0;

                    int addedRow = 0;

                    DiamStd = GetStdFromFile(Application.StartupPath + "\\DiamStd.csv");
                    TolshStd = GetStdFromFile(Application.StartupPath + "\\TolshStd.csv");

                    structTab tab = new structTab();
                    tsPb1.Maximum = cCelRow * cCelCol;
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                Regex regRazmer = new Regex(@"размер", RegexOptions.IgnoreCase);

                                if (regRazmer.IsMatch(temp))
                                {
                                    tab.StartCol = i;
                                    tab.StartRow = j;
                                    tab.listExcelIndexTab = new List<int>();
                                    tab.listdtProductIndexRow = new List<int>();
                                    tabs.Add(tab);
                                }

                                InfoOrganization(temp);

                                #region Адрес
                                Regex regAddr = new Regex(@"^\s*г\..*оф\.\d+\s", RegexOptions.IgnoreCase);
                                if (regAddr.IsMatch(temp))
                                {
                                    Regex sklad = new Regex(@"^\s*г\..*оф\.\d+\s", RegexOptions.IgnoreCase);
                                    string tmp = sklad.Match(temp).Value;
                                    textBoxOrgAdress.Text = tmp;
                                }

                                #endregion

                                #region Склад
                                Regex regSklad = new Regex(@"(?<=Склада\s*:\s).*д\.\d+\s(?=\s;?)", RegexOptions.IgnoreCase);
                                if (regSklad.IsMatch(temp))
                                {
                                    Regex sklad = new Regex(@"(?<=Склада\s*:\s).*д\.\d+\s(?=\s;?)", RegexOptions.IgnoreCase);
                                    string tmp = sklad.Match(temp).Value;
                                    ListViewItem lvi = new ListViewItem(tmp);
                                    bool isIn = false;
                                    if (listViewAdrSklad.Items.Count < 1) listViewAdrSklad.Items.Add(lvi);
                                    for (int ii = 0; ii < listViewAdrSklad.Items.Count; ii++)
                                    {
                                        if (listViewAdrSklad.Items[ii].SubItems[0].Text == tmp) isIn = true;
                                    }
                                    if (!isIn) listViewAdrSklad.Items.Add(lvi);
                                }

                                #endregion

                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }
                    int endIndexRow = 0;
                    tsPb1.Value = 0;
                    tsPb1.Maximum = tabs.Count;
                    for (int k = 0; k < tabs.Count; k++)  //для всех найденных таблиц
                    {

                        if (k < tabs.Count - 1)
                        {
                            if (tabs[k + 1].StartRow > tabs[k].StartRow)
                                endIndexRow = tabs[k + 1].StartRow;
                            else endIndexRow = tabs[k + 2].StartRow;
                        }
                        else endIndexRow = cCelRow;

                        bool stop = false;
                        int firstCol = tabs[k].StartCol;
                        int curCol = tabs[k].StartCol;
                        while (!stop)
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].StartRow, curCol];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"размер", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    //смотрим на строку выше заголовка столбца таблицы и ищем название продукции, указанной в этой таблице

                                    for (int ii = 1; ii <= cCelCol; ii++)
                                    {
                                        int jj = tabs[k].StartRow - 2;
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, ii];
                                        if (cellRange.Value != null)
                                        {
                                            temp = cellRange.Value.ToString().Trim();
                                            dtProduct.Rows.Add();
                                            dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"] = new Regex(@"(?!\w+ые|\w+ие|\w+ая|\d)(?<=^|\s)\w+(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if (dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"].ToString().ToLower() == "стали") dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"] = "Cталь";
                                            if (dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"].ToString().ToLower() == "трубы") dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"] = "Труба";
                                            dtProduct.Rows[dtProduct.Rows.Count - 1]["Тип"] = new Regex(@"(?<=^|\s)\w+(?:ые|ие|ая)(?=,|\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        }
                                    }

                                    //если название все еще не найдено, то добавить просто название Труба
                                    if (dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"].ToString().Trim() == "")
                                        dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"] = "Труба";

                                    string[] strSizes;

                                    for (int jj = tabs[k].StartRow + 1; jj < endIndexRow; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, tabs[k].StartCol];
                                        if (cellRange.Value != null)
                                        {
                                            temp = cellRange.Value.ToString().Trim();
                                            strSizes = temp.Split('-');
                                            if (strSizes.Length == 2)
                                            {
                                                double size1 = Convert.ToDouble(new Regex(@"\d+(?:,\d*)?", RegexOptions.IgnoreCase).Match(strSizes[0]).Value);
                                                double size2 = Convert.ToDouble(new Regex(@"\d+(?:,\d*)?", RegexOptions.IgnoreCase).Match(strSizes[1]).Value);
                                                for (int s = 0; s < DiamStd.Count; s++)
                                                {
                                                    if (new Regex(@"\d+(?:,\d*)?", RegexOptions.IgnoreCase).IsMatch(strSizes[0]))
                                                    {

                                                        if (DiamStd[s] >= size1 && DiamStd[s] <= size2)
                                                        {
                                                            dtProduct.Rows.Add();
                                                            tabs[k].listExcelIndexTab.Add(jj);
                                                            int lastRow = dtProduct.Rows.Count - 1;
                                                            tabs[k].listdtProductIndexRow.Add(lastRow);
                                                            dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = DiamStd[s];
                                                            dtProduct.Rows[lastRow]["Тип"] = "тип не указан";

                                                            cellRange = (Excel.Range)excelworksheet.Cells[jj, tabs[k].StartCol + 1];
                                                            if (cellRange.Value != null)
                                                            {
                                                                temp = cellRange.Value.ToString().Trim();
                                                                if (new Regex(@"-", RegexOptions.IgnoreCase).IsMatch(temp))
                                                                {

                                                                    string[] strTol = temp.Split('-');

                                                                    if (strTol.Length == 2)
                                                                    {
                                                                        double Tsize1 = Convert.ToDouble(new Regex(@"\d+(?:,\d+)?", RegexOptions.IgnoreCase).Match(strTol[0]).Value);
                                                                        double Tsize2 = Convert.ToDouble(new Regex(@"\d+(?:,\d+)?", RegexOptions.IgnoreCase).Match(strTol[1]).Value);
                                                                        if (Tsize1 > Tsize2)
                                                                        {
                                                                            dtProduct.Rows[lastRow + addedRow]["Толщина (ширина), мм"] = temp;
                                                                        }
                                                                        else
                                                                        {
                                                                            bool first = true;
                                                                            for (int t = 0; t < TolshStd.Count; t++)
                                                                            {
                                                                                if (TolshStd[t] >= Tsize1 && TolshStd[t] <= Tsize2)
                                                                                {
                                                                                    if (first)
                                                                                    {
                                                                                        dtProduct.Rows[lastRow + addedRow]["Толщина (ширина), мм"] = TolshStd[t];
                                                                                        first = false;
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        DataRow row = dtProduct.NewRow();
                                                                                        row["Название"] = dtProduct.Rows[lastRow + addedRow]["Название"];
                                                                                        row["Тип"] = dtProduct.Rows[lastRow + addedRow]["Тип"];
                                                                                        row["Диаметр (высота), мм"] = dtProduct.Rows[lastRow + addedRow]["Диаметр (высота), мм"];
                                                                                        row["Толщина (ширина), мм"] = TolshStd[t];
                                                                                        row["Марка"] = dtProduct.Rows[lastRow + addedRow]["Марка"];
                                                                                        row["Стандарт"] = dtProduct.Rows[lastRow + addedRow]["Стандарт"];
                                                                                        row["Примечание"] = dtProduct.Rows[lastRow + addedRow]["Примечание"];
                                                                                        row["Мерность (т, м, мм)"] = dtProduct.Rows[lastRow + addedRow]["Мерность (т, м, мм)"];
                                                                                        row["Цена"] = dtProduct.Rows[lastRow + addedRow]["Цена"];
                                                                                        dtProduct.Rows.Add(row);
                                                                                        tabs[k].listExcelIndexTab.Add(jj);
                                                                                        lastRow = dtProduct.Rows.Count - 1;
                                                                                        tabs[k].listdtProductIndexRow.Add(lastRow);
                                                                                        //addedRow++;
                                                                                    }
                                                                                }
                                                                            }
                                                                        }

                                                                    }
                                                                }
                                                                else if (new Regex(@",", RegexOptions.IgnoreCase).IsMatch(temp))
                                                                {
                                                                    int matches = new Regex(@",", RegexOptions.IgnoreCase).Matches(temp).Count;
                                                                    if (matches == 1)
                                                                    {
                                                                        if (!new Regex(@",\s", RegexOptions.IgnoreCase).IsMatch(temp))
                                                                        {
                                                                            string[] strTol = temp.Split(',');
                                                                            double Tsize1 = Convert.ToDouble(new Regex(@"\d+(?:,\d+)?", RegexOptions.IgnoreCase).Match(strTol[0]).Value);
                                                                            double Tsize2 = Convert.ToDouble(new Regex(@"\d+(?:,\d+)?", RegexOptions.IgnoreCase).Match(strTol[1]).Value);
                                                                            if (Tsize1 > Tsize2)
                                                                            {
                                                                                dtProduct.Rows[lastRow + addedRow]["Толщина (ширина), мм"] = Tsize1;
                                                                                DataRow row = dtProduct.NewRow();
                                                                                row["Название"] = dtProduct.Rows[lastRow + addedRow]["Название"];
                                                                                row["Тип"] = dtProduct.Rows[lastRow + addedRow]["Тип"];
                                                                                row["Диаметр (высота), мм"] = dtProduct.Rows[lastRow + addedRow]["Диаметр (высота), мм"];
                                                                                row["Толщина (ширина), мм"] = Tsize2;
                                                                                row["Марка"] = dtProduct.Rows[lastRow + addedRow]["Марка"];
                                                                                row["Стандарт"] = dtProduct.Rows[lastRow + addedRow]["Стандарт"];
                                                                                row["Примечание"] = dtProduct.Rows[lastRow + addedRow]["Примечание"];
                                                                                row["Мерность (т, м, мм)"] = dtProduct.Rows[lastRow + addedRow]["Мерность (т, м, мм)"];
                                                                                row["Цена"] = dtProduct.Rows[lastRow + addedRow]["Цена"];
                                                                                dtProduct.Rows.Add(row);
                                                                                tabs[k].listExcelIndexTab.Add(jj);
                                                                                lastRow = dtProduct.Rows.Count - 1;
                                                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                                            }
                                                                            else dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"\d+(?:,\d+)?(?=\s|-|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                                        }
                                                                        else
                                                                        {
                                                                            string[] strTol = temp.Split(',');
                                                                            double Tsize1 = Convert.ToDouble(new Regex(@"\d+(?:,\d+)?", RegexOptions.IgnoreCase).Match(strTol[0]).Value);
                                                                            double Tsize2 = Convert.ToDouble(new Regex(@"\d+(?:,\d+)?", RegexOptions.IgnoreCase).Match(strTol[1]).Value);
                                                                            if (Tsize1 > Tsize2)
                                                                            {
                                                                                dtProduct.Rows[lastRow + addedRow]["Толщина (ширина), мм"] = Tsize1;
                                                                                DataRow row = dtProduct.NewRow();
                                                                                row["Название"] = dtProduct.Rows[lastRow + addedRow]["Название"];
                                                                                row["Тип"] = dtProduct.Rows[lastRow + addedRow]["Тип"];
                                                                                row["Диаметр (высота), мм"] = dtProduct.Rows[lastRow + addedRow]["Диаметр (высота), мм"];
                                                                                row["Толщина (ширина), мм"] = Tsize2;
                                                                                row["Марка"] = dtProduct.Rows[lastRow + addedRow]["Марка"];
                                                                                row["Стандарт"] = dtProduct.Rows[lastRow + addedRow]["Стандарт"];
                                                                                row["Примечание"] = dtProduct.Rows[lastRow + addedRow]["Примечание"];
                                                                                row["Мерность (т, м, мм)"] = dtProduct.Rows[lastRow + addedRow]["Мерность (т, м, мм)"];
                                                                                row["Цена"] = dtProduct.Rows[lastRow + addedRow]["Цена"];
                                                                                dtProduct.Rows.Add(row);
                                                                                tabs[k].listExcelIndexTab.Add(jj);
                                                                                lastRow = dtProduct.Rows.Count - 1;
                                                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                                            }
                                                                        }
                                                                    }
                                                                    if (matches > 1)
                                                                    {
                                                                        string[] strTol = temp.Split(',');
                                                                        for (int x = 0; x < strTol.Length; x++)
                                                                        {
                                                                            double Tsize1 = Convert.ToDouble(new Regex(@"\d+(?:,\d+)?", RegexOptions.IgnoreCase).Match(strTol[x]).Value);
                                                                            if (x == 0)
                                                                            {
                                                                                dtProduct.Rows[lastRow + addedRow]["Толщина (ширина), мм"] = Tsize1;
                                                                            }
                                                                            else
                                                                            {
                                                                                DataRow row = dtProduct.NewRow();
                                                                                row["Название"] = dtProduct.Rows[lastRow + addedRow]["Название"];
                                                                                row["Тип"] = dtProduct.Rows[lastRow + addedRow]["Тип"];
                                                                                row["Диаметр (высота), мм"] = dtProduct.Rows[lastRow + addedRow]["Диаметр (высота), мм"];
                                                                                row["Толщина (ширина), мм"] = Tsize1;
                                                                                row["Марка"] = dtProduct.Rows[lastRow + addedRow]["Марка"];
                                                                                row["Стандарт"] = dtProduct.Rows[lastRow + addedRow]["Стандарт"];
                                                                                row["Примечание"] = dtProduct.Rows[lastRow + addedRow]["Примечание"];
                                                                                row["Мерность (т, м, мм)"] = dtProduct.Rows[lastRow + addedRow]["Мерность (т, м, мм)"];
                                                                                row["Цена"] = dtProduct.Rows[lastRow + addedRow]["Цена"];
                                                                                dtProduct.Rows.Add(row);
                                                                                tabs[k].listExcelIndexTab.Add(jj);
                                                                                lastRow = dtProduct.Rows.Count - 1;
                                                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                    dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"\d+(?:,\d+)?(?=\s|-|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                            }
                                                        }

                                                    }
                                                }
                                            }
                                            else if (strSizes.Length == 1)
                                            {
                                                if (new Regex(@"\d+(?:,\d*)?", RegexOptions.IgnoreCase).IsMatch(strSizes[0]))
                                                {
                                                    dtProduct.Rows.Add();
                                                    tabs[k].listExcelIndexTab.Add(jj);
                                                    int lastRow = dtProduct.Rows.Count - 1;
                                                    tabs[k].listdtProductIndexRow.Add(lastRow);
                                                    dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"\d+(?:,\d*)?", RegexOptions.IgnoreCase).Match(strSizes[0]).Value;
                                                    dtProduct.Rows[lastRow]["Тип"] = "тип не указан";

                                                    cellRange = (Excel.Range)excelworksheet.Cells[jj, tabs[k].StartCol + 1];
                                                    if (cellRange.Value != null)
                                                    {
                                                        temp = cellRange.Value.ToString().Trim();
                                                        if (new Regex(@"-", RegexOptions.IgnoreCase).IsMatch(temp))
                                                        {

                                                            string[] strTol = temp.Split('-');

                                                            if (strTol.Length == 2)
                                                            {
                                                                double Tsize1 = Convert.ToDouble(new Regex(@"\d+(?:,\d+)?", RegexOptions.IgnoreCase).Match(strTol[0]).Value);
                                                                double Tsize2 = Convert.ToDouble(new Regex(@"\d+(?:,\d+)?", RegexOptions.IgnoreCase).Match(strTol[1]).Value);
                                                                if (Tsize1 > Tsize2)
                                                                {
                                                                    dtProduct.Rows[lastRow + addedRow]["Толщина (ширина), мм"] = temp;
                                                                }
                                                                else
                                                                {
                                                                    bool first = true;
                                                                    for (int t = 0; t < TolshStd.Count; t++)
                                                                    {
                                                                        if (TolshStd[t] >= Tsize1 && TolshStd[t] <= Tsize2)
                                                                        {
                                                                            if (first)
                                                                            {
                                                                                dtProduct.Rows[lastRow + addedRow]["Толщина (ширина), мм"] = TolshStd[t];
                                                                                first = false;
                                                                            }
                                                                            else
                                                                            {
                                                                                DataRow row = dtProduct.NewRow();
                                                                                row["Название"] = dtProduct.Rows[lastRow + addedRow]["Название"];
                                                                                row["Тип"] = dtProduct.Rows[lastRow + addedRow]["Тип"];
                                                                                row["Диаметр (высота), мм"] = dtProduct.Rows[lastRow + addedRow]["Диаметр (высота), мм"];
                                                                                row["Толщина (ширина), мм"] = TolshStd[t];
                                                                                row["Марка"] = dtProduct.Rows[lastRow + addedRow]["Марка"];
                                                                                row["Стандарт"] = dtProduct.Rows[lastRow + addedRow]["Стандарт"];
                                                                                row["Примечание"] = dtProduct.Rows[lastRow + addedRow]["Примечание"];
                                                                                row["Мерность (т, м, мм)"] = dtProduct.Rows[lastRow + addedRow]["Мерность (т, м, мм)"];
                                                                                row["Цена"] = dtProduct.Rows[lastRow + addedRow]["Цена"];
                                                                                dtProduct.Rows.Add(row);
                                                                                tabs[k].listExcelIndexTab.Add(jj);
                                                                                lastRow = dtProduct.Rows.Count - 1;
                                                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                                                //addedRow++;
                                                                            }
                                                                        }
                                                                    }
                                                                }

                                                            }
                                                        }
                                                        else if (new Regex(@",", RegexOptions.IgnoreCase).IsMatch(temp))
                                                        {
                                                            int matches = new Regex(@",", RegexOptions.IgnoreCase).Matches(temp).Count;
                                                            if (matches == 1)
                                                            {
                                                                if (!new Regex(@",\s", RegexOptions.IgnoreCase).IsMatch(temp))
                                                                {
                                                                    string[] strTol = temp.Split(',');
                                                                    double Tsize1 = Convert.ToDouble(new Regex(@"\d+(?:,\d+)?", RegexOptions.IgnoreCase).Match(strTol[0]).Value);
                                                                    double Tsize2 = Convert.ToDouble(new Regex(@"\d+(?:,\d+)?", RegexOptions.IgnoreCase).Match(strTol[1]).Value);
                                                                    if (Tsize1 > Tsize2)
                                                                    {
                                                                        dtProduct.Rows[lastRow + addedRow]["Толщина (ширина), мм"] = Tsize1;
                                                                        DataRow row = dtProduct.NewRow();
                                                                        row["Название"] = dtProduct.Rows[lastRow + addedRow]["Название"];
                                                                        row["Тип"] = dtProduct.Rows[lastRow + addedRow]["Тип"];
                                                                        row["Диаметр (высота), мм"] = dtProduct.Rows[lastRow + addedRow]["Диаметр (высота), мм"];
                                                                        row["Толщина (ширина), мм"] = Tsize2;
                                                                        row["Марка"] = dtProduct.Rows[lastRow + addedRow]["Марка"];
                                                                        row["Стандарт"] = dtProduct.Rows[lastRow + addedRow]["Стандарт"];
                                                                        row["Примечание"] = dtProduct.Rows[lastRow + addedRow]["Примечание"];
                                                                        row["Мерность (т, м, мм)"] = dtProduct.Rows[lastRow + addedRow]["Мерность (т, м, мм)"];
                                                                        row["Цена"] = dtProduct.Rows[lastRow + addedRow]["Цена"];
                                                                        dtProduct.Rows.Add(row);
                                                                        tabs[k].listExcelIndexTab.Add(jj);
                                                                        lastRow = dtProduct.Rows.Count - 1;
                                                                        tabs[k].listdtProductIndexRow.Add(lastRow);
                                                                    }
                                                                    else dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"\d+(?:,\d+)?(?=\s|-|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                                }
                                                                else
                                                                {
                                                                    string[] strTol = temp.Split(',');
                                                                    double Tsize1 = Convert.ToDouble(new Regex(@"\d+(?:,\d+)?", RegexOptions.IgnoreCase).Match(strTol[0]).Value);
                                                                    double Tsize2 = Convert.ToDouble(new Regex(@"\d+(?:,\d+)?", RegexOptions.IgnoreCase).Match(strTol[1]).Value);
                                                                    if (Tsize1 > Tsize2)
                                                                    {
                                                                        dtProduct.Rows[lastRow + addedRow]["Толщина (ширина), мм"] = Tsize1;
                                                                        DataRow row = dtProduct.NewRow();
                                                                        row["Название"] = dtProduct.Rows[lastRow + addedRow]["Название"];
                                                                        row["Тип"] = dtProduct.Rows[lastRow + addedRow]["Тип"];
                                                                        row["Диаметр (высота), мм"] = dtProduct.Rows[lastRow + addedRow]["Диаметр (высота), мм"];
                                                                        row["Толщина (ширина), мм"] = Tsize2;
                                                                        row["Марка"] = dtProduct.Rows[lastRow + addedRow]["Марка"];
                                                                        row["Стандарт"] = dtProduct.Rows[lastRow + addedRow]["Стандарт"];
                                                                        row["Примечание"] = dtProduct.Rows[lastRow + addedRow]["Примечание"];
                                                                        row["Мерность (т, м, мм)"] = dtProduct.Rows[lastRow + addedRow]["Мерность (т, м, мм)"];
                                                                        row["Цена"] = dtProduct.Rows[lastRow + addedRow]["Цена"];
                                                                        dtProduct.Rows.Add(row);
                                                                        tabs[k].listExcelIndexTab.Add(jj);
                                                                        lastRow = dtProduct.Rows.Count - 1;
                                                                        tabs[k].listdtProductIndexRow.Add(lastRow);
                                                                    }
                                                                }
                                                            }
                                                            if (matches > 1)
                                                            {
                                                                string[] strTol = temp.Split(',');
                                                                for (int x = 0; x < strTol.Length; x++)
                                                                {
                                                                    double Tsize1 = Convert.ToDouble(new Regex(@"\d+(?:,\d+)?", RegexOptions.IgnoreCase).Match(strTol[x]).Value);
                                                                    if (x == 0)
                                                                    {
                                                                        dtProduct.Rows[lastRow + addedRow]["Толщина (ширина), мм"] = Tsize1;
                                                                    }
                                                                    else
                                                                    {
                                                                        DataRow row = dtProduct.NewRow();
                                                                        row["Название"] = dtProduct.Rows[lastRow + addedRow]["Название"];
                                                                        row["Тип"] = dtProduct.Rows[lastRow + addedRow]["Тип"];
                                                                        row["Диаметр (высота), мм"] = dtProduct.Rows[lastRow + addedRow]["Диаметр (высота), мм"];
                                                                        row["Толщина (ширина), мм"] = Tsize1;
                                                                        row["Марка"] = dtProduct.Rows[lastRow + addedRow]["Марка"];
                                                                        row["Стандарт"] = dtProduct.Rows[lastRow + addedRow]["Стандарт"];
                                                                        row["Примечание"] = dtProduct.Rows[lastRow + addedRow]["Примечание"];
                                                                        row["Мерность (т, м, мм)"] = dtProduct.Rows[lastRow + addedRow]["Мерность (т, м, мм)"];
                                                                        row["Цена"] = dtProduct.Rows[lastRow + addedRow]["Цена"];
                                                                        dtProduct.Rows.Add(row);
                                                                        tabs[k].listExcelIndexTab.Add(jj);
                                                                        lastRow = dtProduct.Rows.Count - 1;
                                                                        tabs[k].listdtProductIndexRow.Add(lastRow);
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        else
                                                            dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"\d+(?:,\d+)?(?=\s|-|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                    }
                                                }
                                            }
                                            cellRange = (Excel.Range)excelworksheet.Cells[jj, tabs[k].StartCol + 1];
                                            if (cellRange.Value != null)
                                            {
                                                temp = cellRange.Value.ToString().Trim();
                                            }
                                        }
                                    }
                                }

                            }
                            countIteration++;
                            if (curCol == firstCol && curCol > 1) curCol--;
                            else if (curCol < firstCol) curCol += 2;
                            else if (curCol > firstCol && curCol - firstCol < 6) curCol++;
                            else if (curCol == 1 && firstCol == 1) curCol++;
                            else stop = true;
                        }



                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        else tsPb1.Value = tsPb1.Maximum;
                    }

                    tsPb1.Value = 0;
                    tsPb1.Maximum = tabs.Count;
                    for (int k = 0; k < tabs.Count; k++)
                    {
                        for (int curCol = 1; curCol <= cCelCol; curCol++)
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].StartRow, curCol];
                            if (cellRange.Value != null)
                            {
                                temp = cellRange.Value.ToString().Trim();
                                if (tabs[k].listExcelIndexTab.Count > 0) // сюда обработку других столбцов
                                {
                                    if (new Regex(@"Гост", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        for (int i = 0; i < tabs[k].listExcelIndexTab.Count; i++) //строки
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[i], curCol];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Стандарт"] = temp;
                                            }
                                        }
                                    }

                                    if (new Regex(@"марка", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        for (int i = 0; i < tabs[k].listExcelIndexTab.Count; i++) //строки
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[i], curCol];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Марка"] = temp;
                                            }
                                        }
                                    }

                                    if (new Regex(@"Производитель", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        for (int i = 0; i < tabs[k].listExcelIndexTab.Count; i++) //строки
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[i], curCol];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Примечание"] = "Производитель:ч" + temp;
                                            }
                                        }
                                    }

                                    if (new Regex(@"розница", RegexOptions.IgnoreCase).IsMatch(temp))  //типа цена
                                    {
                                        for (int i = 0; i < tabs[k].listExcelIndexTab.Count; i++) //строки
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[i], curCol];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Цена"] = temp;
                                            }
                                        }
                                    }

                                    if (new Regex(@"стенка", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        for (int i = 0; i < tabs[k].listExcelIndexTab.Count; i++)
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[i], curCol];
                                            if (cellRange.Value != null)
                                            {
                                                temp = cellRange.Value.ToString().Trim();

                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        else tsPb1.Value = tsPb1.Maximum;
                    }
                }
                clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции UralTrubomet\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла СПК
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void SPK110517(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                SetNameFromName(filePath);

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла SPK\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    List<structTab> tabs = new List<structTab>();
                    List<double> DiamStd = new List<double>();
                    List<double> TolshStd = new List<double>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    tsPb1.Maximum = 18 * cCelRow;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    tsLabelClearingTable.Text = "Поиск наименований";
                    tsPb1.Value = 0;

                    DiamStd = GetStdFromFile(Application.StartupPath + "\\DiamStd.csv");
                    TolshStd = GetStdFromFile(Application.StartupPath + "\\TolshStd.csv");

                    structTab tab = new structTab();
                    tsPb1.Maximum = cCelRow * 18;
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= 18; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                Regex regRazmer = new Regex(@"^наиме", RegexOptions.IgnoreCase);

                                if (regRazmer.IsMatch(temp))
                                {
                                    if (cellRange.MergeArea.Count == 2)
                                    {
                                        i++;
                                    }
                                    tab.StartCol = i;
                                    tab.StartRow = j;
                                    tab.listExcelIndexTab = new List<int>();
                                    tab.listdtProductIndexRow = new List<int>();
                                    tabs.Add(tab);

                                }

                                if (tabs.Count > 0)
                                    if (j < tabs[0].StartRow && j > tabs[tabs.Count - 1].StartRow)
                                    {
                                        InfoOrganization(temp);

                                        #region Адрес
                                        Regex regAddr = new Regex(@"^отдел\s*розн.*ж\s*:\s+", RegexOptions.IgnoreCase);
                                        if (regAddr.IsMatch(temp))
                                        {
                                            Regex sklad = new Regex(@"(?<=^отдел\s*розн.*ж\s*:\s+).*\.$", RegexOptions.IgnoreCase);
                                            string tmp = sklad.Match(temp).Value;
                                            textBoxOrgAdress.Text = tmp;
                                        }

                                        #endregion

                                        #region Склад
                                        Regex regSklad = new Regex(@"(?<=металлобаза\s*:\s)", RegexOptions.IgnoreCase);
                                        if (regSklad.IsMatch(temp))
                                        {
                                            Regex sklad = new Regex(@"(?<=металлобаза\s*:\s).*\.$", RegexOptions.IgnoreCase);
                                            string tmp = sklad.Match(temp).Value;
                                            ListViewItem lvi = new ListViewItem(tmp);
                                            bool isIn = false;
                                            if (listViewAdrSklad.Items.Count < 1) listViewAdrSklad.Items.Add(lvi);
                                            for (int ii = 0; ii < listViewAdrSklad.Items.Count; ii++)
                                            {
                                                if (listViewAdrSklad.Items[ii].SubItems[0].Text == tmp) isIn = true;
                                            }
                                            if (!isIn) listViewAdrSklad.Items.Add(lvi);
                                        }

                                        #endregion
                                    }

                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }
                    int endIndexRow = 0;
                    tsPb1.Value = 0;
                    tsPb1.Maximum = tabs.Count;

                    for (int k = 0; k < tabs.Count; k++)  //для всех найденных таблиц
                    {

                        if (k < tabs.Count - 1)
                        {
                            if (tabs[k + 1].StartRow > tabs[k].StartRow)
                                endIndexRow = tabs[k + 1].StartRow;
                            else if (k < tabs.Count - 2)
                            {
                                if (tabs[k + 2].StartRow > tabs[k].StartRow)
                                    endIndexRow = tabs[k + 2].StartRow;
                                else if (k < tabs.Count - 3)
                                {
                                    if (tabs[k + 3].StartRow > tabs[k].StartRow)
                                        endIndexRow = tabs[k + 3].StartRow;
                                }
                                else endIndexRow = cCelRow;
                            }
                            else endIndexRow = cCelRow;
                        }
                        else endIndexRow = cCelRow;
                        if (tabs[k].StartRow == endIndexRow) endIndexRow = cCelCol;
                        bool stop = false;
                        int firstCol = tabs[k].StartCol;
                        int curCol = tabs[k].StartCol;
                        while (!stop)
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].StartRow, curCol];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"наим.*ие", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    for (int jj = tabs[k].StartRow + 1; jj < endIndexRow; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, tabs[k].StartCol];
                                        if (cellRange.Value != null)
                                        {
                                            temp = cellRange.Value.ToString().Trim();
                                            bool isAddedRow = false;
                                            if (new Regex(@"(?!\w+ые|\w+ие|\w+ая|\w+\d|\d|ГОСТ)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                dtProduct.Rows.Add(); isAddedRow = true;
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                dtProduct.Rows[lastRow]["Название"] = new Regex(@"(?!\w+ые|\w+ие|\w+ая|\w+\d|\d|ГОСТ)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "стали") dtProduct.Rows[lastRow]["Название"] = "Cталь";
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "трубы") dtProduct.Rows[lastRow]["Название"] = "Труба";
                                                dtProduct.Rows[lastRow]["Тип"] = new Regex(@"(?<=^|\s)(?:\w+(?:ые|ие|ая|ый))|BP|ВР|[AА]-*I{1,3}|х/к|г/к|Н/У\sТ/О\sчерная|ВГП|э/св|б/ш|оцинк(?=,|\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=^\d+(?:,\d+)[хx])\d+(?:,\d+)(?=\s|-|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Примечание"] = temp;
                                            }
                                            if (new Regex(@"\d+[xх]\d+\s*-\s*\d+", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                if (!isAddedRow) dtProduct.Rows.Add();
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                tabs[k].listExcelIndexTab.Add(jj);
                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\d+(?:[,.]\d+)?[xх]\d+(?:[,.]\d+)?\s*-\s*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"\d+(?:[,.]\d+)?(?=[xх]\d+(?:[,.]\d+)?\s*-\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=\d+(?:[,.]\d+)?[xх])\d+(?:[,.]\d+)?(?=\s*-\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                            }
                                            else if (new Regex(@"^\d+(?:[.,]\d+)?(?:ду)?\s*-\s*\d+(?:[.,]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                if (!isAddedRow) dtProduct.Rows.Add();
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                tabs[k].listExcelIndexTab.Add(jj);
                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"\d+(?:[.,]\d+)?(?=(?:ду)?\s*-\s*\d+(?:[.,]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\d+(?:[.,]\d+)?(?:ду)?\s*-\s*)\d\.\d", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                            }
                                            else if (new Regex(@"^\d+(?:[.,]\d+)?\s+\d+(?:[.,]\d+)?[хx]\d+(?:[.,]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                if (!isAddedRow) dtProduct.Rows.Add();
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                tabs[k].listExcelIndexTab.Add(jj);
                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=\d+(?:[.,]\d+)?\s+)\d+(?:[.,]\d+)?(?=[хx]\d+(?:[.,]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"\d+(?:[.,]\d+)?(?=\s+\d+(?:[.,]\d+)?[хx]\d+(?:[.,]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=\d+(?:[.,]\d+)?\s+\d+(?:[.,]\d+)?[хx])\d+(?:[.,]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                            }
                                            else if (new Regex(@"(?<=^|d|^ДУ)\d+(?:[.,]\d+)?(?=[xх]|\s|$|-\d+$|[бшкм]|ДУ|,\s)", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                if (!isAddedRow) dtProduct.Rows.Add();
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                tabs[k].listExcelIndexTab.Add(jj);
                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=^|d|ДУ)\d+(?:[.,]\d+)?(?=\s|[хx]|-|$|[бшкм]|ДУ|,\s)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=[хx])\d+(?:[.,]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                            }
                                            else if (new Regex(@"^\d+(?:[.,]\d+)?\s*-\s*\d+(?:[.,]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                if (!isAddedRow) dtProduct.Rows.Add();
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                string[] strSizes;
                                                strSizes = temp.Split('-');
                                                if (strSizes.Length == 2)
                                                {
                                                    double size1 = Convert.ToDouble(new Regex(@"\d+(?:[.,]\d*)?", RegexOptions.IgnoreCase).Match(strSizes[0]).Value);
                                                    double size2 = Convert.ToDouble(new Regex(@"\d+(?:[.,]\d*)?", RegexOptions.IgnoreCase).Match(strSizes[1]).Value);
                                                    for (int s = 0; s < DiamStd.Count; s++)
                                                    {
                                                        if (new Regex(@"\d+(?:[.,]\d*)?", RegexOptions.IgnoreCase).IsMatch(strSizes[0]))
                                                        {

                                                            if (DiamStd[s] >= size1 && DiamStd[s] <= size2)
                                                            {
                                                                //dtProduct.Rows.Add();
                                                                tabs[k].listExcelIndexTab.Add(jj);
                                                                lastRow = dtProduct.Rows.Count - 1;
                                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = DiamStd[s];
                                                                dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                                            }
                                                        }
                                                    }
                                                }
                                            }

                                        }
                                    }

                                    //если название все еще не найдено, то добавить просто название Труба
                                    //if(dtProduct.Rows.Count>0)
                                    //if (dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"].ToString().Trim() == "")
                                    //    dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"] = "Труба";
                                }

                            }
                            countIteration++;
                            if (curCol == firstCol && curCol > 1) curCol++;
                            else if (curCol > firstCol && curCol - firstCol < 4) curCol++;
                            else if (curCol == 1 && firstCol == 1) curCol++;
                            else stop = true;
                        }

                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        else tsPb1.Value = tsPb1.Maximum;
                    }

                    tsPb1.Value = 0;
                    tsPb1.Maximum = tabs.Count;
                    for (int k = 0; k < tabs.Count; k++)
                    {
                        for (int curCol = tabs[k].StartCol; curCol <= tabs[k].StartCol + 4; curCol++)
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].StartRow, curCol];
                            if (cellRange.Value != null)
                            {
                                temp = cellRange.Value.ToString().Trim();
                                if (tabs[k].listExcelIndexTab.Count > 0) // сюда обработку других столбцов
                                {
                                    if (new Regex(@"сталь", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        for (int i = 0; i < tabs[k].listExcelIndexTab.Count; i++) //строки
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[i], curCol];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Марка"] = temp;
                                            }
                                        }
                                    }


                                    if (new Regex(@"цена(?=.*тн\.?$|$)", RegexOptions.IgnoreCase).IsMatch(temp))  //типа цена
                                    {
                                        for (int i = 0; i < tabs[k].listExcelIndexTab.Count; i++) //строки
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[i], curCol];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Цена"] = new Regex(@"\d+", RegexOptions.IgnoreCase).Match(temp).Value;
                                                if (dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Цена"].ToString().Length == 1)
                                                {
                                                    cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[i], curCol + 1];
                                                    if (cellRange.Value != null)
                                                        temp = cellRange.Value.ToString().Trim();
                                                    else temp = "";
                                                    if (temp != "")
                                                    {
                                                        dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Цена"] = new Regex(@"\d+", RegexOptions.IgnoreCase).Match(temp).Value;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        else tsPb1.Value = tsPb1.Maximum;


                    }
                    //проверка строк для InfoOrganization

                    if (tabs.Count > 0)
                    {
                        Excel.Shapes Shapes = excelworksheet.Shapes;
                        foreach (Excel.Shape Shape in Shapes)
                        {
                            if (Shape.Nodes.Count > 0)
                            {
                                for (int i = 1; i <= Shape.TextFrame2.TextRange.Lines.Count; i++)
                                {
                                    temp = Shape.TextFrame2.TextRange.Lines[i].Text.ToString();

                                    #region Адрес
                                    Regex regAddr = new Regex(@"отдел\s*опт.*ж\s*:\s+", RegexOptions.IgnoreCase);
                                    if (regAddr.IsMatch(temp))
                                    {
                                        Regex sklad = new Regex(@"отдел\s*опт.*ж\s*:\s+.*\.\s*$", RegexOptions.IgnoreCase);
                                        string tmp = sklad.Match(temp).Value;
                                        textBoxOrgAdress.Text = tmp;

                                    }

                                    #endregion

                                    #region Телефон
                                    Regex regTelefon = new Regex(@"Тел/факс", RegexOptions.IgnoreCase);
                                    if (regTelefon.IsMatch(temp))
                                    {
                                        Regex tel = new Regex(@"(?<=Тел/факс\s*)\+\d\s*\(\d{3,5}\)(?:\s*-*\d{2,3})+", RegexOptions.IgnoreCase);
                                        string tmp = tel.Match(temp).Value;
                                        textBoxOrgTelefon.Text = tmp;

                                    }

                                    #endregion

                                    #region Склад
                                    Regex regSklad = new Regex(@"(?<=металлобаза\s*:\s)", RegexOptions.IgnoreCase);
                                    if (regSklad.IsMatch(temp))
                                    {
                                        Regex sklad = new Regex(@"металлобаза\s*:\s.*\.$", RegexOptions.IgnoreCase);
                                        string tmp = sklad.Match(temp).Value;
                                        ListViewItem lvi = new ListViewItem(tmp);
                                        bool isIn = false;
                                        if (listViewAdrSklad.Items.Count < 1) listViewAdrSklad.Items.Add(lvi);
                                        for (int ii = 0; ii < listViewAdrSklad.Items.Count; ii++)
                                        {
                                            if (listViewAdrSklad.Items[ii].SubItems[0].Text == tmp) isIn = true;
                                        }
                                        if (!isIn) listViewAdrSklad.Items.Add(lvi);
                                    }
                                }
                            }
                            #endregion
                        }
                    }
                }

                clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции SPK\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла Евраз
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void Evraz(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                SetNameFromName(filePath);

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла Evraz\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    List<structTab> tabs = new List<structTab>();
                    List<double> DiamStd = new List<double>();
                    List<double> TolshStd = new List<double>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    tsPb1.Maximum = 18 * cCelRow;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    tsLabelClearingTable.Text = "Поиск наименований";
                    tsPb1.Value = 0;

                    DiamStd = GetStdFromFile(Application.StartupPath + "\\DiamStd.csv");
                    TolshStd = GetStdFromFile(Application.StartupPath + "\\TolshStd.csv");

                    structTab tab = new structTab();
                    tsPb1.Maximum = cCelRow * 18;
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= 18; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                Regex regRazmer = new Regex(@"^профиль", RegexOptions.IgnoreCase);

                                if (regRazmer.IsMatch(temp))
                                {
                                    if (cellRange.MergeArea.Count == 2)
                                    {
                                        tab.StartRow = j + 1;
                                    }
                                    tab.StartCol = i;
                                    tab.listExcelIndexTab = new List<int>();
                                    tab.listdtProductIndexRow = new List<int>();
                                    tabs.Add(tab);
                                }
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }
                    int endIndexRow = 0;
                    tsPb1.Value = 0;
                    tsPb1.Maximum = tabs.Count;

                    for (int k = 0; k < tabs.Count; k++)  //для всех найденных таблиц
                    {

                        if (k < tabs.Count - 1)
                        {
                            if (tabs[k + 1].StartRow > tabs[k].StartRow)
                                endIndexRow = tabs[k + 1].StartRow;
                            else if (k < tabs.Count - 2)
                            {
                                if (tabs[k + 2].StartRow > tabs[k].StartRow)
                                    endIndexRow = tabs[k + 2].StartRow;
                                else if (k < tabs.Count - 3)
                                {
                                    if (tabs[k + 3].StartRow > tabs[k].StartRow)
                                        endIndexRow = tabs[k + 3].StartRow;
                                }
                                else endIndexRow = cCelRow;
                            }
                            else endIndexRow = cCelRow;
                        }
                        else endIndexRow = cCelRow;
                        if (tabs[k].StartRow == endIndexRow) endIndexRow = cCelRow;
                        bool stop = false;
                        int firstCol = tabs[k].StartCol;
                        int curCol = tabs[k].StartCol;
                        string Standart = "";
                        while (!stop)
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].StartRow - 1, curCol];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"профиль", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    for (int jj = tabs[k].StartRow + 1; jj < endIndexRow; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, tabs[k].StartCol];
                                        if (cellRange.Value != null)
                                        {
                                            temp = cellRange.Value.ToString().Trim();
                                            bool isAddedRow = false;
                                            if (nameProd != "") //new Regex(@"(?!\w+ые|\w+ие|\w+ая|\w+\d|\d|ГОСТ)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase).IsMatch(temp) || 
                                            {
                                                dtProduct.Rows.Add(); isAddedRow = true;
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                dtProduct.Rows[lastRow]["Название"] = nameProd; // new Regex(@"(?!\w+ые|\w+ие|\w+ая|\w+\d|\d|ГОСТ)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "стали") dtProduct.Rows[lastRow]["Название"] = "Cталь";
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "трубы") dtProduct.Rows[lastRow]["Название"] = "Труба";
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "листы") dtProduct.Rows[lastRow]["Название"] = "Лист";
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "рельсы") dtProduct.Rows[lastRow]["Название"] = "Рельса";
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "двутавры") dtProduct.Rows[lastRow]["Название"] = "Двутавр";
                                                dtProduct.Rows[lastRow]["Тип"] = new Regex(@"(?<=^|\s)(?:\w+(?:ые|ие|ая|ый))|\bBP\b|\bВР\b|[AА]-*I{1,3}|х/к|г/к|Н/У\sТ/О\sчерная|ВГП|э/св|б/ш|оцинк(?=,|\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=^\d+(?:,\d+)[хx])\d+(?:,\d+)(?=\s|-|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Примечание"] = temp;
                                                dtProduct.Rows[lastRow]["Стандарт"] = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:Г[Оо][Сс][Тт]\s{0,3})(?:[рР]-\s?)?(?:\d{1,5}[-\s]*)+|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])").Match(temp).Value; //шаблон Стандарта
                                                if (dtProduct.Rows[lastRow]["Стандарт"].ToString() == "") dtProduct.Rows[lastRow]["Стандарт"] = Standart;
                                            }
                                            if (new Regex(@"\d+(?:[.,]\d+)?(?=\s*[бмшкпу]\d?)", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                if (!isAddedRow) dtProduct.Rows.Add();
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                tabs[k].listExcelIndexTab.Add(jj);
                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                //dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\d+(?:[,.]\d+)?[xх]\d+(?:[,.]\d+)?\s*-\s*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"\d+(?:[.,]\d+)?(?=\s*[бмшкпу]\d?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                //dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=\d+(?:[,.]\d+)?[xх])\d+(?:[,.]\d+)?(?=\s*-\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                            }
                                            else if (new Regex(@"\d+(?:[.,]\d+)?(?=мм)", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                if (!isAddedRow) dtProduct.Rows.Add();
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                tabs[k].listExcelIndexTab.Add(jj);
                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"\d+(?:[.,]\d+)?(?=мм)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                //dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\d+(?:[.,]\d+)?(?:ду)?\s*-\s*)\d\.\d", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                            }
                                            else if (new Regex(@"\d+(?:[.,]\d+)?(\s*[хx]\s*\d+(?:[.,]\d+)?)+", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                if (!isAddedRow) dtProduct.Rows.Add();
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                tabs[k].listExcelIndexTab.Add(jj);
                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=^|\s)\d+(?:[.,]\d+)?(?=[хx]\d+(?:[.,]\d+)?[хx])", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=[хx]\d+(?:[.,]\d+)?[хx])\d+(?:[.,]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=\d+(?:[.,]\d+)?[хx])\d+(?:[.,]\d+)?(?=[хx])", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                            }
                                            else if (new Regex(@"(?<=\b[KPКР]\b\s*)\d+(?:[.,]\d+)?(?=\s*г)", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                if (!isAddedRow) dtProduct.Rows.Add();
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                tabs[k].listExcelIndexTab.Add(jj);
                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=\b[KPКР]\b\s*)\d+(?:[.,]\d+)?(?=\s*г)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                //dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=[хx])\d+(?:[.,]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                            }
                                            else if (new Regex(@"(?<=\w+\s*)\d+(?:[.,]\d+)?(?=\s*г)", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                if (!isAddedRow) dtProduct.Rows.Add();
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                tabs[k].listExcelIndexTab.Add(jj);
                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=\w+\s*)\d+(?:[.,]\d+)?(?=\s*г)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                //dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=[хx])\d+(?:[.,]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                            }
                                            else if (new Regex(@"(?<=^)\d+(?:[.,]\d+)?(?=,\s|,\d\d|$)", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                if (!isAddedRow) dtProduct.Rows.Add();
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                tabs[k].listExcelIndexTab.Add(jj);
                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=^)\d+(?:[.,]\d+)?(?=,\s|,\d\d|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                //dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=[хx])\d+(?:[.,]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                            }
                                            else if (new Regex(@"(?<=\w+eeee\s*)\d+(?:[.,]\d+)?(?=\s*г)", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                if (!isAddedRow) dtProduct.Rows.Add();
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                string[] strSizes;
                                                strSizes = temp.Split('-');
                                                if (strSizes.Length == 2)
                                                {
                                                    double size1 = Convert.ToDouble(new Regex(@"\d+(?:[.,]\d*)?", RegexOptions.IgnoreCase).Match(strSizes[0]).Value);
                                                    double size2 = Convert.ToDouble(new Regex(@"\d+(?:[.,]\d*)?", RegexOptions.IgnoreCase).Match(strSizes[1]).Value);
                                                    for (int s = 0; s < DiamStd.Count; s++)
                                                    {
                                                        if (new Regex(@"\d+(?:[.,]\d*)?", RegexOptions.IgnoreCase).IsMatch(strSizes[0]))
                                                        {

                                                            if (DiamStd[s] >= size1 && DiamStd[s] <= size2)
                                                            {
                                                                //dtProduct.Rows.Add();
                                                                tabs[k].listExcelIndexTab.Add(jj);
                                                                lastRow = dtProduct.Rows.Count - 1;
                                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = DiamStd[s];
                                                                dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        else if (cellRange.MergeArea.Columns.Count > 1)
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[jj, tabs[k].StartCol - 1];
                                            if (cellRange.Value != null)
                                            {
                                                temp = cellRange.Value.ToString().Trim();
                                                nameProd = new Regex(@"^\w+", RegexOptions.IgnoreCase).Match(temp).Value;
                                                Standart = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:Г[Оо][Сс][Тт]\s{0,3})(?:[рР]-\s?)?(?:\d{1,5}[-\s]*)+|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])").Match(temp).Value; //шаблон Стандарта 
                                            }
                                        }
                                        else if (cellRange.MergeArea.Rows.Count == 2)
                                        {
                                            dtProduct.Rows.Add();
                                            int lastRow = dtProduct.Rows.Count - 1;
                                            dtProduct.Rows[lastRow]["Название"] = nameProd;
                                            dtProduct.Rows[lastRow]["Тип"] = dtProduct.Rows[lastRow - 1]["Тип"];
                                            dtProduct.Rows[lastRow]["Примечание"] = dtProduct.Rows[lastRow - 1]["Примечание"];
                                            dtProduct.Rows[lastRow]["Стандарт"] = dtProduct.Rows[lastRow - 1]["Стандарт"];
                                            dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = dtProduct.Rows[lastRow - 1]["Диаметр (высота), мм"];
                                            if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "двутавры") dtProduct.Rows[lastRow]["Название"] = "Двутавр";
                                            if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "листы") dtProduct.Rows[lastRow]["Название"] = "Лист";
                                            tabs[k].listExcelIndexTab.Add(jj);
                                            tabs[k].listdtProductIndexRow.Add(lastRow);
                                        }
                                    }

                                    //если название все еще не найдено, то добавить просто название Труба
                                    //if(dtProduct.Rows.Count>0)
                                    //if (dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"].ToString().Trim() == "")
                                    //    dtProduct.Rows[dtProduct.Rows.Count - 1]["Название"] = "Труба";
                                }

                            }
                            countIteration++;
                            if (curCol == firstCol && curCol > 1) curCol++;
                            else if (curCol > firstCol && curCol - firstCol < 4) curCol++;
                            else if (curCol == 1 && firstCol == 1) curCol++;
                            else stop = true;
                        }

                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        else tsPb1.Value = tsPb1.Maximum;
                    }

                    tsPb1.Value = 0;
                    tsPb1.Maximum = tabs.Count;
                    for (int k = 0; k < tabs.Count; k++)
                    {
                        for (int curCol = tabs[k].StartCol; curCol <= tabs[k].StartCol + 3; curCol++)
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].StartRow - 1, curCol];
                            if (cellRange.Value != null)
                            {
                                temp = cellRange.Value.ToString().Trim();
                                if (tabs[k].listExcelIndexTab.Count > 0) // сюда обработку других столбцов
                                {
                                    if (new Regex(@"марка", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        for (int i = 0; i < tabs[k].listExcelIndexTab.Count; i++) //строки
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[i], curCol];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Марка"] = temp;
                                            }
                                        }
                                    }


                                    if (new Regex(@"цена", RegexOptions.IgnoreCase).IsMatch(temp))  //типа цена
                                    {
                                        for (int i = 0; i < tabs[k].listExcelIndexTab.Count; i++) //строки
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[i], curCol];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Цена"] = new Regex(@"\d+", RegexOptions.IgnoreCase).Match(temp).Value;
                                                if (dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Цена"].ToString().Length == 1)
                                                {
                                                    cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[i], curCol + 1];
                                                    if (cellRange.Value != null)
                                                        temp = cellRange.Value.ToString().Trim();
                                                    else temp = "";
                                                    if (temp != "")
                                                    {
                                                        dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Цена"] = new Regex(@"\d+", RegexOptions.IgnoreCase).Match(temp).Value;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        else tsPb1.Value = tsPb1.Maximum;


                    }
                    //проверка строк для InfoOrganization
                    if (tabs.Count > 0)
                        for (int j = 1; j <= cCelRow; j++) //строки
                        {
                            for (int i = 1; i <= 18; i++) //столбцы
                            {
                                Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    if (j < tabs[0].StartRow || j > tabs[tabs.Count - 1].StartRow)
                                    {
                                        //InfoOrganization(temp);

                                        #region Склад
                                        Regex regAddr = new Regex(@"^адреса\s+", RegexOptions.IgnoreCase);
                                        if (regAddr.IsMatch(temp))
                                        {
                                            int jj = j + 1;
                                            Excel.Range cellrengeAddr = (Excel.Range)excelworksheet.Cells[jj, i];
                                            if (cellrengeAddr.Value != null)
                                            {
                                                Regex sklad = new Regex(@"челябинск.*", RegexOptions.IgnoreCase);
                                                string tmp = sklad.Match(cellrengeAddr.Value.ToString().Trim()).Value;
                                                while (tmp != "")
                                                {
                                                    ListViewItem lvi = new ListViewItem(tmp);
                                                    bool isIn = false;
                                                    if (listViewAdrSklad.Items.Count < 1) listViewAdrSklad.Items.Add(lvi);
                                                    for (int ii = 0; ii < listViewAdrSklad.Items.Count; ii++)
                                                    {
                                                        if (listViewAdrSklad.Items[ii].SubItems[0].Text == tmp) isIn = true;
                                                    }
                                                    if (!isIn) listViewAdrSklad.Items.Add(lvi);
                                                    jj++;
                                                    cellrengeAddr = (Excel.Range)excelworksheet.Cells[jj, i];
                                                    tmp = sklad.Match(cellrengeAddr.Value.ToString().Trim()).Value;
                                                }
                                            }
                                        }
                                        #endregion

                                        #region Менеджер
                                        Regex regManager = new Regex(@"(?<=^менеджер.*:\s*тел.*)(?:\(\d{3}\)\s*)?(?:-?\d{2,})+", RegexOptions.IgnoreCase);
                                        if (regManager.IsMatch(temp))
                                        {
                                            string tel = regManager.Match(temp).Value;
                                            int jj = j + 1;
                                            Excel.Range cellrengeManager = (Excel.Range)excelworksheet.Cells[jj, i];
                                            if (cellrengeManager.Value != null)
                                            {
                                                Regex sklad = new Regex(@"^(?:\s*\w+\s*){2,3}$", RegexOptions.IgnoreCase);
                                                string tmp = sklad.Match(cellrengeManager.Value.ToString().Trim()).Value;
                                                while (tmp != "")
                                                {
                                                    string nameMan = tmp;
                                                    ListViewItem lvi = new ListViewItem(tmp);
                                                    cellrengeManager = (Excel.Range)excelworksheet.Cells[jj, i + 3];
                                                    if (cellrengeManager.Value != null)
                                                    {
                                                        sklad = new Regex(@"доб.*\d{3,4}", RegexOptions.IgnoreCase);
                                                        tmp = tel + " " + sklad.Match(cellrengeManager.Value.ToString().Trim()).Value;
                                                        lvi.SubItems.Add(tmp);
                                                    }
                                                    bool isIn = false;
                                                    if (listViewManager.Items.Count < 1) listViewManager.Items.Add(lvi);
                                                    for (int ii = 0; ii < listViewManager.Items.Count; ii++)
                                                    {
                                                        if (listViewManager.Items[ii].SubItems[0].Text == nameMan) isIn = true;
                                                    }
                                                    if (!isIn) listViewManager.Items.Add(lvi);
                                                    jj++;
                                                    sklad = new Regex(@"^(?:\s*\w+\s*){2,3}$", RegexOptions.IgnoreCase);
                                                    cellrengeManager = (Excel.Range)excelworksheet.Cells[jj, i];
                                                    tmp = sklad.Match(cellrengeManager.Value.ToString().Trim()).Value;
                                                    if (tmp == "")
                                                    {
                                                        Excel.Range cellrengeManager1 = (Excel.Range)excelworksheet.Cells[jj, i + 1];
                                                        if (cellrengeManager1.Value != null)
                                                        {
                                                            tmp = sklad.Match(cellrengeManager1.Value.ToString().Trim()).Value;
                                                            if (tmp == "")
                                                            {
                                                                Excel.Range cellrengeManager2 = (Excel.Range)excelworksheet.Cells[jj, i + 2];
                                                                if (cellrengeManager2.Value != null)
                                                                {
                                                                    tmp = sklad.Match(cellrengeManager2.Value.ToString().Trim()).Value;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        #endregion

                                        #region Email
                                        Regex regEmail = new Regex(@"(?<=^прием.*)\w+@\w+\.\w{2,4}", RegexOptions.IgnoreCase);
                                        if (regEmail.IsMatch(temp))
                                        {
                                            textBoxOrgEmail.Text = regEmail.Match(temp).Value;
                                        }
                                        #endregion
                                    }
                                }
                            }
                        }
                }

                clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции Evraz\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        ///<summary> 
        ///<remarks> Открытие и чтение экселевского файла Металл-База </remarks>
        ///<param name="path" >путь к файлу</param>
        ///</summary>
        private void Metall_Baza(string path)
        {
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                SetNameFromName(filePath);
                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;
                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла Metall-Baza\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    if (excelworksheet.Name.ToLower() == "оглавление") continue;
                    #region Лист - лист
                    if (excelworksheet.Name.ToLower() == "лист")
                    {
                        tsLabelcurrSheet.Text = excelworksheet.Index.ToString();

                        int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                        int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                        tsPb1.Maximum = cCelRow;// * cCelCol ;

                        listIndexOfNotEmptyName = new List<int>();
                        colForName = 0;

                        tsLabelClearingTable.Text = "Поиск наименований";
                        tsPb1.Value = 0;
                        for (int j = 1; j <= cCelRow; j++) //строки
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, 1];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString();
                            else temp = "";
                            if (temp != "")
                            {
                                #region получение списка непустых строк

                                nameProd = new Regex(@"(?!\s*\w+\d\w+|\s*\w+ий|\s*\w+ой|\s*Г?ост\s)(?<=\s*)\w{3,}", RegexOptions.IgnoreCase).Match(temp).Value;
                                if (nameProd != "")
                                {
                                    dtProduct.Rows.Add();   // добавить строку в результирующую таблицу
                                    int lastRow = dtProduct.Rows.Count - 1; //запомнить индекс последней строки
                                    dtProduct.Rows[lastRow]["Название"] = nameProd; /*записать наименование вручную указанного названия 
                                                                         * в ячейку названия продукции в результирующей таблице*/
                                    listIndexOfNotEmptyName.Add(j);     //добавить в список индексов непустых значений индекс текущей строки
                                    int c = listIndexOfNotEmptyName[0] + countEmpty;    //запомнить в переменную индекс первого значения плюс количество пустых ячеек
                                    //если список непустых значений не пустой
                                    if (listIndexOfNotEmptyName.Count > 0)
                                    {
                                        //то если список сдвигов индексов пустой
                                        if (listShiftIndex.Count < 1)
                                            // то занести в список сдвигов индексов индекс певого значения плюс количество пустых ячеек
                                            listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                                        /*если список сдвигов содержит больше 2х записей и индекс текущей строки меньше чем 
                                          индекс строки в списке непустых значений в предпоследней записи */
                                        else if (listIndexOfNotEmptyName.Count > 2 && j < listIndexOfNotEmptyName[listIndexOfNotEmptyName.Count - 2])
                                        {
                                            countEmpty = 0; //сброс счета количества пустых ячеек
                                            //занести в список сдвигов индексов индекс певого значения плюс количество пустых ячеек
                                            listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                                            //количество строк для сдвига равно текущему количеству строк в результирующей таблице
                                            countRowsForShift = dtProduct.Rows.Count - 1;
                                        }
                                        else listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                                    }
                                }
                                else if (listIndexOfNotEmptyName.Count > 0) countEmpty++;
                                ManualStringNameProd = nameProd;

                                #endregion
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        }
                        tsLabelClearingTable.Text = "Поиск диаметров по имени";
                        tsPb1.Value = 0;
                        for (int i = 0; i < listIndexOfNotEmptyName.Count; i++) //строки
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[i], 1];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString();
                            else temp = "";
                            if (temp != "")
                            {
                                Regex regType = new Regex(@"(?<=\s*)\w+ий|\w+ой|\w+ый", RegexOptions.IgnoreCase);
                                Regex regTolshList = new Regex(@"(?<=^\s?\w+\s*)\d+(?:,\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                Regex regShirina = new Regex(@"(?<=\s|\))\d{3,}(?=[xх]\s?\d)|(?<=\s|\)|\()\d+[,\.]\d+(?=[xх]\d)", RegexOptions.IgnoreCase);
                                Regex regDlina = new Regex(@"(?<=\d[xх]\s?)\d{3,}|(?<=\d[xх])\d+,\d+", RegexOptions.IgnoreCase);

                                #region выделение подстроки типа листа

                                if (regType.IsMatch(temp))
                                {
                                    dtProduct.Rows[i]["Тип"] = regType.Match(temp);
                                }

                                #endregion

                                #region выделение подстроки высоты(толщины) листа
                                if (regTolshList.IsMatch(temp))
                                {
                                    dtProduct.Rows[i]["Диаметр (высота), мм"] = regTolshList.Match(temp);
                                }
                                #endregion

                                #region выделение подстроки ширины листа
                                if (regShirina.IsMatch(temp))
                                {
                                    dtProduct.Rows[i]["Толщина (ширина), мм"] = regShirina.Match(temp);
                                }
                                #endregion

                                #region выделение подстроки длины листа
                                if (regDlina.IsMatch(temp))
                                {
                                    dtProduct.Rows[i]["Метраж, м (длина, мм)"] = regDlina.Match(temp);
                                }
                                #endregion

                                GetRegexTUFromString(temp, i);
                                GetRegexMarkFromString(temp, i);
                            }
                            dtProduct.Rows[i]["Примечание"] = temp;
                            cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[i], 5];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString();
                            else temp = "";
                            if (temp != "")
                            {
                                GetRegexPriceFromString(temp, i);
                            }

                            cellRange = (Excel.Range)excelworksheet.Cells[listIndexOfNotEmptyName[i], 2];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString();
                            else temp = "";
                            if (temp != "" && dtProduct.Rows[i]["Название"].ToString().ToLower() == "тигли")
                            {
                                dtProduct.Rows[i]["Тип"] = temp;
                            }

                            if (dtProduct.Rows[i]["Диаметр (высота), мм"].ToString() != "" || dtProduct.Rows[i]["Толщина (ширина), мм"].ToString() != "")
                            {
                                dtProduct.Rows[i]["Тип"] = "тип не указан";
                            }
                        }
                    }
                    #endregion

                    #region Лист - Сорт
                    if (excelworksheet.Name.ToLower() == "сорт")
                    {
                        structTab tab = new structTab
                        {
                            StartCol = 1,
                            StartRow = 1,
                            listExcelIndexTab = new List<int>(),
                            listdtProductIndexRow = new List<int>()
                        };

                        tsLabelcurrSheet.Text = excelworksheet.Index.ToString();

                        int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                        int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                        tsPb1.Maximum = cCelRow;// * cCelCol ;

                        listIndexOfNotEmptyName = new List<int>();
                        colForName = 0;

                        tsLabelClearingTable.Text = "Поиск наименований";
                        tsPb1.Value = 0;
                        for (int j = 1; j <= cCelRow; j++) //строки
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, 1];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString();
                            else temp = "";
                            if (temp != "")
                            {
                                #region получение списка непустых строк

                                nameProd = new Regex(@"(?!\s*\w+\d\w+|\s*\w+ий|\s*\w+ой|\s*Г?ост\s)(?<=\s*)\w{3,}", RegexOptions.IgnoreCase).Match(temp).Value;
                                if (nameProd != "")
                                {
                                    dtProduct.Rows.Add();   // добавить строку в результирующую таблицу
                                    int lastRow = dtProduct.Rows.Count - 1; // индекс последней строки
                                    dtProduct.Rows[lastRow]["Название"] = nameProd;
                                    tab.listdtProductIndexRow.Add(lastRow);
                                    tab.listExcelIndexTab.Add(j);
                                }

                                ManualStringNameProd = nameProd;

                                #endregion
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        }
                        tsLabelClearingTable.Text = "Поиск параметров по имени";
                        tsPb1.Value = 0;
                        for (int i = 0; i < tab.listExcelIndexTab.Count; i++) //строки
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[tab.listExcelIndexTab[i], 1];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString();
                            else temp = "";
                            if (temp != "")
                            {
                                Regex regType = new Regex(@"(?<=\s*)\w+ий|\w+ой|\w+ый", RegexOptions.IgnoreCase);


                                #region выделение подстроки типа листа

                                if (regType.IsMatch(temp))
                                {
                                    dtProduct.Rows[tab.listdtProductIndexRow[i]]["Тип"] = regType.Match(temp);
                                }
                                else dtProduct.Rows[i]["Тип"] = "тип не указан";

                                #endregion

                                #region уголок
                                if (dtProduct.Rows[tab.listdtProductIndexRow[i]]["Название"].ToString().ToLower() == "уголок")
                                {
                                    Regex regTolshList = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                    Regex regShirina = new Regex(@"(?<=\s\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                    Regex regDlina = new Regex(@"\s\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);

                                    #region выделение подстроки высоты(толщины) листа
                                    if (regDlina.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Диаметр (высота), мм"] = regDlina.Match(temp);
                                    }
                                    #endregion

                                    #region выделение подстроки ширины листа
                                    if (regTolshList.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                    }
                                    #endregion

                                    #region выделение подстроки длины листа
                                    if (regShirina.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Метраж, м (длина, мм)"] = regShirina.Match(temp);
                                    }
                                    #endregion
                                }
                                #endregion

                                #region круг
                                if (dtProduct.Rows[tab.listdtProductIndexRow[i]]["Название"].ToString().ToLower() == "круг")
                                {
                                    Regex regDiam = new Regex(@"(?<=круг\s*)\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);
                                    //Regex regShirina = new Regex(@"(?<=\s\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                    Regex regDlina = new Regex(@"(?<=дл\.\s*)\d+(?:[,\.]\d+)?(?=м)", RegexOptions.IgnoreCase);

                                    #region выделение подстроки метража
                                    if (regDlina.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Метраж, м (длина, мм)"] = regDlina.Match(temp);
                                    }
                                    #endregion

                                    #region выделение подстроки диаметра
                                    if (regDiam.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Диаметр (высота), мм"] = regDiam.Match(temp);
                                    }
                                    #endregion

                                }
                                #endregion

                                #region Арматура
                                if (dtProduct.Rows[tab.listdtProductIndexRow[i]]["Название"].ToString().ToLower() == "арматура")
                                {
                                    Regex regDiam = new Regex(@"(?<=арматура\s*)\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);
                                    //Regex regShirina = new Regex(@"(?<=\s\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                    Regex regDlina = new Regex(@"(?<=(?:дл\.)\s*)\d+(?:[,\.]\d+)?(?=м)", RegexOptions.IgnoreCase);
                                    Regex regTypeArm = new Regex(@"(?<=\().+(?=\))", RegexOptions.IgnoreCase);

                                    #region выделение подстроки высоты(толщины)
                                    if (regDlina.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Метраж, м (длина, мм)"] = regDlina.Match(temp);
                                    }
                                    #endregion

                                    #region выделение подстроки ширины
                                    if (regDiam.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Диаметр (высота), мм"] = regDiam.Match(temp);
                                    }
                                    #endregion

                                    #region выделение подстроки типа
                                    if (regTypeArm.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Тип"] = regTypeArm.Match(temp);
                                    }
                                    #endregion

                                }
                                #endregion

                                #region Швеллер
                                if (dtProduct.Rows[tab.listdtProductIndexRow[i]]["Название"].ToString().ToLower() == "швеллер")
                                {
                                    Regex regDiam = new Regex(@"(?<=швеллер\s*)\d+(?:[,\.]\d+)?(?=у|п|м)|(?<=швеллер\s*)\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                    Regex regTolshList = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                    Regex regDlina = new Regex(@"(?<=\s\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                    Regex regTypeArm = new Regex(@"(?<=швеллер\s*\d+(?:[,\.]\d+)?)[упм]", RegexOptions.IgnoreCase);

                                    #region выделение подстроки высоты(толщины)
                                    if (regDlina.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Метраж, м (длина, мм)"] = regDlina.Match(temp);
                                    }
                                    #endregion

                                    #region выделение подстроки длины листа
                                    if (regTolshList.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                    }
                                    #endregion

                                    #region выделение подстроки ширины
                                    if (regDiam.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Диаметр (высота), мм"] = regDiam.Match(temp);
                                    }
                                    #endregion

                                    #region выделение подстроки типа
                                    if (regTypeArm.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Тип"] = regTypeArm.Match(temp);
                                    }
                                    #endregion

                                }
                                #endregion

                                #region квадрат
                                if (dtProduct.Rows[tab.listdtProductIndexRow[i]]["Название"].ToString().ToLower() == "квадрат")
                                {
                                    Regex regDiam = new Regex(@"(?<=квадрат\s*)\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);
                                    //Regex regShirina = new Regex(@"(?<=\s\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                    Regex regDlina = new Regex(@"(?<=дл\.?\s*)\d+(?:[,\.]\d+)?(?=м)", RegexOptions.IgnoreCase);

                                    #region выделение подстроки метража
                                    if (regDlina.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Метраж, м (длина, мм)"] = regDlina.Match(temp);
                                    }
                                    #endregion

                                    #region выделение подстроки диаметра
                                    if (regDiam.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Диаметр (высота), мм"] = regDiam.Match(temp);
                                    }
                                    #endregion

                                }
                                #endregion

                                #region балка
                                if (dtProduct.Rows[tab.listdtProductIndexRow[i]]["Название"].ToString().ToLower() == "балка")
                                {
                                    Regex regDiam = new Regex(@"(?<=балка\s*)\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);
                                    //Regex regShirina = new Regex(@"(?<=\s\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                    Regex regDlina = new Regex(@"(?<=дл\.?\s*)\d+(?:[,\.]\d+)?(?=м)", RegexOptions.IgnoreCase);
                                    Regex regTypeBalka = new Regex(@"(?<=балка\s*\d+(?:[,\.]\d+)?)\w\d", RegexOptions.IgnoreCase);

                                    #region выделение подстроки метража
                                    if (regDlina.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Метраж, м (длина, мм)"] = regDlina.Match(temp);
                                    }
                                    #endregion

                                    #region выделение подстроки диаметра
                                    if (regDiam.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Диаметр (высота), мм"] = regDiam.Match(temp);
                                    }
                                    #endregion

                                    #region выделение подстроки типа
                                    if (regTypeBalka.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Тип"] = regTypeBalka.Match(temp);
                                    }
                                    #endregion

                                }
                                #endregion

                                #region полоса
                                if (dtProduct.Rows[tab.listdtProductIndexRow[i]]["Название"].ToString().ToLower() == "полоса")
                                {
                                    Regex regDiam = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                    //Regex regShirina = new Regex(@"(?<=\s\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                    Regex regTolshList = new Regex(@"\s\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                    Regex regDlina = new Regex(@"(?<=(?:дл\.=?)\s*)\d+(?:[,\.]\d+)?(?=м)", RegexOptions.IgnoreCase);
                                    Regex regTypePolos = new Regex(@"оцинк", RegexOptions.IgnoreCase);

                                    #region выделение подстроки высоты
                                    if (regDiam.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Диаметр (высота), мм"] = regDiam.Match(temp);
                                    }
                                    #endregion

                                    #region выделение подстроки ширины 
                                    if (regTolshList.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                    }
                                    #endregion

                                    #region выделение подстроки высоты(толщины)
                                    if (regDlina.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Метраж, м (длина, мм)"] = regDlina.Match(temp);
                                    }
                                    #endregion

                                    #region выделение подстроки ширины 
                                    if (regTypePolos.IsMatch(temp))
                                    {
                                        dtProduct.Rows[tab.listdtProductIndexRow[i]]["Тип"] = "Оцинкованная";
                                    }
                                    #endregion
                                }
                                #endregion

                                GetRegexTUFromString(temp, tab.listdtProductIndexRow[i]);
                                GetRegexMarkFromString(temp, tab.listdtProductIndexRow[i]);
                            }
                            dtProduct.Rows[tab.listdtProductIndexRow[i]]["Примечание"] = temp;

                            cellRange = (Excel.Range)excelworksheet.Cells[tab.listExcelIndexTab[i], 4];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString();
                            else temp = "";
                            if (temp != "" && temp != "0")
                            {
                                GetRegexPriceFromString(temp, tab.listdtProductIndexRow[i]);
                            }
                            else
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[tab.listExcelIndexTab[i], 5];
                                if (cellRange.Value != null)
                                {
                                    temp = cellRange.Value.ToString();
                                    if (temp != "")
                                    {
                                        GetRegexPriceFromString(temp, tab.listdtProductIndexRow[i]);
                                    }
                                }
                            }

                            cellRange = (Excel.Range)excelworksheet.Cells[tab.listExcelIndexTab[i], 2];
                            if (cellRange.Value != null)
                            {
                                temp = cellRange.Value.ToString();
                                if (temp != "")
                                {
                                    dtProduct.Rows[tab.listdtProductIndexRow[i]]["Мерность (т, м, мм)"] = temp;
                                }
                            }
                        }
                    }
                    #endregion

                    #region Лист - Трубы
                    if (excelworksheet.Name.ToLower() == "трубы")
                    {
                        List<structTab> tabs = new List<structTab>();

                        tsLabelcurrSheet.Text = excelworksheet.Index.ToString();

                        int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                        int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                        tsPb1.Maximum = cCelRow;// * cCelCol ;

                        listIndexOfNotEmptyName = new List<int>();
                        colForName = 0;

                        tsLabelClearingTable.Text = "Поиск наименований";
                        tsPb1.Value = 0;
                        for (int j = 1; j <= cCelRow; j++) //строки
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, 1];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString();
                            else temp = "";
                            if (temp != "" && cellRange.MergeArea.Count > 1)
                            {
                                #region получение списка непустых строк

                                nameProd = new Regex(@"труб\w|угольник\w?|фитинг|отвод\w|крест\w?", RegexOptions.IgnoreCase).Match(temp).Value;
                                Regex regType = new Regex(@"\w+ые|\w+ая", RegexOptions.IgnoreCase);
                                structTab tab = new structTab();
                                if (nameProd != "")
                                {
                                    tab.StartCol = 1;
                                    tab.StartRow = j;
                                    tab.listExcelIndexTab = new List<int>();
                                    tab.listdtProductIndexRow = new List<int>();
                                    if (new Regex(@"труб", RegexOptions.IgnoreCase).IsMatch(nameProd)) nameProd = "Труба";
                                    if (new Regex(@"угольник", RegexOptions.IgnoreCase).IsMatch(nameProd)) nameProd = "Угольник";
                                    tab.Name = nameProd;
                                    if (regType.IsMatch(temp))
                                        tab.Type = regType.Match(temp).Value;
                                    else tab.Type = "тип не указан";
                                    tab.Standart = new Regex(@"ГОСТ\s*\d+(?:-\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                    tabs.Add(tab);

                                }

                                #endregion
                            }
                            else if (temp != "")
                            {
                                nameProd = new Regex(@"муфт\w|переход|тройник|резьб\w|сгон\w|заглушки|крест\w?", RegexOptions.IgnoreCase).Match(temp).Value;
                                Regex regType = new Regex(@"\w+ые|\w+ая", RegexOptions.IgnoreCase);
                                structTab tab = new structTab();
                                if (nameProd != "")
                                {
                                    tab.StartCol = 1;
                                    tab.StartRow = j;
                                    tab.listExcelIndexTab = new List<int>();
                                    tab.listdtProductIndexRow = new List<int>();
                                    tab.Name = nameProd;
                                    if (regType.IsMatch(temp))
                                        tab.Type = regType.Match(temp).Value;
                                    else tab.Type = "тип не указан";
                                    tab.Standart = new Regex(@"ГОСТ\s*\d+(?:-\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                    tabs.Add(tab);

                                }
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        }
                        tsLabelClearingTable.Text = "Поиск параметров по имени";
                        tsPb1.Value = 0;

                        for (int k = 0; k < tabs.Count; k++)
                        {
                            structTab stab = tabs[k];
                            int endRowForCurTab = 1;
                            if (k < tabs.Count - 2)
                            { endRowForCurTab = tabs[k + 1].StartRow; }
                            else { endRowForCurTab = cCelRow; }

                            for (int i = stab.StartRow + 1; i < endRowForCurTab; i++) //строки
                            {

                                Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[i, 1];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString();
                                else temp = "";
                                if (temp != "")
                                {
                                    #region труба

                                    if (stab.Name.ToLower() == "труба" || stab.Name.ToLower() == "трубы")
                                    {

                                        #region из трех параметров
                                        if (new Regex(@"\d+(?:[,\.]\d+)?[xх]\d+(?:[,\.]\d+)?[xх]\d", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            dtProduct.Rows.Add();
                                            int lastRow = dtProduct.Rows.Count - 1;
                                            dtProduct.Rows[lastRow]["Название"] = stab.Name;
                                            dtProduct.Rows[lastRow]["Тип"] = stab.Type;

                                            dtProduct.Rows[lastRow]["Стандарт"] = stab.Standart;

                                            Regex regTolshList = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                            Regex regShirina = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                            Regex regDlina = new Regex(@"\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);

                                            #region выделение подстроки Диаметра
                                            if (regDlina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDlina.Match(temp);
                                            }
                                            #endregion

                                            #region выделение подстроки толщины
                                            if (regTolshList.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                            }
                                            #endregion

                                            #region выделение подстроки длины
                                            if (regShirina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = regShirina.Match(temp);
                                            }
                                            #endregion


                                            //GetRegexTUFromString(temp, lastRow);
                                            //GetRegexMarkFromString(temp, lastRow);

                                            dtProduct.Rows[lastRow]["Примечание"] = temp;

                                            //поиск цены
                                            cellRange = (Excel.Range)excelworksheet.Cells[i, 4];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                GetRegexPriceFromString(temp, lastRow);
                                            }
                                            else //если цена не найдена в этом столбце, то поискать в следующем
                                            {
                                                cellRange = (Excel.Range)excelworksheet.Cells[i, 5];
                                                if (cellRange.Value != null)
                                                {
                                                    temp = cellRange.Value.ToString();
                                                    if (temp != "")
                                                    {
                                                        GetRegexPriceFromString(temp, lastRow);
                                                    }
                                                }
                                            }

                                            cellRange = (Excel.Range)excelworksheet.Cells[i, 2];
                                            if (cellRange.Value != null)
                                            {
                                                temp = cellRange.Value.ToString();
                                                if (temp != "")
                                                {
                                                    dtProduct.Rows[lastRow]["Мерность (т, м, мм)"] = temp;
                                                }
                                            }
                                        }
                                        #endregion 

                                        #region из двух параметров
                                        if (new Regex(@"\d+(?:[,\.]\d+)?[xх]\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            dtProduct.Rows.Add();
                                            int lastRow = dtProduct.Rows.Count - 1;
                                            dtProduct.Rows[lastRow]["Название"] = stab.Name;
                                            dtProduct.Rows[lastRow]["Тип"] = stab.Type;
                                            dtProduct.Rows[lastRow]["Стандарт"] = stab.Standart;

                                            Regex regTolshList = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                            //Regex regShirina = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                            Regex regDlina = new Regex(@"\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);

                                            #region выделение подстроки Диаметра
                                            if (regDlina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDlina.Match(temp);
                                            }
                                            #endregion

                                            #region выделение подстроки толщины
                                            if (regTolshList.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                            }
                                            #endregion


                                            //GetRegexTUFromString(temp, lastRow);
                                            //GetRegexMarkFromString(temp, lastRow);

                                            dtProduct.Rows[lastRow]["Примечание"] = temp;

                                            //поиск цены
                                            cellRange = (Excel.Range)excelworksheet.Cells[i, 4];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                GetRegexPriceFromString(temp, lastRow);
                                            }
                                            else //если цена не найдена в этом столбце, то поискать в следующем
                                            {
                                                cellRange = (Excel.Range)excelworksheet.Cells[i, 5];
                                                if (cellRange.Value != null)
                                                {
                                                    temp = cellRange.Value.ToString();
                                                    if (temp != "")
                                                    {
                                                        GetRegexPriceFromString(temp, lastRow);
                                                    }
                                                }
                                            }

                                            cellRange = (Excel.Range)excelworksheet.Cells[i, 2];
                                            if (cellRange.Value != null)
                                            {
                                                temp = cellRange.Value.ToString();
                                                if (temp != "")
                                                {
                                                    dtProduct.Rows[lastRow]["Мерность (т, м, мм)"] = temp;
                                                }
                                            }
                                        }
                                        #endregion 

                                    }
                                    #endregion

                                    #region отводы

                                    else if (stab.Name.ToLower() == "фитинг")// || stab.Name.ToLower() == "трубы")
                                    {

                                        #region из двух параметров
                                        if (new Regex(@"\d+(?:[,\.]\d+)?[xх]\d+(?:[,\.]\d+)?|d?\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            dtProduct.Rows.Add();
                                            int lastRow = dtProduct.Rows.Count - 1;
                                            dtProduct.Rows[lastRow]["Название"] = new Regex(@"отвод\w|угольн\w+|", RegexOptions.IgnoreCase).Match(temp).Value;
                                            dtProduct.Rows[lastRow]["Тип"] = new Regex(@"\w+ый|\w+ые|\w+ой", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if (stab.Standart == "")
                                                stab.Standart = new Regex(@"ГОСТ\s*\d+(?:-\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                            dtProduct.Rows[lastRow]["Стандарт"] = stab.Standart;

                                            Regex regTolshList = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                            //Regex regShirina = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                            Regex regDlina = new Regex(@"\d+(?:[,\.]\d+)?(?=[xх]\d)|(?<=d?)\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);

                                            #region выделение подстроки Диаметра
                                            if (regDlina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDlina.Match(temp);
                                            }
                                            #endregion

                                            #region выделение подстроки толщины
                                            if (regTolshList.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                            }
                                            #endregion


                                            //GetRegexTUFromString(temp, lastRow);
                                            //GetRegexMarkFromString(temp, lastRow);

                                            dtProduct.Rows[lastRow]["Примечание"] = temp;

                                            //поиск цены
                                            cellRange = (Excel.Range)excelworksheet.Cells[i, 3];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                GetRegexPriceFromString(temp, lastRow);
                                            }

                                        }
                                        #endregion 
                                    }
                                    #endregion

                                    #region угольник
                                    else if (stab.Name.ToLower() == "угольники")// || stab.Name.ToLower() == "трубы")
                                    {

                                        #region из двух параметров
                                        if (new Regex(@"\d+(?:[,\.]\d+)?[xх]\d+(?:[,\.]\d+)?|d?\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            dtProduct.Rows.Add();
                                            int lastRow = dtProduct.Rows.Count - 1;
                                            dtProduct.Rows[lastRow]["Название"] = new Regex(@"угольн\w+", RegexOptions.IgnoreCase).Match(temp).Value;
                                            dtProduct.Rows[lastRow]["Тип"] = new Regex(@"\w+ый|\w+ые|\w+ой", RegexOptions.IgnoreCase).Match(temp).Value;
                                            dtProduct.Rows[lastRow]["Стандарт"] = stab.Standart;

                                            Regex regTolshList = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                            //Regex regShirina = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                            Regex regDlina = new Regex(@"\d+(?:[,\.]\d+)?(?=[xх]\d)|(?<=d?)\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);

                                            #region выделение подстроки Диаметра
                                            if (regDlina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDlina.Match(temp);
                                            }
                                            #endregion

                                            #region выделение подстроки толщины
                                            if (regTolshList.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                            }
                                            #endregion


                                            //GetRegexTUFromString(temp, lastRow);
                                            //GetRegexMarkFromString(temp, lastRow);

                                            dtProduct.Rows[lastRow]["Примечание"] = temp;

                                            //поиск цены
                                            cellRange = (Excel.Range)excelworksheet.Cells[i, 3];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                GetRegexPriceFromString(temp, lastRow);
                                            }

                                        }
                                        #endregion 
                                    }
                                    #endregion

                                    #region муфта
                                    else if (stab.Name.ToLower() == "муфта")// || stab.Name.ToLower() == "трубы")
                                    {

                                        #region из двух параметров
                                        if (new Regex(@"\d+(?:[,\.]\d+)?[xх]\d+(?:[,\.]\d+)?|d?\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            dtProduct.Rows.Add();
                                            int lastRow = dtProduct.Rows.Count - 1;
                                            dtProduct.Rows[lastRow]["Название"] = stab.Name;
                                            dtProduct.Rows[lastRow]["Стандарт"] = stab.Standart;

                                            Regex regTypeMuft = new Regex(@"\w+ая|\w+ые", RegexOptions.IgnoreCase);
                                            if (regTypeMuft.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Тип"] = regTypeMuft.Match(temp).Value;
                                            }
                                            else
                                                dtProduct.Rows[lastRow]["Тип"] = stab.Type;

                                            Regex regTolshList = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                            //Regex regShirina = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                            Regex regDlina = new Regex(@"\d+(?:[,\.]\d+)?(?=[xх]\d)|(?<=d?)\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);

                                            #region выделение подстроки Диаметра
                                            if (regDlina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDlina.Match(temp);
                                            }
                                            #endregion

                                            #region выделение подстроки толщины
                                            if (regTolshList.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                            }
                                            #endregion


                                            //GetRegexTUFromString(temp, lastRow);
                                            //GetRegexMarkFromString(temp, lastRow);

                                            dtProduct.Rows[lastRow]["Примечание"] = temp;

                                            //поиск цены
                                            cellRange = (Excel.Range)excelworksheet.Cells[i, 3];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                GetRegexPriceFromString(temp, lastRow);
                                            }

                                        }
                                        #endregion
                                    }
                                    #endregion

                                    #region переход
                                    else if (stab.Name.ToLower() == "переход")// || stab.Name.ToLower() == "трубы")
                                    {

                                        #region из двух параметров
                                        if (new Regex(@"\d+(?:[,\.]\d+)?[xх]\d+(?:[,\.]\d+)?|d?\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            dtProduct.Rows.Add();
                                            int lastRow = dtProduct.Rows.Count - 1;
                                            dtProduct.Rows[lastRow]["Название"] = stab.Name;
                                            dtProduct.Rows[lastRow]["Стандарт"] = stab.Standart;

                                            Regex regTypeMuft = new Regex(@"\w+ая|\w+ые", RegexOptions.IgnoreCase);
                                            if (regTypeMuft.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Тип"] = regTypeMuft.Match(temp).Value;
                                            }
                                            else
                                                dtProduct.Rows[lastRow]["Тип"] = stab.Type;

                                            Regex regTolshList = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                            //Regex regShirina = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                            Regex regDlina = new Regex(@"\d+(?:[,\.]\d+)?(?=[xх]\d)|(?<=d?)\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);

                                            #region выделение подстроки Диаметра
                                            if (regDlina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDlina.Match(temp);
                                            }
                                            #endregion

                                            #region выделение подстроки толщины
                                            if (regTolshList.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                            }
                                            #endregion


                                            //GetRegexTUFromString(temp, lastRow);
                                            //GetRegexMarkFromString(temp, lastRow);

                                            dtProduct.Rows[lastRow]["Примечание"] = temp;

                                            //поиск цены
                                            cellRange = (Excel.Range)excelworksheet.Cells[i, 3];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                GetRegexPriceFromString(temp, lastRow);
                                            }

                                        }
                                        #endregion
                                    }
                                    #endregion

                                    #region тройник
                                    else if (stab.Name.ToLower() == "тройник")// || stab.Name.ToLower() == "трубы")
                                    {

                                        #region из двух параметров
                                        if (new Regex(@"\d+(?:[,\.]\d+)?[xх]\d+(?:[,\.]\d+)?|d?\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            dtProduct.Rows.Add();
                                            int lastRow = dtProduct.Rows.Count - 1;
                                            dtProduct.Rows[lastRow]["Название"] = stab.Name;
                                            dtProduct.Rows[lastRow]["Стандарт"] = stab.Standart;

                                            Regex regTypeMuft = new Regex(@"\w+ой|\w+ый", RegexOptions.IgnoreCase);
                                            if (regTypeMuft.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Тип"] = regTypeMuft.Match(temp).Value;
                                            }
                                            else
                                                dtProduct.Rows[lastRow]["Тип"] = stab.Type;

                                            Regex regTolshList = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                            //Regex regShirina = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                            Regex regDlina = new Regex(@"\d+(?:[,\.]\d+)?(?=[xх]\d)|(?<=d?)\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);

                                            #region выделение подстроки Диаметра
                                            if (regDlina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDlina.Match(temp);
                                            }
                                            #endregion

                                            #region выделение подстроки толщины
                                            if (regTolshList.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                            }
                                            #endregion


                                            //GetRegexTUFromString(temp, lastRow);
                                            //GetRegexMarkFromString(temp, lastRow);

                                            dtProduct.Rows[lastRow]["Примечание"] = temp;

                                            //поиск цены
                                            cellRange = (Excel.Range)excelworksheet.Cells[i, 3];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                GetRegexPriceFromString(temp, lastRow);
                                            }

                                        }
                                        #endregion
                                    }
                                    #endregion

                                    #region резьба
                                    else if (stab.Name.ToLower() == "резьба")// || stab.Name.ToLower() == "трубы")
                                    {

                                        #region из двух параметров
                                        if (new Regex(@"\d+(?:[,\.]\d+)?[xх]\d+(?:[,\.]\d+)?|d?\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            dtProduct.Rows.Add();
                                            int lastRow = dtProduct.Rows.Count - 1;
                                            dtProduct.Rows[lastRow]["Название"] = stab.Name;
                                            dtProduct.Rows[lastRow]["Стандарт"] = stab.Standart;

                                            Regex regTypeMuft = new Regex(@"\w+ой|\w+ый", RegexOptions.IgnoreCase);
                                            if (regTypeMuft.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Тип"] = regTypeMuft.Match(temp).Value;
                                            }
                                            else
                                                dtProduct.Rows[lastRow]["Тип"] = stab.Type;

                                            Regex regTolshList = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                            //Regex regShirina = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                            Regex regDlina = new Regex(@"\d+(?:[,\.]\d+)?(?=[xх]\d)|(?<=d?)\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);

                                            #region выделение подстроки Диаметра
                                            if (regDlina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDlina.Match(temp);
                                            }
                                            #endregion

                                            #region выделение подстроки толщины
                                            if (regTolshList.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                            }
                                            #endregion


                                            //GetRegexTUFromString(temp, lastRow);
                                            //GetRegexMarkFromString(temp, lastRow);

                                            dtProduct.Rows[lastRow]["Примечание"] = temp;

                                            //поиск цены
                                            cellRange = (Excel.Range)excelworksheet.Cells[i, 3];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                GetRegexPriceFromString(temp, lastRow);
                                            }

                                        }
                                        #endregion
                                    }
                                    #endregion

                                    #region Сгоны
                                    else if (stab.Name.ToLower() == "сгоны")// || stab.Name.ToLower() == "трубы")
                                    {

                                        #region из двух параметров
                                        if (new Regex(@"\d+(?:[,\.]\d+)?[xх]\d+(?:[,\.]\d+)?|d?\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            dtProduct.Rows.Add();
                                            int lastRow = dtProduct.Rows.Count - 1;
                                            dtProduct.Rows[lastRow]["Название"] = stab.Name;
                                            dtProduct.Rows[lastRow]["Стандарт"] = stab.Standart;

                                            Regex regTypeMuft = new Regex(@"\w+ой|\w+ые", RegexOptions.IgnoreCase);
                                            if (regTypeMuft.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Тип"] = regTypeMuft.Match(temp).Value;
                                            }
                                            else
                                                dtProduct.Rows[lastRow]["Тип"] = stab.Type;

                                            Regex regTolshList = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                            //Regex regShirina = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                            Regex regDlina = new Regex(@"\d+(?:[,\.]\d+)?(?=[xх]\d)|(?<=d?)\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);

                                            #region выделение подстроки Диаметра
                                            if (regDlina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDlina.Match(temp);
                                            }
                                            #endregion

                                            #region выделение подстроки толщины
                                            if (regTolshList.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                            }
                                            #endregion


                                            //GetRegexTUFromString(temp, lastRow);
                                            //GetRegexMarkFromString(temp, lastRow);

                                            dtProduct.Rows[lastRow]["Примечание"] = temp;

                                            //поиск цены
                                            cellRange = (Excel.Range)excelworksheet.Cells[i, 3];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                GetRegexPriceFromString(temp, lastRow);
                                            }

                                        }
                                        #endregion
                                    }
                                    #endregion

                                    #region Заглушки
                                    else if (stab.Name.ToLower() == "заглушки")// || stab.Name.ToLower() == "трубы")
                                    {

                                        #region из двух параметров
                                        if (new Regex(@"\d+(?:[,\.]\d+)?[xх]\d+(?:[,\.]\d+)?|d?\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            dtProduct.Rows.Add();
                                            int lastRow = dtProduct.Rows.Count - 1;
                                            dtProduct.Rows[lastRow]["Название"] = stab.Name;
                                            dtProduct.Rows[lastRow]["Стандарт"] = stab.Standart;

                                            Regex regTypeMuft = new Regex(@"\w+ой|\w+ые", RegexOptions.IgnoreCase);
                                            if (regTypeMuft.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Тип"] = regTypeMuft.Match(temp).Value;
                                            }
                                            else
                                                dtProduct.Rows[lastRow]["Тип"] = stab.Type;

                                            Regex regTolshList = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                            //Regex regShirina = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                            Regex regDlina = new Regex(@"\d+(?:[,\.]\d+)?(?=[xх]\d)|(?<=d?)\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);

                                            #region выделение подстроки Диаметра
                                            if (regDlina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDlina.Match(temp);
                                            }
                                            #endregion

                                            #region выделение подстроки толщины
                                            if (regTolshList.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                            }
                                            #endregion


                                            //GetRegexTUFromString(temp, lastRow);
                                            //GetRegexMarkFromString(temp, lastRow);

                                            dtProduct.Rows[lastRow]["Примечание"] = temp;

                                            //поиск цены
                                            cellRange = (Excel.Range)excelworksheet.Cells[i, 3];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                GetRegexPriceFromString(temp, lastRow);
                                            }

                                        }
                                        #endregion
                                    }
                                    #endregion
                                }
                            }
                        }
                    }
                    #endregion

                    #region Лист - качстали
                    if (excelworksheet.Name.ToLower() == "качстали")
                    {
                        List<structTab> tabs = new List<structTab>();

                        tsLabelcurrSheet.Text = excelworksheet.Index.ToString();

                        int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                        int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                        tsPb1.Maximum = cCelRow;// * cCelCol ;

                        listIndexOfNotEmptyName = new List<int>();
                        colForName = 0;

                        tsLabelClearingTable.Text = "Поиск наименований";
                        tsPb1.Value = 0;
                        for (int j = 1; j <= cCelRow; j++) //строки
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, 1];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString();
                            else temp = "";
                            if (temp != "" && cellRange.MergeArea.Count > 1)
                            {
                                #region получение списка непустых строк

                                nameProd = new Regex(@"(?!\s*\w+\d\w+|\s*\w+ий|\s*\w+ой|\s*Г?ост\s|\w+ые|\w+ая|сталь)(?<=\s*)\w{3,}", RegexOptions.IgnoreCase).Match(temp).Value;
                                Regex regType = new Regex(@"\w+ые|\w+ая", RegexOptions.IgnoreCase);
                                structTab tab = new structTab();
                                if (nameProd != "")
                                {
                                    tab.StartCol = 1;
                                    tab.StartRow = j;
                                    tab.listExcelIndexTab = new List<int>();
                                    tab.listdtProductIndexRow = new List<int>();
                                    tab.Name = nameProd;
                                    if (regType.IsMatch(temp))
                                        tab.Type = regType.Match(temp).Value;
                                    else tab.Type = "тип не указан";
                                    tab.Standart = new Regex(@"ГОСТ\s*\d+(?:-\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                    tabs.Add(tab);

                                }

                                #endregion
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        }
                        tsLabelClearingTable.Text = "Поиск параметров по имени";
                        tsPb1.Value = 0;

                        for (int k = 0; k < tabs.Count; k++)
                        {
                            structTab stab = tabs[k];
                            int endRowForCurTab = 1;
                            if (k < tabs.Count - 2)
                            { endRowForCurTab = tabs[k + 1].StartRow; }
                            else { endRowForCurTab = cCelRow; }

                            for (int i = stab.StartRow + 1; i < endRowForCurTab; i++) //строки
                            {

                                Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[i, 1];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString();
                                else temp = "";
                                if (temp != "")
                                {

                                    #region вложенная таблица

                                    dtProduct.Rows.Add();
                                    int lastRow = dtProduct.Rows.Count - 1;
                                    dtProduct.Rows[lastRow]["Название"] = new Regex(@"(?!\w+ый|\w+ая)круг|профиль|шгр|квадрат\b|заготовка|полоса|профиль", RegexOptions.IgnoreCase).Match(temp.Trim()).Value;
                                    if (dtProduct.Rows[lastRow]["Название"].ToString().Trim().ToLower() == "шгр")
                                        dtProduct.Rows[lastRow]["Название"] = "Шестигранник";
                                    if (dtProduct.Rows[lastRow]["Название"].ToString().Trim().ToLower() == "кург")
                                        dtProduct.Rows[lastRow]["Название"] = "Круг";
                                    if (dtProduct.Rows[lastRow]["Название"].ToString() == "") dtProduct.Rows[lastRow]["Название"] = stab.Name;

                                    dtProduct.Rows[lastRow]["Тип"] = new Regex(@"\w+ые|\w+ая|\w+ый", RegexOptions.IgnoreCase).Match(temp.Trim()).Value;
                                    if (dtProduct.Rows[lastRow]["Тип"].ToString() == "тип не указан" || dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                        dtProduct.Rows[lastRow]["Тип"] = stab.Type;

                                    dtProduct.Rows[lastRow]["Стандарт"] = stab.Standart;

                                    Regex regTolshList = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                    //Regex regShirina = new Regex(@"(?<=\d+(?:[,\.]\d+)?[xх])\d+(?:[,\.]\d+)?(?=[xх]\d)", RegexOptions.IgnoreCase);
                                    Regex regDlina = new Regex(@"\d+(?:[,\.]\d+)?\s*(?=мм)|\d+(?:[,\.]\d+)?(?=[xх]\d)|(?<=/)\d+(?=[рp])|\d{3,4}|\d{3}-\d{3}", RegexOptions.IgnoreCase);

                                    #region выделение подстроки Диаметра
                                    if (regDlina.IsMatch(temp))
                                    {
                                        dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDlina.Match(temp);
                                    }
                                    #endregion

                                    #region выделение подстроки толщины
                                    if (regTolshList.IsMatch(temp))
                                    {
                                        dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                    }
                                    #endregion

                                    //GetRegexTUFromString(temp, lastRow);
                                    GetRegexMarkFromString(temp, lastRow);

                                    dtProduct.Rows[lastRow]["Примечание"] = temp;

                                    //поиск цены
                                    cellRange = (Excel.Range)excelworksheet.Cells[i, 4];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString();
                                    else temp = "";
                                    if (temp != "" && temp != "0")
                                    {
                                        GetRegexPriceFromString(temp, lastRow);
                                    }
                                    else //если цена не найдена в этом столбце, то поискать в следующем
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[i, 5];
                                        if (cellRange.Value != null)
                                        {
                                            temp = cellRange.Value.ToString();
                                            if (temp != "")
                                            {
                                                GetRegexPriceFromString(temp, lastRow);
                                            }
                                        }
                                    }

                                    cellRange = (Excel.Range)excelworksheet.Cells[i, 2];
                                    if (cellRange.Value != null)
                                    {
                                        temp = cellRange.Value.ToString();
                                        if (temp != "")
                                        {
                                            dtProduct.Rows[lastRow]["Мерность (т, м, мм)"] = temp;
                                        }
                                    }

                                    #endregion


                                }
                            }
                        }
                    }
                    #endregion

                    #region Лист - метизы
                    if (excelworksheet.Name.ToLower() == "метизы")
                    {
                        List<structTab> tabs = new List<structTab>();

                        tsLabelcurrSheet.Text = excelworksheet.Index.ToString();

                        int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                        int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                        tsPb1.Maximum = cCelRow;// * cCelCol ;

                        listIndexOfNotEmptyName = new List<int>();
                        colForName = 0;

                        tsLabelClearingTable.Text = "Поиск наименований";
                        tsPb1.Value = 0;
                        for (int j = 1; j <= cCelRow; j++) //строки
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, 1];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString();
                            else temp = "";
                            if (temp != "" && cellRange.MergeArea.Count > 1)
                            {
                                #region получение списка непустых строк

                                nameProd = new Regex(@"(?!\s*\w+\d\w+|\s*\w+ий|\s*\w+ой|\s*Г?ост\s|\w+ые|\w+ая|сталь|наименование)(?<=\s*)\w{3,}(?:-\w+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                Regex regType = new Regex(@"\w+ые|\w+ая|c полупотайной головкой|с\s*полукруглой\s*головкой|с\s*потайной\s*головкой|с\s*шестигранной\s*головкой|с\s*плоской\s*головкой|холодной\s*вытяжки|для\s*пружинных\s*шайб|егоза|[BPВР]\s*-\s*1|рабица|из\s*рифленой\s*проволоки|для\s*фрезеровки\s*древесины", RegexOptions.IgnoreCase);
                                structTab tab = new structTab();
                                if (nameProd != "")
                                {
                                    tab.StartCol = 1;
                                    tab.StartRow = j;
                                    tab.listExcelIndexTab = new List<int>();
                                    tab.listdtProductIndexRow = new List<int>();
                                    tab.Name = nameProd;
                                    if (regType.IsMatch(temp))
                                        tab.Type = regType.Match(temp).Value;
                                    else tab.Type = "тип не указан";
                                    tab.Standart = new Regex(@"ГОСТ\s*\d+(?:-\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                    tabs.Add(tab);

                                }

                                #endregion
                            }
                            else if (temp != "")
                            {
                                nameProd = new Regex(@"сетка\s+\w{4,}|канат\s*гост|лента(?:\s+\w+\s*)*гост|сита\s+\w{3,}|электроды\s+для|электроды\s+\w{4,}|болт\s+\w{4,}|^рым-болт$|гайки|шпильки|гвозди\s+\w{4,}|дюбель-гвоздь|дюбель-винт|шплинты\s+гост|заклепки|шайба(?:\s+\w+\s*)*гост(?:\s+[\w+\d+]\s*)*\d$|ножи|пилы", RegexOptions.IgnoreCase).Match(temp).Value;
                                Regex regType = new Regex(@"\w+ые|\w+ая", RegexOptions.IgnoreCase);
                                structTab tab = new structTab();
                                if (nameProd != "")
                                {
                                    nameProd = new Regex(@"(?!\s*\w+\d\w+|\s*\w+ий|\s*\w+ой|\s*Г?ост\s|\w+ые|\w+ая|сталь|наименование)(?<=\s*)\w{3,}(?:-\w+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                    tab.StartCol = 1;
                                    tab.StartRow = j;
                                    tab.listExcelIndexTab = new List<int>();
                                    tab.listdtProductIndexRow = new List<int>();
                                    tab.Name = nameProd;
                                    if (regType.IsMatch(temp))
                                        tab.Type = regType.Match(temp).Value;
                                    else tab.Type = "тип не указан";
                                    tab.Standart = new Regex(@"ГОСТ\s*\d+(?:-\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                    tabs.Add(tab);

                                }
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        }
                        tsLabelClearingTable.Text = "Поиск параметров по имени";
                        tsPb1.Value = 0;
                        tsPb1.Maximum = tabs.Count;
                        for (int k = 0; k < tabs.Count; k++)
                        {
                            structTab stab = tabs[k];
                            int endRowForCurTab = 1;
                            if (k < tabs.Count - 2)
                            { endRowForCurTab = tabs[k + 1].StartRow; }
                            else { endRowForCurTab = cCelRow; }

                            for (int i = stab.StartRow + 1; i < endRowForCurTab; i++) //строки
                            {

                                Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[i, 1];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString();
                                else temp = "";
                                if (temp != "")
                                {

                                    #region вложенная таблица

                                    if (new Regex(@"(?!\w+ый|\w+ая|^\d|\s*\w+\d\w+|\s*\w+ий|\s*\w+ой|гост|\w+ые|\w+ая|сталь|наименование)^\w{3,}", RegexOptions.IgnoreCase).IsMatch(temp.Trim()))
                                    {
                                        dtProduct.Rows.Add();
                                        int lastRow = dtProduct.Rows.Count - 1;
                                        dtProduct.Rows[lastRow]["Название"] = new Regex(@"(?!\w+ый|\w+ая)^\w{3,}", RegexOptions.IgnoreCase).Match(temp.Trim()).Value;

                                        if (dtProduct.Rows[lastRow]["Название"].ToString() == "") dtProduct.Rows[lastRow]["Название"] = stab.Name;

                                        dtProduct.Rows[lastRow]["Тип"] = new Regex(@"\w+ые|\w+ая|\w+ый", RegexOptions.IgnoreCase).Match(temp.Trim()).Value;
                                        if (dtProduct.Rows[lastRow]["Тип"].ToString() == "тип не указан" || dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                            dtProduct.Rows[lastRow]["Тип"] = stab.Type;

                                        dtProduct.Rows[lastRow]["Стандарт"] = stab.Standart;

                                        if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "сетка")
                                        {
                                            Regex regTolshList = new Regex(@"(?<=\s)\d+(?:[,\.]\d+)?(?=\s?[xх]\s?\d+(?:[,\.]\d+)?)", RegexOptions.IgnoreCase);
                                            Regex regDlina = new Regex(@"(?<=\d+(?:[,\.]\d+)?\s?[/xх]\s?)\d+(?:[,\.]\d+)?(?=\s|-|$)", RegexOptions.IgnoreCase);
                                            Regex regShirina = new Regex(@"(?<=\d+(?:[,\.]\d+)?\s?[xх]\s?)\d+(?:[,\.]\d+)?(?=\s?[xх]\s?\d+(?:[,\.]\d+)?)", RegexOptions.IgnoreCase);

                                            #region выделение подстроки Диаметра
                                            if (regDlina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDlina.Match(temp);
                                            }
                                            #endregion

                                            #region выделение подстроки толщины
                                            if (regTolshList.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                            }
                                            #endregion

                                            #region ширина
                                            if (regShirina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = regShirina.Match(temp);
                                            }
                                            #endregion
                                        }
                                        else if (new Regex(@"(?<=^\w+\s+)\d+(?:[,\.]\d+)?|\d+(?:[,\.]\d+)?(?=\s?[xх]\s?\d+(?:[,\.]\d+)?(?:\s|$))|d\s?\d+(?:[,\.]\d+)?|\d+(?:[,\.]\d+)?\s?(?=мм)", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            Regex regTolshList = new Regex(@"(?<=\s\d+(?:[,\.]\d+)?\s?[xх]\s?)\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                            Regex regDlina = new Regex(@"(?<=^\w+\s+)\d+(?:[,\.]\d+)?|\d+(?:[,\.]\d+)?(?=\s?[xх]\s?\d+(?:[,\.]\d+)?(?:\s|$))|(?<=d\s?)\d+(?:[,\.]\d+)?|\d+(?:[,\.]\d+)?\s?(?=мм)", RegexOptions.IgnoreCase);

                                            #region выделение подстроки Диаметра
                                            if (regDlina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDlina.Match(temp);
                                            }
                                            #endregion

                                            #region выделение подстроки толщины
                                            if (regTolshList.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                            }
                                            #endregion
                                        }

                                        else if (new Regex(@"\d+(?:[,\.]\d+)?\s?[xх]\s?\d+(?:[,\.]\d+)?\s?[xх]\s?\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            Regex regTolshList = new Regex(@"(?<=\s\d+(?:[,\.]\d+)?\s?[xх]\s?)\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                                            Regex regShirina = new Regex(@"(?<=\d+(?:[,\.]\d+)?\s?[xх]\s?)\d+(?:[,\.]\d+)?(?=\s?[xх]\s?\d+(?:[,\.]\d+)?)", RegexOptions.IgnoreCase);
                                            Regex regDlina = new Regex(@"(?<=^\w+\s+)\d+(?:[,\.]\d+)?|\d+(?:[,\.]\d+)?(?=\s?[xх]\s?\d+(?:[,\.]\d+)?(?:\s|$))", RegexOptions.IgnoreCase);

                                            #region выделение подстроки Диаметра
                                            if (regDlina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDlina.Match(temp);
                                            }
                                            #endregion

                                            #region выделение подстроки толщины
                                            if (regTolshList.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolshList.Match(temp);
                                            }
                                            #endregion

                                            #region выделение подстроки длины
                                            if (regShirina.IsMatch(temp))
                                            {
                                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = regShirina.Match(temp);
                                            }
                                            #endregion
                                        }

                                        GetRegexTUFromString(temp, lastRow);
                                        GetRegexMarkFromString(temp, lastRow);

                                        dtProduct.Rows[lastRow]["Примечание"] = temp;

                                        //поиск цены
                                        cellRange = (Excel.Range)excelworksheet.Cells[i, 2];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "" && temp != "0")
                                        {
                                            GetRegexPriceFromString(temp, lastRow);
                                        }
                                        else //если цена не найдена в этом столбце, то поискать в следующем
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[i, 3];
                                            if (cellRange.Value != null)
                                            {
                                                temp = cellRange.Value.ToString().Trim();
                                                if (temp != "")
                                                {
                                                    GetRegexPriceFromString(temp, lastRow);
                                                }
                                            }
                                        }

                                    }
                                    #endregion

                                }
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        }
                    }
                    #endregion

                }
                clearingTable();
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции Metall-Baza\n\n" + ex.ToString()); }
        }

        ///<summary> 
        ///<remarks> Открытие и чтение экселевского файла ММК </remarks>
        ///<param name="path" >путь к файлу</param>
        ///</summary>
        private void MMK(string path)
        {
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                SetNameFromName(filePath);
                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                List<string> s = new List<string>();
                excelappworkbooks = excelapp.Workbooks;
                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла MMK\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();
                List<structTab> tabs = new List<structTab>();
                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = 14;//excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                    tsPb1.Maximum = cCelRow * cCelCol;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    tsLabelClearingTable.Text = "Поиск наименований";
                    tsPb1.Value = 0;
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= cCelCol; i++)
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString();
                            else temp = "";
                            if (temp != "")
                            {
                                #region получение списка вложенных таблиц

                                if (new Regex(@"размер,\s*м", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    structTab tab = new structTab
                                    {
                                        StartCol = i,
                                        StartRow = j,
                                        listExcelIndexTab = new List<int>(),
                                        listdtProductIndexRow = new List<int>()
                                    };
                                    tabs.Add(tab);
                                }

                                ManualStringNameProd = nameProd;

                                #endregion
                            }

                            InfoOrganization(temp);

                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        }
                    }
                    tsLabelClearingTable.Text = "Поиск параметров по имени";
                    tsPb1.Value = 0;
                    nameProd = "";
                    string Mark = "";
                    string Standart = "";
                    string Type = "";
                    string TU = "";
                    string tmp = "";
                    int lastRow = 0;
                    Regex regName = new Regex(@"лента\b|лист\b|арматура\b|полоса\b|угол(?:ок)?\b|швеллер\w?\b|труб\w?\b|круг\w?\b|шестигранник|шгр\b|квадрат\b|сталь\b|катанка|быстрорез|колесо|заготовка|блок\b|^\s*вал\s*|втулка|поковка", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex("");
                    Regex regTolsh = new Regex("");
                    Regex regMetr = new Regex("");

                    //для каждой таблицы из списка произвести поиск содержания
                    for (int k = 0; k < tabs.Count; k++)
                    {
                        Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].StartRow, tabs[k].StartCol];
                        if (cellRange.Value != null)
                            temp = cellRange.Value.ToString();
                        else temp = "";
                        if (new Regex(@"размер", RegexOptions.IgnoreCase).IsMatch(temp))
                        {
                            for (int i = tabs[k].StartRow + 1; i < cCelRow; i++)
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[i, tabs[k].StartCol];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString();
                                else temp = "";
                                if (cellRange.MergeArea.Count > 1)
                                {
                                    nameProd = "";
                                    Excel.Range cellRangeForName = (Excel.Range)excelworksheet.Cells[i, tabs[k].StartCol - 1];
                                    if (cellRangeForName.Value != null)
                                        temp = cellRangeForName.Value.ToString();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        nameProd = new Regex(@"(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        Mark = new Regex(@"[AА]500[CС]|[aа]-?i(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        Standart = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:Г[Оо][Сс][Тт]\s{0,3})(?:[рР]-\s?)?(?:\d{1,5}[-\s]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])", RegexOptions.IgnoreCase).Match(temp).Value;
                                        Type = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|г/к|х/к|ВГП|пруток|моток|(?<=\d)[а-яa-z](?=\s|[xх*])", RegexOptions.IgnoreCase).Match(temp).Value;
                                        TU = new Regex(@"ТУ(?:\s*-\s*у)?\s*\d+\s*-\s*\d+[рp]?\s*-\s*\d+(\s*-\s*\d+)?|ГОСТ\s*(?:[РP]\s*)?\d+(?:-\d+)?|zn\s*\d+", RegexOptions.IgnoreCase).Match(temp).Value;
                                    }
                                }
                                else if (nameProd != "")
                                {
                                    dtProduct.Rows.Add();
                                    tabs[k].listExcelIndexTab.Add(i);
                                    lastRow = dtProduct.Rows.Count - 1;
                                    tabs[k].listdtProductIndexRow.Add(lastRow);
                                    dtProduct.Rows[lastRow]["Название"] = nameProd;
                                    if (dtProduct.Rows[lastRow]["Название"].ToString() == "")
                                        dtProduct.Rows[lastRow]["Название"] = "Труба";
                                    if (Type == "")
                                        Type = new Regex(@"(?<=\d)[а-яa-z](?=\s|[xх*])", RegexOptions.IgnoreCase).Match(temp).Value;
                                    if (Type.ToLower().Contains("г/к")) Type = "горячекатаный";
                                    if (Type.ToLower().Contains("х/к")) Type = "холоднокатанный";
                                    dtProduct.Rows[lastRow]["Тип"] = Type;
                                    if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                        dtProduct.Rows[lastRow]["Тип"] = "тип не указан";

                                    Excel.Range cellRangeMark = (Excel.Range)excelworksheet.Cells[i, tabs[k].StartCol - 1];
                                    if (cellRangeMark.Value != null)
                                        tmp = cellRangeMark.Value.ToString();
                                    else tmp = "";
                                    if (tmp != "") dtProduct.Rows[lastRow]["Марка"] = tmp;
                                    else
                                        dtProduct.Rows[lastRow]["Марка"] = Mark;

                                    Excel.Range cellRangePrice = (Excel.Range)excelworksheet.Cells[i, tabs[k].StartCol + 3];
                                    if (cellRangePrice.Value != null)
                                        tmp = cellRangePrice.Value.ToString();
                                    else tmp = "";
                                    if (tmp != "") dtProduct.Rows[lastRow]["Цена"] = new Regex(@"\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).Match(tmp.Trim()).Value;

                                    dtProduct.Rows[lastRow]["Стандарт"] = Standart;
                                    dtProduct.Rows[lastRow]["Примечание"] = temp;

                                    if (nameProd.ToLower() == "круг" && !new Regex(@"\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[i, tabs[k].StartCol - 1];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString();
                                        else temp = "";
                                        dtProduct.Rows[lastRow]["Стандарт"] = "";
                                        dtProduct.Rows[lastRow]["Примечание"] += "  " + temp;
                                    }


                                    string[] diam, tolsh, metraj;
                                    List<double> Ddiam = new List<double>(), Dtolsh = new List<double>(), Dmetraj = new List<double>();
                                    List<double> ch = new List<double>();
                                    string tempo = "";
                                    tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(temp).Value;
                                    if (tempo != "")
                                    {
                                        diam = tempo.Split('-');
                                        foreach (string e in diam)
                                            Ddiam.Add(Convert.ToDouble(e));
                                        ch.Clear();
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
                                            }
                                            if (ch.Count > 0) Ddiam = ch;
                                        }
                                    }
                                    else
                                    {
                                        tempo = "";
                                        tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(temp).Value;
                                        if (tempo != "")
                                        {
                                            diam = new string[] { tempo };
                                        }
                                        else
                                        {
                                            tempo = "";
                                            tempo = new Regex(@"(?<=^|dy\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if (tempo != "")
                                            {
                                                diam = tempo.Split('-');
                                                foreach (string e in diam)
                                                    Ddiam.Add(Convert.ToDouble(e));
                                                ch.Clear();
                                                double increment = 0;
                                                if (Ddiam[1] >= 1 && Ddiam[1] < 4) increment = 0.5;
                                                if (Ddiam[1] >= 4 && Ddiam[1] < 50) increment = 2;
                                                if (Ddiam[1] >= 50) increment = 10;
                                                if (increment > 0)
                                                {
                                                    for (double d = Ddiam[0]; d <= Ddiam[1]; d += increment)
                                                    {
                                                        if (d != Ddiam[0] && (d - 0.1) % 1 == 0)
                                                            d -= 0.1;
                                                        ch.Add(d);
                                                        if (d + increment > Ddiam[1] && d != Ddiam[1]) ch.Add(Ddiam[1]);
                                                    }
                                                    if (ch.Count > 0) Ddiam = ch;
                                                }
                                            }
                                            else
                                            {
                                                tempo = "";
                                                tempo = new Regex(@"(?<=^|dy\s*)\d+(?:[,\.]\d+)?(?=\w?\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                if (tempo != "")
                                                {
                                                    diam = new string[] { tempo };
                                                }
                                                else diam = new string[] { "" };
                                            }
                                        }
                                    }

                                    tempo = new Regex(@"(?<=^|dy\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(temp).Value;
                                    if (tempo != "")
                                    {
                                        tolsh = tempo.Split('-');
                                    }
                                    else
                                    {
                                        tempo = "";

                                        tempo = new Regex(@"(?<=^|dy\s*)\d+(?:[,\.]\d+)?(?=\w?\s*[xх*])", RegexOptions.IgnoreCase).Match(temp).Value;
                                        if (tempo != "")
                                        {
                                            tolsh = new string[] { tempo };
                                        }
                                        else tolsh = new string[] { "" };
                                    }
                                    tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                    if (tempo != "")
                                    {
                                        metraj = tempo.Split('-');
                                    }
                                    else
                                    {
                                        tempo = "";
                                        tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        if (tempo != "")
                                        {
                                            metraj = new string[] { tempo };
                                        }
                                        else metraj = new string[] { "" };

                                    }

                                    if (Ddiam.Count == 0)
                                        foreach (string e in diam)
                                        {
                                            if (e != "")
                                                Ddiam.Add(Convert.ToDouble(e));
                                            else Ddiam.Add(0);
                                        }

                                    if (Dtolsh.Count == 0)
                                        foreach (string e in tolsh)
                                            if (e != "")
                                                Dtolsh.Add(Convert.ToDouble(e));
                                            else Dtolsh.Add(0);

                                    if (Dmetraj.Count == 0)
                                        foreach (string e in metraj)
                                            if (e != "")
                                                Dmetraj.Add(Convert.ToDouble(e));
                                            else Dmetraj.Add(0);

                                    if (Ddiam[0] < Dtolsh[0])
                                    {
                                        ch = Ddiam;
                                        Ddiam = Dtolsh;
                                        Dtolsh = ch;
                                    }


                                    for (int d = 0; d < Ddiam.Count; d++)
                                        for (int t = 0; t < Dtolsh.Count; t++)
                                            for (int m = 0; m < Dmetraj.Count; m++)
                                            {
                                                lastRow = dtProduct.Rows.Count - 1;
                                                if (d == 0 && t == 0 && m == 0)
                                                {
                                                    if (Ddiam[0] != 0) dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = Ddiam[0];
                                                    if (Dtolsh[0] != 0) dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = Dtolsh[0];
                                                    if (Dmetraj[0] != 0) dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = Dmetraj[0];
                                                }
                                                else
                                                {
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
                            }
                        }

                    }
                    for (int i = 1; i < tabs[0].StartRow; i++)
                    {
                        for (int j = 1; j < 19; j++)
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[i, j];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString();
                            else temp = "";

                            if (textBoxOrgAdress.Text == "")
                                textBoxOrgAdress.Text = new Regex(@"\w+\.\w+,\s*\w+\.\s*\w+,\s*\w+\.\s*\d+\w+(?=\s+)", RegexOptions.IgnoreCase).Match(temp).Value;

                            if (textBoxOrgTelefon.Text == "")
                                textBoxOrgTelefon.Text = new Regex(@"(?<=Тел/факс\s*:\s*)\d+\s*\(?\d+\)?(?:\s*\d+-\d+-\d+,?)+", RegexOptions.IgnoreCase).Match(temp).Value;

                        }
                    }
                }
                clearingTable();
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции MMK\n\n" + ex.ToString()); }
        }

        ///<summary> 
        ///<remarks> Открытие и чтение экселевского файла Металлург </remarks>
        ///<param name="path" >путь к файлу</param>
        ///</summary>
        private void Metallurg(string path)
        {
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                SetNameFromName(filePath);
                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;
                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла Metallurg\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    #region Вложенная таблица

                    List<structTab> tabs = new List<structTab>();

                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                    tsPb1.Maximum = cCelRow;// * cCelCol ;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    tsLabelClearingTable.Text = "Поиск наименований";
                    tsPb1.Value = 0;
                    int lastRow = 0;

                    Regex razmer = new Regex(@"\d+(?:[,\.]\d+)?\s*[xх]\s*\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);

                    Regex regName = new Regex(@"^\w+(?=\s*(?:\w+\s*)?\d+)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ые|\w+ая|в\s*вус\s*изоляции", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"\d+(?:[,\.]\d+)?(?=\s*[xх]\s*\d+(?:[,\.]\d+)?)", RegexOptions.IgnoreCase);
                    Regex regTolsh = new Regex(@"(?<=\d+(?:[,\.]\d+)?\s*[xх]\s*)\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"(?:\d{,3}[ШСТУ]+\d{,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+\d{,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+\d{,3})(?=\s+|$)|(?:(?:Ст.)|(?:ст.)(?:\s*\d{1,2})\b)(?=\s+|$)|(?:\d{,3}[ХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)(?=\s+|$)|(?:[АA]-?\d)(?=\s+|$)|(?<=[Сс][Тт]\.\s?)\d{1,2}[гГ]\d{1,2}[cCсС]|(?<=ст\.)\d{1,2}[хфа]+(?=\s|$)|(?<=\s)[сС][тТ]\.?\s?\d{1,2}[_\w]+");
                    Regex regGost = new Regex(@"ТУ(?:\s*-\s*у)?\s*\d+\s*-\s*\d+[рp]?\s*-\s*\d+(\s*-\s*\d+)?|ГОСТ\s*(?:[РP]\s*)?\d+-\d+|ТУ\s*\d+\.\d+\s*-\s*\d+\s*-\s*\d+\s*:\s*\d+\s*|\d+\s*-\s*\d+\s*-\s*\d+(\s*-\s*\d+)?", RegexOptions.IgnoreCase);
                    Regex regVes = new Regex(@"(?<=вес\s*)\d+(?:[,\.]\d+)?(?=\s*т?)", RegexOptions.IgnoreCase);
                    Regex regPrice = new Regex(@"(?<=(?:цена)?\s*)\d+(?:[,\.]\d+)?(?=\s*р[/\\]т?)", RegexOptions.IgnoreCase);

                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, 1];
                        if (cellRange.Value != null)
                            temp = cellRange.Value.ToString();
                        else temp = "";
                        if (temp != "")
                        {
                            if (razmer.IsMatch(temp))
                            {
                                dtProduct.Rows.Add();
                                lastRow = dtProduct.Rows.Count - 1;
                                dtProduct.Rows[lastRow]["Название"] = regName.Match(temp).Value;
                                dtProduct.Rows[lastRow]["Тип"] = regType.Match(temp).Value;
                                if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                    dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDiam.Match(temp).Value;
                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolsh.Match(temp).Value;
                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = "";
                                dtProduct.Rows[lastRow]["Мерность (т, м, мм)"] = regVes.Match(temp).Value;
                                dtProduct.Rows[lastRow]["Марка"] = regMark.Match(temp).Value;
                                dtProduct.Rows[lastRow]["Стандарт"] = regGost.Match(temp).Value;
                                dtProduct.Rows[lastRow]["Класс"] = "";
                                dtProduct.Rows[lastRow]["Цена"] = regPrice.Match(temp).Value;
                                dtProduct.Rows[lastRow]["Примечание"] = temp;

                            }
                        }

                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                    }

                    #endregion
                }
                clearingTable();
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции Metallurg\n\n" + ex.ToString()); }
        }

        ///<summary> 
        ///<remarks> Открытие и чтение экселевского файла ЧЗПТ </remarks>
        ///<param name="path" >путь к файлу</param>
        ///</summary>
        private void CHZPT(string path)
        {
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                SetNameFromName(filePath);
                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;
                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла CHZPT\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {

                    List<structTab> tabs = new List<structTab>();

                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                    tsPb1.Maximum = cCelRow * cCelCol;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    tsLabelClearingTable.Text = "Поиск наименований";
                    tsPb1.Value = 0;
                    int lastRow = 0;

                    Regex razmer = new Regex(@"\d+(?:[,\.]\d+)?\s*[xх*]\s*\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);

                    Regex regName = new Regex(@"^\w+(?=\s*(?:\w+\s*)?\d+)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ые|\w+ая|в\s*вус\s*изоляции|э/с|х/к", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"(?<=\s|^)\d+(?:[,\.]\d+)?(?=\s*[xх*]\s*\d+(?:[,\.]\d+)?)", RegexOptions.IgnoreCase);
                    Regex regTolsh = new Regex(@"(?<=\d+(?:[,\.]\d+)?\s*[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regShirina = new Regex(@"(?<=\d+(?:[,\.]\d+)?\s*[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s*[xх*]\s*\d+(?:[,\.]\d+)?)", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"(?:\d{,3}[ШСТУ]+\d{,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+\d{,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+\d{,3})(?=\s+|$)|(?:(?:Ст.)|(?:ст.)(?:\s*\d{1,2})\b)(?=\s+|$)|(?:\d{,3}[ХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)(?=\s+|$)|(?:[АA]-?\d)(?=\s+|$)|(?<=[Сс][Тт]\.\s?)\d{1,2}[гГ]\d{1,2}[cCсС]|(?<=ст\.)\d{1,2}[хфа]+(?=\s|$)|(?<=\s)[сС][тТ]\.?\s?\d{1,2}[_\w]+");
                    Regex regGost = new Regex(@"ТУ(?:\s*-\s*у)?\s*\d+\s*-\s*\d+[рp]?\s*-\s*\d+(\s*-\s*\d+)?|ГОСТ\s*(?:[РP]\s*)?\d+-\d+|ТУ\s*\d+\.\d+\s*-\s*\d+\s*-\s*\d+\s*:\s*\d+\s*|\d+\s*-\s*\d+\s*-\s*\d+(\s*-\s*\d+)?", RegexOptions.IgnoreCase);
                    Regex regVes = new Regex(@"(?<=вес\s*)\d+(?:[,\.]\d+)?(?=\s*т?)", RegexOptions.IgnoreCase);
                    Regex regPrice = new Regex(@"\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);

                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= cCelCol; i++)
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"наим.*ие", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    nameProd = "Труба";
                                    structTab tab = new structTab
                                    {
                                        StartCol = i,
                                        StartRow = j,
                                        listExcelIndexTab = new List<int>(),
                                        listdtProductIndexRow = new List<int>(),
                                        Name = nameProd
                                    };
                                    if (regType.IsMatch(temp))
                                        tab.Type = regType.Match(temp).Value;
                                    else tab.Type = "тип не указан";
                                    tab.Standart = "";
                                    tabs.Add(tab);
                                }

                                if (textBoxOrgAdress.Text == "")
                                    textBoxOrgAdress.Text = new Regex(@"\d+,?\s*\w+\.\w+,?\s*\w+\s*\w*,\s*[\w\d]+(?=,)", RegexOptions.IgnoreCase).Match(temp).Value;

                                if (textBoxOrgEmail.Text == "")
                                    textBoxOrgEmail.Text = new Regex(@"(?<=E-MAIL:\s)[\w\d-]+@[\w\d-]+\.\w\w\w?", RegexOptions.IgnoreCase).Match(temp).Value;

                                if (textBoxOrgSite.Text == "")
                                    textBoxOrgSite.Text = new Regex(@"\w+\.\w+\.\w\w", RegexOptions.IgnoreCase).Match(temp).Value;
                                if (textBoxOrgTelefon.Text == "")
                                    textBoxOrgTelefon.Text = new Regex(@"(?<=ТЕЛ\./ФАКС:\s*)\+\d+\(\d{3,5}\)\d+-\d+-\d+,\s*\w+\.\d+-\d+-\d+-\d+-\d+(?=,?)", RegexOptions.IgnoreCase).Match(temp).Value;
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        }
                    }

                    tsPb1.Maximum = tabs.Count;
                    tsPb1.Value = 0;

                    #region обработка вложенной таблицы
                    for (int k = 0; k < tabs.Count; k++)
                    {
                        structTab tab = tabs[k];

                        int endRowForCurTab = 1;
                        if (k < tabs.Count - 2)
                        { endRowForCurTab = tabs[k + 1].StartRow; }
                        else { endRowForCurTab = cCelRow; }
                        for (int i = tab.StartCol; i < cCelCol; i++)
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[tab.StartRow, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"наим.*ие", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    for (int j = tab.StartRow + 1; j < endRowForCurTab; j++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            if (razmer.IsMatch(temp))
                                            {
                                                dtProduct.Rows.Add();
                                                tabs[k].listExcelIndexTab.Add(j);
                                                lastRow = dtProduct.Rows.Count - 1;
                                                tabs[k].listdtProductIndexRow.Add(lastRow);
                                                dtProduct.Rows[lastRow]["Название"] = tab.Name;
                                                if (dtProduct.Rows[lastRow]["Название"].ToString() == "")
                                                    dtProduct.Rows[lastRow]["Название"] = "Труба";
                                                dtProduct.Rows[lastRow]["Тип"] = regType.Match(temp).Value;
                                                if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                                    dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDiam.Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolsh.Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = regShirina.Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Марка"] = regMark.Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Стандарт"] = regGost.Match(temp).Value;
                                                dtProduct.Rows[lastRow]["Примечание"] = temp;
                                            }
                                        }
                                    }
                                }
                                if (tabs[k].listExcelIndexTab.Count > 0)
                                {
                                    if (new Regex(@"от\s*20\s*тн", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        for (int j = 0; j < tabs[k].listExcelIndexTab.Count; j++)
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[j], i];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                GetRegexPriceFromString(temp, tabs[k].listdtProductIndexRow[j]);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;

                    }
                    #endregion
                }
                clearingTable();
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции CHZPT\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла МеталлИнвест до 07.17
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void MetallInvestOld(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                SetNameFromName(filePath);

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла MetallInvest\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    List<structTab> tabs = new List<structTab>();
                    List<double> DiamStd = new List<double>();
                    List<double> TolshStd = new List<double>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;
                    tsPb1.Maximum = cCelCol * cCelRow;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    tsLabelClearingTable.Text = "Поиск наименований";
                    tsPb1.Value = 0;

                    DiamStd = GetStdFromFile(Application.StartupPath + "\\DiamStd.csv");
                    TolshStd = GetStdFromFile(Application.StartupPath + "\\TolshStd.csv");

                    structTab tab = new structTab();
                    tsPb1.Maximum = cCelRow * cCelCol;

                    Regex regName = new Regex(@"(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|э/с|х/к", RegexOptions.IgnoreCase);
                    Regex regSklad = new Regex(@"(?<=\w*\s*склад.*\s*:\s*)\w+\.\s*\w+,\s*\w+(?:\s*\w+)?,\s*\d+\.\s*.Металл-база.", RegexOptions.IgnoreCase);

                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                Regex regRazmer = new Regex(@"толщина", RegexOptions.IgnoreCase);

                                if (regRazmer.IsMatch(temp))
                                {


                                    tab.listExcelIndexTab = new List<int>();
                                    tab.listdtProductIndexRow = new List<int>();
                                    cellRange = (Excel.Range)excelworksheet.Cells[j - 1, i - 2];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        tab.Name = regName.Match(temp).Value;
                                        tab.Type = regType.Match(temp).Value;
                                    }
                                    if (cellRange.MergeArea.Count == 2)
                                    {
                                        i++;
                                    }
                                    tab.StartCol = i;
                                    tab.StartRow = j;
                                    tabs.Add(tab);

                                }
                            }

                            #region Склад

                            if (regSklad.IsMatch(temp))
                            {
                                Regex sklad = new Regex(@"(?<=\w*\s*склад.*\s*:\s*)\w+\.\s*\w+,\s*\w+(?:\s*\w+)?,\s*\d+\.\s*.Металл-база.", RegexOptions.IgnoreCase);
                                string tmp = sklad.Match(temp).Value;
                                ListViewItem lvi = new ListViewItem(tmp);
                                bool isIn = false;
                                if (listViewAdrSklad.Items.Count < 1) listViewAdrSklad.Items.Add(lvi);
                                for (int ii = 0; ii < listViewAdrSklad.Items.Count; ii++)
                                {
                                    if (listViewAdrSklad.Items[ii].SubItems[0].Text == tmp) isIn = true;
                                }
                                if (!isIn) listViewAdrSklad.Items.Add(lvi);
                            }

                            #endregion

                            InfoOrganization(temp);


                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }
                    int endIndexRow = 0;

                    tsPb1.Value = 0;
                    tsPb1.Maximum = tabs.Count;

                    Regex regDiam = new Regex(@"(?<=\s|^)\d+(?:[,\.]\d+)?(?=\s*[xх*]\s*\d+(?:[,\.]\d+)?)|(?<=\s|^)\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase);
                    //Regex regTolsh = new Regex(@"(?<=\d+(?:[,\.]\d+)?\s*[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s*[xх*]\s*\d+(?:[,\.]\d+)?)", RegexOptions.IgnoreCase);
                    Regex regMetr = new Regex(@"(?<=\d+(?:[,\.]\d+)?\s*[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase);

                    tsPb1.Value = 0;
                    tsPb1.Maximum = tabs.Count;
                    for (int k = 0; k < tabs.Count; k++)
                    {
                        if (k < tabs.Count - 2)
                        {
                            if (tabs[k + 1].StartRow > tabs[k].StartRow && tabs[k + 1].StartCol <= tabs[k + 1].StartCol + 3)
                                endIndexRow = tabs[k + 1].StartRow;
                            else if (tabs[k + 2].StartRow > tabs[k].StartRow && tabs[k + 2].StartCol <= tabs[k].StartCol + 3)
                                endIndexRow = tabs[k + 2].StartRow;
                            else if (tabs[k + 3].StartRow > tabs[k].StartRow && tabs[k + 3].StartCol <= tabs[k].StartCol + 3)
                                endIndexRow = tabs[k + 3].StartRow;
                            else endIndexRow = cCelRow;
                        }
                        else { endIndexRow = cCelRow; }

                        Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].StartRow, tabs[k].StartCol];
                        if (cellRange.Value != null)
                        {
                            temp = cellRange.Value.ToString().Trim();
                            if (new Regex(@"толщина", RegexOptions.IgnoreCase).IsMatch(temp))
                            {
                                string memoryMergeCell = "";
                                for (int i = tabs[k].StartRow + 1; i < endIndexRow - 1; i++) //строки
                                {
                                    cellRange = (Excel.Range)excelworksheet.Cells[i, tabs[k].StartCol];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        if (new Regex(@"\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            #region dd-dd
                                            string[] strSizes = temp.Split('-');

                                            for (int s = 0; s < strSizes.Length; s++)
                                            {
                                                dtProduct.Rows.Add();
                                                tabs[k].listExcelIndexTab.Add(i);
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                tabs[k].listdtProductIndexRow.Add(lastRow);

                                                dtProduct.Rows[lastRow]["Название"] = tabs[k].Name;
                                                dtProduct.Rows[lastRow]["Тип"] = tabs[k].Type;
                                                if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                                    dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).Match(strSizes[s]).Value;
                                            }
                                            #endregion
                                        }
                                        else if (new Regex(@"\d+(?:[,\.]\d+)?\s*;\s*\d+(?:[,\.]\d+)?(?:;\s*\d+(?:[,\.]\d+)?)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                        {
                                            #region dd;dd...
                                            string[] strSizes = temp.Split(';');

                                            for (int s = 0; s < strSizes.Length; s++)
                                            {
                                                strSizes[s] = strSizes[s].Trim();
                                                dtProduct.Rows.Add();
                                                tabs[k].listExcelIndexTab.Add(i);
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                tabs[k].listdtProductIndexRow.Add(lastRow);

                                                dtProduct.Rows[lastRow]["Название"] = tabs[k].Name;
                                                dtProduct.Rows[lastRow]["Тип"] = tabs[k].Type;
                                                if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                                    dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).Match(strSizes[s]).Value;
                                            }
                                            #endregion
                                        }

                                        //dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDiam.Match(temp).Value;
                                        ////dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolsh.Match(temp).Value;
                                        //dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = regMetr.Match(temp).Value;
                                        //dtProduct.Rows[lastRow]["Примечание"] = temp;
                                    }
                                    if (cellRange.MergeArea.Count > 1)
                                    {
                                        if (temp != "") memoryMergeCell = temp;
                                        if (new Regex(@"(?!-)\d+(?:[,\.]\d+)?\s*,\s*\d+(?:[,\.]\d+)?(?:,\s*\d+(?:[,\.]\d+)?)?", RegexOptions.IgnoreCase).IsMatch(memoryMergeCell))
                                        {
                                            #region dd,dd...
                                            string[] strSizes = memoryMergeCell.Split(',');

                                            for (int s = 0; s < strSizes.Length; s++)
                                            {
                                                strSizes[s] = strSizes[s].Trim();
                                                dtProduct.Rows.Add();
                                                tabs[k].listExcelIndexTab.Add(i);
                                                int lastRow = dtProduct.Rows.Count - 1;
                                                tabs[k].listdtProductIndexRow.Add(lastRow);

                                                dtProduct.Rows[lastRow]["Название"] = tabs[k].Name;
                                                dtProduct.Rows[lastRow]["Тип"] = tabs[k].Type;
                                                if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                                    dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).Match(strSizes[s]).Value;
                                            }

                                            #endregion
                                        }
                                    }
                                }
                            }

                        }
                    }
                    int curExcelRow = 0;
                    for (int k = 0; k < tabs.Count; k++)
                    {
                        for (int curCol = tabs[k].StartCol - 2; curCol < tabs[k].StartCol + 5; curCol++)
                        {
                            string memoryMergePrice = "";
                            string memoryMergeGost = "";
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].StartRow, curCol];
                            if (cellRange.Value != null)
                            {
                                temp = cellRange.Value.ToString().Trim();
                                if (new Regex(@"наруж.*размер", RegexOptions.IgnoreCase).IsMatch(temp))  //типа цена
                                {
                                    for (int i = 0; i < tabs[k].listExcelIndexTab.Count; i++) //строки
                                    {
                                        curExcelRow = tabs[k].listExcelIndexTab[i];
                                        cellRange = (Excel.Range)excelworksheet.Cells[curExcelRow, curCol];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Диаметр (высота), мм"] = regDiam.Match(temp).Value;
                                            //dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = regTolsh.Match(temp).Value;
                                            dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Метраж, м (длина, мм)"] = regMetr.Match(temp).Value;
                                            dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Примечание"] = temp;
                                        }
                                    }
                                }

                                if (new Regex(@"Цена", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    string mark1 = "", mark2 = "";
                                    cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].StartRow + 1, curCol];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        mark1 = temp;
                                    }

                                    cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].StartRow + 1, curCol + 1];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        mark2 = temp;
                                    }


                                    for (int i = 0; i < tabs[k].listdtProductIndexRow.Count; i++) //строки
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[i], curCol];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            memoryMergePrice = temp;
                                        }
                                        GetRegexPriceFromString(memoryMergePrice, tabs[k].listdtProductIndexRow[i]);
                                        dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Марка"] = mark1;
                                    }
                                    #region вторая цена
                                    //int p = 0;
                                    //while (p < tabs[k].listExcelIndexTab.Count) //строки
                                    //{
                                    //    cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[p], curCol + 1];
                                    //    if (cellRange.Value != null)
                                    //        temp = cellRange.Value.ToString().Trim();
                                    //    else temp = "";
                                    //    if (temp != "")
                                    //    {
                                    //        memoryMergePrice = temp;
                                    //    }
                                    //    int lastRow = tabs[k].listdtProductIndexRow[p];
                                    //    DataRow row = dtProduct.NewRow();
                                    //    row["Название"] = dtProduct.Rows[lastRow]["Название"];
                                    //    row["Тип"] = dtProduct.Rows[lastRow]["Тип"];
                                    //    row["Диаметр (высота), мм"] = dtProduct.Rows[lastRow]["Диаметр (высота), мм"];
                                    //    row["Толщина (ширина), мм"] = dtProduct.Rows[lastRow]["Толщина (ширина), мм"];
                                    //    row["Марка"] = dtProduct.Rows[lastRow]["Марка"];
                                    //    row["Стандарт"] = mark2;// dtProduct.Rows[lastRow]["Стандарт"];
                                    //    row["Примечание"] = dtProduct.Rows[lastRow]["Примечание"];
                                    //    row["Мерность (т, м, мм)"] = dtProduct.Rows[lastRow]["Мерность (т, м, мм)"];
                                    //    row["Цена"] = memoryMergePrice;// dtProduct.Rows[lastRow]["Цена"];
                                    //    int countdoubles = 0;
                                    //    for (int i = p; i < tabs[k].listExcelIndexTab.Count - 1; i++)
                                    //        if (tabs[k].listExcelIndexTab[i] == tabs[k].listExcelIndexTab[i + 1])
                                    //        {
                                    //            countdoubles++;
                                    //        }
                                    //        else break;
                                    //    dtProduct.Rows.InsertAt(row, lastRow);
                                    //    tabs[k].listExcelIndexTab.Insert(p, tabs[k].listExcelIndexTab[p]);
                                    //    tabs[k].listdtProductIndexRow.Insert(p, lastRow);
                                    //    p += countdoubles+1;


                                    //}
                                    #endregion

                                }
                                if (new Regex(@"гост", RegexOptions.IgnoreCase).IsMatch(temp))
                                {

                                    for (int i = 0; i < tabs[k].listdtProductIndexRow.Count; i++) //строки
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[tabs[k].listExcelIndexTab[i], curCol];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            memoryMergeGost = temp;
                                        }
                                        dtProduct.Rows[tabs[k].listdtProductIndexRow[i]]["Стандарт"] = memoryMergeGost;
                                    }
                                }
                            }
                        }


                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        else tsPb1.Value = tsPb1.Maximum;

                    }
                }
                clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции MetallInvest\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла МеталлИнвест до 2018г
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void MetallInvestOld2(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+(?:г\.?)?\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?)").Match(Path.GetFileName(filePath)).Value;
                if (new Regex(@"метал.*инвест", RegexOptions.IgnoreCase).IsMatch(orgname))
                    orgname = "МеталлИнвест";
                textBoxOrgName.Text = orgname;

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла MetallInvest\n\n" + ex.ToString()); }

                string temp = "";
                //this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                List<structTab> tabs = new List<structTab>();

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    structTab tab = new structTab();
                    tab.listdtProductIndexRow = new List<int>();
                    tab.listExcelIndexTab = new List<int>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    Regex regName = new Regex(@"двутавр\b|лента\b|лист\b|арматура\b|полоса\b|угол(?:ок)?\b|швеллер\w?\b|труб\w?\b|круг\w?\b|шестигранник|шгр\b|квадрат\b|сталь\b|катанка|быстрорез|колесо|заготовка|блок\b|^\s*вал\s*|втулка|поковка|шпунт|шахтная\s+стойка|штрипс", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой(?:\s*проч)|\w+ий|г[\/]к|(?<=\d)[бмшку]|р[\/]п|рифл\w*|(?:\s*,?\s*[АA][54]00(?:\w+)?)+", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"(?<=\s+)\d+(?:[,.]\d+)?(?=мм|\s|$|\s*[xх*]\s*\d+(?:[,\.]\d+)?\s*мм|[бмшку]|\()|(?<=[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase);
                    Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"(?:\d{,3}[ШСТУ]+\d{,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+\d{,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+\d{,3})(?=\s+|$)|(?:(?:Ст.)|(?:ст.)(?:\s*\d{1,2})\b)(?=\s+|$)|(?:\d{,3}[ХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)(?=\s+|$)|(?:[АA]-?\d)(?=\s+|$)|(?<=[Сс][Тт]\.\s?)\d{1,2}[гГ]\d{1,2}[cCсС]|(?<=ст\.)\d{1,2}[хфа]+(?=\s|$)|(?<=\s)[сС][тТ]\.?\s?\d{1,2}[_\w]+", RegexOptions.IgnoreCase);
                    Regex regSklad = new Regex(@"(?<=\w*\s*склад.*\s*:\s*)\w+\.\s*\w+,\s*\w+(?:\s*\w+)?,\s*\d+\.\s*.Металл-база.", RegexOptions.IgnoreCase);

                    int lastRow = 0;
                    int ColNomenkl = 1;
                    nameProd = "";
                    string type = "", mera = "", tmp = "";
                    string[] diam;
                    List<int> priceList = new List<int>();
                    List<string> markName = new List<string>();
                    List<int> diamList = new List<int>();

                    tsLabelClearingTable.Text = "Поиск заголовков";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;

                    #region поиск заголовков
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        int memoryJ = j;
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            priceList.Clear();
                            markName.Clear();
                            int ii = i;
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[memoryJ, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                #region Склад

                                if (regSklad.IsMatch(temp))
                                {
                                    Regex sklad = new Regex(@"(?<=\w*\s*склад.*\s*:\s*)\w+\.\s*\w+,\s*\w+(?:\s*\w+)?,\s*\d+\.\s*.Металл-база.", RegexOptions.IgnoreCase);
                                    tmp = sklad.Match(temp).Value;
                                    ListViewItem lvi = new ListViewItem(tmp);
                                    bool isIn = false;
                                    if (listViewAdrSklad.Items.Count < 1) listViewAdrSklad.Items.Add(lvi);
                                    for (int skl = 0; skl < listViewAdrSklad.Items.Count; skl++)
                                    {
                                        if (listViewAdrSklad.Items[skl].SubItems[0].Text == tmp) isIn = true;
                                    }
                                    if (!isIn) listViewAdrSklad.Items.Add(lvi);
                                }

                                #endregion

                                InfoOrganization(temp);
                                if (new Regex(@"Цен.", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColNomenkl = ii - 1;
                                    i = ColNomenkl + cellRange.MergeArea.Columns.Count + 1;
                                    for (int jj = memoryJ + 1; jj <= cCelRow; jj++)
                                    {
                                        for (int k = ColNomenkl; k < i; k++)
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[jj, k];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            else temp = "";
                                            if (temp != "")
                                            {
                                                if (regMark.IsMatch(temp))
                                                {
                                                    priceList.Add(k);
                                                    markName.Add(temp);
                                                }
                                            }
                                        }
                                        if (priceList.Count > 1)
                                        {
                                            j = jj + 1;
                                            break;
                                        }
                                    }
                                    for (int jj = j; jj < cCelRow; jj++)
                                    {
                                        diamList.Clear();
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, ColNomenkl];
                                        if (cellRange.Value != null)
                                            temp = cellRange.Value.ToString().Trim();
                                        else temp = "";
                                        if (temp != "")
                                        {
                                            if (regName.IsMatch(temp))
                                            {
                                                nameProd = regName.Match(temp).Value;
                                                type = regType.Match(temp).Value;
                                                if (new Regex(@"\d\s*-\s*\d", RegexOptions.IgnoreCase).IsMatch(temp))
                                                {
                                                    diam = new Regex(@"\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value.Split('-');
                                                    if (Convert.ToInt32(diam[1]) - Convert.ToInt32(diam[0]) == 1)
                                                        for (int d = Convert.ToInt32(diam[0]); d <= Convert.ToInt32(diam[1]); d += 1)
                                                        {
                                                            diamList.Add(d);
                                                        }
                                                    else
                                                        for (int d = Convert.ToInt32(diam[0]); d <= Convert.ToInt32(diam[1]); d += 2)
                                                        {
                                                            diamList.Add(d);
                                                        }
                                                }
                                                else diamList.Add(1);
                                                foreach (int intDiam in diamList)
                                                    for (int k = 0; k < priceList.Count; k++)
                                                    {
                                                        dtProduct.Rows.Add();
                                                        lastRow = dtProduct.Rows.Count - 1;
                                                        tab.listExcelIndexTab.Add(j);
                                                        tab.listdtProductIndexRow.Add(lastRow);

                                                        if (new Regex(@"ш\s*г\s*[рp]", RegexOptions.IgnoreCase).IsMatch(nameProd))
                                                            dtProduct.Rows[lastRow]["Название"] = "Шестигранник";
                                                        else
                                                            dtProduct.Rows[lastRow]["Название"] = nameProd;

                                                        dtProduct.Rows[lastRow]["Примечание"] = temp;
                                                        if (new Regex(@"г[\/]к", RegexOptions.IgnoreCase).IsMatch(type)) type = "горячекатаный";
                                                        if (new Regex(@"р[\/]п", RegexOptions.IgnoreCase).IsMatch(type)) type = "равнополочный";
                                                        if (new Regex(@"рифл", RegexOptions.IgnoreCase).IsMatch(type)) type = "рифленый";
                                                        dtProduct.Rows[lastRow]["Тип"] = type;
                                                        if (dtProduct.Rows[lastRow]["Тип"].ToString() == "") dtProduct.Rows[lastRow]["Тип"] = "тип не указан";


                                                        if (diamList.Count == 1)
                                                        {
                                                            dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDiam.Match(temp).Value;
                                                            dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\s+)\d+(?:[,\.]\d+)?(?=\s*[xх*]\s*\d+(?:[,\.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                            dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=[xх*]\s*\d+(?:[,\.]\d+)?\s*[xх*])\d+(?:[,\.]\d+)?(?=\s|\s*мм)", RegexOptions.IgnoreCase).Match(temp).Value;

                                                            if (nameProd.ToLower() == "уголок")
                                                            {
                                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=\s+)\d+(?:[,\.]\d+)?(?=\s*[xх*]\s*\d+(?:[,\.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\d+(?:[,\.]\d+)?\s*[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=\d+(?:[,\.]\d+)?\s*[xх*])\d+(?:[,\.]\d+)?(?=\s*[xх*]\s*)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                            }
                                                        }
                                                        else
                                                        {
                                                            dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = intDiam;
                                                        }

                                                        dtProduct.Rows[lastRow]["Мерность (т, м, мм)"] = mera;

                                                        dtProduct.Rows[lastRow]["Марка"] = markName[k];
                                                        Excel.Range cellRangePrice = (Excel.Range)excelworksheet.Cells[jj, priceList[k]];
                                                        if (cellRangePrice.Value != null)
                                                            tmp = cellRangePrice.Value.ToString().Trim();
                                                        else tmp = "";
                                                        if (tmp != "")
                                                        {
                                                            dtProduct.Rows[lastRow]["Цена"] = new Regex(@"(?<=^\s*)\d+(?:[,\.]\d+)?(?=\s*$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                        }
                                                    }
                                            }
                                        }
                                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                                        else tsPb1.Value = tsPb1.Maximum;
                                    }

                                }
                                j = memoryJ;
                            }
                        }
                    }
                    #endregion
                }


                clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции Apogei\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла МеталлКомплект
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void MettallKomplekt(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                SetNameFromName(filePath);

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла MettallKomplekt\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    structTab tab = new structTab();
                    tab.listdtProductIndexRow = new List<int>();
                    tab.listExcelIndexTab = new List<int>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;
                    tsPb1.Maximum = cCelCol * cCelRow;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    tsLabelClearingTable.Text = "Поиск наименований";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;

                    Regex regName = new Regex(@"(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой(?:\s*проч)|\w+ий", RegexOptions.IgnoreCase);
                    Regex regSklad = new Regex(@"(?<=\w*\s*склад.*\s*:\s*)\w+\.\s*\w+,\s*\w+(?:\s*\w+)?,\s*\d+\.\s*.Металл-база.", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"(?<=\s)\d+(?:[,.]\d+)?(?=\s|[бшкм])", RegexOptions.IgnoreCase);
                    Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"\bст[\s\.]*[\dгсп]+(?:\s*-\s*мд)?|[aа]\s*-\s*i{1,3}(?:\s*-\s*мд)?|[aа]т?\d+[cс]?(?:\s*-\s*мд)?|(?<=\d)[бмшк]\d?", RegexOptions.IgnoreCase);

                    int lastRow = 0, columnForOstatok = 0, columnForRezerv = 0;
                    string tmp = "";

                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"металл", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    tab.StartRow = j;

                                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                                    tsPb1.Value = 0;
                                    tsPb1.Maximum = cCelRow - j;

                                    for (int jj = j + 1; jj < cCelRow; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                                        if (cellRange.Value != null)
                                            tmp = cellRange.Value.ToString().Trim();
                                        if (regName.IsMatch(tmp))
                                        {
                                            dtProduct.Rows.Add();
                                            lastRow = dtProduct.Rows.Count - 1;
                                            tab.listExcelIndexTab.Add(jj);
                                            tab.listdtProductIndexRow.Add(lastRow);
                                            dtProduct.Rows[lastRow]["Название"] = regName.Match(tmp).Value;
                                            dtProduct.Rows[lastRow]["Примечание"] = tmp;
                                            dtProduct.Rows[lastRow]["Тип"] = regType.Match(tmp).Value;
                                            if (dtProduct.Rows[lastRow]["Тип"].ToString() == "") dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                            if (regDiam.IsMatch(tmp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDiam.Match(tmp).Value;
                                            }
                                            else if (new Regex(@"\d+(?:[,.]\d+)?(?:\s*[хx]\s*\d+(?:[,.]\d+)?)+", RegexOptions.IgnoreCase).IsMatch(tmp))
                                            {
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "лист")
                                                {
                                                    dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\s)\d+(?:[,.]\d+)?(?=\s*[хx]\s*\d+)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx]\s*)\d+(?:[,.]\d+)?(?=\s*[хx]\s*\d+(?:[,.]\d+)?|\s)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx]\s*)\d+(?:[,.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                }
                                                else
                                                {
                                                    dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"\d+(?:[,.]\d+)?(?=\s*[хx]\s*\d+)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx]\s*)\d+(?:[,.]\d+)?(?=\s*[хx]\s*\d+(?:[,.]\d+)?|\s)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx]\s*)\d+(?:[,.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                }
                                            }
                                            dtProduct.Rows[lastRow]["Марка"] = regMark.Match(tmp).Value;
                                            dtProduct.Rows[lastRow]["Стандарт"] = regTU.Match(tmp).Value;

                                            if (dtProduct.Rows[lastRow]["Стандарт"].ToString() == "")
                                            {
                                                tmp = "";
                                                cellRange = (Excel.Range)excelworksheet.Cells[jj, i + 2];
                                                if (cellRange.Value != null)
                                                    tmp = cellRange.Value.ToString().Trim();
                                                if (tmp != "")
                                                {
                                                    dtProduct.Rows[lastRow]["Стандарт"] = regTU.Match(tmp).Value;
                                                }
                                            }

                                            if (dtProduct.Rows[lastRow]["Марка"].ToString() == "")
                                            {
                                                tmp = "";
                                                cellRange = (Excel.Range)excelworksheet.Cells[jj, i + 2];
                                                if (cellRange.Value != null)
                                                    tmp = cellRange.Value.ToString().Trim();
                                                if (tmp != "")
                                                {
                                                    dtProduct.Rows[lastRow]["Марка"] = regMark.Match(tmp).Value;
                                                }
                                            }

                                            tmp = "";
                                            cellRange = (Excel.Range)excelworksheet.Cells[jj + 1, i];
                                            if (cellRange.Value != null)
                                                tmp = cellRange.Value.ToString().Trim();
                                            if (tmp == "")
                                            {
                                                cellRange = (Excel.Range)excelworksheet.Cells[jj + 1, i + 2];
                                                if (cellRange.Value != null)
                                                    tmp = cellRange.Value.ToString().Trim();
                                                if (tmp != "")
                                                {
                                                    dtProduct.Rows.Add();
                                                    lastRow = dtProduct.Rows.Count - 1;
                                                    tab.listExcelIndexTab.Add(jj);
                                                    tab.listdtProductIndexRow.Add(lastRow);
                                                    dtProduct.Rows[lastRow]["Название"] = dtProduct.Rows[lastRow - 1]["Название"];
                                                    dtProduct.Rows[lastRow]["Тип"] = dtProduct.Rows[lastRow - 1]["Тип"];
                                                    dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = dtProduct.Rows[lastRow - 1]["Диаметр (высота), мм"];
                                                    dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = dtProduct.Rows[lastRow - 1]["Толщина (ширина), мм"];
                                                    dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = dtProduct.Rows[lastRow - 1]["Метраж, м (длина, мм)"];
                                                    dtProduct.Rows[lastRow]["Стандарт"] = regTU.Match(tmp).Value;
                                                }

                                                cellRange = (Excel.Range)excelworksheet.Cells[jj + 1, i + 2];
                                                if (cellRange.Value != null)
                                                    tmp = cellRange.Value.ToString().Trim();
                                                if (tmp != "")
                                                {
                                                    cellRange = (Excel.Range)excelworksheet.Cells[jj + 1, i + 2];
                                                    if (cellRange.Value != null)
                                                        tmp = cellRange.Value.ToString().Trim();
                                                    dtProduct.Rows[lastRow]["Марка"] = regMark.Match(tmp).Value;
                                                    jj++;
                                                }
                                            }
                                        }
                                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                                        else tsPb1.Value = tsPb1.Maximum;
                                    }
                                }

                                if (dtProduct.Rows.Count > 0)
                                {
                                    if (new Regex(@"цена.*5.*до.*25.*тн", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        tmp = "";
                                        tsLabelClearingTable.Text = "Поиск цен";
                                        tsPb1.Value = 0;
                                        tsPb1.Maximum = tab.listExcelIndexTab.Count;

                                        for (int jj = 0; jj < tab.listdtProductIndexRow.Count; jj++)
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[tab.listExcelIndexTab[jj], i];
                                            if (cellRange.Value != null)
                                                tmp = cellRange.Value.ToString().Trim();
                                            else tmp = "";
                                            if (tmp != "")
                                            {
                                                dtProduct.Rows[tab.listdtProductIndexRow[jj]]["Цена"] = tmp;
                                            }
                                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                                            else tsPb1.Value = tsPb1.Maximum;
                                        }
                                    }
                                    if (new Regex(@"остаток", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        columnForOstatok = i;

                                        tsLabelClearingTable.Text = "Поиск резерва и остатков";
                                        tsPb1.Value = 0;
                                        tsPb1.Maximum = tab.listExcelIndexTab.Count;

                                        for (int ii = i + 1; ii < cCelCol; ii++)
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[j, ii];
                                            if (cellRange.Value != null)
                                                temp = cellRange.Value.ToString().Trim();
                                            if (new Regex(@"резерв", RegexOptions.IgnoreCase).IsMatch(temp))
                                            {
                                                columnForRezerv = ii;
                                                break;
                                            }
                                        }
                                        if (columnForOstatok > 0 && columnForRezerv > 0)
                                        {
                                            double ostatok = -10000, rezerv = -10000;

                                            for (int jj = 0; jj < tab.listdtProductIndexRow.Count; jj++)
                                            {
                                                try
                                                {
                                                    tmp = "";
                                                    cellRange = (Excel.Range)excelworksheet.Cells[tab.listExcelIndexTab[jj], columnForOstatok];
                                                    if (cellRange.Value != null)
                                                    {
                                                        tmp = cellRange.Text.ToString().Trim();
                                                        if (tmp != "")
                                                            ostatok = Convert.ToDouble(cellRange.Text.ToString().Trim());
                                                    }
                                                    else ostatok = 0;

                                                    tmp = "";
                                                    cellRange = (Excel.Range)excelworksheet.Cells[tab.listExcelIndexTab[jj], columnForRezerv];
                                                    if (cellRange.Value != null)
                                                    {
                                                        tmp = cellRange.Text.ToString().Trim();
                                                        if (tmp != "")
                                                            rezerv = Convert.ToDouble(cellRange.Text.ToString().Trim());
                                                    }
                                                    else rezerv = 0;
                                                }
                                                catch
                                                {
                                                    ostatok = 0; rezerv = 0;
                                                    //MessageBox.Show("Не удалось преобразовать стринги в инты");
                                                }
                                                if (ostatok != -10000 && rezerv != -10000)
                                                {
                                                    dtProduct.Rows[tab.listdtProductIndexRow[jj]]["Мерность (т, м, мм)"] = rezerv + ostatok;
                                                }
                                                if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                                                else tsPb1.Value = tsPb1.Maximum;
                                            }
                                        }
                                    }
                                    if (j > tab.StartRow + 2) j = cCelRow;
                                }
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }


                }

                for (int i = 0; i < dtProduct.Rows.Count; i++)
                {
                    if (dtProduct.Rows[i]["Мерность (т, м, мм)"].ToString() == "" | dtProduct.Rows[i]["Мерность (т, м, мм)"].ToString() == "0")
                    {
                        if (dtProduct.Rows[i]["Цена"].ToString() == "" && dtProduct.Rows[i]["Стандарт"].ToString() == "")
                            dtProduct.Rows[i].Delete();
                    }
                }

                clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции MettallKomplekt\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла УСТЧ
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void USTCH(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                SetNameFromName(filePath);

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла MettallKomplekt\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    structTab tab = new structTab();
                    tab.listdtProductIndexRow = new List<int>();
                    tab.listExcelIndexTab = new List<int>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;
                    tsPb1.Maximum = cCelCol * cCelRow;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    tsLabelClearingTable.Text = "Поиск наименований";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;

                    Regex regName = new Regex(@"лента|лист|арматура|полоса|угол|швеллер|труб", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой(?:\s*проч)|\w+ий", RegexOptions.IgnoreCase);
                    Regex regSklad = new Regex(@"(?<=\w*\s*склад.*\s*:\s*)\w+\.\s*\w+,\s*\w+(?:\s*\w+)?,\s*\d+\.\s*.Металл-база.", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"(?<=\s)\d+(?:[,.]\d+)?(?=\s|[бшкм])", RegexOptions.IgnoreCase);
                    Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"\bст[\s\.]*[\dгсп]+(?:\s*-\s*мд)?|[aа]\s*-\s*i{1,3}(?:\s*-\s*мд)?|[aа]т?\d+[cс]?(?:\s*-\s*мд)?|(?<=\d)[бмшк]\d?", RegexOptions.IgnoreCase);

                    int lastRow = 0;
                    string tmp = "";

                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"наимен", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    tab.StartRow = j;

                                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                                    tsPb1.Value = 0;
                                    tsPb1.Maximum = cCelRow - j;

                                    for (int jj = j + 1; jj < cCelRow; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                                        if (cellRange.Value != null)
                                            tmp = cellRange.Value.ToString().Trim();
                                        if (regName.IsMatch(tmp) || new Regex(@"\d+(?:[,.]\d+)?(?:\s*[хx\*]\s*\d+(?:[,.]\d+)?)+", RegexOptions.IgnoreCase).IsMatch(tmp) || regDiam.IsMatch(tmp))
                                        {
                                            dtProduct.Rows.Add();
                                            lastRow = dtProduct.Rows.Count - 1;
                                            tab.listExcelIndexTab.Add(jj);
                                            tab.listdtProductIndexRow.Add(lastRow);
                                            dtProduct.Rows[lastRow]["Название"] = regName.Match(tmp).Value;
                                            if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "труб") dtProduct.Rows[lastRow]["Название"] = "Труба";
                                            if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "угол") dtProduct.Rows[lastRow]["Название"] = "Уголок";
                                            if (dtProduct.Rows[lastRow]["Название"].ToString().Length > 1)
                                                dtProduct.Rows[lastRow]["Название"] = dtProduct.Rows[lastRow]["Название"].ToString().Substring(0, 1).ToUpper() + dtProduct.Rows[lastRow]["Название"].ToString().Substring(1, dtProduct.Rows[lastRow]["Название"].ToString().Length - 1);

                                            if (dtProduct.Rows[lastRow]["Название"].ToString() == "")
                                            {
                                                if (dtProduct.Rows.Count > 1)
                                                    if (dtProduct.Rows[lastRow - 1]["Название"].ToString() != "")
                                                        dtProduct.Rows[lastRow]["Название"] = dtProduct.Rows[lastRow - 1]["Название"];
                                            }

                                            dtProduct.Rows[lastRow]["Примечание"] = tmp;
                                            dtProduct.Rows[lastRow]["Тип"] = regType.Match(tmp).Value;
                                            if (dtProduct.Rows[lastRow]["Тип"].ToString() == "") dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                            if (regDiam.IsMatch(tmp))
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDiam.Match(tmp).Value;
                                            }
                                            else if (new Regex(@"\d+(?:[,.]\d+)?(?:\s*[хx\*]\s*\d+(?:[,.]\d+)?)+", RegexOptions.IgnoreCase).IsMatch(tmp))
                                            {
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "лист" || dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "полоса")
                                                {
                                                    dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=^|\s)\d+(?:[,.]\d+)?(?=\s*[хx\*]\s*\d+)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s*[хx\*]\s*\d+(?:[,.]\d+)?|\s)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                }
                                                else if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "уголок")
                                                {
                                                    dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=^|\s)\d+(?:[,.]\d+)?(?=\s*[хx\*]\s*\d+)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\s*[хx\*]\s*\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s*[хx\*])", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                }
                                                else
                                                {
                                                    dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=^|\s)\d+(?:[,.]\d+)?(?=\s*[хx\*]\s*\d+)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s*[хx\*]|\s)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                }
                                            }
                                            dtProduct.Rows[lastRow]["Марка"] = regMark.Match(tmp).Value;

                                            foreach (Match m in regTU.Matches(tmp))
                                            {
                                                if (dtProduct.Rows[lastRow]["Стандарт"].ToString() == "") dtProduct.Rows[lastRow]["Стандарт"] = m.Value;
                                                else dtProduct.Rows[lastRow]["Стандарт"] += "; " + m.Value;
                                            }

                                        }
                                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                                        else tsPb1.Value = tsPb1.Maximum;
                                    }
                                }

                                if (dtProduct.Rows.Count > 0)
                                {
                                    if (new Regex(@"цена", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        tmp = "";
                                        tsLabelClearingTable.Text = "Поиск цен";
                                        tsPb1.Value = 0;
                                        tsPb1.Maximum = tab.listExcelIndexTab.Count;

                                        for (int jj = 0; jj < tab.listdtProductIndexRow.Count; jj++)
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[tab.listExcelIndexTab[jj], i];
                                            if (cellRange.Value != null)
                                                tmp = cellRange.Value.ToString().Trim();
                                            else tmp = "";
                                            if (tmp != "")
                                            {
                                                dtProduct.Rows[tab.listdtProductIndexRow[jj]]["Цена"] = tmp;
                                            }
                                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                                            else tsPb1.Value = tsPb1.Maximum;
                                        }
                                    }
                                    if (new Regex(@"кол.*во", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {

                                        tsLabelClearingTable.Text = "Поиск резерва и остатков";
                                        tsPb1.Value = 0;
                                        tsPb1.Maximum = tab.listExcelIndexTab.Count;

                                        for (int jj = 0; jj < tab.listdtProductIndexRow.Count; jj++)
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[tab.listExcelIndexTab[jj], i];
                                            if (cellRange.Value != null)
                                                tmp = cellRange.Value.ToString().Trim();
                                            else tmp = "";
                                            if (tmp != "")
                                            {
                                                dtProduct.Rows[tab.listdtProductIndexRow[jj]]["Мерность (т, м, мм)"] = tmp;
                                            }
                                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                                            else tsPb1.Value = tsPb1.Maximum;
                                        }
                                    }
                                    if (j > tab.StartRow + 2) j = cCelRow;
                                }
                                Regex regOrgAdr = new Regex(@"Адрес\s*:\s*\w+,(?:\s*\d+\s*,)?(?:\s*[\w+\s*]+,)?\s*[\w+\s*]+,\s*[\w+\s*]+,\s*\d+(?:,\s*офис\s*[\d\w]+)?", RegexOptions.IgnoreCase);
                                if (regOrgAdr.IsMatch(temp))
                                    textBoxOrgAdress.Text = regOrgAdr.Match(temp).Value;

                                Regex regOrgTel = new Regex(@"Телефоны\s*:\s*(?:\d*\s*\(\d+\)\s*)?(\s*\d+(?:-\d+)+,?\s*)+", RegexOptions.IgnoreCase);
                                if (regOrgTel.IsMatch(temp))
                                    textBoxOrgTelefon.Text = regOrgTel.Match(temp).Value;

                                Regex regManager = new Regex(@"контактное\s*лицо", RegexOptions.IgnoreCase);
                                if (regManager.IsMatch(temp))
                                {
                                    ListViewItem lvi = new ListViewItem(new Regex(@"(?<=лицо\s*:\s*)(?:\w+\s*)+(?=,)").Match(temp).Value); //имя менеджера
                                    lvi.SubItems.Add(new Regex(@"телефон\s*:\s*.*\d\d(?=,)").Match(temp).Value);           //телефон менеджера
                                    if (lvi.SubItems[0].Text != "icq") listViewManager.Items.Add(lvi);
                                }

                                Regex regMail = new Regex(@"(?<=эл.*почта\s*:\s*)[\w\d\.-]+@[\w\d\.-]+.(?:ru|com|рф|info)", RegexOptions.IgnoreCase);
                                if (regMail.IsMatch(temp))
                                    textBoxOrgEmail.Text = regMail.Match(temp).Value;
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }


                }

                clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции MettallKomplekt\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла MetallServisCentr
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void MetallServisCentr(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?)").Match(Path.GetFileName(filePath)).Value;
                if (new Regex(@"МСЦ", RegexOptions.IgnoreCase).IsMatch(orgname))
                    orgname = "МеталлСервисЦентр";
                textBoxOrgName.Text = orgname;

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла MetallServisCentr\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    structTab tab = new structTab();
                    tab.listdtProductIndexRow = new List<int>();
                    tab.listExcelIndexTab = new List<int>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    Regex regName = new Regex(@"лента|лист|арматура|полоса|угол|швеллер|труб|круг|шестигранник|шгр", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой(?:\s*проч)|\w+ий", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"(?<=ф)\d+(?:[,.]\d+)?(?=\s|-|$)|(?<=^)\d+(?:[,.]\d+)?(?:\s*[хx*]\s*\d+(?:[,.]\d+)?)+", RegexOptions.IgnoreCase);
                    Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"\bст[\s\.]*[\dгсп]+(?:\s*-\s*мд)?|[aа]\s*-\s*i{1,3}(?:\s*-\s*мд)?|[aа]т?\d+[cс]?(?:\s*-\s*мд)?|(?<=\d)[бмшк]\d?", RegexOptions.IgnoreCase);
                    string type = "";
                    string mark = "";

                    int lastRow = 0;
                    string tmp = "";
                    string stmp = "";
                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;
                    for (int i = 1; i <= cCelCol; i++) //столбцы
                    {
                        for (int j = 1; j <= cCelRow; j++) //строки
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (regDiam.IsMatch(temp))
                                {
                                    tab.StartRow = j;
                                    cellRange = (Excel.Range)excelworksheet.Cells[j - 1, i];
                                    if (cellRange.MergeArea.Count > 1)
                                    {
                                        for (int ma = i - 1; ma <= i + 1; ma++)
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[j - 1, ma];
                                            if (cellRange.Value != null)
                                                tmp = cellRange.Value.ToString().Trim();
                                            if (tmp != "")
                                            {
                                                nameProd = regName.Match(tmp).Value;
                                                if (nameProd.ToLower() == "труб") nameProd = "Труба";
                                                if (nameProd.ToLower() == "угол") nameProd = "Уголок";
                                                if (nameProd.Length > 1)
                                                    nameProd = nameProd.Substring(0, 1).ToUpper() + nameProd.Substring(1, nameProd.Length - 1).ToLower();
                                                type = regType.Match(tmp).Value;
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[j - 1, i];
                                        if (cellRange.Value != null)
                                            tmp = cellRange.Value.ToString().Trim();
                                        else tmp = "";
                                        if (tmp != "")
                                        {
                                            nameProd = regName.Match(tmp).Value;
                                            if (nameProd.ToLower() == "труб") nameProd = "Труба";
                                            if (nameProd.ToLower() == "угол") nameProd = "Уголок";
                                            if (nameProd.Length > 1)
                                                nameProd = nameProd.Substring(0, 1).ToUpper() + nameProd.Substring(1, nameProd.Length - 1);
                                            type = regType.Match(tmp).Value;
                                        }
                                        else
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[j - 1, i - 1];
                                            if (cellRange.Value != null)
                                                tmp = cellRange.Value.ToString().Trim();
                                            if (tmp != "")
                                            {
                                                nameProd = regName.Match(tmp).Value;
                                                if (nameProd.ToLower() == "труб") nameProd = "Труба";
                                                if (nameProd.ToLower() == "угол") nameProd = "Уголок";
                                                if (nameProd.Length > 1)
                                                    nameProd = nameProd.Substring(0, 1).ToUpper() + nameProd.Substring(1, nameProd.Length - 1);
                                                type = regType.Match(tmp).Value;
                                            }
                                            else
                                            {
                                                j = cCelRow;
                                                break;
                                            }
                                        }
                                    }

                                    for (int jj = j; jj <= cCelRow; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                                        if (cellRange.Value != null)
                                            tmp = cellRange.Value.ToString().Trim();
                                        else tmp = "";
                                        if (tmp != "")
                                        {
                                            if (regDiam.IsMatch(tmp) && nameProd != "")
                                            {
                                                dtProduct.Rows.Add();
                                                lastRow = dtProduct.Rows.Count - 1;
                                                tab.listExcelIndexTab.Add(jj);
                                                tab.listdtProductIndexRow.Add(lastRow);
                                                dtProduct.Rows[lastRow]["Название"] = regName.Match(tmp).Value;
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "труб") dtProduct.Rows[lastRow]["Название"] = "Труба";
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "угол") dtProduct.Rows[lastRow]["Название"] = "Уголок";
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().Length > 1)
                                                    dtProduct.Rows[lastRow]["Название"] = dtProduct.Rows[lastRow]["Название"].ToString().Substring(0, 1).ToUpper() + dtProduct.Rows[lastRow]["Название"].ToString().Substring(1, dtProduct.Rows[lastRow]["Название"].ToString().Length - 1);

                                                if (dtProduct.Rows[lastRow]["Название"].ToString() == "")
                                                {
                                                    dtProduct.Rows[lastRow]["Название"] = nameProd;
                                                }

                                                dtProduct.Rows[lastRow]["Примечание"] = tmp;
                                                dtProduct.Rows[lastRow]["Тип"] = regType.Match(tmp).Value;

                                                if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                                {
                                                    dtProduct.Rows[lastRow]["Тип"] = type;
                                                    if (dtProduct.Rows[lastRow]["Тип"].ToString() == "") dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                                }


                                                //ищет марку в ячейке слева
                                                cellRange = (Excel.Range)excelworksheet.Cells[jj, i - 1];
                                                if (cellRange.Value != null)
                                                    stmp = cellRange.Value.ToString().Trim();
                                                else stmp = "";
                                                if (stmp != "")
                                                {
                                                    mark = stmp;//regMark.Match(tmp).Value;
                                                    dtProduct.Rows[lastRow]["Марка"] = mark;
                                                }
                                                else
                                                {
                                                    dtProduct.Rows[lastRow]["Марка"] = mark;
                                                }

                                                //ищет цену в ячейке справа
                                                cellRange = (Excel.Range)excelworksheet.Cells[jj, i + 1];
                                                if (cellRange.Value != null)
                                                    stmp = cellRange.Value.ToString().Trim();
                                                else stmp = "";
                                                if (stmp != "")
                                                {
                                                    dtProduct.Rows[lastRow]["Цена"] = stmp.Trim();
                                                }

                                                foreach (Match m in regTU.Matches(tmp))
                                                {
                                                    if (dtProduct.Rows[lastRow]["Стандарт"].ToString() == "") dtProduct.Rows[lastRow]["Стандарт"] = m.Value;
                                                    else dtProduct.Rows[lastRow]["Стандарт"] += "; " + m.Value;
                                                }

                                                //ищет параметры изделия
                                                if (new Regex(@"(?<=ф)\d+(?:[,.]\d+)?(?=\s|-|$)", RegexOptions.IgnoreCase).IsMatch(tmp))
                                                {
                                                    if (new Regex(@"\d\s*-\s*\d", RegexOptions.IgnoreCase).IsMatch(tmp))
                                                    {
                                                        string[] parms = new Regex(@"\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).Match(tmp).Value.Split('-');
                                                        dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = parms[0];
                                                        DataRow row = dtProduct.NewRow();
                                                        row["Название"] = dtProduct.Rows[lastRow]["Название"];
                                                        row["Тип"] = dtProduct.Rows[lastRow]["Тип"];
                                                        row["Диаметр (высота), мм"] = parms[1];
                                                        row["Толщина (ширина), мм"] = dtProduct.Rows[lastRow]["Толщина (ширина), мм"];
                                                        row["Метраж, м (длина, мм)"] = dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"];
                                                        row["Мерность (т, м, мм)"] = dtProduct.Rows[lastRow]["Мерность (т, м, мм)"];
                                                        row["Марка"] = dtProduct.Rows[lastRow]["Марка"];
                                                        row["Стандарт"] = dtProduct.Rows[lastRow]["Стандарт"];
                                                        row["Класс"] = dtProduct.Rows[lastRow]["Класс"];
                                                        row["Цена"] = dtProduct.Rows[lastRow]["Цена"];
                                                        row["Примечание"] = dtProduct.Rows[lastRow]["Примечание"];
                                                        dtProduct.Rows.Add(row);
                                                    }
                                                    else if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "круг" || dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "шестигранник")
                                                    {
                                                        dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=ф)\d+(?:[,.]\d+)?(?=\s|-|&)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    }
                                                    else
                                                    {
                                                        dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=^|\s)\d+(?:[,.]\d+)?(?=\s*[хx\*]\s*\d+)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                        dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s*[хx\*]\s*\d+(?:[,.]\d+)?|\s)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                        dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    }
                                                }
                                                else
                                                {
                                                    dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=^|\s)\d+(?:[,.]\d+)?(?=.*[хx\*].*\d+)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=\d+(?:[,.]\d+)?.*[хx\*].*)\d+(?:[,.]\d+)?(?=.*[хx\*].*\d+(?:[,.]\d+)?|\s)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=\d+(?:[,.]\d+)?.*[хx\*].*)\d+(?:[,.]\d+)?(?=мм|\s|$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                }

                                            }
                                            else
                                            {
                                                j = jj - 1;
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            j = jj - 1;
                                            break;
                                        }
                                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                                        else tsPb1.Value = tsPb1.Maximum;
                                    }
                                }

                                Regex regOrgAdr = new Regex(@"^\s*\d+,?\s*г.*\w,\s*\w+.*(?:\w+|\d+)\b", RegexOptions.IgnoreCase);
                                if (regOrgAdr.IsMatch(temp))
                                    textBoxOrgAdress.Text = regOrgAdr.Match(temp).Value;

                                Regex regOrgTel = new Regex(@"(?<=тел.*)(?:\d\s)?(?:\(\d+\))?\s*\d+(?:-\d+)+", RegexOptions.IgnoreCase);
                                if (regOrgTel.IsMatch(temp))
                                    foreach (Match m in regOrgTel.Matches(temp))
                                        if (textBoxOrgTelefon.Text == "")
                                            textBoxOrgTelefon.Text = m.Value;
                                        else textBoxOrgTelefon.Text += "; " + m.Value;

                                Regex regManager = new Regex(@"контактное\s*лицо", RegexOptions.IgnoreCase);
                                if (regManager.IsMatch(temp))
                                {
                                    ListViewItem lvi = new ListViewItem(new Regex(@"(?<=лицо\s*:\s*)(?:\w+\s*)+(?=,)").Match(temp).Value); //имя менеджера
                                    lvi.SubItems.Add(new Regex(@"телефон\s*:\s*.*\d\d(?=,)").Match(temp).Value);           //телефон менеджера
                                    if (lvi.SubItems[0].Text != "icq") listViewManager.Items.Add(lvi);
                                }

                                Regex regSklad = new Regex(@"склад", RegexOptions.IgnoreCase);
                                if (regSklad.IsMatch(temp))
                                {
                                    ListViewItem lvi = new ListViewItem(new Regex(@"(?<=скл(?:ада)?\s*:\s*)\w.*,\s\d+(?:\s*\(.*\))?").Match(temp).Value); //имя менеджера
                                    listViewAdrSklad.Items.Add(lvi);
                                }


                                Regex regMail = new Regex(@"(?<=эл.*почта\s*:\s*)[\w\d\.-]+@[\w\d\.-]+.(?:ru|com|рф|info)", RegexOptions.IgnoreCase);
                                if (regMail.IsMatch(temp))
                                    textBoxOrgEmail.Text = regMail.Match(temp).Value;
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }


                }

                clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции MetallServisCentr\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла PromSnab
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void PromSnab(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                orgname = "ПромСнабжение";
                textBoxOrgName.Text = orgname;

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла PromSnab\n\n" + ex.ToString()); }

                string temp = "";
                this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    List<structTab> tabs = new List<structTab>();


                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;

                    tsLabelClearingTable.Text = "Поиск наименований";
                    tsPb1.Value = 0;

                    Regex regName = new Regex(@"лента|лист|арматура|полоса|угол|швеллер|труб|круг|шестигранник|шгр|квадрат|полоса|сталь|катанка", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой(?:\s*проч)|\w+ий", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"(?<=ф)\d+(?:[,.]\d+)?(?=\s|-|$)|(?<=^)\d+(?:[,.]\d+)?(?:\s*[хx*]\s*\d+(?:[,.]\d+)?)+", RegexOptions.IgnoreCase);
                    Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"(?:\d{,3}[ШСТУ]+\d{,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+\d{,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+\d{,3})(?=\s+|$)|(?:(?:Ст.)|(?:ст.)(?:\s*\d{1,2})\b)(?=\s+|$)|(?:\d{,3}[ХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)(?=\s+|$)|(?:[АA]-?\d)(?=\s+|$)|(?<=[Сс][Тт]\.\s?)\d{1,2}[гГ]\d{1,2}[cCсС]|(?<=ст\.)\d{1,2}[хфа]+(?=\s|$)|(?<=\s)[сС][тТ]\.?\s?\d{1,2}[_\w]+", RegexOptions.IgnoreCase);
                    Regex regManagerName = new Regex(@"(?<=[Тт]ел.*)(?:[А-Я]\w+\s*){2,3}");

                    int lastRow = 0;
                    string tmp = "";
                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;
                    List<int> milimetr = new List<int>();
                    List<int> colichestvo = new List<int>();

                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        int jj = j;
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"^mm|^мм", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    milimetr.Add(i);
                                    j = cCelRow;
                                }
                                if (temp.ToLower().Contains("кол") || temp.ToLower().Contains("кол-во"))
                                {
                                    colichestvo.Add(i);
                                }

                                if (regManagerName.IsMatch(temp))
                                {
                                    ListViewItem lvi = new ListViewItem(regManagerName.Match(temp).Value); //имя менеджера
                                    lvi.SubItems.Add(new Regex(@"(?<=[тТ]ел.*:\s)[\d-(),\s]+(?=,\s+[А-Я]\w+)").Match(temp).Value);           //телефон менеджера
                                    if (lvi.SubItems[0].Text != "icq") listViewManager.Items.Add(lvi);
                                }

                                if (new Regex(@"\d+,\s+г\.\w+(?:,[\s*\w\.\s\d]+)?,?\s*\d", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    textBoxOrgAdress.Text = new Regex(@"\d+,\s+г\.\w+(?:,[\s*\w\.\s\d]+)?,?\s*\d", RegexOptions.IgnoreCase).Match(temp).Value;
                                }

                                if (new Regex(@"(?<=Тел.*ф.*,\s*)\d+(?:-\d+)+", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    textBoxOrgTelefon.Text = new Regex(@"(?<=Тел.*ф.*,\s*)\d+(?:-\d+)+", RegexOptions.IgnoreCase).Match(temp).Value;
                                }

                                if (new Regex(@"[^-*_][\w\d-*_\.]+@[\w\d]+\.\w{2,4}", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    textBoxOrgEmail.Text = new Regex(@"[^-*_][\w\d-*_\.]+@[\w\d]+\.\w{2,4}", RegexOptions.IgnoreCase).Match(temp).Value;
                                }
                            }
                        }
                    }

                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {

                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.MergeArea.Columns.Count > 3)
                            {
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (temp != "")
                                {
                                    structTab tab = new structTab();
                                    tab.listdtProductIndexRow = new List<int>();
                                    tab.listExcelIndexTab = new List<int>();

                                    tab.StartRow = j + 1;
                                    tab.Name = regName.Match(temp).Value;
                                    tab.Type = regType.Match(temp).Value;
                                    if (tab.Name == "Сталь" && tab.Type == "угловая")
                                    { tab.Type = ""; tab.Name = "Уголок"; }
                                    tab.Standart = regMark.Match(temp).Value;

                                    tabs.Add(tab);
                                }
                            }

                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }

                    tsLabelClearingTable.Text = "Обработка";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = tabs.Count * milimetr.Count;
                    int endRow = 1;
                    for (int k = 0; k < tabs.Count; k++)
                    {
                        structTab tab = tabs[k];
                        if (k < tabs.Count - 1)
                            endRow = tabs[k + 1].StartRow - 1;
                        else endRow = cCelRow;
                        for (int mm = 0; mm < milimetr.Count; mm++)
                        {
                            for (int jj = tab.StartRow; jj < endRow; jj++)
                            {
                                Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[jj, milimetr[mm]];
                                if (cellRange.Value != null)
                                    temp = cellRange.Value.ToString().Trim();
                                else temp = "";
                                if (new Regex(@"\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    dtProduct.Rows.Add();
                                    lastRow = dtProduct.Rows.Count - 1;
                                    if (dtProduct.Rows[lastRow]["Название"].ToString().Length > 1)
                                        dtProduct.Rows[lastRow]["Название"] = tab.Name.Substring(0, 1).ToUpper() + tab.Name.Substring(1, tab.Name.Length - 1).ToLower(); ;
                                    dtProduct.Rows[lastRow]["Тип"] = tab.Type.ToLower();
                                    if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                    {
                                        dtProduct.Rows[lastRow]["Тип"] = new Regex(@"(?<=\d+(?:[,.]\d+)?)[пус][\sxх*]", RegexOptions.IgnoreCase).Match(temp).Value.ToUpper();
                                        if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                            dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                    }
                                    dtProduct.Rows[lastRow]["Стандарт"] = tab.Standart;
                                    dtProduct.Rows[lastRow]["Примечание"] = temp;
                                    tmp = temp;
                                    if (tab.Name.ToLower().Contains("сталь") && tab.Type.ToLower().Contains("угловая") || tab.Name.ToLower().Contains("уголок"))
                                    {
                                        dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=^|\s)\d+(?:[,.]\d+)?(?=\s*[хx\*]\s*\d+)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                        dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(tmp).Value;
                                        dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=L=\s*)\d+(?:[,.]\d+)?\s*м?(,\s*\d+(?:[,.]\d+)?\s*м?)*", RegexOptions.IgnoreCase).Match(tmp).Value;
                                    }
                                    else if (tab.Name.ToLower().Contains("полоса"))
                                    {
                                        dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=^|\s)\d+(?:[,.]\d+)?(?=\s*[хx\*]\s*\d+)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                        dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(tmp).Value;
                                        dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=L=\s*)\d+(?:[,.]\d+)?\s*м?(,\s*\d+(?:[,.]\d+)?\s*м?)*", RegexOptions.IgnoreCase).Match(tmp).Value;
                                    }
                                    else if (tab.Name.ToLower().Contains("лист"))
                                    {
                                        dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=^|\s)\d+(?:[,.]\d+)?(?=\s*[хx\*]\s*\d+)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                        dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s*[хx\*]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                        dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[хx\*]\s*)\d+(?:[,.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                    }
                                    else
                                    {
                                        dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                        dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=L=\s*)\d+(?:[,.]\d+)?\s*м?(,\s*\d+(?:[,.]\d+)?\s*м?)*", RegexOptions.IgnoreCase).Match(tmp).Value;
                                    }
                                    cellRange = (Excel.Range)excelworksheet.Cells[jj, colichestvo[mm]];
                                    if (cellRange.Value != null)
                                        temp = cellRange.Value.ToString().Trim();
                                    else temp = "";
                                    if (temp != "")
                                    {
                                        dtProduct.Rows[lastRow]["Мерность (т, м, мм)"] = new Regex(@"\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                    }
                                }
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }
                }

                //clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции PromSnab\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла MetallServisCentr
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void ChelyabinskProfit(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+(?:г\.?)?\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?)").Match(Path.GetFileName(filePath)).Value;
                textBoxOrgName.Text = orgname;

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла ChelyabinskProfit\n\n" + ex.ToString()); }

                string temp = "";
                //this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    structTab tab = new structTab();
                    tab.listdtProductIndexRow = new List<int>();
                    tab.listExcelIndexTab = new List<int>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    Regex regName = new Regex(@"лента|лист|арматура|полоса|угол|швеллер|труб|круг|шестигранник|шгр|квадрат|полоса|сталь|катанка", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой(?:\s*проч)|\w+ий|г[\/]к", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"\d+(?:[,.]\d+)?(?:[xх-]\d+(?:[,.]\d+)?)?(?:\s*[x*х]\s*\d+(?:[,.]\d+)?(?:-\d+(?:[,.]\d+)?)?)+", RegexOptions.IgnoreCase);
                    Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"(?:\d{,3}[ШСТУ]+\d{,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+\d{,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+\d{,3})(?=\s+|$)|(?:(?:Ст.)|(?:ст.)(?:\s*\d{1,2})\b)(?=\s+|$)|(?:\d{,3}[ХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)(?=\s+|$)|(?:[АA]-?\d)(?=\s+|$)|(?<=[Сс][Тт]\.\s?)\d{1,2}[гГ]\d{1,2}[cCсС]|(?<=ст\.)\d{1,2}[хфа]+(?=\s|$)|(?<=\s)[сС][тТ]\.?\s?\d{1,2}[_\w]+", RegexOptions.IgnoreCase);
                    string type = "";
                    string mark = "";

                    int lastRow = 0;
                    string tmp = "";
                    string stmp = "";
                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;
                    for (int i = 1; i <= cCelCol; i++) //столбцы
                    {
                        for (int j = 1; j <= cCelRow; j++) //строки
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (regDiam.IsMatch(temp))
                                {
                                    tab.StartRow = j;
                                    cellRange = (Excel.Range)excelworksheet.Cells[j - 1, i];
                                    if (cellRange.MergeArea.Count > 1)
                                    {
                                        for (int ma = i - 1; ma <= i + 1; ma++)
                                        {
                                            if (ma < 1) ma = 1;
                                            cellRange = (Excel.Range)excelworksheet.Cells[j - 1, ma];
                                            if (cellRange.Value != null)
                                                tmp = cellRange.Value.ToString().Trim();
                                            else tmp = "";
                                            if (tmp != "")
                                            {
                                                nameProd = regName.Match(tmp).Value;
                                                if (nameProd.ToLower() == "труб") nameProd = "Труба";
                                                if (nameProd.ToLower() == "угол") nameProd = "Уголок";
                                                if (nameProd.Length > 1)
                                                    nameProd = nameProd.Substring(0, 1).ToUpper() + nameProd.Substring(1, nameProd.Length - 1).ToLower();
                                                type = regType.Match(tmp).Value;
                                                mark = regMark.Match(tmp).Value;
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[j - 1, i];
                                        if (cellRange.Value != null)
                                            tmp = cellRange.Value.ToString().Trim();
                                        else tmp = "";
                                        if (tmp != "")
                                        {
                                            nameProd = regName.Match(tmp).Value;
                                            if (nameProd.ToLower() == "труб") nameProd = "Труба";
                                            if (nameProd.ToLower() == "угол") nameProd = "Уголок";
                                            if (nameProd.Length > 1)
                                                nameProd = nameProd.Substring(0, 1).ToUpper() + nameProd.Substring(1, nameProd.Length - 1);
                                            type = regType.Match(tmp).Value;
                                            mark = regMark.Match(tmp).Value;
                                        }
                                        else
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[j - 1, i - 1];
                                            if (cellRange.Value != null)
                                                tmp = cellRange.Value.ToString().Trim();
                                            if (tmp != "")
                                            {
                                                nameProd = regName.Match(tmp).Value;
                                                if (nameProd.ToLower() == "труб") nameProd = "Труба";
                                                if (nameProd.ToLower() == "угол") nameProd = "Уголок";
                                                if (nameProd.Length > 1)
                                                    nameProd = nameProd.Substring(0, 1).ToUpper() + nameProd.Substring(1, nameProd.Length - 1);
                                                type = regType.Match(tmp).Value;
                                                mark = regMark.Match(tmp).Value;
                                            }
                                            else
                                            {
                                                j = cCelRow;
                                                break;
                                            }
                                        }
                                    }

                                    for (int jj = j; jj <= cCelRow; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                                        if (cellRange.Value != null)
                                            tmp = cellRange.Value.ToString().Trim();
                                        else tmp = "";
                                        if (tmp != "")
                                        {
                                            if (regDiam.IsMatch(tmp) && nameProd != "")
                                            {
                                                dtProduct.Rows.Add();
                                                lastRow = dtProduct.Rows.Count - 1;
                                                tab.listExcelIndexTab.Add(jj);
                                                tab.listdtProductIndexRow.Add(lastRow);
                                                dtProduct.Rows[lastRow]["Название"] = regName.Match(tmp).Value;
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "труб") dtProduct.Rows[lastRow]["Название"] = "Труба";
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "угол") dtProduct.Rows[lastRow]["Название"] = "Уголок";
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().Length > 1)
                                                    dtProduct.Rows[lastRow]["Название"] = dtProduct.Rows[lastRow]["Название"].ToString().Substring(0, 1).ToUpper() + dtProduct.Rows[lastRow]["Название"].ToString().Substring(1, dtProduct.Rows[lastRow]["Название"].ToString().Length - 1);

                                                if (dtProduct.Rows[lastRow]["Название"].ToString() == "")
                                                {
                                                    dtProduct.Rows[lastRow]["Название"] = nameProd;
                                                }

                                                dtProduct.Rows[lastRow]["Примечание"] = tmp;
                                                dtProduct.Rows[lastRow]["Тип"] = regType.Match(tmp).Value;

                                                if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                                {
                                                    if (new Regex(@"г[\/]к", RegexOptions.IgnoreCase).IsMatch(type)) type = "горячекатаный";
                                                    dtProduct.Rows[lastRow]["Тип"] = type;
                                                    if (dtProduct.Rows[lastRow]["Тип"].ToString() == "") dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                                }

                                                dtProduct.Rows[lastRow]["Марка"] = regMark.Match(tmp).Value;
                                                if (dtProduct.Rows[lastRow]["Марка"].ToString() == "")
                                                {
                                                    dtProduct.Rows[lastRow]["Марка"] = mark;
                                                    if (dtProduct.Rows[lastRow]["Марка"].ToString() == "") dtProduct.Rows[lastRow]["Марка"] = "";
                                                }

                                                //ищет цену в ячейке справа
                                                cellRange = (Excel.Range)excelworksheet.Cells[jj, i + 1];
                                                if (cellRange.Value != null)
                                                    stmp = cellRange.Value.ToString().Trim();
                                                else stmp = "";
                                                if (stmp != "")
                                                {
                                                    dtProduct.Rows[lastRow]["Цена"] = stmp.Trim();
                                                }

                                                foreach (Match m in regTU.Matches(tmp))
                                                {
                                                    if (dtProduct.Rows[lastRow]["Стандарт"].ToString() == "") dtProduct.Rows[lastRow]["Стандарт"] = m.Value;
                                                    else dtProduct.Rows[lastRow]["Стандарт"] += "; " + m.Value;
                                                }

                                                //ищет параметры изделия
                                                string[] diam, tolsh, metraj;
                                                string tempo = "";
                                                tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                if (tempo != "")
                                                {
                                                    diam = tempo.Split('-');
                                                }
                                                else
                                                {
                                                    tempo = "";
                                                    tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    if (tempo != "")
                                                    {
                                                        diam = new string[] { tempo };
                                                    }
                                                    else diam = new string[] { "" };
                                                }

                                                tempo = new Regex(@"(?<=^\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                if (tempo != "")
                                                {
                                                    tolsh = tempo.Split('-');
                                                }
                                                else
                                                {
                                                    tempo = "";
                                                    tempo = new Regex(@"(?<=^\s*)\d+(?:[,\.]\d+)?\s*[xх]\s*\d+(?:[,\.]\d+)?(?=\s*[x*х]\s*\d+(?:[,\.]\d+)?)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    if (tempo != "")
                                                    {
                                                        tolsh = tempo.Split('х', 'x');
                                                    }
                                                    else
                                                    {
                                                        tempo = "";
                                                        tempo = new Regex(@"(?<=^\s*)\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                        if (tempo != "")
                                                        {
                                                            tolsh = new string[] { tempo };
                                                        }
                                                        else tolsh = new string[] { "" };
                                                    }
                                                }

                                                tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                if (tempo != "")
                                                {
                                                    metraj = tempo.Split('-');
                                                }
                                                else
                                                {
                                                    tempo = "";
                                                    tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s*$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    if (tempo != "")
                                                    {
                                                        metraj = new string[] { tempo };
                                                    }
                                                    else metraj = new string[] { "" };
                                                }

                                                for (int d = 0; d < diam.Length; d++)
                                                    for (int t = 0; t < tolsh.Length; t++)
                                                        for (int m = 0; m < metraj.Length; m++)
                                                        {
                                                            lastRow = dtProduct.Rows.Count - 1;
                                                            if (d == 0 && t == 0 && m == 0)
                                                            {
                                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = diam[0];
                                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = tolsh[0];
                                                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = metraj[0];
                                                            }
                                                            else
                                                            {
                                                                DataRow row = dtProduct.NewRow();
                                                                row["Название"] = dtProduct.Rows[lastRow]["Название"];
                                                                row["Тип"] = dtProduct.Rows[lastRow]["Тип"];
                                                                row["Диаметр (высота), мм"] = diam[d];
                                                                row["Толщина (ширина), мм"] = tolsh[t];
                                                                row["Метраж, м (длина, мм)"] = metraj[m];
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
                                            else
                                            {
                                                j = jj;
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            j = jj - 1;
                                            break;
                                        }
                                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                                        else tsPb1.Value = tsPb1.Maximum;
                                    }
                                }

                                InfoOrganization(temp);

                                #region адрес организации
                                if (!new Regex(@"склад", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    if (new Regex(@"(?:[Аа]дрес(?!\s*(?:склад)))|(?:\d+[\s,]\s?[Гг][\.\s]+\w+[,\s]+(?:(?:ул\.\s?)|(?:улица[\s:]+))\w+[,\s]+[\w\d.\\/]+\b(?:,?\s?(?:оф|офис)\.?\s?\d+(?:\w+)?)?)").IsMatch(temp))
                                    {
                                        if (textBoxOrgAdress.Text == "")
                                            textBoxOrgAdress.Text = new Regex(@"(?:(?<=[Аа]дрес\sофиса\s)[\w+\d+]+[\s*\w\d,.]*)|(?:\d+[\s,]\s?[Гг][\.\s]+\w+[,\s]+(?:(?:ул\.\s?)|(?:улица[\s:]+))\w+[,\s]+[\w\d.\s\\/]+(?:,?\s?(?:оф|офис)\.?\s?\d+(?:\w+)?)?)").Match(temp).Value;
                                        else
                                            textBoxOrgAdress.Text += ";" + new Regex(@"(?:(?<=[Аа]дрес\sофиса\s)[\w+\d+]+[\s*\w\d,.]*)|(?:\d+[\s,]\s?[Гг][\.\s]+\w+[,\s]+(?:(?:ул\.\s?)|(?:улица[\s:]+))\w+[,\s]+[\w\d.\s\\/]+(?:,?\s?(?:оф|офис)\.?\s?\d+(?:\w+)?)?)").Match(temp).Value;
                                    }
                                    else if (new Regex(@"адрес\s*офис\w+\s*$", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[j, i + 1];
                                        if (cellRange.Value != null)
                                        {
                                            temp = cellRange.Value.ToString();
                                            if (textBoxOrgAdress.Text == "")
                                                textBoxOrgAdress.Text = temp;
                                            else
                                                textBoxOrgAdress.Text += ";" + temp;
                                        }
                                    }
                                    else if (new Regex(@"г\.[\s+\w+,\.]+(?=,\s?тел)", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        if (textBoxOrgAdress.Text == "")
                                            textBoxOrgAdress.Text = new Regex(@"г\.[\s+\w+,\.]+(?=,\s?тел)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        else
                                            textBoxOrgAdress.Text += new Regex(@"г\.[\s+\w+,\.]+(?=,\s?тел)", RegexOptions.IgnoreCase).Match(temp).Value;
                                    }
                                }
                                else if (new Regex(@"(?<=офис.*)г.*\sоф.\d{1,3}", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    if (textBoxOrgAdress.Text == "")
                                        textBoxOrgAdress.Text = new Regex(@"(?<=офис.*)г.*\sоф.\d{1,3}", RegexOptions.IgnoreCase).Match(temp).Value;
                                    else
                                        textBoxOrgAdress.Text += new Regex(@"(?<=офис.*)г.*\sоф.\d{1,3}", RegexOptions.IgnoreCase).Match(temp).Value;
                                }
                                #endregion

                                Regex regOrgTel = new Regex(@"(?<=тел\.?\s*:\s*)(?:\d\s*)?\(\d+\)\s*[-\d]+(?:,\s*\(\d+\)\s*[-\d]+)+", RegexOptions.IgnoreCase);
                                if (regOrgTel.IsMatch(temp))
                                    foreach (Match m in regOrgTel.Matches(temp))
                                        if (textBoxOrgTelefon.Text == "")
                                            textBoxOrgTelefon.Text = m.Value;
                                        else textBoxOrgTelefon.Text += "; " + m.Value;

                                //Regex regManager = new Regex(@"контактное\s*лицо", RegexOptions.IgnoreCase);
                                //if (regManager.IsMatch(temp))
                                //{
                                //    ListViewItem lvi = new ListViewItem(new Regex(@"(?<=лицо\s*:\s*)(?:\w+\s*)+(?=,)").Match(temp).Value); //имя менеджера
                                //    lvi.SubItems.Add(new Regex(@"телефон\s*:\s*.*\d\d(?=,)").Match(temp).Value);           //телефон менеджера
                                //    if (lvi.SubItems[0].Text != "icq") listViewManager.Items.Add(lvi);
                                //}

                                Regex regSklad = new Regex(@"базы:\s*\d+,\s*г?\.?\s*.+,\s*\d+", RegexOptions.IgnoreCase);
                                if (regSklad.IsMatch(temp))
                                {
                                    ListViewItem lvi = new ListViewItem(new Regex(@"(?<=базы:\s*)\d+,\s*г?\.?\s*.+,\s*\d+").Match(temp).Value); //имя менеджера
                                    listViewAdrSklad.Items.Add(lvi);
                                }


                                //Regex regMail = new Regex(@"(?<=эл.*почта\s*:\s*)[\w\d\.-]+@[\w\d\.-]+.(?:ru|com|рф|info)", RegexOptions.IgnoreCase);
                                //if (regMail.IsMatch(temp))
                                //    textBoxOrgEmail.Text = regMail.Match(temp).Value;
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }


                }

                clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции ChelyabinskProfit\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла MetallKom
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void MetallKom(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+(?:г\.?)?\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?)").Match(Path.GetFileName(filePath)).Value;
                textBoxOrgName.Text = orgname;

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла MetallKom\n\n" + ex.ToString()); }

                string temp = "";
                //this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    structTab tab = new structTab();
                    tab.listdtProductIndexRow = new List<int>();
                    tab.listExcelIndexTab = new List<int>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    Regex regName = new Regex(@"лента|лист|арматура|полоса|угол|швеллер|труб|круг|шестигранник|шгр|квадрат|полоса|сталь|катанка", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой(?:\s*проч)|\w+ий|г[\/]к", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"\d+(?:[,.]\d+)?(?:[xх-]\d+(?:[,.]\d+)?)?(?:\s*[x*х]\s*\d+(?:[,.]\d+)?(?:-\d+(?:[,.]\d+)?)?)+(?:мм)?", RegexOptions.IgnoreCase);
                    Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"(?<=\s+)\d+[xcхсндг\dпа]*(\s*$|\s*-\s*\d+\s*$)|(?<=\s+)\d+[сп]+(?=\s+)", RegexOptions.IgnoreCase);
                    string type = "";
                    string mark = "";

                    int lastRow = 0;
                    string tmp = "";
                    string stmp = "";
                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;
                    for (int i = 1; i <= cCelCol; i++) //столбцы
                    {
                        for (int j = 1; j <= cCelRow; j++) //строки
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (regDiam.IsMatch(temp))
                                {
                                    tab.StartRow = j;
                                    cellRange = (Excel.Range)excelworksheet.Cells[j - 1, i];
                                    if (cellRange.MergeArea.Count > 1)
                                    {
                                        for (int ma = i - 1; ma <= i + 1; ma++)
                                        {
                                            if (ma < 1) ma = 1;
                                            cellRange = (Excel.Range)excelworksheet.Cells[j - 1, ma];
                                            if (cellRange.Value != null)
                                                tmp = cellRange.Value.ToString().Trim();
                                            else tmp = "";
                                            if (tmp != "")
                                            {
                                                nameProd = regName.Match(tmp).Value;
                                                if (nameProd.ToLower() == "труб") nameProd = "Труба";
                                                if (nameProd.ToLower() == "угол") nameProd = "Уголок";
                                                if (nameProd.Length > 1)
                                                    nameProd = nameProd.Substring(0, 1).ToUpper() + nameProd.Substring(1, nameProd.Length - 1).ToLower();
                                                type = regType.Match(tmp).Value;
                                                mark = regMark.Match(tmp).Value;
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[j - 1, i];
                                        if (cellRange.Value != null)
                                            tmp = cellRange.Value.ToString().Trim();
                                        else tmp = "";
                                        if (tmp != "")
                                        {
                                            nameProd = regName.Match(tmp).Value;
                                            if (nameProd.ToLower() == "труб") nameProd = "Труба";
                                            if (nameProd.ToLower() == "угол") nameProd = "Уголок";
                                            if (nameProd.Length > 1)
                                                nameProd = nameProd.Substring(0, 1).ToUpper() + nameProd.Substring(1, nameProd.Length - 1);
                                            type = regType.Match(tmp).Value;
                                            mark = regMark.Match(tmp).Value;
                                        }
                                        else
                                        {
                                            cellRange = (Excel.Range)excelworksheet.Cells[j - 1, i - 1];
                                            if (cellRange.Value != null)
                                                tmp = cellRange.Value.ToString().Trim();
                                            if (tmp != "")
                                            {
                                                nameProd = regName.Match(tmp).Value;
                                                if (nameProd.ToLower() == "труб") nameProd = "Труба";
                                                if (nameProd.ToLower() == "угол") nameProd = "Уголок";
                                                if (nameProd.Length > 1)
                                                    nameProd = nameProd.Substring(0, 1).ToUpper() + nameProd.Substring(1, nameProd.Length - 1);
                                                type = regType.Match(tmp).Value;
                                                mark = regMark.Match(tmp).Value;
                                            }
                                            else
                                            {
                                                j = cCelRow;
                                                break;
                                            }
                                        }
                                    }

                                    for (int jj = j; jj <= cCelRow; jj++)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                                        if (cellRange.Value != null)
                                            tmp = cellRange.Value.ToString().Trim();
                                        else tmp = "";
                                        if (tmp != "")
                                        {
                                            if (regDiam.IsMatch(tmp) && nameProd != "")
                                            {
                                                dtProduct.Rows.Add();
                                                lastRow = dtProduct.Rows.Count - 1;
                                                tab.listExcelIndexTab.Add(jj);
                                                tab.listdtProductIndexRow.Add(lastRow);
                                                dtProduct.Rows[lastRow]["Название"] = regName.Match(tmp).Value;
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "труб") dtProduct.Rows[lastRow]["Название"] = "Труба";
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().ToLower() == "угол") dtProduct.Rows[lastRow]["Название"] = "Уголок";
                                                if (dtProduct.Rows[lastRow]["Название"].ToString().Length > 1)
                                                    dtProduct.Rows[lastRow]["Название"] = dtProduct.Rows[lastRow]["Название"].ToString().Substring(0, 1).ToUpper() + dtProduct.Rows[lastRow]["Название"].ToString().Substring(1, dtProduct.Rows[lastRow]["Название"].ToString().Length - 1);

                                                if (dtProduct.Rows[lastRow]["Название"].ToString() == "")
                                                {
                                                    dtProduct.Rows[lastRow]["Название"] = nameProd;
                                                }

                                                dtProduct.Rows[lastRow]["Примечание"] = tmp;
                                                dtProduct.Rows[lastRow]["Тип"] = regType.Match(tmp).Value;

                                                if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                                {
                                                    if (new Regex(@"г[\/]к", RegexOptions.IgnoreCase).IsMatch(type)) type = "горячекатаный";
                                                    dtProduct.Rows[lastRow]["Тип"] = type;
                                                    if (dtProduct.Rows[lastRow]["Тип"].ToString() == "") dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                                }

                                                dtProduct.Rows[lastRow]["Марка"] = regMark.Match(tmp).Value;
                                                if (dtProduct.Rows[lastRow]["Марка"].ToString() == "")
                                                {
                                                    dtProduct.Rows[lastRow]["Марка"] = mark;
                                                    if (dtProduct.Rows[lastRow]["Марка"].ToString() == "") dtProduct.Rows[lastRow]["Марка"] = "";
                                                }

                                                //ищет цену в ячейке справа
                                                cellRange = (Excel.Range)excelworksheet.Cells[jj, i + 1];
                                                if (cellRange.Value != null)
                                                    stmp = cellRange.Value.ToString().Trim();
                                                else stmp = "";
                                                if (stmp != "")
                                                {
                                                    dtProduct.Rows[lastRow]["Цена"] = stmp.Trim();
                                                }

                                                foreach (Match m in regTU.Matches(tmp))
                                                {
                                                    if (dtProduct.Rows[lastRow]["Стандарт"].ToString() == "") dtProduct.Rows[lastRow]["Стандарт"] = m.Value;
                                                    else dtProduct.Rows[lastRow]["Стандарт"] += "; " + m.Value;
                                                }

                                                //ищет параметры изделия
                                                string[] diam, tolsh, metraj;
                                                string tempo = "";
                                                tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*[xх*])|(?<=\d+(?:[,\.]\d+)?\s*мм\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                if (tempo != "")
                                                {
                                                    diam = tempo.Split('-');
                                                }
                                                else
                                                {
                                                    tempo = "";
                                                    tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s*[xх*])|(?<=\d+(?:[,\.]\d+)?\s*мм\s*)\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    if (tempo != "")
                                                    {
                                                        diam = new string[] { tempo };
                                                    }
                                                    else diam = new string[] { "" };

                                                }

                                                tempo = new Regex(@"(?<=^\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                if (tempo != "")
                                                {
                                                    tolsh = tempo.Split('-');
                                                }
                                                else
                                                {
                                                    tempo = "";
                                                    tempo = new Regex(@"(?<=^\s*)\d+(?:[,\.]\d+)?\s*[xх]\s*\d+(?:[,\.]\d+)?(?=\s*[x*х]\s*\d+(?:[,\.]\d+)?(?:\s*-\s*\d+(?:[,\.]\d+)?)?\s*[x*х]\s*\d+(?:[,\.]\d+)?(?:\s*-\s*\d+(?:[,\.]\d+)?)?)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    if (tempo != "")
                                                    {
                                                        tolsh = tempo.Split('х', 'x');
                                                    }
                                                    else
                                                    {
                                                        tempo = "";
                                                        tempo = new Regex(@"(?<=^\s*)\d+(?:[,\.]\d+)?(?=\s*[xх*])|(?<=^\s*свыше\s*)\d+(?:[,\.]\d+)?(?=\s*мм)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                        if (tempo != "")
                                                        {
                                                            tolsh = new string[] { tempo };
                                                        }
                                                        else tolsh = new string[] { "" };
                                                    }
                                                }

                                                tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*(?:[мm]+)?\s*$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                if (tempo != "")
                                                {
                                                    metraj = tempo.Split('-');
                                                }
                                                else
                                                {
                                                    tempo = "";
                                                    tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s*(?:мм)?\s*$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                                    if (tempo != "")
                                                    {
                                                        metraj = new string[] { tempo };
                                                    }
                                                    else metraj = new string[] { "" };
                                                }

                                                for (int d = 0; d < diam.Length; d++)
                                                    for (int t = 0; t < tolsh.Length; t++)
                                                        for (int m = 0; m < metraj.Length; m++)
                                                        {
                                                            lastRow = dtProduct.Rows.Count - 1;
                                                            if (d == 0 && t == 0 && m == 0)
                                                            {
                                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = diam[0];
                                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = tolsh[0];
                                                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = metraj[0];
                                                            }
                                                            else
                                                            {
                                                                DataRow row = dtProduct.NewRow();
                                                                row["Название"] = dtProduct.Rows[lastRow]["Название"];
                                                                row["Тип"] = dtProduct.Rows[lastRow]["Тип"];
                                                                row["Диаметр (высота), мм"] = diam[d];
                                                                row["Толщина (ширина), мм"] = tolsh[t];
                                                                row["Метраж, м (длина, мм)"] = metraj[m];
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
                                            else
                                            {
                                                j = jj;
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            j = jj - 1;
                                            break;
                                        }
                                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                                        else tsPb1.Value = tsPb1.Maximum;
                                    }
                                }

                                InfoOrganization(temp);

                                #region адрес организации
                                if (!new Regex(@"склад", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    if (new Regex(@"\d+\s*,?\s*г\.\s*\w+\s*,?\s*[\w\s]+\d+\w?\s").IsMatch(temp))
                                    {
                                        if (textBoxOrgAdress.Text == "")
                                            textBoxOrgAdress.Text = new Regex(@"\d+\s*,?\s*г\.\s*\w+\s*,?\s*[\w\s]+\d+\w?\s").Match(temp).Value;
                                        else
                                            textBoxOrgAdress.Text += ";" + new Regex(@"\d+\s*,?\s*г\.\s*\w+\s*,?\s*[\w\s]+\d+\w?\s").Match(temp).Value;
                                    }
                                    else if (new Regex(@"адрес\s*офис\w+\s*$", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[j, i + 1];
                                        if (cellRange.Value != null)
                                        {
                                            temp = cellRange.Value.ToString();
                                            if (textBoxOrgAdress.Text == "")
                                                textBoxOrgAdress.Text = temp;
                                            else
                                                textBoxOrgAdress.Text += ";" + temp;
                                        }
                                    }
                                    else if (new Regex(@"г\.[\s+\w+,\.]+(?=,\s?тел)", RegexOptions.IgnoreCase).IsMatch(temp))
                                    {
                                        if (textBoxOrgAdress.Text == "")
                                            textBoxOrgAdress.Text = new Regex(@"г\.[\s+\w+,\.]+(?=,\s?тел)", RegexOptions.IgnoreCase).Match(temp).Value;
                                        else
                                            textBoxOrgAdress.Text += new Regex(@"г\.[\s+\w+,\.]+(?=,\s?тел)", RegexOptions.IgnoreCase).Match(temp).Value;
                                    }
                                }
                                else if (new Regex(@"(?<=офис.*)г.*\sоф.\d{1,3}", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    if (textBoxOrgAdress.Text == "")
                                        textBoxOrgAdress.Text = new Regex(@"(?<=офис.*)г.*\sоф.\d{1,3}", RegexOptions.IgnoreCase).Match(temp).Value;
                                    else
                                        textBoxOrgAdress.Text += new Regex(@"(?<=офис.*)г.*\sоф.\d{1,3}", RegexOptions.IgnoreCase).Match(temp).Value;
                                }
                                #endregion

                                Regex regOrgTel = new Regex(@"(?<=\s)8[89]\d{9}(?=\s)", RegexOptions.IgnoreCase);
                                if (regOrgTel.IsMatch(temp))
                                    foreach (Match m in regOrgTel.Matches(temp))
                                        if (textBoxOrgTelefon.Text == "")
                                            textBoxOrgTelefon.Text = m.Value;
                                        else textBoxOrgTelefon.Text += "; " + m.Value;

                                Regex regManager = new Regex(@"тел.*8[89]\d{9}\s*\w+", RegexOptions.IgnoreCase);
                                if (regManager.IsMatch(temp))
                                {
                                    ListViewItem lvi = new ListViewItem(new Regex(@"(?<=тел.*8[89]\d{9}\s*)(?:\w+\s*)+", RegexOptions.IgnoreCase).Match(temp).Value); //имя менеджера
                                    lvi.SubItems.Add(new Regex(@"(?<=тел.*)8[89]\d{9}(?=\s*\w+)", RegexOptions.IgnoreCase).Match(temp).Value);           //телефон менеджера
                                    if (lvi.SubItems[0].Text != "icq") listViewManager.Items.Add(lvi);
                                }

                                //Regex regSklad = new Regex(@"базы:\s*\d+,\s*г?\.?\s*.+,\s*\d+", RegexOptions.IgnoreCase);
                                //if (regSklad.IsMatch(temp))
                                //{
                                //    ListViewItem lvi = new ListViewItem(new Regex(@"(?<=базы:\s*)\d+,\s*г?\.?\s*.+,\s*\d+").Match(temp).Value); //имя менеджера
                                //    listViewAdrSklad.Items.Add(lvi);
                                //}


                                //Regex regMail = new Regex(@"(?<=эл.*почта\s*:\s*)[\w\d\.-]+@[\w\d\.-]+.(?:ru|com|рф|info)", RegexOptions.IgnoreCase);
                                //if (regMail.IsMatch(temp))
                                //    textBoxOrgEmail.Text = regMail.Match(temp).Value;
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }


                }

                clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции MetallKom\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла PromMet
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void PromMet(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+(?:г\.?)?\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?)").Match(Path.GetFileName(filePath)).Value;
                if (new Regex(@"Пром.?груп", RegexOptions.IgnoreCase).IsMatch(Path.GetFileName(filePath)))
                    orgname = "Промгруппа";
                if (new Regex(@"Пром.?груп.*круг", RegexOptions.IgnoreCase).IsMatch(Path.GetFileName(filePath)))
                    orgname = "Промгруппа круг";
                textBoxOrgName.Text = orgname;

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла PromMet\n\n" + ex.ToString()); }

                string temp = "";
                //this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    structTab tab = new structTab();
                    tab.listdtProductIndexRow = new List<int>();
                    tab.listExcelIndexTab = new List<int>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    Regex regName = new Regex(@"лента|лист|арматура|полоса|угол|швеллер|труб|круг|шестигранник|шгр|квадрат|полоса|сталь|катанка|быстрорез", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой(?:\s*проч)|\w+ий|г[\/]к", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"\d+(?:[,.]\d+)?(?=\s*$)|(?<=пол\s*\.?\s*\d+(?:[,.]\d+)?\s*\*)\d+(?:[,.]\d+)?|\d+(?:[,.]\d+)?(?:\s*-\s*\d+(?:[,.]\d+)?)?(?:\s*[xх\*]\s*\d+(?:[,.]\d+)?(?:\s*-\s*\d+(?:[,.]\d+)?)?)+", RegexOptions.IgnoreCase);
                    Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"(?:\d{1,3}[ШСТУХ]+\d{0,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+(?:\d{0,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+)*\d{0,3})(?=\s+|$)|(?:(?:Ст.)|(?:ст.)(?:\s*\d{1,2})\b)(?=\s+|$)|(?:\d{,3}[ХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)(?=\s+|$)|(?:[АA]-?\d)(?=\s+|$)|(?<=[Сс][Тт]\.\s?)\d{1,2}[гГ]\d{1,2}[cCсС]|(?<=ст\.)\d{1,2}[хфа]+(?=\s|$)|(?<=\s)[сС][тТ]\.?\s?\d{1,2}[_\w]+|(?<=\s)[0-4]\d[хx]\d\d[ТН]*(?=\s)|AISI\s*\d+\w*\s|(?<=\s+)\d+[xcхсндг\dпа]*(\s*$|\s*-\s*\d+\s*$)|(?<=\s+)\d+[сп]+(?=\s+)", RegexOptions.IgnoreCase);

                    int lastRow = 0;
                    int ColNumbPP = 0, ColDiam = 0, ColMark = 0, ColTreb = 0, ColCol = 0;
                    nameProd = "";
                    string type = "", tmp = "";

                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;

                    #region поиск заголовков
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        int jj = j;
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"№\s*п/п", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    tab.StartRow = j + 1; //отсюда начинать поиск данных
                                    ColNumbPP = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"Марка", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColMark = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"требования", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColTreb = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"размер", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColDiam = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"кол.*во", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColCol = i;
                                    j = cCelRow;
                                }

                                #region сайт
                                if (new Regex(@"(?:www\.)[\w\d-]{2,}\.[A-Za-zА-Яа-я]+").IsMatch(temp)) // поиск сайта
                                {
                                    textBoxOrgSite.Text = new Regex(@"(?:www\.)[\w\d-]{2,}\.[A-Za-zА-Яа-я]+").Match(temp).Value;
                                    //break;
                                }
                                #endregion

                                #region Email
                                Regex regEmail = new Regex(@"[\d\.\w-\*\\]+@[\w-]+\.\w{1,5}");
                                if (regEmail.IsMatch(temp)) // поиск Email
                                {
                                    if (textBoxOrgEmail.Text == "") textBoxOrgEmail.Text = regEmail.Match(temp).Value;
                                    else textBoxOrgEmail.Text += "; " + regEmail.Match(temp).Value;
                                    //break;
                                }
                                #endregion

                                #region адрес организации
                                if (new Regex(@"(?:\d{6}\s*,\s*)?г.\s*\w+.*\d+(?=\s*$)", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    if (textBoxOrgAdress.Text == "")
                                        textBoxOrgAdress.Text = new Regex(@"(?:\d{6}\s*,\s*)?г.\s*\w+.*\d+(?=\s*$)").Match(temp).Value;
                                    else
                                        textBoxOrgAdress.Text += ";" + new Regex(@"(?:\d{6}\s*,\s*)?г.\s*\w+.*\d+(?=\s*$)").Match(temp).Value;
                                }
                                #endregion

                                Regex regOrgTel = new Regex(@"(?<=тел\.?.*:?\s*)(?:\(\s*\d{3,5}\s*\)\s*)?[\d-]+", RegexOptions.IgnoreCase);
                                if (regOrgTel.IsMatch(temp))
                                {
                                    foreach (Match m in regOrgTel.Matches(temp))
                                        if (textBoxOrgTelefon.Text == "")
                                            textBoxOrgTelefon.Text = m.Value;
                                        else textBoxOrgTelefon.Text += "; " + m.Value;
                                }
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }
                    #endregion

                    tsLabelClearingTable.Text = "Обработка";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow - tab.StartRow;
                    int predStop = 0;
                    bool Stop = false;

                    for (int j = tab.StartRow - 1; j <= cCelRow; j++) //строки
                    {
                        Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, ColNumbPP];
                        if (cellRange.Value != null)
                            tmp = cellRange.Value.ToString().Trim();
                        else tmp = "";
                        if (tmp != "")
                        {
                            if (regName.IsMatch(tmp))
                            {
                                if (new Regex(@"круг", RegexOptions.IgnoreCase).IsMatch(orgname))
                                { nameProd = "Круг"; }
                                else nameProd = regName.Match(tmp).Value;
                                type = regType.Match(tmp).Value;
                                predStop = 0;
                                Stop = false;
                            }
                            else predStop++;
                        }
                        else
                        {
                            cellRange = (Excel.Range)excelworksheet.Cells[j, ColNumbPP + 1];
                            if (cellRange.Value != null)
                                tmp = cellRange.Value.ToString().Trim();
                            else tmp = "";
                            if (tmp != "")
                            {
                                if (regName.IsMatch(tmp))
                                {
                                    if (new Regex(@"круг", RegexOptions.IgnoreCase).IsMatch(orgname))
                                    { nameProd = "Круг"; }
                                    else nameProd = regName.Match(tmp).Value;
                                    type = regType.Match(tmp).Value;
                                    predStop = 0;
                                    Stop = false;
                                }
                                else predStop++;
                            }
                            else predStop++;
                        }
                        cellRange = (Excel.Range)excelworksheet.Cells[j, ColDiam];
                        if (cellRange.Value != null)
                            temp = cellRange.Value.ToString().Trim();
                        else temp = "";
                        if (temp != "")
                        {
                            if (regDiam.IsMatch(temp) && nameProd != "")
                            {
                                predStop = 0;
                                Stop = false;
                                dtProduct.Rows.Add();
                                lastRow = dtProduct.Rows.Count - 1;
                                tab.listExcelIndexTab.Add(j);
                                tab.listdtProductIndexRow.Add(lastRow);

                                if (new Regex(@"ш\s*г\s*[рp]", RegexOptions.IgnoreCase).IsMatch(temp))
                                    dtProduct.Rows[lastRow]["Название"] = "Шестигранник";
                                else if (new Regex(@"пол\s*\.?", RegexOptions.IgnoreCase).IsMatch(temp))
                                    dtProduct.Rows[lastRow]["Название"] = "Полоса";
                                else if (new Regex(@"кв\s*\.?", RegexOptions.IgnoreCase).IsMatch(temp))
                                    dtProduct.Rows[lastRow]["Название"] = "Квадрат";
                                else
                                    dtProduct.Rows[lastRow]["Название"] = nameProd;


                                dtProduct.Rows[lastRow]["Примечание"] = temp;
                                dtProduct.Rows[lastRow]["Тип"] = regType.Match(temp).Value;

                                if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                {
                                    if (new Regex(@"г[\/]к", RegexOptions.IgnoreCase).IsMatch(type)) type = "горячекатаный";
                                    dtProduct.Rows[lastRow]["Тип"] = type;
                                    if (dtProduct.Rows[lastRow]["Тип"].ToString() == "") dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                }

                                cellRange = (Excel.Range)excelworksheet.Cells[j, ColMark];
                                if (cellRange.Value != null)
                                    tmp = cellRange.Value.ToString().Trim();
                                else tmp = "";
                                if (tmp != "")
                                {
                                    dtProduct.Rows[lastRow]["Марка"] = tmp;
                                    if (new Regex(@"ш\s*г\s*[рp]", RegexOptions.IgnoreCase).IsMatch(tmp))
                                        dtProduct.Rows[lastRow]["Название"] = "Шестигранник";
                                    else if (new Regex(@"пол\s*\.?", RegexOptions.IgnoreCase).IsMatch(tmp))
                                        dtProduct.Rows[lastRow]["Название"] = "Полоса";
                                    else if (new Regex(@"кв\s*\.?", RegexOptions.IgnoreCase).IsMatch(tmp))
                                        dtProduct.Rows[lastRow]["Название"] = "Квадрат";
                                }

                                cellRange = (Excel.Range)excelworksheet.Cells[j, ColTreb];
                                if (cellRange.Value != null)
                                    tmp = cellRange.Value.ToString().Trim();
                                else tmp = "";
                                if (tmp != "")
                                {
                                    dtProduct.Rows[lastRow]["Стандарт"] = tmp;
                                }

                                cellRange = (Excel.Range)excelworksheet.Cells[j, ColCol];
                                if (cellRange.Value != null)
                                    tmp = cellRange.Value.ToString().Trim();
                                else tmp = "";
                                if (tmp != "")
                                {
                                    dtProduct.Rows[lastRow]["Мерность (т, м, мм)"] = tmp;
                                }

                                //dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDiam.Match(temp).Value;
                                //dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=пол\s*\.?\s*)\d+(?:[,.]\d+)?(?=\s*\*)", RegexOptions.IgnoreCase).Match(temp).Value;
                                //dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = "";
                                string[] diam, tolsh, metraj;
                                string tempo = "";
                                tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(temp).Value;
                                if (tempo != "")
                                {
                                    diam = tempo.Split('-');
                                }
                                else
                                {
                                    tempo = "";
                                    tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(temp).Value;
                                    if (tempo != "")
                                    {
                                        diam = new string[] { tempo };
                                    }
                                    else
                                    {
                                        tempo = "";
                                        tempo = new Regex(@"\d+(?:[,.]\d+)?(?=\s*$)|(?<=пол\s*\.?\s*\d+(?:[,.]\d+)?\s*\*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                        if (tempo != "")
                                        {
                                            diam = new string[] { tempo };
                                        }
                                        else diam = new string[] { "" };
                                    }
                                }

                                tempo = new Regex(@"(?<=^\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(temp).Value;
                                if (tempo != "")
                                {
                                    tolsh = tempo.Split('-');
                                }
                                else
                                {
                                    tempo = "";
                                    tempo = new Regex(@"(?<=^\s*)\d+(?:[,\.]\d+)?\s*[xх]\s*\d+(?:[,\.]\d+)?(?=\s*[x*х]\s*\d+(?:[,\.]\d+)?\s*[x*х]\s*)", RegexOptions.IgnoreCase).Match(temp).Value;
                                    if (tempo != "")
                                    {
                                        tolsh = tempo.Split('х', 'x');
                                    }
                                    else
                                    {
                                        tempo = "";
                                        tempo = new Regex(@"(?<=^\s*)\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(temp).Value;
                                        if (tempo != "")
                                        {
                                            tolsh = new string[] { tempo };
                                        }
                                        else
                                        {
                                            tempo = "";
                                            tempo = new Regex(@"(?<=пол\s*\.?\s*)\d+(?:[,.]\d+)?(?=\s*\*)", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if (tempo != "")
                                            {
                                                tolsh = new string[] { tempo };
                                            }
                                            else tolsh = new string[] { "" };
                                        }
                                    }
                                }

                                tempo = new Regex(@"(?<=[xх*]\s*\d+(?:[,\.]\d+)?(?:\s*-\s*\d+(?:[,\.]\d+)?)?[xх*]\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                if (tempo != "")
                                {
                                    metraj = tempo.Split('-');
                                }
                                else
                                {
                                    tempo = "";
                                    tempo = new Regex(@"(?<=[xх*]\s*\d+(?:[,\.]\d+)?(?:\s*-\s*\d+(?:[,\.]\d+)?)?[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s*$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                    if (tempo != "")
                                    {
                                        metraj = new string[] { tempo };
                                    }
                                    else metraj = new string[] { "" };
                                }

                                for (int d = 0; d < diam.Length; d++)
                                    for (int t = 0; t < tolsh.Length; t++)
                                        for (int m = 0; m < metraj.Length; m++)
                                        {
                                            lastRow = dtProduct.Rows.Count - 1;
                                            if (d == 0 && t == 0 && m == 0)
                                            {
                                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = diam[0];
                                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = tolsh[0];
                                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = metraj[0];
                                            }
                                            else
                                            {
                                                DataRow row = dtProduct.NewRow();
                                                row["Название"] = dtProduct.Rows[lastRow]["Название"];
                                                row["Тип"] = dtProduct.Rows[lastRow]["Тип"];
                                                row["Диаметр (высота), мм"] = diam[d];
                                                row["Толщина (ширина), мм"] = tolsh[t];
                                                row["Метраж, м (длина, мм)"] = metraj[m];
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
                        }
                        else Stop = true;
                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        else tsPb1.Value = tsPb1.Maximum;
                        if (predStop > 1 && Stop) { nameProd = ""; }//predStop = 0; Stop = false; break; }
                    }
                }


                clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции PromMet\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла PromMet
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void PromGrup(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+(?:г\.?)?\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?)").Match(Path.GetFileName(filePath)).Value;
                textBoxOrgName.Text = orgname;

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла PromMet\n\n" + ex.ToString()); }

                string temp = "";
                //this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    structTab tab = new structTab();
                    tab.listdtProductIndexRow = new List<int>();
                    tab.listExcelIndexTab = new List<int>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    Regex regName = new Regex(@"лента|лист|арматура|полоса|угол|швеллер|труб|круг|шестигранник|шгр|квадрат|полоса|сталь|катанка|быстрорез", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой(?:\s*проч)|\w+ий|г[\/]к", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase);
                    Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"(?<=\s+)\d+[xcхсндг\dпа]*(\s*$|\s*-\s*\d+\s*$)|(?<=\s+)\d+[сп]+(?=\s+)", RegexOptions.IgnoreCase);

                    int lastRow = 0;
                    int ColNumbPP = 0, ColDiam = 0, ColMark = 0, ColTreb = 0, ColCol = 0;
                    nameProd = "";
                    string type = "", tmp = "";

                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;

                    #region поиск заголовков
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        int jj = j;
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"№\s*п/п", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    tab.StartRow = j + 1; //отсюда начинать поиск данных
                                    ColNumbPP = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"Марка", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColMark = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"требования", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColTreb = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"размер", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColDiam = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"кол.*во", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColCol = i;
                                    j = cCelRow;
                                }

                                #region сайт
                                if (new Regex(@"(?:www\.)[\w\d-]{2,}\.[A-Za-zА-Яа-я]+").IsMatch(temp)) // поиск сайта
                                {
                                    textBoxOrgSite.Text = new Regex(@"(?:www\.)[\w\d-]{2,}\.[A-Za-zА-Яа-я]+").Match(temp).Value;
                                    //break;
                                }
                                #endregion

                                #region Email
                                Regex regEmail = new Regex(@"[\d\.\w-\*\\]+@[\w-]+\.\w{1,5}");
                                if (regEmail.IsMatch(temp)) // поиск Email
                                {
                                    if (textBoxOrgEmail.Text == "") textBoxOrgEmail.Text = regEmail.Match(temp).Value;
                                    else textBoxOrgEmail.Text += "; " + regEmail.Match(temp).Value;
                                    //break;
                                }
                                #endregion

                                #region адрес организации
                                if (new Regex(@"г.\s*\w+.*\d+(?=\s*$)", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    if (textBoxOrgAdress.Text == "")
                                        textBoxOrgAdress.Text = new Regex(@"г.\s*\w+.*\d+(?=\s*$)").Match(temp).Value;
                                    else
                                        textBoxOrgAdress.Text += ";" + new Regex(@"г.\s*\w+.*\d+(?=\s*$)").Match(temp).Value;
                                }
                                #endregion

                                Regex regOrgTel = new Regex(@"тел\.?.*:\s*[\s()\d-;]+|^\s*\d[-\d]+", RegexOptions.IgnoreCase);
                                if (regOrgTel.IsMatch(temp))
                                    foreach (Match m in regOrgTel.Matches(temp))
                                        if (textBoxOrgTelefon.Text == "")
                                            textBoxOrgTelefon.Text = m.Value;
                                        else textBoxOrgTelefon.Text += "; " + m.Value;

                                //Regex regManager = new Regex(@"тел.*8[89]\d{9}\s*\w+", RegexOptions.IgnoreCase);
                                //if (regManager.IsMatch(temp))
                                //{
                                //    ListViewItem lvi = new ListViewItem(new Regex(@"(?<=тел.*8[89]\d{9}\s*)(?:\w+\s*)+", RegexOptions.IgnoreCase).Match(temp).Value); //имя менеджера
                                //    lvi.SubItems.Add(new Regex(@"(?<=тел.*)8[89]\d{9}(?=\s*\w+)", RegexOptions.IgnoreCase).Match(temp).Value);           //телефон менеджера
                                //    if (lvi.SubItems[0].Text != "icq") listViewManager.Items.Add(lvi);
                                //}
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }
                    #endregion

                    tsLabelClearingTable.Text = "Обработка";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow - tab.StartRow;
                    for (int j = tab.StartRow; j <= cCelRow; j++) //строки
                    {
                        Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, ColDiam];
                        if (cellRange.Value != null)
                            temp = cellRange.Value.ToString().Trim();
                        else temp = "";
                        if (temp != "")
                        {
                            if (nameProd == "")
                            {
                                cellRange = (Excel.Range)excelworksheet.Cells[j - 1, ColNumbPP];
                                if (cellRange.Value != null)
                                    tmp = cellRange.Value.ToString().Trim();
                                else tmp = "";
                                if (tmp != "")
                                {
                                    if (regName.IsMatch(tmp))
                                    {
                                        nameProd = "Круг";
                                    }
                                    type = regType.Match(tmp).Value;
                                }
                            }
                            if (regDiam.IsMatch(temp) && nameProd != "")
                            {
                                dtProduct.Rows.Add();
                                lastRow = dtProduct.Rows.Count - 1;
                                tab.listExcelIndexTab.Add(j);
                                tab.listdtProductIndexRow.Add(lastRow);

                                if (new Regex(@"ш\s*г\s*[рp]", RegexOptions.IgnoreCase).IsMatch(temp))
                                    dtProduct.Rows[lastRow]["Название"] = "Шестигранник";
                                else
                                    dtProduct.Rows[lastRow]["Название"] = nameProd;


                                dtProduct.Rows[lastRow]["Примечание"] = temp;
                                dtProduct.Rows[lastRow]["Тип"] = regType.Match(temp).Value;

                                if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                {
                                    if (new Regex(@"г[\/]к", RegexOptions.IgnoreCase).IsMatch(type)) type = "горячекатаный";
                                    dtProduct.Rows[lastRow]["Тип"] = type;
                                    if (dtProduct.Rows[lastRow]["Тип"].ToString() == "") dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                }

                                cellRange = (Excel.Range)excelworksheet.Cells[j, ColMark];
                                if (cellRange.Value != null)
                                    tmp = cellRange.Value.ToString().Trim();
                                else tmp = "";
                                if (tmp != "")
                                {
                                    dtProduct.Rows[lastRow]["Марка"] = tmp;
                                }

                                cellRange = (Excel.Range)excelworksheet.Cells[j, ColTreb];
                                if (cellRange.Value != null)
                                    tmp = cellRange.Value.ToString().Trim();
                                else tmp = "";
                                if (tmp != "")
                                {
                                    dtProduct.Rows[lastRow]["Стандарт"] = tmp;
                                }

                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = regDiam.Match(temp).Value;
                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = "";
                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = "";
                            }
                        }
                        else nameProd = "";
                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        else tsPb1.Value = tsPb1.Maximum;
                    }

                }


                clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции PromMet\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла Apogei
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void Apogei(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+(?:г\.?)?\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?)").Match(Path.GetFileName(filePath)).Value;
                textBoxOrgName.Text = orgname;

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла Apogei\n\n" + ex.ToString()); }

                string temp = "";
                //this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    structTab tab = new structTab();
                    tab.listdtProductIndexRow = new List<int>();
                    tab.listExcelIndexTab = new List<int>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    Regex regName = new Regex(@"лента\b|лист\b|арматура\b|полоса\b|угол(?:ок)?\b|швеллер\w?\b|труб\w?\b|круг\w?\b|шестигранник|шгр\b|квадрат\b|сталь\b|катанка|быстрорез|колесо|заготовка|блок\b|^\s*вал\s*|втулка|поковка", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой(?:\s*проч)|\w+ий|г[\/]к", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"(?<=\s+)\d+(?:[,.]\d+)?(?=мм|\s|$|\s*[xх*]\s*\d+(?:[,\.]\d+)?\s*мм)|(?<=[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase);
                    Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"(?:\d{,3}[ШСТУ]+\d{,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+\d{,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+\d{,3})(?=\s+|$)|(?:(?:Ст.)|(?:ст.)(?:\s*\d{1,2})\b)(?=\s+|$)|(?:\d{,3}[ХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)(?=\s+|$)|(?:[АA]-?\d)(?=\s+|$)|(?<=[Сс][Тт]\.\s?)\d{1,2}[гГ]\d{1,2}[cCсС]|(?<=ст\.)\d{1,2}[хфа]+(?=\s|$)|(?<=\s)[сС][тТ]\.?\s?\d{1,2}[_\w]+", RegexOptions.IgnoreCase);

                    int lastRow = 0;
                    int ColNomenkl = 1, ColPrice = 1, ColCol = 1;
                    nameProd = "";
                    string type = "", diam = "", mera = "", marka = "", price = "", tmp = "";

                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;

                    #region поиск заголовков
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        int jj = j;
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"^\s*Номенклатура", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    tab.StartRow = j + 1; //отсюда начинать поиск данных
                                    ColNomenkl = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"Цен.", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColPrice = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"(?:\w+)?\s*остат\w+\s*$", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColCol = i;
                                    j = cCelRow;
                                }

                                #region сайт
                                if (new Regex(@"(?:www\.)[\w\d-]{2,}\.[A-Za-zА-Яа-я]+").IsMatch(temp)) // поиск сайта
                                {
                                    textBoxOrgSite.Text = new Regex(@"(?:www\.)[\w\d-]{2,}\.[A-Za-zА-Яа-я]+").Match(temp).Value;
                                    //break;
                                }
                                #endregion

                                #region Email
                                Regex regEmail = new Regex(@"[\d\.\w-\*\\]+@[\w-]+\.\w{1,5}");
                                if (regEmail.IsMatch(temp)) // поиск Email
                                {
                                    if (textBoxOrgEmail.Text == "") textBoxOrgEmail.Text = regEmail.Match(temp).Value;
                                    else textBoxOrgEmail.Text += "; " + regEmail.Match(temp).Value;
                                    //break;
                                }
                                #endregion

                                #region адрес организации
                                if (new Regex(@"г.\s*\w+.*\d+(?=\s*$)", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    if (textBoxOrgAdress.Text == "")
                                        textBoxOrgAdress.Text = new Regex(@"г.\s*\w+.*\d+(?=\s*$)").Match(temp).Value;
                                    else
                                        textBoxOrgAdress.Text += ";" + new Regex(@"г.\s*\w+.*\d+(?=\s*$)").Match(temp).Value;
                                }
                                #endregion

                                Regex regOrgTel = new Regex(@"тел\.?(?:ефон)?\s*:\s*[\s()\d-;]+|^\s*\d[-\d]+", RegexOptions.IgnoreCase);
                                if (regOrgTel.IsMatch(temp))
                                    foreach (Match m in regOrgTel.Matches(temp))
                                        if (textBoxOrgTelefon.Text == "")
                                            textBoxOrgTelefon.Text = m.Value;
                                        else textBoxOrgTelefon.Text += "; " + m.Value;

                                //Regex regManager = new Regex(@"тел.*8[89]\d{9}\s*\w+", RegexOptions.IgnoreCase);
                                //if (regManager.IsMatch(temp))
                                //{
                                //    ListViewItem lvi = new ListViewItem(new Regex(@"(?<=тел.*8[89]\d{9}\s*)(?:\w+\s*)+", RegexOptions.IgnoreCase).Match(temp).Value); //имя менеджера
                                //    lvi.SubItems.Add(new Regex(@"(?<=тел.*)8[89]\d{9}(?=\s*\w+)", RegexOptions.IgnoreCase).Match(temp).Value);           //телефон менеджера
                                //    if (lvi.SubItems[0].Text != "icq") listViewManager.Items.Add(lvi);
                                //}
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }
                    #endregion

                    tsLabelClearingTable.Text = "Обработка";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow - tab.StartRow;
                    #region Обработка
                    for (int j = tab.StartRow; j <= cCelRow; j++) //строки
                    {
                        nameProd = ""; type = ""; diam = ""; price = ""; mera = ""; marka = "";
                        Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, ColNomenkl];
                        if (cellRange.Value != null)
                            temp = cellRange.Value.ToString().Trim();
                        else temp = "";
                        if (temp != "")
                        {
                            nameProd = regName.Match(temp).Value;
                            type = regType.Match(temp).Value;
                            diam = regDiam.Match(temp).Value;
                            Excel.Range MeraRange = (Excel.Range)excelworksheet.Cells[j, ColCol];
                            if (MeraRange.Value != null)
                                tmp = MeraRange.Value.ToString().Trim();
                            else tmp = "";
                            if (tmp != "")
                            {
                                mera = new Regex(@"\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).Match(tmp).Value;
                            }
                            marka = regMark.Match(temp).Value;
                            Excel.Range PriceRange = (Excel.Range)excelworksheet.Cells[j, ColPrice];
                            if (PriceRange.Value != null)
                                tmp = PriceRange.Value.ToString().Trim();
                            else tmp = "";
                            if (tmp != "")
                            {
                                price = tmp;
                            }

                            if (price == "") if (new Regex(@"^\s*[А-Яа-я]+\s+\d+\s*$|^[А-Яа-я]+\s+[А-Яа-я]+\s+\d+\s*$|^[А-Яа-я]+\s*$", RegexOptions.IgnoreCase).IsMatch(temp)) continue;

                            if ((nameProd != "" && diam != "" && (mera != "" || marka != "")) || (nameProd != "" && price != "" && (mera != "" || marka != "")))
                            {
                                dtProduct.Rows.Add();
                                lastRow = dtProduct.Rows.Count - 1;
                                tab.listExcelIndexTab.Add(j);
                                tab.listdtProductIndexRow.Add(lastRow);

                                if (new Regex(@"ш\s*г\s*[рp]", RegexOptions.IgnoreCase).IsMatch(nameProd))
                                    dtProduct.Rows[lastRow]["Название"] = "Шестигранник";
                                else
                                    dtProduct.Rows[lastRow]["Название"] = nameProd;

                                dtProduct.Rows[lastRow]["Примечание"] = temp;
                                if (new Regex(@"г[\/]к", RegexOptions.IgnoreCase).IsMatch(type)) type = "горячекатаный";
                                dtProduct.Rows[lastRow]["Тип"] = type;
                                if (dtProduct.Rows[lastRow]["Тип"].ToString() == "") dtProduct.Rows[lastRow]["Тип"] = "тип не указан";


                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = diam;
                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\s+)\d+(?:[,\.]\d+)?(?=\s*[xх*]\s*\d+(?:[,\.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=[xх*]\s*\d+(?:[,\.]\d+)?\s*[xх*])\d+(?:[,\.]\d+)?(?=\s|\s*мм)", RegexOptions.IgnoreCase).Match(temp).Value;

                                if (nameProd.ToLower() == "колесо")
                                {
                                    dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=\s+)\d+(?:[,\.]\d+)?(?=\s*[xх*]\s*\d+(?:[,\.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                                    dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\d+(?:[,\.]\d+)?\s*[xх*]\s*)\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                                }

                                dtProduct.Rows[lastRow]["Мерность (т, м, мм)"] = mera;

                                dtProduct.Rows[lastRow]["Марка"] = marka;
                                dtProduct.Rows[lastRow]["Цена"] = new Regex(@"\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).Match(price).Value;
                            }

                        }
                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        else tsPb1.Value = tsPb1.Maximum;
                    }
                    #endregion
                }


                //clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции Apogei\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла Ileko
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void Ileko(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+(?:г\.?)?\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?)").Match(Path.GetFileName(filePath)).Value;
                textBoxOrgName.Text = orgname;

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла Ileko\n\n" + ex.ToString()); }

                string temp = "";
                //this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    structTab tab = new structTab();
                    tab.listdtProductIndexRow = new List<int>();
                    tab.listExcelIndexTab = new List<int>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    Regex regName = new Regex(@"лента|лист|арматура|полоса|угол|швеллер|труб|круг|шестигранник|шгр|квадрат|полоса|сталь|катанка|быстрорез", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой(?:\s*проч)|\w+ий|г[\/]к", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"(?<=\s)\d+(?:[,.]\d+)?(?:[xх/-]\d+(?:[,.]\d+)?)?(?:\s*[x*х]\s*\d+(?:[,.]\d+)?(?:[/-]\d+(?:[,.]\d+)?)?){2,}(?:[/-]\d+(?:[,.]\d+)?)?(?=\s|;|$|-\d\s)", RegexOptions.IgnoreCase);
                    Regex regDiam2 = new Regex(@"(?<=\s)\d+(?:[,.]\d+)?(?:[xх/-]\d+(?:[,.]\d+)?)?(?:\s*[x*х]\s*\d+(?:[,.]\d+)?(?:[/-]\d+(?:[,.]\d+)?)?)(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"(?:\d{1,3}[ШСТУХ]+\d{0,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+(?:\d{0,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+)*\d{0,3})(?=\s+|$)|(?:(?:Ст.)|(?:ст.)(?:\s*\d{1,2})\b)(?=\s+|$)|(?:\d{,3}[ХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)(?=\s+|$)|(?:[АA]-?\d)(?=\s+|$)|(?<=[Сс][Тт]\.\s?)\d{1,2}[гГ]\d{1,2}[cCсС]|(?<=ст\.)\d{1,2}[хфа]+(?=\s|$)|(?<=\s)[сС][тТ]\.?\s?\d{1,2}[_\w]+|(?<=\s)[0-4]\d[хx]\d\d[ТН]*(?=\s)|AISI\s*\d+\w*\s", RegexOptions.IgnoreCase);

                    int lastRow = 0;
                    int ColName = 0, ColCol = 0;
                    nameProd = "";
                    string type = "", tmp = "", regDiamString = "";

                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;

                    #region поиск заголовков
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        int jj = j;
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"Название", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    tab.StartRow = j + 1; //отсюда начинать поиск данных
                                    ColName = i;
                                    j = cCelRow;
                                }

                                if (new Regex(@"Остаток", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColCol = i;
                                    j = cCelRow;
                                }
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }
                    #endregion

                    tsLabelClearingTable.Text = "Обработка";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow - tab.StartRow;
                    if (ColName > 0)
                        for (int j = tab.StartRow; j <= cCelRow; j++) //строки
                        {
                            regDiamString = "";
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, ColName];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                nameProd = regName.Match(temp).Value;
                                regDiamString = regDiam.Match(temp).Value;
                                if (regDiamString == "")
                                    regDiamString = regDiam2.Match(temp).Value;
                                if (nameProd != "" && regDiamString != "")
                                {
                                    dtProduct.Rows.Add();
                                    lastRow = dtProduct.Rows.Count - 1;
                                    tab.listExcelIndexTab.Add(j);
                                    tab.listdtProductIndexRow.Add(lastRow);

                                    if (new Regex(@"ш\s*г\s*[рp]", RegexOptions.IgnoreCase).IsMatch(temp))
                                        dtProduct.Rows[lastRow]["Название"] = "Шестигранник";
                                    else
                                        dtProduct.Rows[lastRow]["Название"] = nameProd;

                                    dtProduct.Rows[lastRow]["Примечание"] = temp;
                                    dtProduct.Rows[lastRow]["Тип"] = regType.Match(temp).Value;
                                    if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                    {
                                        if (new Regex(@"г[\/]к", RegexOptions.IgnoreCase).IsMatch(type)) type = "горячекатаный";
                                        dtProduct.Rows[lastRow]["Тип"] = type;
                                        if (dtProduct.Rows[lastRow]["Тип"].ToString() == "") dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                                    }

                                    GetRegexMarkFromString(temp, lastRow);
                                    //dtProduct.Rows[lastRow]["Марка"] = regMark.Match(temp).Value;
                                    Excel.Range cellRangeOst = (Excel.Range)excelworksheet.Cells[j, ColCol];
                                    if (cellRangeOst.Value != null)
                                        tmp = cellRangeOst.Value.ToString().Trim();
                                    else tmp = "";
                                    if (tmp != "")
                                    {
                                        dtProduct.Rows[lastRow]["Мерность (т, м, мм)"] = tmp;
                                    }

                                    dtProduct.Rows[lastRow]["Марка"] = regMark.Match(temp).Value;

                                    string[] diam, tolsh, metraj;
                                    string tempo = "";
                                    if (regDiam.IsMatch(temp))
                                        tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(regDiamString).Value;
                                    else if (regDiam2.IsMatch(temp))
                                        tempo = new Regex(@"(?<=^)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(regDiamString).Value;
                                    if (tempo != "")
                                    {
                                        diam = tempo.Split('-');
                                    }
                                    else
                                    {
                                        tempo = "";
                                        if (regDiam.IsMatch(temp))
                                            tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(regDiamString).Value;
                                        else if (regDiam2.IsMatch(temp))
                                            tempo = new Regex(@"(?<=^)\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(regDiamString).Value;
                                        if (tempo != "")
                                        {
                                            diam = new string[] { tempo };
                                        }
                                        else diam = new string[] { "" };
                                    }
                                    if (regDiam.IsMatch(temp))
                                        tempo = new Regex(@"(?<=^)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(regDiamString).Value;
                                    else if (regDiam2.IsMatch(temp))
                                        tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*$)", RegexOptions.IgnoreCase).Match(regDiamString).Value;
                                    if (tempo != "")
                                    {
                                        tolsh = tempo.Split('-');
                                    }
                                    else
                                    {
                                        tempo = "";
                                        if (regDiam.IsMatch(temp))
                                            tempo = new Regex(@"(?<=^)\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(regDiamString).Value;
                                        else if (regDiam2.IsMatch(temp))
                                            tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s*$)", RegexOptions.IgnoreCase).Match(regDiamString).Value;
                                        if (tempo != "")
                                        {
                                            tolsh = new string[] { tempo };
                                        }
                                        else tolsh = new string[] { "" };
                                    }
                                    if (regDiam.IsMatch(temp))
                                        tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(regDiamString).Value;
                                    else if (regDiam2.IsMatch(temp))
                                        tempo = "";
                                    if (tempo != "")
                                    {
                                        metraj = tempo.Split('-');
                                    }
                                    else
                                    {
                                        if (regDiam.IsMatch(temp))
                                            tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?(?:\s*/\s*\d+(?:[,\.]\d+)?)+(?=\s|$|-\d\s)", RegexOptions.IgnoreCase).Match(regDiamString).Value;
                                        else if (regDiam2.IsMatch(temp))
                                            tempo = "";
                                        if (tempo != "")
                                        {
                                            metraj = tempo.Split('/');
                                        }
                                        else
                                        {
                                            tempo = "";
                                            if (regDiam.IsMatch(temp))
                                                tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(regDiamString).Value;
                                            else if (regDiam2.IsMatch(temp))
                                                tempo = "";
                                            if (tempo != "")
                                            {
                                                metraj = new string[] { tempo };
                                            }
                                            else metraj = new string[] { "" };
                                        }
                                    }

                                    for (int d = 0; d < diam.Length; d++)
                                        for (int t = 0; t < tolsh.Length; t++)
                                            for (int m = 0; m < metraj.Length; m++)
                                            {
                                                lastRow = dtProduct.Rows.Count - 1;
                                                if (d == 0 && t == 0 && m == 0)
                                                {
                                                    dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = diam[0];
                                                    dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = tolsh[0];
                                                    dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = metraj[0];
                                                }
                                                else
                                                {
                                                    DataRow row = dtProduct.NewRow();
                                                    row["Название"] = dtProduct.Rows[lastRow]["Название"];
                                                    row["Тип"] = dtProduct.Rows[lastRow]["Тип"];
                                                    row["Диаметр (высота), мм"] = diam[d];
                                                    row["Толщина (ширина), мм"] = tolsh[t];
                                                    row["Метраж, м (длина, мм)"] = metraj[m];
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
                            }
                            else nameProd = "";
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }

                }

                //clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции Ileko\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла Demidov
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void Demidov(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+(?:г\.?)?\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?)").Match(Path.GetFileName(filePath)).Value;
                orgname = "Демидов";
                textBoxOrgName.Text = "Демидов";

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла Demidov\n\n" + ex.ToString()); }

                string temp = "";
                //this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                Excel.Worksheet excelworksheet = (Excel.Worksheet)excelsheets[excelsheets.Count];
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    structTab tab = new structTab();
                    tab.listdtProductIndexRow = new List<int>();
                    tab.listExcelIndexTab = new List<int>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    Regex regName = new Regex(@"лента|лист|арматура|полоса|угол|швеллер|труб|круг|шестигранник|шгр|квадрат|полоса|сталь|катанка|быстрорез", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой(?:\s*проч)|\w+ий|г[\/]к", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"(?<=\s)\d+(?:[,.]\d+)?(?:[xх/-]\d+(?:[,.]\d+)?)?(?:(?:\s*[x*х]\s*\d+(?:[,.]\d+)?(?:[/-]\d+(?:[,.]\d+)?)?){2,}(?:[/-]\d+(?:[,.]\d+)?)?)?(?=\s|;|$|-\d\s)", RegexOptions.IgnoreCase);
                    Regex regDiam2 = new Regex(@"(?<=\s)\d+(?:[,.]\d+)?(?:[xх/-]\d+(?:[,.]\d+)?)?(?:\s*[x*х]\s*\d+(?:[,.]\d+)?(?:[/-]\d+(?:[,.]\d+)?)?)(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"(?:\d{1,3}[ШСТУХ]+\d{0,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+(?:\d{0,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+)*\d{0,3})(?=\s+|$)|(?:(?:Ст.)|(?:ст.)(?:\s*\d{1,2})\b)(?=\s+|$)|(?:\d{,3}[ХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)(?=\s+|$)|(?:[АA]-?\d)(?=\s+|$)|(?<=[Сс][Тт]\.\s?)\d{1,2}[гГ]\d{1,2}[cCсС]|(?<=ст\.)\d{1,2}[хфа]+(?=\s|$)|(?<=\s)[сС][тТ]\.?\s?\d{1,2}[_\w]+|(?<=\s)[0-4]\d[хx]\d\d[ТН]*(?=\s)|AISI\s*\d+\w*\s", RegexOptions.IgnoreCase);

                    int lastRow = 0;

                    nameProd = "";
                    string name = "", type = "", tmp = "", regDiamString = "", mark = "", tempPrice = "";

                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;

                    #region обработка
                    for (int i = 1; i <= cCelCol; i++) //столбцы
                    {
                        for (int j = 1; j <= cCelRow; j++) //строки
                        {
                            int jj = j;
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (cellRange.MergeArea.Columns.Count > 1)
                                {
                                    nameProd = regName.Match(temp).Value;
                                    mark = regMark.Match(temp).Value;
                                    type = regType.Match(temp).Value;
                                }
                                else
                                {
                                    cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                                    if (cellRange.Value != null)
                                        tmp = cellRange.Value.ToString().Trim();
                                    else tmp = "";
                                    if (tmp != "")
                                    {
                                        if (regName.IsMatch(tmp))
                                            name = regName.Match(tmp).Value;
                                        else
                                            name = nameProd;

                                        if (regDiam.IsMatch(tmp))
                                            regDiamString = regDiam.Match(tmp).Value;
                                        else regDiamString = "";

                                        if (name != "" && regDiamString != "")
                                        {
                                            dtProduct.Rows.Add();
                                            tab.listExcelIndexTab.Add(i);
                                            lastRow = dtProduct.Rows.Count - 1;
                                            tab.listdtProductIndexRow.Add(lastRow);
                                            dtProduct.Rows[lastRow]["Название"] = name;

                                            if (regType.IsMatch(tmp))
                                                dtProduct.Rows[lastRow]["Тип"] = regType.Match(tmp).Value;
                                            else dtProduct.Rows[lastRow]["Тип"] = type;
                                            if (type.ToLower().Contains("г/к")) type = "горячекатаный";
                                            if (type.ToLower().Contains("х/к")) type = "холоднокатанный";
                                            dtProduct.Rows[lastRow]["Тип"] = type;
                                            if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                                dtProduct.Rows[lastRow]["Тип"] = "тип не указан";

                                            if (regMark.IsMatch(tmp))
                                                dtProduct.Rows[lastRow]["Марка"] = regMark.Match(tmp).Value;
                                            else dtProduct.Rows[lastRow]["Марка"] = mark;

                                            dtProduct.Rows[lastRow]["Примечание"] = tmp;

                                            Excel.Range cellRangePrice = (Excel.Range)excelworksheet.Cells[jj, i + 1];
                                            if (cellRangePrice.Value != null)
                                                tempPrice = cellRangePrice.Value.ToString().Trim();
                                            else tempPrice = "";
                                            if (tempPrice != "")
                                            {
                                                dtProduct.Rows[lastRow]["Цена"] = tempPrice;
                                            }

                                            string[] diam, tolsh, metraj;
                                            List<double> Ddiam = new List<double>(), Dtolsh = new List<double>(), Dmetraj = new List<double>();
                                            List<double> ch = new List<double>();
                                            string tempo = "";
                                            tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if (tempo != "")
                                            {
                                                diam = tempo.Split('-');
                                                foreach (string e in diam)
                                                    Ddiam.Add(Convert.ToDouble(e));
                                                ch.Clear();
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
                                                    }
                                                    if (ch.Count > 0) Ddiam = ch;
                                                }
                                            }
                                            else
                                            {
                                                tempo = "";
                                                tempo = new Regex(@"(?<=[xх*]\s*)\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(temp).Value;
                                                if (tempo != "")
                                                {
                                                    diam = new string[] { tempo };
                                                }
                                                else
                                                {
                                                    tempo = "";
                                                    tempo = new Regex(@"(?<=\s)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                    if (tempo != "")
                                                    {
                                                        diam = tempo.Split('-');
                                                        foreach (string e in diam)
                                                            Ddiam.Add(Convert.ToDouble(e));
                                                        ch.Clear();
                                                        double increment = 0;
                                                        if (Ddiam[1] >= 1 && Ddiam[1] < 4) increment = 0.5;
                                                        if (Ddiam[1] >= 4 && Ddiam[1] < 50) increment = 2;
                                                        if (Ddiam[1] >= 50) increment = 10;
                                                        if (increment > 0)
                                                        {
                                                            for (double d = Ddiam[0]; d <= Ddiam[1]; d += increment)
                                                            {
                                                                if (d != Ddiam[0] && (d - 0.1) % 1 == 0)
                                                                    d -= 0.1;
                                                                ch.Add(d);
                                                                if (d + increment > Ddiam[1] && d != Ddiam[1]) ch.Add(Ddiam[1]);
                                                            }
                                                            if (ch.Count > 0) Ddiam = ch;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        tempo = "";
                                                        tempo = new Regex(@"(?<=\s)\d+(?:[,\.]\d+)?(?=\w?\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                        if (tempo != "")
                                                        {
                                                            diam = new string[] { tempo };
                                                        }
                                                        else diam = new string[] { "" };
                                                    }
                                                }
                                            }

                                            tempo = new Regex(@"(?<=^|dy\s*)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s*[xх*])", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if (tempo != "")
                                            {
                                                tolsh = tempo.Split('-');
                                            }
                                            else
                                            {
                                                tempo = "";

                                                tempo = new Regex(@"(?<=^|dy\s*)\d+(?:[,\.]\d+)?(?=\w?\s*[xх*])", RegexOptions.IgnoreCase).Match(temp).Value;
                                                if (tempo != "")
                                                {
                                                    tolsh = new string[] { tempo };
                                                }
                                                else tolsh = new string[] { "" };
                                            }
                                            tempo = new Regex(@"(?<=[xх*]\s)\d+(?:[,\.]\d+)?\s*-\s*\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                            if (tempo != "")
                                            {
                                                metraj = tempo.Split('-');
                                            }
                                            else
                                            {
                                                tempo = "";
                                                tempo = new Regex(@"(?<=[xх*]\s)\d+(?:[,\.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
                                                if (tempo != "")
                                                {
                                                    metraj = new string[] { tempo };
                                                }
                                                else metraj = new string[] { "" };

                                            }

                                            if (Ddiam.Count == 0)
                                                foreach (string e in diam)
                                                {
                                                    if (e != "")
                                                        Ddiam.Add(Convert.ToDouble(e));
                                                    else Ddiam.Add(0);
                                                }

                                            if (Dtolsh.Count == 0)
                                                foreach (string e in tolsh)
                                                    if (e != "")
                                                        Dtolsh.Add(Convert.ToDouble(e));
                                                    else Dtolsh.Add(0);

                                            if (Dmetraj.Count == 0)
                                                foreach (string e in metraj)
                                                    if (e != "")
                                                        Dmetraj.Add(Convert.ToDouble(e));
                                                    else Dmetraj.Add(0);

                                            if (Ddiam[0] < Dtolsh[0])
                                            {
                                                ch = Ddiam;
                                                Ddiam = Dtolsh;
                                                Dtolsh = ch;
                                            }


                                            for (int d = 0; d < Ddiam.Count; d++)
                                                for (int t = 0; t < Dtolsh.Count; t++)
                                                    for (int m = 0; m < Dmetraj.Count; m++)
                                                    {
                                                        lastRow = dtProduct.Rows.Count - 1;
                                                        if (d == 0 && t == 0 && m == 0)
                                                        {
                                                            if (Ddiam[0] != 0) dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = Ddiam[0];
                                                            if (Dtolsh[0] != 0) dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = Dtolsh[0];
                                                            if (Dmetraj[0] != 0) dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = Dmetraj[0];
                                                        }
                                                        else
                                                        {
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
                                    }
                                }
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }

                    #endregion
                }

                //clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции Demidov\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла DemidovTruba
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void DemidovTruba(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+(?:г\.?)?\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?)").Match(Path.GetFileName(filePath)).Value;
                orgname = "Демидов_Труба";
                textBoxOrgName.Text = "Демидов_Труба";

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла Demidov\n\n" + ex.ToString()); }

                string temp = "";
                //this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    structTab tab = new structTab();
                    tab.listdtProductIndexRow = new List<int>();
                    tab.listExcelIndexTab = new List<int>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    Regex regName = new Regex(@"лента|лист|арматура|полоса|угол|швеллер|труба|круг|шестигранник|шгр|квадрат|полоса|сталь|катанка|быстрорез", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой(?:\s*проч)|\w+ий|г[\/]к|ВГ|проф|ЭС", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"(?<=\s)\d+(?:[,.]\d+)?(?:[xх/-]\d+(?:[,.]\d+)?)?(?:(?:\s*[x*х]\s*\d+(?:[,.]\d+)?(?:[/-]\d+(?:[,.]\d+)?)?){2,}(?:[/-]\d+(?:[,.]\d+)?)?)?(?=\s|;|$|-\d\s)", RegexOptions.IgnoreCase);
                    Regex regDiam2 = new Regex(@"(?<=\s)\d+(?:[,.]\d+)?(?:[xх/-]\d+(?:[,.]\d+)?)?(?:\s*[x*х]\s*\d+(?:[,.]\d+)?(?:[/-]\d+(?:[,.]\d+)?)?)(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"(?:\d{1,3}[ШСТУХ]+\d{0,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+(?:\d{0,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+)*\d{0,3})(?=\s+|$)|(?:(?:Ст.)|(?:ст.)(?:\s*\d{1,2})\b)(?=\s+|$)|(?:\d{,3}[ХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)(?=\s+|$)|(?:[АA]-?\d)(?=\s+|$)|(?<=[Сс][Тт]\.\s?)\d{1,2}[гГ]\d{1,2}[cCсС]|(?<=ст\.)\d{1,2}[хфа]+(?=\s|$)|(?<=\s)[сС][тТ]\.?\s?\d{1,2}[_\w]+|(?<=\s)[0-4]\d[хx]\d\d[ТН]*(?=\s)|AISI\s*\d+\w*\s", RegexOptions.IgnoreCase);

                    int lastRow = 0, ColName = 1, ColRaz = 1, ColGost = 1, ColDlina = 2, ColPrice = 1;

                    nameProd = "";
                    string type = "", tmp = "", regDiamString = "", tempPrice = "";

                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;

                    #region поиск заголовков
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        int jj = j;
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"^\s*метал", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    tab.StartRow = j + 1; //отсюда начинать поиск данных
                                    ColName = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"Цена.*от.*5", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColPrice = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"^\s*размер", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColRaz = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"^\s*сталь", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColGost = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"^\s*длина", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColDlina = i;
                                    j = cCelRow;
                                }

                                #region сайт
                                if (new Regex(@"(?:www\.)[\w\d-]{2,}\.[A-Za-zА-Яа-я]+").IsMatch(temp)) // поиск сайта
                                {
                                    textBoxOrgSite.Text = new Regex(@"(?:www\.)[\w\d-]{2,}\.[A-Za-zА-Яа-я]+").Match(temp).Value;
                                    //break;
                                }
                                #endregion

                                #region Email
                                Regex regEmail = new Regex(@"[\d\.\w-\*\\]+@[\w-]+\.\w{1,5}");
                                if (regEmail.IsMatch(temp)) // поиск Email
                                {
                                    if (textBoxOrgEmail.Text == "") textBoxOrgEmail.Text = regEmail.Match(temp).Value;
                                    else textBoxOrgEmail.Text += "; " + regEmail.Match(temp).Value;
                                    //break;
                                }
                                #endregion

                                #region адрес организации
                                if (new Regex(@"г.\s*\w+.*\d+(?=\s*$)", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    if (textBoxOrgAdress.Text == "")
                                        textBoxOrgAdress.Text = new Regex(@"г.\s*\w+.*\d+(?=\s*$)").Match(temp).Value;
                                    else
                                        textBoxOrgAdress.Text += ";" + new Regex(@"г.\s*\w+.*\d+(?=\s*$)").Match(temp).Value;
                                }
                                #endregion

                                Regex regOrgTel = new Regex(@"(?<=тел.*)\+7\(\d+\)\d+-\d+-\d+(?=\s*\w+)", RegexOptions.IgnoreCase);
                                if (regOrgTel.IsMatch(temp))
                                    foreach (Match m in regOrgTel.Matches(temp))
                                        if (textBoxOrgTelefon.Text == "")
                                            textBoxOrgTelefon.Text = m.Value;
                                        else textBoxOrgTelefon.Text += "; " + m.Value;

                                Regex regManager = new Regex(@"тел.*\+7\(\d+\)\d+-\d+-\d+", RegexOptions.IgnoreCase);
                                if (regManager.IsMatch(temp))
                                {
                                    ListViewItem lvi = new ListViewItem(new Regex(@"(?<=тел.*\+7\(\d+\)\d+-\d+-\d+\s*)(?:[а-яА-Я]\s*)+(?=\s|\s*$)", RegexOptions.IgnoreCase).Match(temp).Value); //имя менеджера
                                    lvi.SubItems.Add(new Regex(@"(?<=сот.*тел.*)8(?:\s?-?\s?\d+)+", RegexOptions.IgnoreCase).Match(temp).Value);           //телефон менеджера
                                    if (lvi.SubItems[0].Text != "icq") listViewManager.Items.Add(lvi);
                                }
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }
                    #endregion


                    #region обработка
                    for (int j = tab.StartRow; j <= cCelRow; j++) //строки
                    {

                        Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, ColName];
                        if (cellRange.Value != null)
                            temp = cellRange.Value.ToString().Trim();
                        else temp = "";
                        if (temp != "")
                        {
                            nameProd = regName.Match(temp).Value;
                            type = regType.Match(temp).Value;
                            regDiamString = regDiam.Match(temp).Value;
                        }

                        if (nameProd != "" && regDiamString != "")
                        {
                            dtProduct.Rows.Add();
                            tab.listExcelIndexTab.Add(j);
                            lastRow = dtProduct.Rows.Count - 1;
                            tab.listdtProductIndexRow.Add(lastRow);
                            dtProduct.Rows[lastRow]["Название"] = nameProd;

                            if (regType.IsMatch(tmp))
                                dtProduct.Rows[lastRow]["Тип"] = regType.Match(tmp).Value;
                            else dtProduct.Rows[lastRow]["Тип"] = type;
                            if (type.ToLower().Contains("г/к")) type = "горячекатанная";
                            if (type.ToLower().Contains("х/к")) type = "холоднокатанная";
                            if (type.ToLower().Contains("эс")) type = "электросварная";
                            if (type.ToLower().Contains("проф")) type = "профильная";
                            dtProduct.Rows[lastRow]["Тип"] = type;
                            if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                            dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"(?<=\w\s+)\d+(?:[,\.]\d+)?(?=\s|\s*$)", RegexOptions.IgnoreCase).Match(temp).Value;
                            cellRange = (Excel.Range)excelworksheet.Cells[j, ColRaz];
                            if (cellRange.Value != null)
                                tmp = cellRange.Value.ToString().Trim();
                            else tmp = "";

                            if (tmp != "")
                            {
                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\*\s?)\d+(?:[,\.]\d+)?(?=\s|\s*$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=\*\s?)\d+(?:[,\.]\d+)?(?=\*)", RegexOptions.IgnoreCase).Match(tmp).Value;

                                dtProduct.Rows[lastRow]["Марка"] = regMark.Match(tmp).Value;
                            }

                            dtProduct.Rows[lastRow]["Примечание"] = temp + " " + tmp;

                            Excel.Range cellRangePrice = (Excel.Range)excelworksheet.Cells[j, ColPrice];
                            if (cellRangePrice.Value != null)
                                tempPrice = cellRangePrice.Value.ToString().Trim();
                            else tempPrice = "";
                            if (tempPrice != "")
                            {
                                dtProduct.Rows[lastRow]["Цена"] = tempPrice;
                            }

                            Excel.Range cellRangeGost = (Excel.Range)excelworksheet.Cells[j, ColGost];
                            if (cellRangeGost.Value != null)
                                tempPrice = cellRangeGost.Value.ToString().Trim();
                            else tempPrice = "";
                            if (tempPrice != "")
                            {
                                dtProduct.Rows[lastRow]["Стандарт"] = tempPrice;
                            }

                            Excel.Range cellRangeDlina = (Excel.Range)excelworksheet.Cells[j, ColDlina];
                            if (cellRangeDlina.Value != null)
                                tempPrice = cellRangeDlina.Value.ToString().Trim();
                            else tempPrice = "";
                            if (tempPrice != "")
                            {
                                dtProduct.Rows[lastRow]["Мерность (т, м, мм)"] = tempPrice;
                            }

                        }

                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        else tsPb1.Value = tsPb1.Maximum;
                    }


                    #endregion
                }

                //clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции Demidov\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла StalMaksimum
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void StalMaksimum(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+(?:г\.?)?\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?)").Match(Path.GetFileName(filePath)).Value;
                orgname = "СтальМаксимум";
                textBoxOrgName.Text = orgname;

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла StalMaksimum\n\n" + ex.ToString()); }

                string temp = "";
                //this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;
                countEmpty = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    structTab tab = new structTab();
                    tab.listdtProductIndexRow = new List<int>();
                    tab.listExcelIndexTab = new List<int>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    Regex regName = new Regex(@"лента|лист|арматура|полоса|угол|швеллер|труба|круг|шестигранник|шгр|квадрат|полоса|сталь|катанка|быстрорез|проволока", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой(?:\s*проч)|\w+ий|г[\/]к|ВГ|проф|ЭС", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"\d+(?:[,.]\d+)?(?:[xх/-]\d+(?:[,.]\d+)?)?(?:(?:\s*[x*х]\s*\d+(?:[,.]\d+)?(?:[/-]\d+(?:[,.]\d+)?)?){2,}(?:[/-]\d+(?:[,.]\d+)?)?)?(?=\s|;|$|-\d\s)|\d+(?:[,.]\d+)?\s?\*\s?\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase);
                    Regex regDiam2 = new Regex(@"(?<=\s)\d+(?:[,.]\d+)?(?:[xх/-]\d+(?:[,.]\d+)?)?(?:\s*[x*х]\s*\d+(?:[,.]\d+)?(?:[/-]\d+(?:[,.]\d+)?)?)(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"(?:\d{1,3}[ШСТУХ]+\d{0,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+(?:\d{0,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+)*\d{0,3})(?=\s+|$)|(?:(?:Ст.)|(?:ст.)(?:\s*\d{1,2})\b)(?=\s+|$)|(?:\d{,3}[ХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)(?=\s+|$)|(?:[АA]-?\d)(?=\s+|$)|(?<=[Сс][Тт]\.\s?)\d{1,2}[гГ]\d{1,2}[cCсС]|(?<=ст\.)\d{1,2}[хфа]+(?=\s|$)|(?<=\s)[сС][тТ]\.?\s?\d{1,2}[_\w]+|(?<=\s)[0-4]\d[хx]\d\d[ТН]*(?=\s)|AISI\s*\d+\w*\s", RegexOptions.IgnoreCase);

                    int lastRow = 0, ColMark = 1, ColRaz = 1, ColGost = 1, ColObem = 2, ColPrice = 1;

                    nameProd = "";
                    string type = "", tmp = "", regDiamString = "", tempPrice = "";

                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;

                    #region поиск заголовков
                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        int jj = j;
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[jj, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (new Regex(@"^\s*марка", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    tab.StartRow = j + 1; //отсюда начинать поиск данных
                                    ColMark = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"Цена", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColPrice = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"^\s*размер", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColRaz = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"^\s*гост", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColGost = i;
                                    j = cCelRow;
                                }
                                if (new Regex(@"^\s*объем", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    ColObem = i;
                                    j = cCelRow;
                                }

                                #region сайт
                                if (new Regex(@"(?:www\.)[\w\d-]{2,}\.[A-Za-zА-Яа-я]+").IsMatch(temp)) // поиск сайта
                                {
                                    textBoxOrgSite.Text = new Regex(@"(?:www\.)[\w\d-]{2,}\.[A-Za-zА-Яа-я]+").Match(temp).Value;
                                    //break;
                                }
                                #endregion

                                #region Email
                                Regex regEmail = new Regex(@"[\d\.\w-\*\\]+@[\w-]+\.\w{1,5}");
                                if (regEmail.IsMatch(temp)) // поиск Email
                                {
                                    if (textBoxOrgEmail.Text == "") textBoxOrgEmail.Text = regEmail.Match(temp).Value;
                                    else textBoxOrgEmail.Text += "; " + regEmail.Match(temp).Value;
                                    //break;
                                }
                                #endregion

                                #region адрес организации
                                if (new Regex(@"г.\s*\w+.*\d+(?=\s*$)", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    if (textBoxOrgAdress.Text == "")
                                        textBoxOrgAdress.Text = new Regex(@"г.\s*\w+.*\d+(?=\s*$)").Match(temp).Value;
                                    else
                                        textBoxOrgAdress.Text += ";" + new Regex(@"г.\s*\w+.*\d+(?=\s*$)").Match(temp).Value;
                                }
                                #endregion

                                #region телефон организации
                                Regex regOrgTel = new Regex(@"(?<=тел.*)\+7\(\d+\)\d+-\d+-\d+(?=\s*\w+)", RegexOptions.IgnoreCase);
                                if (regOrgTel.IsMatch(temp))
                                    foreach (Match m in regOrgTel.Matches(temp))
                                        if (textBoxOrgTelefon.Text == "")
                                            textBoxOrgTelefon.Text = m.Value;
                                        else textBoxOrgTelefon.Text += "; " + m.Value;
                                #endregion

                                #region менеджер и телефон менеджера
                                Regex regManager = new Regex(@"тел.*\+7\(\d+\)\d+-\d+-\d+", RegexOptions.IgnoreCase);
                                if (regManager.IsMatch(temp))
                                {
                                    ListViewItem lvi = new ListViewItem(new Regex(@"(?<=тел.*\+7\(\d+\)\d+-\d+-\d+\s*)(?:[а-яА-Я]\s*)+(?=\s|\s*$)", RegexOptions.IgnoreCase).Match(temp).Value); //имя менеджера
                                    lvi.SubItems.Add(new Regex(@"(?<=сот.*тел.*)8(?:\s?-?\s?\d+)+", RegexOptions.IgnoreCase).Match(temp).Value);           //телефон менеджера
                                    if (lvi.SubItems[0].Text != "icq") listViewManager.Items.Add(lvi);
                                }
                                #endregion
                            }
                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }
                    #endregion


                    #region обработка
                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow - tab.StartRow;
                    for (int j = tab.StartRow; j <= cCelRow; j++) //строки
                    {

                        Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, ColRaz];
                        if (cellRange.Value != null)
                            temp = cellRange.Value.ToString().Trim();
                        else temp = "";
                        if (temp != "")
                        {
                            if (new Regex(@"\bкв\.?\b|квадрат", RegexOptions.IgnoreCase).IsMatch(temp))
                                nameProd = "Квадрат";
                            else if (new Regex(@"\bшгр\b|квадрат", RegexOptions.IgnoreCase).IsMatch(temp))
                                nameProd = "Шестигранник";
                            else if (regName.IsMatch(excelworksheet.Name.ToString()) && !excelworksheet.Name.ToLower().Contains("лист"))
                                nameProd = regName.Match(excelworksheet.Name.ToString()).Value;
                            else nameProd = "Труба";

                            regDiamString = regDiam.Match(temp).Value;
                        }

                        if (nameProd != "" && regDiamString != "")
                        {
                            nameProd = nameProd.Substring(0, 1).ToUpper() + nameProd.Substring(1, nameProd.Length - 1).ToLower();
                            regDiamString = "";
                            dtProduct.Rows.Add();
                            tab.listExcelIndexTab.Add(j);
                            lastRow = dtProduct.Rows.Count - 1;
                            tab.listdtProductIndexRow.Add(lastRow);
                            dtProduct.Rows[lastRow]["Название"] = nameProd;

                            if (regType.IsMatch(tmp))
                                dtProduct.Rows[lastRow]["Тип"] = regType.Match(tmp).Value;
                            else dtProduct.Rows[lastRow]["Тип"] = type;
                            if (type.ToLower().Contains("г/к")) type = "горячекатанная";
                            if (type.ToLower().Contains("х/к")) type = "холоднокатанная";
                            if (type.ToLower().Contains("эс")) type = "электросварная";
                            if (type.ToLower().Contains("проф")) type = "профильная";
                            dtProduct.Rows[lastRow]["Тип"] = type;
                            if (dtProduct.Rows[lastRow]["Тип"].ToString() == "")
                                dtProduct.Rows[lastRow]["Тип"] = "тип не указан";

                            cellRange = (Excel.Range)excelworksheet.Cells[j, ColRaz];
                            if (cellRange.Value != null)
                                tmp = cellRange.Value.ToString().Trim();
                            else tmp = "";
                            if (tmp != "")
                            {
                                dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = new Regex(@"\d+(?:[,\.]\d+)?(?=\s|\s*$|\*)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = new Regex(@"(?<=\*\s?)\d+(?:[,\.]\d+)?(?=\s|\s*$)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = new Regex(@"(?<=\*\s?)\d+(?:[,\.]\d+)?(?=\*)", RegexOptions.IgnoreCase).Match(tmp).Value;
                            }

                            dtProduct.Rows[lastRow]["Примечание"] = temp;

                            Excel.Range cellRangePrice = (Excel.Range)excelworksheet.Cells[j, ColPrice];
                            if (cellRangePrice.Value != null)
                                tempPrice = cellRangePrice.Value.ToString().Trim();
                            else tempPrice = "";
                            if (tempPrice != "")
                            {
                                dtProduct.Rows[lastRow]["Цена"] = new Regex(@"\d+(?:[\-,\.]\d+)?", RegexOptions.IgnoreCase).Match(tempPrice).Value;
                                dtProduct.Rows[lastRow]["Цена"] = dtProduct.Rows[lastRow]["Цена"].ToString().Replace('-', ',');
                            }

                            Excel.Range cellRangeGost = (Excel.Range)excelworksheet.Cells[j, ColGost];
                            if (cellRangeGost.Value != null)
                                tempPrice = cellRangeGost.Value.ToString().Trim();
                            else tempPrice = "";
                            if (tempPrice != "")
                            {
                                dtProduct.Rows[lastRow]["Стандарт"] = tempPrice;
                            }

                            Excel.Range cellRangeDlina = (Excel.Range)excelworksheet.Cells[j, ColObem];
                            if (cellRangeDlina.Value != null)
                                tempPrice = cellRangeDlina.Value.ToString().Trim();
                            else tempPrice = "";
                            if (tempPrice != "")
                            {
                                dtProduct.Rows[lastRow]["Мерность (т, м, мм)"] = new Regex(@"\d+(?:[,.]\d+)?\s*\w*", RegexOptions.IgnoreCase).Match(tempPrice).Value;
                            }

                            Excel.Range cellRangeMarka = (Excel.Range)excelworksheet.Cells[j, ColMark];
                            if (cellRangeMarka.Value != null)
                                tempPrice = cellRangeMarka.Value.ToString().Trim();
                            else tempPrice = "";
                            if (tempPrice != "")
                            {
                                dtProduct.Rows[lastRow]["Марка"] = tempPrice;
                            }
                        }

                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        else tsPb1.Value = tsPb1.Maximum;
                    }


                    #endregion
                }

                //clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка в основной функции StalMaksimum\n\ncountIteration = " + countIteration + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла Perm
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void Perm(string path)
        {
            int countIteration = 0;
            try
            {
                if (excelapp != null || excelappworkbook != null)
                {
                    System.Threading.Thread.Sleep(100);
                }
                textBoxPath.Text = path;
                filePath = path;

                orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+(?:г\.?)?\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?)").Match(Path.GetFileName(filePath)).Value;
                orgname = "Пермь";
                textBoxOrgName.Text = orgname;

                SetDateFromName(filePath);

                excelapp = new Excel.Application();
                //excelapp.Visible = true;

                isExcelOpen = true;
                excelappworkbooks = excelapp.Workbooks;

                try
                {
                    excelappworkbook = excelapp.Workbooks.Open(filePath,
            0, true, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

                    excelsheets = excelappworkbook.Worksheets;
                }
                catch (Exception ex) { MessageBox.Show("Ошибка при открытии файла StalMaksimum\n\n" + ex.ToString()); }

                string temp = "";
                //this.Focus();
                //int countRowsIndt = 0; //общее количество строк в результирующей таблице, используется для продолжения результирующей таблицы при переходе к след листу екселя

                isTelefon = false;

                listViewAdrSklad.Items.Clear();
                listViewManager.Items.Clear();

                countRowsForShift = 0;

                tsLabeltotalSheets.Text = excelsheets.Count.ToString();
                foreach (Excel.Worksheet excelworksheet in excelsheets)
                {
                    countIteration++;
                    tsLabelcurrSheet.Text = excelworksheet.Index.ToString();
                    structTab tab = new structTab();
                    tab.listdtProductIndexRow = new List<int>();
                    tab.listExcelIndexTab = new List<int>();

                    int cCelRow = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int cCelCol = excelworksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                    if (cCelCol < 10) cCelCol = 10;
                    if (cCelCol > 20) cCelCol = 20;

                    listIndexOfNotEmptyName = new List<int>();
                    colForName = 0;

                    Regex regName = new Regex(@"лента|(?<=\s|^)лист\b|арматура|полоса|угол|швеллер|труб\w|балк\w|болт\w|круг|шестигранник|шгр|квадрат|полоса|катанка|быстрорез|проволока|шуруп|дюбел\w|сетка|электрод|гвозд\w|заклепк\w|шпильк\w|цеп\w|гайк\w|винт\w|\bпрофиль\b", RegexOptions.IgnoreCase);//(?!\w+ое|\w+ые|\w+ый|\w+ая|\w+ой|\w+ий|\w+\d\w*)(?<=^|\D\s)\w{3,}(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regName2 = new Regex(@"сталь", RegexOptions.IgnoreCase);
                    Regex regType = new Regex(@"\w+ое|\w+ые|\w+ый|\w+ая|\w+ой(?:\s*проч)|\w+ий|г[\/]к|ВГ|проф|ЭС", RegexOptions.IgnoreCase);
                    Regex regDiam = new Regex(@"\d+(?:[,.]\d+)?(?:[xх/-]\d+(?:[,.]\d+)?)?(?:(?:\s*[x*х]\s*\d+(?:[,.]\d+)?(?:[/-]\d+(?:[,.]\d+)?)?){2,}(?:[/-]\d+(?:[,.]\d+)?)?)?(?=\s|;|$|-\d\s)|\d+(?:[,.]\d+)?\s?\*\s?\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase);
                    Regex regDiam2 = new Regex(@"(?<=\s)\d+(?:[,.]\d+)?(?:[xх/-]\d+(?:[,.]\d+)?)?(?:\s*[x*х]\s*\d+(?:[,.]\d+)?(?:[/-]\d+(?:[,.]\d+)?)?)(?=\s|$)", RegexOptions.IgnoreCase);
                    Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:гост\s*)(?:[рР]-?\s*)?(?:\d{1,5}[-\s*]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])|асчм\s*\d+(?:\s*-\s*\d+)*", RegexOptions.IgnoreCase);
                    Regex regMark = new Regex(@"(?:\d{1,3}[ШСТУХ]+\d{0,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+(?:\d{0,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+)*\d{0,3})(?=\s+|$)|(?:(?:Ст.)|(?:ст.)(?:\s*\d{1,2})\b)(?=\s+|$)|(?:\d{,3}[ХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)(?=\s+|$)|(?:[АA]-?\d)(?=\s+|$)|(?<=[Сс][Тт]\.\s?)\d{1,2}[гГ]\d{1,2}[cCсС]|(?<=ст\.)\d{1,2}[хфа]+(?=\s|$)|(?<=\s)[сС][тТ]\.?\s?\d{1,2}[_\w]+|(?<=\s)[0-4]\d[хx]\d\d[ТН]*(?=\s)|AISI\s*\d+\w*\s", RegexOptions.IgnoreCase);


                    int ColRaz = 0, ColTolsh = 0, ColPrice = 0, ColMark = 0, ColDlina = 0;
                    nameProd = "";
                    string tmp = "", regDiamString = "", regTolshString = "", primechanie = "";

                    tsLabelClearingTable.Text = "Поиск имен и их параметров";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = cCelRow * cCelCol;
                    string[] diam, tolsh, metraj;
                    diam = new string[] { "" };
                    tolsh = new string[] { "" };
                    metraj = new string[] { "" };

                    List<structTab> tabs = new List<structTab>();
                    List<double> Ddiam = new List<double>(), Dtolsh = new List<double>(), Dmetraj = new List<double>();
                    List<double> ch = new List<double>();

                    for (int j = 1; j <= cCelRow; j++) //строки
                    {
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange;
                            if (j < 1) j = 1;
                            cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                if (cellRange.MergeArea.Columns.Count == 4)
                                {
                                    if (i < 5)
                                        cellRange = (Excel.Range)excelworksheet.Cells[j, 1];
                                    else cellRange = (Excel.Range)excelworksheet.Cells[j, 6];
                                    if (cellRange.Value != null)
                                        tmp = cellRange.Value.ToString().Trim();
                                    else tmp = "";
                                    if (tmp != "")
                                    {
                                        tab = new structTab();
                                        tab.Name = regName.Match(tmp).Value;
                                        if (tab.Name == "") tab.Name = regName2.Match(tmp).Value;
                                        if (tab.Name != "")
                                        {
                                            tab.Name = NameProdUpLower(tab.Name);
                                            tab.StartRow = j + 1;
                                            if (i < 5)
                                                tab.StartCol = 1;
                                            else tab.StartCol = 6;
                                            tab.Type = regType.Match(tmp).Value;

                                            int ii = 0;
                                            for (int k = tab.StartCol; k < tab.StartCol + 4; k++)
                                            {
                                                cellRange = (Excel.Range)excelworksheet.Cells[j - 1, k];
                                                if (cellRange.Value != null)
                                                    tmp = cellRange.Value.ToString().Trim();
                                                else tmp = "";
                                                if (tmp == "")
                                                {
                                                    ii++;
                                                }
                                            }
                                            if (ii == 4) tab.Razriv = j - 1;
                                            else tab.Razriv = 0;

                                            tabs.Add(tab);
                                            i += 5;
                                        }
                                    }
                                }
                            }

                            if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                            else tsPb1.Value = tsPb1.Maximum;
                        }
                    }

                    for (int j = 1; j <= tabs[0].StartRow; j++) //строки
                    {
                        for (int i = 1; i <= cCelCol; i++) //столбцы
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, i];
                            if (cellRange.Value != null)
                                temp = cellRange.Value.ToString().Trim();
                            else temp = "";
                            if (temp != "")
                            {
                                #region общая информация
                                #region сайт
                                if (new Regex(@"(?:www\.)[\w\d-]{2,}\.[A-Za-zА-Яа-я]+").IsMatch(temp)) // поиск сайта
                                {
                                    textBoxOrgSite.Text = new Regex(@"(?:www\.)[\w\d-]{2,}\.[A-Za-zА-Яа-я]+").Match(temp).Value;
                                    //break;
                                }
                                #endregion

                                #region Email
                                Regex regEmail = new Regex(@"[\d\.\w-\*\\]+@[\w-]+\.\w{1,5}");
                                if (regEmail.IsMatch(temp)) // поиск Email
                                {
                                    if (textBoxOrgEmail.Text == "") textBoxOrgEmail.Text = regEmail.Match(temp).Value;
                                    else textBoxOrgEmail.Text += "; " + regEmail.Match(temp).Value;
                                    //break;
                                }
                                #endregion

                                #region адрес организации
                                if (new Regex(@"г.\s*\w+.*\d+(?=\s*$)", RegexOptions.IgnoreCase).IsMatch(temp))
                                {
                                    if (textBoxOrgAdress.Text == "")
                                        textBoxOrgAdress.Text = new Regex(@"г.\s*\w+.*\d+(?=\s*$)").Match(temp).Value;
                                    else
                                        textBoxOrgAdress.Text += ";" + new Regex(@"г.\s*\w+.*\d+(?=\s*$)").Match(temp).Value;
                                }
                                #endregion

                                #region телефон организации
                                Regex regOrgTel = new Regex(@"\d(?:[-\s]\d+){4,}", RegexOptions.IgnoreCase);
                                if (regOrgTel.IsMatch(temp))
                                    foreach (Match m in regOrgTel.Matches(temp))
                                        if (textBoxOrgTelefon.Text == "")
                                            textBoxOrgTelefon.Text = m.Value;
                                        else textBoxOrgTelefon.Text += "; " + m.Value;
                                #endregion

                                #region менеджер и телефон менеджера
                                Regex regManager = new Regex(@"(?:тел.*)?(?:\+7)?\(\d+\)\d+-\d+-\d+", RegexOptions.IgnoreCase);
                                if (regManager.IsMatch(temp))
                                {
                                    ListViewItem lvi = new ListViewItem(new Regex(@"(?:[а-яА-Я]\s*)+(?=(?:тел.*)?(?:\+7)?\(\d+\)\d+-\d+-\d+\s*)", RegexOptions.IgnoreCase).Match(temp).Value); //имя менеджера
                                    lvi.SubItems.Add(new Regex(@"\(\d+\)\d+-\d+-\d+", RegexOptions.IgnoreCase).Match(temp).Value);           //телефон менеджера
                                    if (lvi.SubItems[0].Text != "icq") listViewManager.Items.Add(lvi);
                                }
                                #endregion
                                #endregion
                            }
                        }
                    }

                    tsLabelClearingTable.Text = "Обработка найденных таблиц...";
                    tsPb1.Value = 0;
                    tsPb1.Maximum = tabs.Count;
                    for (int k = 0; k < tabs.Count; k++)
                    {
                        tab = tabs[k];
                        int endRow = cCelRow;

                        //поиск последней строки для текущей мини-таблицы
                        for (int z = k + 1; z < tabs.Count - 1; z++)
                            if (tab.StartCol == tabs[z].StartCol)
                            { endRow = tabs[z].StartRow - 1; break; }

                        //обнуление позиций столбцов заголовков
                        ColRaz = 0; ColTolsh = 0; ColPrice = 0; ColMark = 0; ColDlina = 0;
                        regDiamString = "";
                        regTolshString = "";
                        //поиск позиций столбцов заголовков для текущей мини-таблицы

                        for (int i = tab.StartCol; i < tab.StartCol + 4; i++)
                        {
                            Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[tab.StartRow, i];
                            if (cellRange.Value != null)
                                tmp = cellRange.Value.ToString().Trim();
                            else tmp = "";
                            if (tmp != "")
                            {
                                #region Определение заголовков
                                if (new Regex(@"диаметр|полка|размер", RegexOptions.IgnoreCase).IsMatch(tmp))
                                    ColRaz = i;
                                else if (new Regex(@"толщина|ширина", RegexOptions.IgnoreCase).IsMatch(tmp))
                                    ColTolsh = i;
                                else if (new Regex(@"Цена", RegexOptions.IgnoreCase).IsMatch(tmp))
                                    ColPrice = i;
                                else if (new Regex(@"Марка", RegexOptions.IgnoreCase).IsMatch(tmp))
                                    ColMark = i;
                                else if (new Regex(@"длина", RegexOptions.IgnoreCase).IsMatch(tmp))
                                    ColDlina = i;
                                #endregion
                            }
                            if (cellRange.MergeArea.Rows.Count == 2) tab.StartRow++;
                        }

                        if ((ColRaz != 0 || ColTolsh != 0) && ColPrice != 0)
                            for (int j = tab.StartRow + 1; j < endRow; j++)
                            {
                                Ddiam = new List<double>(); Dtolsh = new List<double>(); Dmetraj = new List<double>();
                                ch = new List<double>();
                                diam = new string[] { "" }; tolsh = new string[] { "" }; metraj = new string[] { "" };
                                primechanie = "";

                                #region размер
                                if (ColRaz != 0)
                                {
                                    Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, ColRaz];
                                    if (cellRange.Value != null)
                                        tmp = cellRange.Value.ToString().Trim();
                                    else tmp = "";
                                    if (tmp != "")
                                    {
                                        regDiamString = new Regex(@"\d+(?:[,\.]\d+)?(?=\s|\s*$|\*|;)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                        if (new Regex(@"\d+(?:[,\.]\d+)?(?:\s*;\s*\d+(?:[,\.]\d+)?)+", RegexOptions.IgnoreCase).IsMatch(tmp))
                                        {
                                            diam = tmp.Split(';');
                                        }
                                        //else if (new Regex(@"\d+(?:[,\.]\d+)?\s*[xх]\s*\d+(?:[,\.]\d+)?(?=\s|\s*$)", RegexOptions.IgnoreCase).IsMatch(tmp))
                                        //{
                                        //    diam = new string[] { new Regex(@"\d+(?:[,\.]\d+)?(?=\s*[xх]\s*\d+(?:[,\.]\d+)?\s|\s*$)", RegexOptions.IgnoreCase).Match(tmp).Value };
                                        //}
                                        else if (new Regex(@"\d+(?:[,\.]\d+)?(?=\s*[xх]\s*\d+(?:[,\.]\d+)?(?:\s|\s*$))", RegexOptions.IgnoreCase).IsMatch(tmp))
                                        {
                                            diam = new string[] { new Regex(@"\d+(?:[,\.]\d+)?(?=\s*[xх]\s*\d+(?:[,\.]\d+)?(?:\s|\s*$))", RegexOptions.IgnoreCase).Match(tmp).Value };
                                            metraj = new string[] { new Regex(@"(?<=\d+(?:[,\.]\d+)?\s*[xх]\s*)\d+(?:[,\.]\d+)?(?=\s|\s*$)", RegexOptions.IgnoreCase).Match(tmp).Value };
                                        }
                                        else if (new Regex(@"\d+(?:[,\.]\d+)?(?=\s|\s*$|\*)", RegexOptions.IgnoreCase).IsMatch(tmp))
                                        {
                                            diam = new string[] { new Regex(@"\d+(?:[,\.]\d+)?(?=\s|\s*$|\*)", RegexOptions.IgnoreCase).Match(tmp).Value };
                                        }
                                        else diam = new string[] { "" };
                                        primechanie = "Диаметр/ширина = " + tmp;
                                    }
                                    else
                                    {
                                        bool stop = false;
                                        //если пустая ячейка в размерах, то идем в обратно вверх и проверяем на пустные ячейки
                                        if (ColRaz < 5)
                                            for (int z = k - 1; z >= 0; z--)
                                            {
                                                if (z < 1) z = 0;
                                                if (tab.StartCol == tabs[z].StartCol)
                                                {
                                                    if (tabs[z].Razriv != 0 || z == 0)
                                                    {
                                                        ColRaz += 5;
                                                        if (z == 0) j = tabs[z].StartRow - 1;
                                                        else j = tabs[z].Razriv - 1;
                                                        int n = j;
                                                        while (true)
                                                        {
                                                            cellRange = (Excel.Range)excelworksheet.Cells[n, ColRaz];
                                                            if (cellRange.MergeArea.Columns.Count == 4)
                                                            {
                                                                endRow = n;
                                                                stop = true;
                                                                break;
                                                            }
                                                            if (cellRange.Value != null)
                                                                tmp = cellRange.Value.ToString().Trim();
                                                            else tmp = "";
                                                            if (tmp != "")
                                                            {
                                                                j = n;
                                                                stop = true;
                                                                break;
                                                            }
                                                            else
                                                            {
                                                                if (cellRange.MergeArea.Rows.Count == 3) n += 2;
                                                                if (cellRange.MergeArea.Rows.Count == 2) n++;
                                                                n++;
                                                            }
                                                        }
                                                        break;
                                                    }
                                                }
                                            }
                                        else
                                        {
                                            if (k == tabs.Count - 1) break;
                                            int n;
                                            j++;
                                            n = j;
                                            while (true)
                                            {
                                                cellRange = (Excel.Range)excelworksheet.Cells[n, ColRaz];
                                                if (cellRange.MergeArea.Columns.Count == 4)
                                                {
                                                    ColRaz -= 5;
                                                    j = n;
                                                    break;
                                                }
                                                else
                                                {
                                                    if (cellRange.Value != null)
                                                        tmp = cellRange.Value.ToString().Trim();
                                                    else tmp = "";
                                                    if (tmp != "")
                                                    {
                                                        ColRaz -= 5;
                                                        j = n;
                                                        break;
                                                    }
                                                }
                                                if (cellRange.MergeArea.Rows.Count == 3) n += 2;
                                                if (cellRange.MergeArea.Rows.Count == 2) n++;
                                                n++;
                                            }
                                            while (true)
                                            {
                                                cellRange = (Excel.Range)excelworksheet.Cells[n, ColRaz];
                                                if (cellRange.MergeArea.Columns.Count == 4)
                                                {
                                                    stop = true;
                                                    break;
                                                }
                                                else
                                                {
                                                    if (cellRange.Value != null)
                                                        tmp = cellRange.Value.ToString().Trim();
                                                    else tmp = "";
                                                    if (tmp != "")
                                                    {
                                                        if (n <= j)
                                                        {
                                                            if (cellRange.MergeArea.Rows.Count == 3) n -= 2;
                                                            if (cellRange.MergeArea.Rows.Count == 2) n--;
                                                            n--;
                                                        }
                                                        if (n > j)
                                                        {
                                                            j = n;
                                                            stop = true;
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (n >= j)
                                                        {
                                                            if (cellRange.MergeArea.Rows.Count == 3) n += 2;
                                                            if (cellRange.MergeArea.Rows.Count == 2) n++;
                                                            n++;
                                                        }
                                                        if (n < j)
                                                        {
                                                            j = n;
                                                            stop = true;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        if (stop)
                                        {
                                            //поиск последней строки для текущей мини-таблицы
                                            for (int z = 0; z < tabs.Count - 1; z++)
                                                if (ColRaz < 5 && tabs[z].StartRow >= j)
                                                {
                                                    if (tabs[z].StartCol == 1)
                                                    { j--; endRow = tabs[z].StartRow - 1; break; }
                                                }
                                                else if (ColRaz > 5 && tabs[z].StartRow >= j)
                                                {
                                                    if (tabs[z].StartCol == 6)
                                                    { j--; endRow = tabs[z].StartRow - 1; break; }
                                                }

                                            continue;
                                        }
                                    }
                                }
                                #endregion

                                #region толщина
                                if (ColTolsh != 0)
                                {
                                    Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, ColTolsh];
                                    if (cellRange.Value != null)
                                        tmp = cellRange.Value.ToString().Trim();
                                    else tmp = "";
                                    if (tmp != "")
                                    {
                                        regTolshString = new Regex(@"\d+(?:[,\.]\d+)?", RegexOptions.IgnoreCase).Match(tmp).Value;
                                        if (new Regex(@"\d+(?:[,\.]\d+)?(?:\s*;\s*\d+(?:[,\.]\d+)?)+", RegexOptions.IgnoreCase).IsMatch(tmp))
                                        {
                                            tolsh = tmp.Split(';');
                                        }
                                        else if (new Regex(@"\d+(?:[,\.]\d+)?(?=\s|\s*$|\*)", RegexOptions.IgnoreCase).IsMatch(tmp))
                                        {
                                            tolsh = new string[] { tmp };
                                        }
                                        else tolsh = new string[] { "" };
                                    }
                                    else
                                    {
                                        bool stop = false;
                                        //если пустая ячейка в размерах, то идем в обратно вверх и проверяем на пустные ячейки
                                        if (ColTolsh < 5)
                                            for (int z = k - 1; z >= 0; z--)
                                            {
                                                if (z < 1) z = 0;
                                                if (tab.StartCol == tabs[z].StartCol)
                                                {
                                                    if (tabs[z].Razriv != 0 || z == 0)
                                                    {
                                                        ColTolsh += 5;
                                                        if (z == 0) j = tabs[z].StartRow - 1;
                                                        else j = tabs[z].Razriv - 1;
                                                        int n = j;
                                                        while (true)
                                                        {
                                                            cellRange = (Excel.Range)excelworksheet.Cells[n, ColTolsh];
                                                            if (cellRange.MergeArea.Columns.Count == 4)
                                                            {
                                                                endRow = n;
                                                                stop = true;
                                                                break;
                                                            }
                                                            if (cellRange.Value != null)
                                                                tmp = cellRange.Value.ToString().Trim();
                                                            else tmp = "";
                                                            if (tmp != "")
                                                            {
                                                                j = n;
                                                                stop = true;
                                                                break;
                                                            }
                                                            else
                                                            {
                                                                if (cellRange.MergeArea.Rows.Count == 3) n += 2;
                                                                if (cellRange.MergeArea.Rows.Count == 2) n++;
                                                                n++;
                                                            }
                                                        }
                                                        break;
                                                    }
                                                }
                                            }
                                        else
                                        {
                                            if (k == tabs.Count - 1) break;
                                            int n;
                                            j++;
                                            n = j;
                                            while (true)
                                            {
                                                cellRange = (Excel.Range)excelworksheet.Cells[n, ColTolsh];
                                                if (cellRange.MergeArea.Columns.Count == 4)
                                                {
                                                    ColTolsh -= 5;
                                                    j = n;
                                                    break;
                                                }
                                                else
                                                {
                                                    if (cellRange.Value != null)
                                                        tmp = cellRange.Value.ToString().Trim();
                                                    else tmp = "";
                                                    if (tmp != "")
                                                    {
                                                        ColTolsh -= 5;
                                                        j = n;
                                                        break;
                                                    }
                                                }
                                                if (cellRange.MergeArea.Rows.Count == 3) n += 2;
                                                if (cellRange.MergeArea.Rows.Count == 2) n++;
                                                n++;
                                            }
                                            while (true)
                                            {
                                                cellRange = (Excel.Range)excelworksheet.Cells[n, ColTolsh];
                                                if (cellRange.MergeArea.Columns.Count == 4)
                                                {
                                                    stop = true;
                                                    break;
                                                }
                                                else
                                                {
                                                    if (cellRange.Value != null)
                                                        tmp = cellRange.Value.ToString().Trim();
                                                    else tmp = "";
                                                    if (tmp != "")
                                                    {
                                                        if (n <= j)
                                                        {
                                                            if (cellRange.MergeArea.Rows.Count == 3) n -= 2;
                                                            if (cellRange.MergeArea.Rows.Count == 2) n--;
                                                            n--;
                                                        }
                                                        if (n > j)
                                                        {
                                                            j = n;
                                                            stop = true;
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (n >= j)
                                                        {
                                                            if (cellRange.MergeArea.Rows.Count == 3) n += 2;
                                                            if (cellRange.MergeArea.Rows.Count == 2) n++;
                                                            n++;
                                                        }
                                                        if (n < j)
                                                        {
                                                            j = n;
                                                            stop = true;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        if (stop)
                                        {
                                            //поиск последней строки для текущей мини-таблицы
                                            for (int z = 0; z < tabs.Count - 1; z++)
                                                if (ColRaz < 5 && tabs[z].StartRow >= j)
                                                {
                                                    if (tabs[z].StartCol == 1)
                                                    { j--; endRow = tabs[z].StartRow - 1; break; }
                                                }
                                                else if (ColRaz > 5 && tabs[z].StartRow >= j)
                                                {
                                                    if (tabs[z].StartCol == 6)
                                                    { j--; endRow = tabs[z].StartRow - 1; break; }
                                                }

                                            continue;
                                        }
                                    }
                                    primechanie += "   толщина = " + tmp;
                                }
                                #endregion

                                #region длина
                                if (ColDlina != 0)
                                {
                                    Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, ColDlina];
                                    if (cellRange.Value != null)
                                        tmp = cellRange.Value.ToString().Trim();
                                    else tmp = "";
                                    if (tmp != "")
                                    {
                                        regDiamString = new Regex(@"\d+(?:[,\.]\d+)?(?=\s|\s*$|\*|;)", RegexOptions.IgnoreCase).Match(tmp).Value;
                                        if (new Regex(@"\d+(?:[,\.]\d+)?(?:\s*;\s*\d+(?:[,\.]\d+)?)+", RegexOptions.IgnoreCase).IsMatch(tmp))
                                        {
                                            metraj = tmp.Split(';');
                                        }
                                        else if (new Regex(@"\d+(?:[,\.]\d+)?(?=\s|\s*$|\*)", RegexOptions.IgnoreCase).IsMatch(tmp))
                                        {
                                            metraj = new string[] { new Regex(@"\d+(?:[,\.]\d+)?(?=\s|\s*$|\*)", RegexOptions.IgnoreCase).Match(tmp).Value };
                                        }
                                        else metraj = new string[] { "" };
                                        primechanie = "Диаметр/ширина = " + tmp;
                                    }
                                    else metraj = new string[] { "" };
                                }

                                #endregion

                                if (regDiamString != "" || regTolshString != "")
                                {
                                    Excel.Range cellRange;
                                    dtProduct.Rows.Add();
                                    int lastRow = dtProduct.Rows.Count - 1;
                                    if (new Regex(@"балк", RegexOptions.IgnoreCase).IsMatch(tab.Name))
                                        tab.Name = "Балка";
                                    if (new Regex(@"труб", RegexOptions.IgnoreCase).IsMatch(tab.Name))
                                        tab.Name = "Труба";
                                    if (new Regex(@"гвозд", RegexOptions.IgnoreCase).IsMatch(tab.Name))
                                        tab.Name = "Гвоздь";
                                    if (new Regex(@"круг", RegexOptions.IgnoreCase).IsMatch(tab.Name))
                                        tab.Name = "Круг";
                                    if (new Regex(@"уголок", RegexOptions.IgnoreCase).IsMatch(tab.Name))
                                        tab.Name = "Угол";
                                    if (new Regex(@"канат", RegexOptions.IgnoreCase).IsMatch(tab.Name))
                                        tab.Name = "Канат";
                                    if (new Regex(@"гайк", RegexOptions.IgnoreCase).IsMatch(tab.Name))
                                        tab.Name = "Гайка";
                                    if (new Regex(@"шайб", RegexOptions.IgnoreCase).IsMatch(tab.Name))
                                        tab.Name = "Шайба";
                                    if (new Regex(@"шпильк", RegexOptions.IgnoreCase).IsMatch(tab.Name))
                                        tab.Name = "Шпилька";
                                    if (new Regex(@"закл[её]пк", RegexOptions.IgnoreCase).IsMatch(tab.Name))
                                        tab.Name = "Заклепка";
                                    if (new Regex(@"Электрод", RegexOptions.IgnoreCase).IsMatch(tab.Name))
                                        tab.Name = "Электрод";
                                    if (new Regex(@"винт", RegexOptions.IgnoreCase).IsMatch(tab.Name))
                                        tab.Name = "Винт";
                                    if (new Regex(@"цепи", RegexOptions.IgnoreCase).IsMatch(tab.Name))
                                        tab.Name = "Цепь";
                                    if (new Regex(@"шурупы", RegexOptions.IgnoreCase).IsMatch(tab.Name))
                                        tab.Name = "Шуруп";
                                    if (new Regex(@"дюбел", RegexOptions.IgnoreCase).IsMatch(tab.Name))
                                        tab.Name = "Дюбель";
                                    if (new Regex(@"болты", RegexOptions.IgnoreCase).IsMatch(tab.Name))
                                        tab.Name = "Болт";

                                    dtProduct.Rows[lastRow]["Название"] = tab.Name;
                                    dtProduct.Rows[lastRow]["Примечание"] = primechanie;
                                    dtProduct.Rows[lastRow]["Тип"] = tab.Type;

                                    //поиск цены в текущей строке Excel
                                    if (ColPrice != 0)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[j, ColPrice];
                                        if (cellRange.Value != null)
                                            tmp = cellRange.Value.ToString().Trim();
                                        else tmp = "";
                                        if (tmp != "")
                                        {
                                            GetRegexPriceFromString(tmp, lastRow);
                                        }
                                    }
                                    if (ColMark != 0)
                                    {
                                        cellRange = (Excel.Range)excelworksheet.Cells[j, ColMark];
                                        if (cellRange.Value != null)
                                            tmp = cellRange.Value.ToString().Trim();
                                        else tmp = "";
                                        if (tmp != "")
                                        {
                                            dtProduct.Rows[lastRow]["Марка"] = tmp;
                                        }

                                    }
                                    if (dtProduct.Rows[lastRow]["Марка"].ToString().Trim() == "" && dtProduct.Rows[lastRow]["Цена"].ToString().Trim() == "")
                                    {
                                        if (lastRow > 0)
                                        {
                                            dtProduct.Rows[lastRow]["Цена"] = dtProduct.Rows[lastRow - 1]["Цена"];
                                            dtProduct.Rows[lastRow]["Марка"] = dtProduct.Rows[lastRow - 1]["Марка"];
                                        }
                                    }


                                    cellRange = (Excel.Range)excelworksheet.Cells[j, tab.StartCol];
                                    if (cellRange.MergeArea.Rows.Count == 2 && cellRange.MergeArea.Columns.Count == 1) j++;

                                    #region преобразование массива строк в лист десятичных дробей
                                    if (Ddiam.Count == 0)
                                        foreach (string e in diam)
                                        {
                                            if (e != "")
                                                Ddiam.Add(Convert.ToDouble(e.Replace('.', ',')));
                                            else Ddiam.Add(0);
                                        }

                                    if (Dtolsh.Count == 0)
                                        foreach (string e in tolsh)
                                            if (e != "")
                                                Dtolsh.Add(Convert.ToDouble(e.Replace('.', ',')));
                                            else Dtolsh.Add(0);

                                    if (Dmetraj.Count == 0)
                                        foreach (string e in metraj)
                                            if (e != "")
                                                Dmetraj.Add(Convert.ToDouble(e.Replace('.', ',')));
                                            else Dmetraj.Add(0);

                                    if (Ddiam[0] < Dtolsh[0])
                                    {
                                        ch = Ddiam;
                                        Ddiam = Dtolsh;
                                        Dtolsh = ch;
                                    }
                                    #endregion

                                    #region заполнение строк итоговой таблицы из обрабатываемой строки Excel

                                    for (int d = 0; d < Ddiam.Count; d++)
                                        for (int t = 0; t < Dtolsh.Count; t++)
                                            for (int m = 0; m < Dmetraj.Count; m++)
                                            {
                                                if (d == 0 && t == 0 && m == 0)
                                                {
                                                    if (Ddiam[0] != 0) dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = Ddiam[0];
                                                    if (Dtolsh[0] != 0) dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = Dtolsh[0];
                                                    if (Dmetraj[0] != 0) dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = Dmetraj[0];
                                                }
                                                else
                                                {
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
                                                    if (row["Диаметр (высота), мм"].ToString() != "" ||
                                                        row["Толщина (ширина), мм"].ToString() != "" ||
                                                        row["Метраж, м (длина, мм)"].ToString() != "")
                                                        dtProduct.Rows.Add(row);
                                                }
                                            }

                                }
                                #endregion
                            }
                        if (tsPb1.Value < tsPb1.Maximum) tsPb1.Value++;
                        else tsPb1.Value = tsPb1.Maximum;
                    }

                }

                //clearingTable();

                tsPb1.Value = tsPb1.Maximum;
                dataGridView1.DataSource = dtProduct;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка при обработке файла " + Path.GetFileName(path) + "\n\n" + ex.ToString()); }
        }

        /// <summary>
        /// Открытие и чтение экселевского файла А-групп трубы
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void AGroupTrub(string path)
        {
            SetDateFromName(path);
            var file = new Class_A_Group_Trub();
            textBoxOrgName.Text = "А_групп_трубы";
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.workCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));

        }

        /// <summary>
        /// Открытие и чтение экселевского файла А-групп профиль
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void AGroupTrubProf(string path)
        {
            SetDateFromName(path);
            var file = new ClassA_Group_TrubProf();
            textBoxOrgName.Text = "А_групп_трубы_профильные";
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.workCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));

        }

        /// <summary>
        /// Открытие и чтение экселевского файла Атом Рос
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void AtomRos(string path)
        {
            SetDateFromName(path);
            var file = new ClassAtomRos();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.workCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение экселевского файла Инком Металл
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void InkomMetal(string path)
        {
            SetDateFromName(path);
            var file = new Class_inkomMetal();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение вордовского файла КСМ
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void KSM(string path)
        {
            SetDateFromName(path);
            var file = new Class_KSM();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение вордовского файла Гранд_Универсал
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void GrandUniversal(string path)
        {
            SetDateFromName(path);
            var file = new Class_GrandUniversal();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение вордовского файла Кузнецов
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void Kuznetsov(string path)
        {
            SetDateFromName(path);
            var file = new Class_Kuznetsov();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение экселевского файла Гарус
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void Garus(string path)
        {
            SetDateFromName(path);
            var file = new Class_Garus();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение экселевского файла УПТК
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void UPTK(string path)
        {
            SetDateFromName(path);
            var file = new Class_UPTK();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение экселевского файла УралМетСтрой
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void UralMetStroi(string path)
        {
            SetDateFromName(path);
            var file = new Class_UralMetStroi();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение экселевского файла СпецСталь-М
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void SpecStal(string path)
        {
            SetDateFromName(path);
            var file = new Class_SpecStal_M();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение экселевского файла РосПромЦентр
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void RosPromCentr(string path)
        {
            SetDateFromName(path);
            var file = new Class_RosPromCentr();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение экселевского файла Теплообменные трубы
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void TeploobmenTrub(string path)
        {
            SetDateFromName(path);
            var file = new Class_TeploobmenTrub();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение экселевского файла Золотой век
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void ZolotoyVek(string path)
        {
            SetDateFromName(path);
            var file = new Class_ZolotoyVek();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение экселевского файла Золотой век
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void Metchiv(string path)
        {
            SetDateFromName(path);
            var file = new Class_Metchiv();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void MedGora(string path)
        {
            SetDateFromName(path);
            var file = new Class_MedGora();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void Prommet(string path)
        {
            SetDateFromName(path);
            var file = new Class_Prommetal();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void SibMetal(string path)
        {
            SetDateFromName(path);
            var file = new Class_SibMetal();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }
        Thread thread;
        private void Skat(string path)
        {
            SetDateFromName(path);
            var file = new Class_Skat();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void StalMarket(string path)
        {
            SetDateFromName(path);
            var file = new Class_StalMarket();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void StroiTehCentr(string path)
        {
            SetDateFromName(path);
            var file = new Class_StroiTehCentr();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void TEL(string path)
        {
            SetDateFromName(path);
            var file = new Class_TelExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void UMPC(string path)
        {
            SetDateFromName(path);
            var file = new Class_UMPC();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void UralPromMetal(string path)
        {
            SetDateFromName(path);
            var file = new Class_UralPromMetal();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void UralPromMetal120719(string path)
        {
            SetDateFromName(path);
            var file = new Class_UralPromMetal_120719_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void Inplano(string path)
        {
            SetDateFromName(path);
            var file = new Class_InplanoExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void EgidaProm(string path)
        {
            SetDateFromName(path);
            var file = new Class_EgidaPromExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void ProfMet250118(string path)
        {
            SetDateFromName(path);
            var file = new Class_ProfMet_250118_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void ProfMet140819(string path)
        {
            SetDateFromName(path);
            var file = new Class_ProfMet_140819_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void AlfaMetall(string path)
        {
            SetDateFromName(path);
            var file = new Class_AlfaMetallExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void UralskayaMetallobaza(string path)
        {
            SetDateFromName(path);
            var file = new Class_UralskayaMetallobazaExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void EnergoAlyans(string path)
        {
            SetDateFromName(path);
            var file = new Class_EnergoAlyansExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void SpecTruba(string path)
        {
            SetDateFromName(path);
            var file = new Class_SpecTrubaExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void SpecMetKomplekt(string path)
        {
            SetDateFromName(path);
            if (new Regex(@"[\s\d\w]\.xlsx?", RegexOptions.IgnoreCase).IsMatch(path))
            {
                var file = new Class_SpecMetKomplektExcel();
                file.Set(path);
                file.SetMaxValProgressBar += setMaximumTsPb1;
                file.ProcessChanged += incrementTsPb1;
                file.WorkCompleted += dataSourceFunc;
                file.SetInfoOrganization += setInfoOrg;

                thread = new Thread(file.GetTableFromExcel);
                thread.Start();
            }
            else
            {
                var file = new Class_SpecMetKomplekt();
                file.Set(path);
                file.SetMaxValProgressBar += setMaximumTsPb1;
                file.ProcessChanged += incrementTsPb1;
                file.WorkCompleted += dataSourceFunc;
                file.SetInfoOrganization += setInfoOrg;

                thread = new Thread(file.GetTableFromExcel);
                thread.Start();
            }
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void SnabMetalServis(string path)
        {
            SetDateFromName(path);
            var file = new Class_SnabMetalServis();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void StroiTehnolog(string path)
        {
            SetDateFromName(path);
            var file = new Class_StroiTehnolog();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void UralCentrStal(string path)
        {
            SetDateFromName(path);
            var file = new Class_UralCentrStal();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void UralTeploEnergoService(string path)
        {
            SetDateFromName(path);
            if (new Regex(@"[\s\d\w]\.xlsx?", RegexOptions.IgnoreCase).IsMatch(path))
            {
                var file = new Class_UralTeploEnergoServiceExcel();
                file.Set(path);
                file.SetMaxValProgressBar += setMaximumTsPb1;
                file.ProcessChanged += incrementTsPb1;
                file.WorkCompleted += dataSourceFunc;
                file.SetInfoOrganization += setInfoOrg;

                thread = new Thread(file.GetTableFromExcel);
                thread.Start();
            }
            else
            {
                var file = new Class_UralTeploEnergoServiceWord();
                file.Set(path);
                file.SetMaxValProgressBar += setMaximumTsPb1;
                file.ProcessChanged += incrementTsPb1;
                file.WorkCompleted += dataSourceFunc;
                file.SetInfoOrganization += setInfoOrg;

                thread = new Thread(file.GetTableFromExcel);
                thread.Start();
            }
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void MetallGrad(string path)
        {
            SetDateFromName(path);
            if (new Regex(@"\.xlsx?(?=\s*$)", RegexOptions.IgnoreCase).IsMatch(path))
            {
                var file = new Class_MetallGradExcel();
                file.Set(path);
                file.SetMaxValProgressBar += setMaximumTsPb1;
                file.ProcessChanged += incrementTsPb1;
                file.WorkCompleted += dataSourceFunc;
                file.SetInfoOrganization += setInfoOrg;

                thread = new Thread(file.GetTableFromExcel);
                thread.Start();
                //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
            }
            else if (new Regex(@"\.docx?(?=\s*$)", RegexOptions.IgnoreCase).IsMatch(path))
            {
                var file = new Class_MetallGradWord();
                file.Set(path);
                file.SetMaxValProgressBar += setMaximumTsPb1;
                file.ProcessChanged += incrementTsPb1;
                file.WorkCompleted += dataSourceFunc;
                file.SetInfoOrganization += setInfoOrg;

                thread = new Thread(file.GetTableFromExcel);
                thread.Start();
                //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
            }
        }

        private void RegionMetProm(string path)
        {
            SetDateFromName(path);
            var file = new Class_RegionMetPromExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void TD_Metiz(string path)
        {
            SetDateFromName(path);
            var file = new Class_TD_MetizExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void Dakor(string path)
        {
            SetDateFromName(path);
            var file = new Class_DakorExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void Dakor110719(string path)
        {
            SetDateFromName(path);
            var file = new Class_DakorExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void Rostehcom(string path)
        {
            SetDateFromName(path);
            var file = new Class_RostehcomExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void Stalcom(string path)
        {
            SetDateFromName(path);
            var file = new Class_StalcomExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void StalMashUral(string path)
        {
            SetDateFromName(path);
            var file = new Class_StalMashUralExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void MetalBrendUral(string path)
        {
            SetDateFromName(path);
            var file = new Class_MetalBrendUralExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void Almas(string path)
        {
            SetDateFromName(path);
            var file = new Class_AlmasExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void Amet(string path)
        {
            SetDateFromName(path);
            var file = new Class_AmetExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void Prommet_MSK(string path)
        {
            SetDateFromName(path);
            var file = new Class_PromMetMSK_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void MetallInvest(string path)
        {
            SetDateFromName(path);
            var file = new Class_MetallInvestExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void PervNerjKom(string path)
        {
            SetDateFromName(path);
            var file = new Class_perv_nerj_kom_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void StalnoyProfil(string path)
        {
            SetDateFromName(path);
            var file = new Class_StalnoyProfil_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void MetallBasa3(string path)
        {
            SetDateFromName(path);
            var file = new Class_MetallBasa3_EKB_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void TrubaNaSklade(string path)
        {
            SetDateFromName(path);
            var file = new Class_Truba_Na_Sklade_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void SPK080719(string path)
        {
            SetDateFromName(path);
            var file = new Class_SPK080719();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void MetallSnabUral220719(string path)
        {
            SetDateFromName(path);
            var file = new Class_MetallSnabUral_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void StalTranzitZlat(string path)
        {
            SetDateFromName(path);
            var file = new Class_StalTranzit_Zlat_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void UralcherMet190819(string path)
        {
            SetDateFromName(path);
            var file = new Class_UralCherMet190819_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void ChelyabinskProfit290719(string path)
        {
            SetDateFromName(path);
            var file = new Class_ChelyabinskProfit_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void ZolotoyVek240719(string path)
        {
            SetDateFromName(path);
            var file = new Class_ZolotoyVek240719_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void Metallurg220719(string path)
        {
            SetDateFromName(path);
            var file = new Class_Metallurg220719_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void KontinetChel(string path)
        {
            SetDateFromName(path);
            var file = new Class_KontinentChel_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void UTK_Stal_EKB(string path)
        {
            SetDateFromName(path);
            var file = new Class_UTK_Stal_EKB_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void A_Grup120719(string path)
        {
            SetDateFromName(path);
            var file = new Class_A_Grup_120719_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        private void RosTehKom160819(string path)
        {
            SetDateFromName(path);
            var file = new Class_RosTehKom_160819_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }


        private void Atlantic(string path)
        {
            SetDateFromName(path);
            var file = new Class_AtlanticExcel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение вордовского файла Техномет
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void TehnoMet(string path)
        {
            SetDateFromName(path);
            if (new Regex(@"[\s\d\w]\.xlsx?", RegexOptions.IgnoreCase).IsMatch(path))
            {
                var file = new Class_TehnoMetExcel();
                file.Set(path);
                file.SetMaxValProgressBar += setMaximumTsPb1;
                file.ProcessChanged += incrementTsPb1;
                file.WorkCompleted += dataSourceFunc;
                file.SetInfoOrganization += setInfoOrg;

                thread = new Thread(file.GetTableFromExcel);
                thread.Start();
            }
            else
            {
                var file = new Class_TehnoMetWord();
                file.Set(path);
                file.SetMaxValProgressBar += setMaximumTsPb1;
                file.ProcessChanged += incrementTsPb1;
                file.WorkCompleted += dataSourceFunc;
                file.SetInfoOrganization += setInfoOrg;

                thread = new Thread(file.GetTableFromExcel);
                thread.Start();
                //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
            }
        }

        /// <summary>
        /// Открытие и чтение вордовского файла Бинг
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void Bing(string path)
        {
            SetDateFromName(path);
            var file = new Class_BingWord();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение вордовского файла МаксМет
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void MaxMet(string path)
        {
            SetDateFromName(path);
            var file = new Class_MaksMet_Word();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение вордовского файла МаксМет2
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void MaxMet2(string path)
        {
            SetDateFromName(path);
            var file = new Class_MaxMet2Word();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение вордовского файла МеталлПромСнаб
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void MetallPromSnab(string path)
        {
            SetDateFromName(path);
            var file = new Class_MetallPromSnab();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение вордовского файла УралЧерМет
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void UralCherMet_do19082019(string path)
        {
            SetDateFromName(path);
            if (new Regex(@"[\s\d\w]\.xlsx?", RegexOptions.IgnoreCase).IsMatch(path))
            {
                var file = new Class_UralCherMetExcel();
                file.Set(path);
                file.SetMaxValProgressBar += setMaximumTsPb1;
                file.ProcessChanged += incrementTsPb1;
                file.WorkCompleted += dataSourceFunc;
                file.SetInfoOrganization += setInfoOrg;

                thread = new Thread(file.GetTableFromExcel);
                thread.Start();
            }
            else
            {
                var file = new Class_UralCherMet();
                file.Set(path);
                file.SetMaxValProgressBar += setMaximumTsPb1;
                file.ProcessChanged += incrementTsPb1;
                file.WorkCompleted += dataSourceFunc;
                file.SetInfoOrganization += setInfoOrg;

                thread = new Thread(file.GetTableFromExcel);
                thread.Start();
                //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
            }
        }

        /// <summary>
        /// Открытие и чтение вордовского файла TrubaMet
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void TrubaMet(string path)
        {
            SetDateFromName(path);
            var file = new Class_TrubaMetWord();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            Thread thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Открытие и чтение Excel файла АтомПром Комплекс
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        private void AtomPromKomp(string path)
        {
            SetDateFromName(path);
            var file = new Class_Atom_Prom_Kompleks_Excel();
            file.Set(path);
            file.SetMaxValProgressBar += setMaximumTsPb1;
            file.ProcessChanged += incrementTsPb1;
            file.WorkCompleted += dataSourceFunc;
            file.SetInfoOrganization += setInfoOrg;

            thread = new Thread(file.GetTableFromExcel);
            thread.Start();
            //dtProduct = file.GetTableFromExcel(new tsPbDelegate(setMaximumTsPb1), new tsPbDelegate(incrementTsPb1));
        }

        /// <summary>
        /// Метод выполняется по завершении метода GetTableFromExcel класса обрабатываемого шаблона
        /// </summary>
        /// <param name="dt">Передается результирующая таблица</param>
        private void dataSourceFunc(DataTable dt)
        {
            Action action = () =>
                {
                    //dataGridView1.DataSource = dt;
                    dataGridView1.DataSource = AddNerjType(dt);
                    tsLabelClearingTable.Text = "Готово" + ", найдено " + dt.Rows.Count + " строк";
                };
            Invoke(action);
        }

        private DataTable AddNerjType(DataTable dt)
        {
            DataTable dtMarks = new DataTable();
            SqlConnection conn = new SqlConnection(sqlConnString);
            string query = "select mark from marksnerj";
            SqlDataAdapter sda = new SqlDataAdapter(query, conn);
            sda.Fill(dtMarks);
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    if (row["Название"].ToString() == "Трубы") row["Название"] = "Труба";
                    foreach (DataRow rowMark in dtMarks.Rows)
                    {
                        if (new Regex(rowMark[0].ToString(), RegexOptions.IgnoreCase).IsMatch(row["Марка"].ToString()))
                        {
                            row["Тип"] = regexParam.GetTypeIfMarkNerj(row["Название"].ToString());
                        }
                    }
                }
            }
            return dt;
        }

        private void incrementTsPb1(int intValue)
        {
            Action action = () =>
                {
                    tsPb1.Value = intValue;
                };
            Invoke(action);
        }

        private void setMaximumTsPb1(int intMaximumValue)
        {
            Action action = () =>
                {
                    tsPb1.Maximum = intMaximumValue;
                };
            Invoke(action);
        }

        private void setInfoOrg(InfoOrganization infoOrg)
        {
            Action action = () =>
                {
                    textBoxOrgName.Text = infoOrg.OrgName;
                    textBoxOrgAdress.Text = infoOrg.OrgAdress;
                    textBoxOrgTelefon.Text = infoOrg.OrgTel;
                    textBoxOrgEmail.Text = infoOrg.Email;
                    textBoxOrgINN.Text = infoOrg.Inn_Kpp;
                    textBoxOrgSite.Text = infoOrg.Site;
                    textBoxOrgRS.Text = infoOrg.r_s;
                    textBoxOrgKS.Text = infoOrg.k_s;
                    textBoxBIK.Text = infoOrg.BIK;
                    for (int i = 0; i < infoOrg.Manager.Count; i++)
                    {
                        if (infoOrg.Manager[i].Length == 2)
                        {
                            ListViewItem lvi = new ListViewItem(infoOrg.Manager[i][0]);
                            lvi.SubItems.Add(infoOrg.Manager[i][1]);
                            listViewManager.Items.Add(lvi);
                        }
                        if (infoOrg.Manager[i].Length == 3)
                        {
                            ListViewItem lvi = new ListViewItem(infoOrg.Manager[i][0]);
                            lvi.SubItems.Add(infoOrg.Manager[i][1]);
                            lvi.SubItems.Add(infoOrg.Manager[i][2]);
                            listViewManager.Items.Add(lvi);
                        }
                    }
                    for (int i = 0; i < infoOrg.SkladAdr.Count; i++)
                    {
                        ListViewItem lvi = new ListViewItem(infoOrg.SkladAdr[i]);
                        listViewAdrSklad.Items.Add(lvi);
                    }
                };
            Invoke(action);
        }

        /// <summary>
        /// Функция чтения файлов со стандартными данными для диаметров или толщин стенок
        /// </summary>
        /// <param name="PathFile">Путь к файлу</param>
        /// <returns>Возвращает список List-double </returns>
        private List<double> GetStdFromFile(string PathFile)
        {
            List<double> result = new List<double>();
            try
            {
                using (StreamReader sr = new StreamReader(PathFile, System.Text.Encoding.Default))
                {
                    string[] str = sr.ReadToEnd().Split(';');
                    foreach (string s in str)
                        result.Add(Convert.ToDouble(s));
                }
            }
            catch (Exception ex) { MessageBox.Show("Ошибка при преобразовании файла DiamStd or TolshStd\nНевозможно прочитать или привести тип\nОшибка №1111\n\n" + ex.ToString()); }
            return result;
        }

        /// <summary>
        /// Функция записи даты из имени файла
        /// </summary>
        /// <param name="path">путь к файлу</param>
        private void SetDateFromName(string filePath)
        {
            dateTimePicker1.Value = regexParam.GetDateTimeFromName(filePath);
        }

        /// <summary>
        /// Функция установки названия организации из имени файла
        /// </summary>
        /// <param name="filePath">путь к файлу</param>
        private void SetNameFromName(string filePath)
        {
            if (new Regex(@"\.{2,}(?=xlsx?|docx?)", RegexOptions.IgnoreCase).IsMatch(filePath))
            {
                filePath = new Regex(@"\.{2,}(?=xlsx?|docx?)", RegexOptions.IgnoreCase).Replace(filePath, @".");
            }
            orgname = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+\.[\w\d]{3,4}$)|(?<=[\\/]|^)[\w\s]+(?=\.xlsx?|\.docx?)").Match(Path.GetFileName(filePath)).Value;
            textBoxOrgName.Text = orgname;

        }

        //функция поиска в результирующей таблице пустых строк и изъятия из их названия типа продукции
        private void clearingTable()
        {
            tsLabelClearingTable.Text = "Подготовка результирующей таблицы";
            tsPb1.Maximum = dtProduct.Rows.Count; // * dtProduct.Columns.Count;
            tsPb1.Value = 0;
            DataRow subRow;
            DataRow row;
            bool isChange = false;
            bool isChangeDiam = false;
            bool isChangeType = false;
            string strPrice = "";

            for (int j = 0; j < dtProduct.Rows.Count; j++)
            {
                row = dtProduct.Rows[j];


                string diam = row["Диаметр (высота), мм"].ToString().Trim(),
                    tol = row["Толщина (ширина), мм"].ToString().Trim(),
                    met = row["Метраж, м (длина, мм)"].ToString().Trim(),
                    price = row["Цена"].ToString().Trim(),
                    type = row["Тип"].ToString().Trim();

                // если в Цену случайно попало неправильное значение то обнулить ячейку Цены
                strPrice = row["Цена"].ToString();
                Regex regPrice = new Regex(@"(?!,)\D", RegexOptions.IgnoreCase);
                if (regPrice.IsMatch(strPrice))
                {
                    row["Цена"] = "";
                    price = row["Цена"].ToString();
                }

                // если в диаметр случайно попало значение от госта то обнулить ячейку диаметра
                foreach (string s in row["Стандарт"].ToString().Trim().Split(';'))
                {
                    Regex r = new Regex("\\d+", RegexOptions.IgnoreCase);
                    foreach (Match m in r.Matches(row["Стандарт"].ToString()))
                    {
                        if (diam == m.Value)
                        {
                            row["Диаметр (высота), мм"] = "";
                            isChangeDiam = true;
                        }
                    }
                }
                //
                if (isChangeDiam) { diam = ""; isChangeDiam = false; }

                //округление цены
                regPrice = new Regex(@"\d,\d", RegexOptions.IgnoreCase);
                if (regPrice.IsMatch(strPrice))
                {
                    try
                    {
                        row["Цена"] = System.Math.Round(Convert.ToDouble(strPrice), 0);
                        price = row["Цена"].ToString();
                    }
                    catch { }
                }

                //
                foreach (string s in row["Тип"].ToString().Trim().Split(';'))
                {
                    Regex r = new Regex("\\w+", RegexOptions.IgnoreCase);
                    foreach (Match m in r.Matches(row["Стандарт"].ToString()))
                    {
                        if (type == m.Value)
                        {
                            row["Тип"] = "";
                            isChangeType = true;
                        }
                    }
                }
                if (isChangeType) { type = ""; isChangeType = false; }


                if (diam == "" && tol == "" && met == "" && price == "") // если в этой строке нет диаметра, толщины, метража и цены
                {
                    int i = j;
                    if (j < dtProduct.Rows.Count) i = j + 1;
                    if (row["Тип"].ToString() != "тип не указан" && j < dtProduct.Rows.Count - 2) //проверить и заполнить тип из строки-подзаголовка
                    {
                        #region тип
                        subRow = dtProduct.Rows[i];
                        while (dtProduct.Rows[i]["Тип"].ToString() == "тип не указан")
                        {
                            diam = dtProduct.Rows[i]["Диаметр (высота), мм"].ToString();
                            tol = dtProduct.Rows[i]["Толщина (ширина), мм"].ToString();
                            met = dtProduct.Rows[i]["Метраж, м (длина, мм)"].ToString();
                            price = dtProduct.Rows[i]["Цена"].ToString();

                            if (diam == "" && tol == "" && met == "" && price == "")
                            {
                                break;
                            }
                            dtProduct.Rows[i]["Тип"] = row["Тип"];

                            if (dtProduct.Rows[i]["Тип"].ToString().ToLower() == "б/ш") dtProduct.Rows[i]["Тип"] = "бесшовная";
                            if (dtProduct.Rows[i]["Тип"].ToString().ToLower() == "э/св") dtProduct.Rows[i]["Тип"] = "электросварная";
                            //if (dtProduct.Rows[i]["Тип"].ToString().ToLower() == "оцинк") dtProduct.Rows[i]["Тип"] = "оцинкованный";
                            Regex reg = new Regex(@"(?<=\w+)ые");
                            string str = dtProduct.Rows[i]["Тип"].ToString();
                            if (reg.IsMatch(str))//(str.Substring(str.Length - 2, 2) == "ые")
                            {
                                //str = str.Remove(str.Length - 2, 2);
                                dtProduct.Rows[i]["Тип"] = reg.Replace(str, "ая");//str + "ая";
                            }

                            if (i < dtProduct.Rows.Count - 1) i++;
                            else break;
                        }
                        isChange = true;
                        #endregion
                    }
                    i = j;
                    if (j < dtProduct.Rows.Count) i = j + 1;
                    if (row["Название"].ToString() != "" && j < dtProduct.Rows.Count - 2) //проверить и заполнить тип из строки-подзаголовка
                    {
                        #region Название
                        subRow = dtProduct.Rows[i];
                        while (dtProduct.Rows[i]["Название"].ToString() == "")
                        {
                            diam = dtProduct.Rows[i]["Диаметр (высота), мм"].ToString();
                            tol = dtProduct.Rows[i]["Толщина (ширина), мм"].ToString();
                            met = dtProduct.Rows[i]["Метраж, м (длина, мм)"].ToString();
                            price = dtProduct.Rows[i]["Цена"].ToString();

                            if (diam == "" && tol == "" && met == "" && price == "")
                            {
                                break;
                            }
                            dtProduct.Rows[i]["Название"] = row["Название"];
                            if (dtProduct.Rows[i]["Название"].ToString().ToLower() == "трубы") dtProduct.Rows[i]["Название"] = "Труба";
                            if (i < dtProduct.Rows.Count - 1) i++;
                            else break;
                        }
                        isChange = true;
                        #endregion
                    }

                    if (row["Название"].ToString().Trim() != "" && j < dtProduct.Rows.Count - 2) //проверить и заполнить название из строки-подзаголовка
                    {
                        #region название
                        i = j + 1;
                        subRow = dtProduct.Rows[i];

                        while (subRow["Название"].ToString().Trim() == "")
                        {

                            diam = subRow["Диаметр (высота), мм"].ToString();
                            tol = subRow["Толщина (ширина), мм"].ToString();
                            met = subRow["Метраж, м (длина, мм)"].ToString();
                            price = subRow["Цена"].ToString();
                            if (diam == new Regex("\\d+", RegexOptions.IgnoreCase).Match(row["Стандарт"].ToString().Trim()).Value)
                            {
                                row["Диаметр (высота), мм"] = "";
                            }
                            if (diam == "" && tol == "" && met == "" && price == "")
                            {
                                break;
                            }
                            subRow["Название"] = row["Название"];
                            if (row["Название"].ToString().ToLower() == "угол") row["Название"] = "Уголок";

                            subRow = dtProduct.Rows[i];
                            if (subRow["Название"].ToString() == "обсадная")
                            {
                                subRow["Тип"] = subRow["Название"];
                                subRow["Название"] = "";
                            }
                            if (i < dtProduct.Rows.Count - 1) i++;
                            else break;
                        }
                        isChange = true;
                        #endregion
                    }
                    if (row["Стандарт"].ToString().Trim() != "" && j < dtProduct.Rows.Count - 2) //проверить и заполнить название из строки-подзаголовка
                    {
                        #region стандарт
                        i = j + 1;
                        subRow = dtProduct.Rows[i];

                        while (subRow["Стандарт"].ToString().Trim() == "")
                        {

                            diam = subRow["Диаметр (высота), мм"].ToString();
                            tol = subRow["Толщина (ширина), мм"].ToString();
                            met = subRow["Метраж, м (длина, мм)"].ToString();
                            price = subRow["Цена"].ToString();

                            if (diam == "" && tol == "" && met == "" && price == "")
                            {
                                break;
                            }

                            subRow["Стандарт"] = row["Стандарт"];
                            i++;
                            if (i < dtProduct.Rows.Count) subRow = dtProduct.Rows[i];
                            //if (subRow["Стандарт"].ToString() == "обсадная")
                            //{
                            //    subRow["Тип"] = subRow["Название"];
                            //    subRow["Название"] = "";
                            //}
                        }
                        isChange = true;
                        #endregion
                    }

                    if (isChange)
                    {
                        if (i > 2)
                            j = i - 2;
                        else if (i == 1) j = i - 1;
                        else j = i;
                    }

                    j--;
                    row.Delete();
                    isChange = false;
                }
                if (tsPb1.Value < tsPb1.Maximum)
                    tsPb1.Value++;


            }
            tsPb1.Value = 0;
            for (int j = 0; j < dtProduct.Rows.Count; j++)
            {
                row = dtProduct.Rows[j];

                if (row["Название"].ToString().Trim() == "Швеллер")
                {
                    if (row["Диаметр (высота), мм"].ToString().Trim() == "")
                    {
                        Regex r = new Regex(@"\w+\s+\d{1,2}(?:[,\.]\d+)?\w", RegexOptions.IgnoreCase);
                        if (r.IsMatch(row["Примечание"].ToString()))
                        {
                            row["Диаметр (высота), мм"] = new Regex(@"(?<=\s)\d{1,2}(?:[,\.]\d+)?(?=\w)", RegexOptions.IgnoreCase).Match(row["Примечание"].ToString()).Value;
                            row["Тип"] = new Regex(@"(?<=\d{1,2}(?:[,\.]\d+)?\s*)\w(?=\s|$)", RegexOptions.IgnoreCase).Match(row["Примечание"].ToString()).Value;
                        }
                    }
                }
                if (row["Название"].ToString().ToLower() == "угол") row["Название"] = "Уголок";
                Regex rTol = new Regex(@"(?<=\s)\d+\*\d+\*\d+(?=\s)", RegexOptions.IgnoreCase);
                if (row["Название"].ToString().ToLower() == "уголок" && rTol.IsMatch(row["Примечание"].ToString()))
                {
                    row["Толщина (ширина), мм"] = new Regex(@"(?<=\*)\d+(?=\s)", RegexOptions.IgnoreCase).Match(row["Примечание"].ToString());
                }
                if (row["Название"].ToString().ToLower() == "прокат" && row["Тип"].ToString().ToLower() != "листовой") row["Название"] = "Труба";
                if (row["Тип"].ToString().ToLower() == "б/ш") row["Тип"] = "бесшовная";
                if (dtProduct.Rows[j]["Название"].ToString().Length > 1)
                    dtProduct.Rows[j]["Название"] = row["Название"].ToString().Substring(0, 1).ToUpper() + row["Название"].ToString().Substring(1, row["Название"].ToString().Length - 1).ToLower();
                else if (textBoxOrgName.Text == "ТО Терминал") dtProduct.Rows[j]["Название"] = "Труба";
                if (dtProduct.Rows[j]["Название"].ToString().ToLower() == "вгп") { dtProduct.Rows[j]["Название"] = "Труба"; dtProduct.Rows[j]["Тип"] += " ВГП"; }
                dtProduct.Rows[j]["Тип"] = row["Тип"].ToString().ToLower();
                if (tsPb1.Value < tsPb1.Maximum)
                    tsPb1.Value++;
            }
            tsPb1.Value = tsPb1.Maximum;
            tsLabelClearingTable.Text = "";
        }

        //функция ручного добавления названия продукции
        private void manualNameProd(Excel.Worksheet excelworksheet, int startRow, int endRow, int startCol, int endCol, string temp)
        {
            conn = new SqlConnection(sqlConnString);
            if (conn.State == ConnectionState.Closed) conn.Open();
            SqlDataAdapter sda = new SqlDataAdapter("select NameProd from dbo.ManualNameProd where NameOrg = '" + orgname + "'", conn);
            DataTable dt = new DataTable();
            sda.Fill(dt);
            if (dt.Rows.Count > 0)
                nameProd = dt.Rows[0][0].ToString();
            else
            {
                AddNameProduct anp = new AddNameProduct(new MyDelegate(getNameProduct));
                anp.Tag = Path.GetFileName(listView1.SelectedItems[0].SubItems[1].Text);
                anp.ShowDialog();
            }
            if (conn.State == ConnectionState.Open) conn.Close();
            for (int j = startRow + 1; j <= endRow; j++) //строки
            {

                Excel.Range cellRange = (Excel.Range)excelworksheet.Cells[j, startCol];
                if (cellRange.Value != null)
                    temp = cellRange.Value.ToString();
                else temp = "";
                if (temp != "")
                {

                    string tmp = new Regex(@"\d+|\d\s?[xх]\s?\d").Match(temp).Value;
                    if (tmp != "")
                    {
                        dtProduct.Rows.Add();   // добавить строку в результирующую таблицу
                        int lastRow = dtProduct.Rows.Count - 1; //запомнить индекс последней строки
                        dtProduct.Rows[lastRow]["Название"] = nameProd; /*записать наименование вручную указанного названия 
                                                                         * в ячейку названия продукции в результирующей таблице*/
                        listIndexOfNotEmptyName.Add(j);     //добавить в список индексов непустых значений индекс текущей строки
                        int c = listIndexOfNotEmptyName[0] + countEmpty;    //запомнить в переменную индекс первого значения плюс количество пустых ячеек
                        //если список непустых значений не пустой
                        if (listIndexOfNotEmptyName.Count > 0)
                        {
                            //то если список сдвигов индексов пустой
                            if (listShiftIndex.Count < 1)
                                // то занести в список сдвигов индексов индекс певого значения плюс количество пустых ячеек
                                listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                            /*если список сдвигов содержит больше 2х записей и индекс текущей строки меньше чем 
                              индекс строки в списке непустых значений в предпоследней записи */
                            else if (listIndexOfNotEmptyName.Count > 2 && j < listIndexOfNotEmptyName[listIndexOfNotEmptyName.Count - 2])
                            {
                                countEmpty = 0; //сброс счета количества пустых ячеек
                                //занести в список сдвигов индексов индекс певого значения плюс количество пустых ячеек
                                listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                                //количество строк для сдвига равно текущему количеству строк в результирующей таблице
                                countRowsForShift = dtProduct.Rows.Count - 1;
                            }
                            else listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                        }
                    }
                    else if (listIndexOfNotEmptyName.Count > 0) countEmpty++;
                }
                else if (listIndexOfNotEmptyName.Count > 0) countEmpty++;
            }
            ManualStringNameProd = nameProd;

        }

        private string NameProdUpLower(string nameProdUpLower)
        {
            nameProdUpLower = nameProdUpLower.Substring(0, 1).ToUpper() + nameProdUpLower.Substring(1, nameProdUpLower.Length - 1).ToLower();
            return nameProdUpLower;
        }

        //функции поиска в строке различных шаблонов и их обработки
        #region Regex'ы

        /// <summary>
        /// Функия поиска телефона организации, адреса сайта, Email'а, инн, кпп, рс, кс, бик, телефонов менеджеров их имен в заданной строке
        /// </summary>
        /// <param name="str">Строка для поиска</param>
        /// <param name="indexOfRow">Номер изучаемой строки</param>
        private void InfoOrganization(string temp)
        {
            temp = temp.Trim();

            #region телефон организации
            if (new Regex(@"[Тт]елефон|^\+?\d\(\d{3}\)\d{3}-\d{2}-\d{2}|^[тТ]ел.").IsMatch(temp)) //поиск телефона организации
            {
                if (!isTelefon)
                    if (textBoxOrgTelefon.Text == "")
                    {
                        textBoxOrgTelefon.Text = new Regex(@"\b\d[\s\d(),-]*(?=\s|$)|\d?\(\d{3}\)\s?\d{3}-\d{2}-\d{2}").Match(temp).Value;
                        isTelefon = true;
                    }
                    else textBoxOrgTelefon.Text += "; " + new Regex(@"\b\d[\s\d(),-]*(?=\s|$)|\d?\(\d{3}\)\s?\d{3}-\d{2}-\d{2}").Match(temp).Value;
                //break;
            }
            else if (new Regex(@"тел\.?\s*\(\d{3,5}\)\s*\d{2,3}-\d{2,3}(?:-\d{2,3})?", RegexOptions.IgnoreCase).IsMatch(temp))
            {

                if (textBoxOrgTelefon.Text == "")
                {
                    textBoxOrgTelefon.Text = new Regex(@"тел\.?\s*\(\d{3,5}\)\s*\d{2,3}-\d{2,3}(?:-\d{2,3})?").Match(temp).Value;
                    isTelefon = true;
                }
                else textBoxOrgTelefon.Text += "; " + new Regex(@"тел\.?\s*\(\d{3,5}\)\s*\d{2,3}-\d{2,3}(?:-\d{2,3})?").Match(temp).Value;

            }
            else if (new Regex(@"тел\./факс:", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                Regex r = new Regex(@"(?<=тел\./факс:\s+(?:\(\d{3}\))?\s*;\s+(?:\d{3}-\d{2}-\d{2})?(?:,\s+)?)\d{3}-\d{2}-\d{2}", RegexOptions.IgnoreCase);
                foreach (Match m in r.Matches(temp))
                    if (textBoxOrgTelefon.Text == "")
                    {
                        textBoxOrgTelefon.Text = m.Value;
                        isTelefon = true;
                    }
                    else textBoxOrgTelefon.Text += "; " + m.Value;
            }
            if (new Regex(@"(?<=\s*)\+?\d-\d{3}-\d{3}-\d{2}-\d{2}(?=\s*icq)").IsMatch(temp))
            {
                if (textBoxOrgTelefon.Text == "")
                {
                    textBoxOrgTelefon.Text = new Regex(@"(?<=\s*)\+?\d-\d{3}-\d{3}-\d{2}-\d{2}(?=\s*icq)").Match(temp).Value;
                    isTelefon = true;
                }
                else textBoxOrgTelefon.Text += "; " + new Regex(@"(?<=\s*)\+?\d-\d{3}-\d{3}-\d{2}-\d{2}(?=\s*icq)").Match(temp).Value;
            }
            if (new Regex(@"^\+?\d-\d{3}-\d{3}-\d{2}-\d{2}\s*\[[\w\s]*\]").IsMatch(temp)) //поиск телефона организации
            {

                if (!isTelefon)
                    if (textBoxOrgTelefon.Text == "")
                    {
                        textBoxOrgTelefon.Text = new Regex(@"^\+?\d-\d{3}-\d{3}-\d{2}-\d{2}\s*\[[\w\s]*\]").Match(temp).Value;
                        isTelefon = true;
                    }
                    else textBoxOrgTelefon.Text += "; " + new Regex(@"^\+?\d-\d{3}-\d{3}-\d{2}-\d{2}\s*\[[\w\s]*\]").Match(temp).Value;
                //break;

            }
            if (new Regex(@"\d(?:-\d{3})+(?:-\d{2})+").IsMatch(temp)) //поиск телефона организации
            {

                if (!isTelefon)
                    if (textBoxOrgTelefon.Text == "")
                    {
                        textBoxOrgTelefon.Text = new Regex(@"\d(?:-\d{3})+(?:-\d{2})+").Match(temp).Value;
                        isTelefon = true;
                    }
                    else textBoxOrgTelefon.Text += "; " + new Regex(@"\d(?:-\d{3})+(?:-\d{2})+").Match(temp).Value;
                //break;

            }
            #endregion

            #region сайт
            if (new Regex(@"(?:www\.)[\w\d-]{2,}\.[A-Za-zА-Яа-я]+").IsMatch(temp)) // поиск сайта
            {
                textBoxOrgSite.Text = new Regex(@"(?:www\.)[\w\d-]{2,}\.[A-Za-zА-Яа-я]+").Match(temp).Value;
                //break;
            }
            #endregion

            #region Email
            Regex regEmail = new Regex(@"[\d\.\w-\*\\]+@[\w-]+\.\w{1,5}");
            if (regEmail.IsMatch(temp)) // поиск Email
            {
                if (textBoxOrgEmail.Text == "") textBoxOrgEmail.Text = regEmail.Match(temp).Value;
                else textBoxOrgEmail.Text += "; " + regEmail.Match(temp).Value;
                //break;
            }
            #endregion

            #region инн/кпп
            if (new Regex(@"[иИ][Нн]{2}\s*/\s*[Кк][Пп]{2}").IsMatch(temp)) // поиск инн/кпп
            {
                textBoxOrgINN.Text = new Regex(@"(?<=ИНН/[\s\w\d:]*)\d+(?:/\d+)?").Match(temp).Value;
                //break;
            }
            else
            {
                var regex = new Regex(@"(?<=ИНН\s*)\d+", RegexOptions.IgnoreCase);
                if (regex.IsMatch(temp))
                {
                    var tmp = regex.Match(temp).Value;
                    textBoxOrgINN.Text = tmp;
                    regex = new Regex(@"(?<=кпп\s*)\d+", RegexOptions.IgnoreCase);
                    if (regex.IsMatch(temp)) { textBoxOrgINN.Text += " / " + regex.Match(temp).Value; }
                }
            }
            #endregion

            #region менеджеры
            if (new Regex(@"(?!icq)(?:моб\w*[\.\s]?(?:тел\w*[:\s]?)?\s?[\d-]+\s+\w+\b)|(?:\d-\d+-\d+-\d+-\d+\s+\(?\w+\)?|(?<=\s\s+)\d-\d{10}\s+\w+)").IsMatch(temp)) // имя менеджера и его телефон
            {
                if (new Regex(@"^\d-\d+-\d+-\d+-\d+\s+\(?\w+\)?").IsMatch(temp)) //для стальной мир
                {
                    ListViewItem lvi = new ListViewItem(new Regex(@"(?<=\s+\(?)\w+(?=\)?)").Match(temp).Value); //имя менеджера
                    lvi.SubItems.Add(new Regex(@"^\d-\d+-\d+-\d+-\d+(?=\s+)").Match(temp).Value);           //телефон в формате 8-888-888-88-88
                    if (lvi.SubItems[0].Text != "icq") listViewManager.Items.Add(lvi);
                }
                else if (new Regex(@"(?<=\s\s+)\d-\d{10}\s+\w+").IsMatch(temp)) //для Вектор
                {
                    ListViewItem lvi = new ListViewItem(new Regex(@"(?<=\d{10}\s+\(?)\w+(?=\)?)").Match(temp).Value);   //имя менеджера
                    lvi.SubItems.Add(new Regex(@"(?<=\s\s+)\d-\d{10}(?=\s+\w+|$)").Match(temp).Value);      //телефон в формате 8-8888888888
                    if (lvi.SubItems[0].Text != "icq") listViewManager.Items.Add(lvi);
                }
                else
                {
                    ListViewItem lvi = new ListViewItem(new Regex(@"(?<=\d+\s)\w+\b").Match(temp).Value);
                    lvi.SubItems.Add(new Regex(@"\s[\d+-]+\s").Match(temp).Value);
                    if (lvi.SubItems[0].Text != "icq") listViewManager.Items.Add(lvi);
                }
            }

            else if (new Regex(@"(?<=менеджер\s*:?\s+)[\w\s]+(?=\s+\d|\s\+)", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                ListViewItem lvi = new ListViewItem(new Regex(@"(?<=[Мм]енеджер\s*:?\s*)\w+(?=\s\d)|(?<=[мМ]енеджер\s*:?\s*)(?:\w+\s+){1,3}(?=\+)", RegexOptions.IgnoreCase).Match(temp).Value); //телефон в формате 8-888-888-88-88
                lvi.SubItems.Add(new Regex(@"(?<=[Мм]енеджер\s*:?\s+[\w\s]+)(?:\d+|\+\d\s\(\s?\d{1,3}\s?\)\s?\d+-?\d+-?\d+,?\s*(?:доб[\.\w+]+\s+\d+)(?:,?\s+\+\d\s?\(?\d{1,3}\)\s?\d+-?\d+-?\d+)*)", RegexOptions.IgnoreCase).Match(temp).Value);
                listViewManager.Items.Add(lvi);
            }

            else if (new Regex(@"(?<=менеджер\s+).+тел.+моб.+\d\d(?=\.\s+)", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                ListViewItem lvi = new ListViewItem(new Regex(@"(?<=менеджер\s+)[\w\s]+(?=,\s?тел.+моб.+\d\d\.\s+)", RegexOptions.IgnoreCase).Match(temp).Value);
                lvi.SubItems.Add(new Regex(@"(?<=менеджер\s+.+,\s?)тел.+моб.+\d\d(?=\.\s+)", RegexOptions.IgnoreCase).Match(temp).Value);
                bool isInList = false;
                for (int i = 0; i < listViewManager.Items.Count; i++)
                {
                    if (lvi.SubItems[0].Text == listViewManager.Items[i].SubItems[0].Text) isInList = true;
                }
                if (!isInList) listViewManager.Items.Add(lvi);
            }
            #endregion

            #region р/с
            var reg = new Regex(@"(?<=[рР][\\/](?:с|счет)\s*)\d+", RegexOptions.IgnoreCase);// поиск рас/счет
            if (reg.IsMatch(temp)) // поиск рас/счет
            {
                textBoxOrgRS.Text = reg.Match(temp).Value;
                //break;
            }
            #endregion

            #region к/с
            reg = new Regex(@"(?<=(?:[Кк]|корр?)[\\/](?:с|счет)\s*)\d+", RegexOptions.IgnoreCase);
            if (reg.IsMatch(temp)) // поиск кор/счет
            {
                textBoxOrgKS.Text = reg.Match(temp).Value;
                //break;
            }
            #endregion

            #region бик
            reg = new Regex(@"(?<=бик\s*)\d+", RegexOptions.IgnoreCase);
            if (reg.IsMatch(temp)) // поиск бик
            {
                textBoxBIK.Text = reg.Match(temp).Value;
                //break;
            }
            #endregion
        }

        /// <summary>
        /// Функия поиска номера телефона организации в заданной строке
        /// </summary>
        /// <param name="str">Строка для поиска</param>
        /// <param name="indexOfRow">Номер изучаемой строки</param>
        private void GetTelefonFromString(string str, int indexOfRow)
        {
            try
            {
                Regex regDiam = new Regex(@"(?<=^|\s+\b)\d+[,.]?\d*(?=(?:[xXхХ\*])|(?:\b\s)|(?:$))|(?<=\b\.)\d+(?=[xXхХ\*])|(?<=ф\s?)\d+|(?<=\s)\.\d{1,3}|^\d+(?=[xXхХ]\d+\s+)"); //шаблон диаметра
                if (regDiam.IsMatch(str))
                    dtProduct.Rows[indexOfRow]["Диаметр (высота), мм"] = regDiam.Match(str).Value;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка №1035\n" + ex.ToString()); }
        }

        /// <summary>
        /// Функия поиска диаметра в заданной строке
        /// </summary>
        /// <param name="str">Строка для поиска</param>
        /// <param name="indexOfRow">Номер изучаемой строки</param>
        private void GetRegexDiamFromString(string str, int indexOfRow)
        {
            try
            {
                Regex regDiam = new Regex(@"(?<=^|\s+\b)\d+[,.]?\d*(?=(?:[xXхХ\*])|(?:\b\s)|(?:$))|(?<=\b\.)\d+(?=[xXхХ\*])|(?<=ф\s?)\d+|(?<=\s)\.\d{1,3}|^\d+(?=[xXхХ]\d+\s+)"); //шаблон диаметра
                if (regDiam.IsMatch(str))
                    dtProduct.Rows[indexOfRow]["Диаметр (высота), мм"] = regDiam.Match(str).Value;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка №1021\n" + ex.ToString()); }
        }

        /// <summary>
        /// Функия поиска Наименования товара в заданной строке
        /// </summary>
        /// <param name="str">Строка для поиска</param>
        /// <param name="indexOfRow">Номер изучаемой строки</param>
        private void GetRegexNaimenovanieFromString(string str, int indexOfRow)
        {
            try
            {
                Regex regName = new Regex(@""); //шаблон диаметра
                if (regName.IsMatch(str))
                    dtProduct.Rows[indexOfRow]["Диаметр (высота), мм"] = regName.Match(str).Value;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка №1021\n" + ex.ToString()); }
        }

        /// <summary>
        /// Функия поиска типа продукции в заданной строке
        /// </summary>
        /// <param name="str">Строка для поиска</param>
        /// <param name="indexOfRow">Номер изучаемой строки</param>
        private void GetRegexTypeFromString(string str, int indexOfRow)
        {
            try
            {
                Regex regType = new Regex(@".+ая|.+ое|.+ый"); //шаблон диаметра
                if (regType.IsMatch(str))
                    dtProduct.Rows[indexOfRow]["Тип"] = regType.Match(str).Value;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка №1022\n" + ex.ToString()); }
        }

        /// <summary>
        /// Функия поиска толщины стенки в заданной строке
        /// </summary>
        /// <param name="str">Строка для поиска</param>
        /// <param name="indexOfRow">Номер изучаемой строки</param>
        private void GetRegexTolshFromString(string str, int indexOfRow)
        {
            try
            {
                Regex regTolsh = new Regex(@"(?<=[хХxX]\.?)\d{1,2}[,.]*\d*\b|(?<=[хХxX\*])\d{1,2}[,.]*\d+\b"); //шаблон толщины стенки
                if (regTolsh.IsMatch(str))
                    dtProduct.Rows[indexOfRow]["Толщина (ширина), мм"] = regTolsh.Match(str).Value;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка №1023\n" + ex.ToString()); }
        }

        /// <summary>
        /// Функия поиска толщины стенки в заданной строке, если там есть лишние пробелы
        /// </summary>
        /// <param name="str">Строка для поиска</param>
        /// <param name="indexOfRow">Номер изучаемой строки</param>
        private void GetRegexTolshFromStrWithSpases(string str, int indexOfRow)
        {
            try
            {
                Regex regTolsh = new Regex(@"(?<=[хХxX]\.?\s?)\d{1,2}[,.]*\d*\b"); //шаблон толщины стенки
                if (regTolsh.IsMatch(str))
                    dtProduct.Rows[indexOfRow]["Толщина (ширина), мм"] = regTolsh.Match(str).Value;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка №1023\n" + ex.ToString()); }
        }

        /// <summary>
        /// Функия поиска марки продукции в заданной строке
        /// </summary>
        /// <param name="str">Строка для поиска</param>
        /// <param name="indexOfRow">Номер изучаемой строки</param>
        private void GetRegexMarkFromString(string str, int indexOfRow)
        {
            try
            {
                Regex regMark = new Regex(@"(?:\d{,3}[ШСТУ]+\d{,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+\d{,3}[ХНКМВТДГСФРАБЕЦЮЧПС]+\d{,3})(?=\s+|$)|(?:(?:Ст.)|(?:ст.)(?:\s*\d{1,2})\b)(?=\s+|$)|(?:\d{,3}[ХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)(?=\s+|$)|(?:[АA]-?\d)(?=\s+|$)|(?<=[Сс][Тт]\.\s?)\d{1,2}[гГ]\d{1,2}[cCсС]|(?<=ст\.)\d{1,2}[хфа]+(?=\s|$)|(?<=\s)[сС][тТ]\.?\s?\d{1,2}[_\w]+"); //шаблон марки стали
                if (regMark.IsMatch(str))
                    if (dtProduct.Rows[indexOfRow]["Марка"].ToString() == "") dtProduct.Rows[indexOfRow]["Марка"] = regMark.Match(str).Value;
                    else dtProduct.Rows[indexOfRow]["Марка"] += regMark.Match(str).Value;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка №1024\n" + ex.ToString()); }
        }

        /// <summary>
        /// Функия поиска стандарта в заданной строке
        /// </summary>
        /// <param name="str">Строка для поиска</param>
        /// <param name="indexOfRow">Номер изучаемой строки</param>
        private void GetRegexTUFromString(string str, int indexOfRow)
        {
            try
            {
                str = str.Trim();
                Regex regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|ТУ\s*\d+(?:\s|$)|(?:Г[Оо][Сс][Тт]\s{0,3})(?:[рР]-\s?)?(?:\d{1,5}[-\s]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])"); //шаблон Стандарта
                if (regTU.IsMatch(str))
                {
                    if (dtProduct.Rows[indexOfRow]["Стандарт"] == null) dtProduct.Rows[indexOfRow]["Стандарт"] = regTU.Match(str).Value;
                    else dtProduct.Rows[indexOfRow]["Стандарт"] = regTU.Match(str).Value;
                    isGost = true;
                }
            }
            catch (Exception ex) { MessageBox.Show("Ошибка №1025\n" + ex.ToString()); }
        }

        /// <summary>
        /// Функия поиска Цены в заданной строке
        /// </summary>
        /// <param name="str">Строка для поиска</param>
        /// <param name="indexOfRow">Номер изучаемой строки</param>
        private void GetRegexPriceFromString(string str, int indexOfRow)
        {
            try
            {
                Regex regPrice = new Regex(@"\d+[,.\s]*\d*"); //шаблон цены
                if (regPrice.IsMatch(str))
                    dtProduct.Rows[indexOfRow]["Цена"] = regPrice.Match(str).Value;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка №1026\n" + ex.ToString()); }
        }

        /// <summary>
        /// Функия поиска Длины в заданной строке
        /// </summary>
        /// <param name="str">Строка для поиска</param>
        /// <param name="indexOfRow">Номер изучаемой строки</param>
        private void GetRegexDlinaFromString(string str, int indexOfRow)
        {
            try
            {
                Regex regDlina = new Regex(@"\d+[,.]*\d*"); //шаблон длины
                if (regDlina.IsMatch(str))
                    dtProduct.Rows[indexOfRow]["Метраж, м (длина, мм)"] = regDlina.Match(str).Value;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка №1027\n" + ex.ToString()); }
        }

        /// <summary>
        /// Функия поиска Весов продукции в заданной строке
        /// </summary>
        /// <param name="str">Строка для поиска</param>
        /// <param name="indexOfRow">Номер изучаемой строки</param>
        private void GetRegexVesFromString(string str, int indexOfRow)
        {
            try
            {
                string temp = "";
                Regex regVes = new Regex(@"\d+[,.]*\d*");  //шаблон Мерность (т, м, мм)а
                if (regVes.IsMatch(str))
                {
                    temp = regVes.Match(str).Value;
                    dtProduct.Rows[indexOfRow]["Мерность (т, м, мм)"] = temp;
                }
                //}
                //catch { MessageBox.Show("Ошибка изъятия регулярного выражения из столбца Мерность (т, м, мм)ов"); }
            }
            catch (Exception ex) { MessageBox.Show("Ошибка №1028\n" + ex.ToString()); }
        }

        /// <summary>
        /// Функия разбора по-умолчанию в заданной строке
        /// </summary>
        /// <param name="str">Строка для поиска</param>
        /// <param name="indexOfRow">Номер изучаемой строки</param>
        private void GetRegexNameFromString(string str, int currentRow)
        {
            try
            {
                if (!(new Regex(@"(?:^[тТ][еЕ][лЛ])|(?:^[мМ][оО][бБ])|(?:^[Ii][Cc][Qq])").IsMatch(str.Trim())))
                {
                    dtProduct.Rows.Add();
                    int lastRow = dtProduct.Rows.Count - 1;
                    string temp = "";
                    str = str.Trim();
                    Regex regName = new Regex("");
                    try { regName = new Regex(@"^\w+\s*\D+"); }  //поиск шаблона названия в строке
                    catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                    temp = regName.Match(str).Value;
                    if (temp != "")
                    {
                        temp = temp.Trim();
                        dtProduct.Rows[lastRow]["Название"] = new Regex(@"(?!\w+\d\w+)^\w{3,}").Match(temp).Value; //выделение названия
                        if (new Regex(@"(?<=\s)[^ф\s][\w\\/]+").IsMatch(temp))
                            dtProduct.Rows[lastRow]["Тип"] = new Regex(@"(?<=\s)[^ф\s][\w\\/]+").Match(temp).Value; //выделение типа
                        else if (new Regex(@"б/ш", RegexOptions.IgnoreCase).IsMatch(temp))
                            dtProduct.Rows[lastRow]["Тип"] += new Regex(@"б/ш", RegexOptions.IgnoreCase).Match(temp).Value;
                        else dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                        listIndexOfNotEmptyName.Add(currentRow);
                        int c = listIndexOfNotEmptyName[0] + countEmpty;
                        if (listIndexOfNotEmptyName.Count > 0)
                        {
                            if (listShiftIndex.Count < 1)
                                listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                            else if (listIndexOfNotEmptyName.Count > 2 && currentRow < listIndexOfNotEmptyName[listIndexOfNotEmptyName.Count - 2]
                                && true)
                            {
                                countEmpty = 0;
                                listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                                countRowsForShift = dtProduct.Rows.Count - 1;
                            }
                            else listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                        }
                    }


                    temp = "";
                    Regex regDiam = new Regex("");
                    try { regDiam = new Regex(@"(?<=\s+\b)\d+[,.]?\d*(?=(?:[xXхХ\*])|(?:\b\s)|(?:$))|(?<=\b\.)\d+(?=[xXхХ\*])|(?<=ф\s?)\d+|(?<=\s)\.\d{1,3}|^\d+(?=[xXхХ]\d+\s+)"); } //
                    catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                    temp = regDiam.Match(str).Groups[0].Value;
                    if (temp != "") { dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = temp; temp = ""; }

                    Regex regTolsh = new Regex("");
                    try { regTolsh = new Regex(@"(?<=[хХxX]\.?)\d{1,2}[,.]*\d*\b|(?<=[хХxX\*])\d{1,2}[,.]*\d+\b"); }
                    catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                    temp = regTolsh.Match(str).Groups[0].Value;
                    if (temp != "") { dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = temp; temp = ""; }

                    Regex regMark = new Regex("");
                    try { regMark = new Regex(@"(?:[\dШСТУ]+[ХНКМВТДГСФРАБЕЦЮЧПС]+[\dХНКМВТДГСФРАБЕЦЮЧПС]+)(?=\s+|$)|(?:(?:Ст.)|(?:ст.)(?:\s*\d{1,2})\b)(?=\s+|$)|(?<=^|\s)(?:\d+[ХХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)(?=\s+|$)|(?:[АA]-?\d{1,3}\w+)(?=\s+|$)|(?<=[Сс][Тт]\.\s?)\d{1,2}[гГ]\d{1,2}[cCсС]|(?<=ст\.)\d{1,2}[хфа]+(?=\s|$)|(?<=\s)[сС][тТ]\.?\s?\d{1,2}[_\w]+|А-?III|[АA]-?\|\|\|"); }
                    catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                    temp = regMark.Match(str).Value;
                    if (temp != "") { dtProduct.Rows[lastRow]["Марка"] = temp; temp = ""; isMark = true; }

                    Regex regTU = new Regex("");
                    try { regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|(?:ГОСТ\s{0,3})(?:[рР]-\s?)?(?:\d{1,5}[-\s]*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])"); }
                    catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                    foreach (Match m in regTU.Matches(str))
                    {
                        temp = m.Value;
                        if (temp != "")
                        {
                            if (dtProduct.Rows[lastRow]["Стандарт"].ToString().Trim() == "") dtProduct.Rows[lastRow]["Стандарт"] = temp;
                            else dtProduct.Rows[lastRow]["Стандарт"] += "; " + temp;
                            temp = ""; isGost = true;
                        }
                    }

                    dtProduct.Rows[lastRow]["Примечание"] = str;

                    //if (empty) dtProduct.Rows[lastRow][0] = "~~~";
                }
            }
            catch (Exception ex) { MessageBox.Show("Ошибка №1029\n" + ex.ToString()); }
        }

        /// <summary>
        /// Функия разбора по-умолчанию в заданной строке, если заголовок содержит "номенклатуру"
        /// </summary>
        /// <param name="str">Строка для поиска</param>
        /// <param name="indexOfRow">Номер изучаемой строки</param>
        private void GetRegexNameNomeklFromString(string str, int currentRow, bool isShift)//для Метал.Снаб.Урал
        {
            try
            {
                if (!(new Regex(@"(?:^[тТ][еЕ][лЛ])|(?:^[мМ][оО][бБ])|(?:^[Ii][Cc][Qq])").IsMatch(str.Trim())))
                {
                    bool isList = false; //для наименования "Лист" нужно выбирать ширину длину и высоту особым образом
                    bool isGK = false; //где находится г/к, рядом с названием или дальше в строке. False - дальше в строке
                    dtProduct.Rows.Add();
                    int lastRow = dtProduct.Rows.Count - 1;
                    string temp = "";

                    Regex regName = new Regex("");
                    try { regName = new Regex(@"^(\w*\s*\D+)"); }  //поиск шаблона названия в строке
                    catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                    temp = regName.Match(str).Value;
                    if (temp != "")
                    {

                        dtProduct.Rows[lastRow]["Название"] = new Regex(@"^\w+").Match(temp).Value; //выделение названия
                        if (new Regex("лист", RegexOptions.IgnoreCase).IsMatch(dtProduct.Rows[lastRow]["Название"].ToString()))
                        {
                            isList = true;
                        }
                        if (new Regex(@"г[\\/]к", RegexOptions.IgnoreCase).IsMatch(str))
                        { dtProduct.Rows[lastRow]["Тип"] = new Regex(@"г[\\/]к").Match(str).Value; isGK = true; }
                        if (new Regex(@"(?<=\s)[^ф\s][\w\\/]+").IsMatch(temp))
                            dtProduct.Rows[lastRow]["Тип"] = new Regex(@"(?<=\s)[^ф\s][\w\\/]+").Match(temp).Value; //выделение типа
                        else dtProduct.Rows[lastRow]["Тип"] = "тип не указан";
                        listIndexOfNotEmptyName.Add(currentRow);
                        int c = listIndexOfNotEmptyName[0] + countEmpty;
                        if (listIndexOfNotEmptyName.Count > 0)
                        {
                            if (listShiftIndex.Count < 1)
                                listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                            else if (listIndexOfNotEmptyName.Count > 2 && currentRow < listIndexOfNotEmptyName[listIndexOfNotEmptyName.Count - 2]
                                && true)
                            {
                                countEmpty = 0;
                                listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                                countRowsForShift = dtProduct.Rows.Count - 1;
                            }
                            else listShiftIndex.Add(listIndexOfNotEmptyName[0] + countEmpty);
                        }
                    }


                    temp = "";
                    if (!isList)
                    {
                        Regex regDiam = new Regex("");
                        try { regDiam = new Regex(@"(?<=\b\s)\d{1,3}[,.]?\d*(?=(?:[xXхХ\*])|(?:\b\s)|(?:$))|(?<=\b\.)\d+(?=[xXхХ\*])|(?<=ф\s?)\d+"); } //
                        catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                        temp = regDiam.Match(str).Groups[0].Value;
                        if (temp != "") { dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = temp; temp = ""; }

                        Regex regTolsh = new Regex("");
                        try { regTolsh = new Regex(@"(?<=[хХxX\*])\d{1,2}[,.]*\d*\b"); }
                        catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                        temp = regTolsh.Match(str).Groups[0].Value;
                        if (temp != "") { dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = temp; temp = ""; }

                        Regex regMark = new Regex("");
                        try { regMark = new Regex(@"(?:[\dШСТУ]+[ХНКМВТДГСФРАБЕЦЮЧПС]+[\dХНКМВТДГСФРАБЕЦЮЧПС]+)|(?:(?:Ст.)|(?:ст.)(?:\s*\d{1,2})\b)|(?:\d+[ХХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)|(?:[АA]-?\d)"); }
                        catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                        temp = regMark.Match(str).Groups[0].Value;
                        if (temp != "") { dtProduct.Rows[lastRow]["Марка"] = temp; temp = ""; isMark = true; }

                        Regex regTU = new Regex("");
                        try { regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|(?:ГОСТ\s{0,3})(?:[рР]-)?(?:\d{1,5}-*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])"); }
                        catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                        temp = regTU.Match(str).Groups[0].Value;
                        if (temp != "") { dtProduct.Rows[lastRow]["Стандарт"] = temp; temp = ""; isGost = true; }

                        if (new Regex(@"г[\\/]к", RegexOptions.IgnoreCase).IsMatch(str) && !isGK)
                            dtProduct.Rows[lastRow]["Тип"] = new Regex(@"г[\\/]к").Match(str).Value;

                        dtProduct.Rows[lastRow]["Примечание"] = str;
                    }
                    else
                    {
                        Regex regList = new Regex(@"(?<=\s)\d+,?\d*(?=[\*xXхХ])", RegexOptions.IgnoreCase);
                        temp = regList.Match(str).Value;
                        if (regList.IsMatch(str))
                        { dtProduct.Rows[lastRow]["Диаметр (высота), мм"] = temp; temp = ""; }

                        regList = new Regex(@"(?<=[\*xXхХ])\d+,?\d*(?=[\*xXхХ])", RegexOptions.IgnoreCase);
                        temp = regList.Match(str).Value;
                        if (regList.IsMatch(str))
                        { dtProduct.Rows[lastRow]["Толщина (ширина), мм"] = temp; temp = ""; }

                        regList = new Regex(@"(?<=\d[\*xXхХ])\d+,?\d*(?=[\s$]+)", RegexOptions.IgnoreCase);
                        temp = regList.Match(str).Value;
                        if (regList.IsMatch(str))
                        { dtProduct.Rows[lastRow]["Метраж, м (длина, мм)"] = temp; temp = ""; }

                        Regex regMark = new Regex("");
                        try { regMark = new Regex(@"(?:[\dШСТУ]+[ХНКМВТДГСФРАБЕЦЮЧПС]+[\dХНКМВТДГСФРАБЕЦЮЧПС]+)|(?:(?:Ст.)|(?:ст.)(?:\s*\d{1,2})\b)|(?:\d+[ХХхXxНКМВТДГСФРАБЕЦЮЧПС]+\b)|(?:[АA]-?\d)"); }
                        catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                        temp = regMark.Match(str).Groups[0].Value;
                        if (temp != "") { dtProduct.Rows[lastRow]["Марка"] = temp; temp = ""; isMark = true; }

                        Regex regTU = new Regex("");
                        try { regTU = new Regex(@"(?:ТУ\s{0,3}\d+-[\d\w.]+-[\d.]+(?:-[\d.])*)|(?:ГОСТ\s{0,3})(?:[рР]-)?(?:\d{1,5}-*)*|(?:[Вв]торой\s+сорт)|(?:[Бб]/[Уу])"); }
                        catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                        temp = regTU.Match(str).Groups[0].Value;
                        if (temp != "") { dtProduct.Rows[lastRow]["Стандарт"] = temp; temp = ""; isGost = true; }

                        if (new Regex(@"г[\\/]к", RegexOptions.IgnoreCase).IsMatch(str) && !isGK)
                            dtProduct.Rows[lastRow]["Тип"] = new Regex(@"г[\\/]к").Match(str).Value;

                        dtProduct.Rows[lastRow]["Примечание"] = str;
                    }
                    //if (empty) dtProduct.Rows[lastRow][0] = "~~~";
                }
            }
            catch (Exception ex) { MessageBox.Show("Ошибка №1030\n" + ex.ToString()); }
        }

        #endregion

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                conn.Close();
            }
            catch { }
            if (isExcelOpen && excelappworkbook != null && excelapp != null)
            {
                try
                {
                    excelappworkbook.Close(false, Type.Missing, Type.Missing);
                    excelapp.Quit();
                }
                catch { isExcelOpen = false; }// просто игнорим это исключение, видимо книга уже как-то кем-то закрыта)

                isExcelOpen = false;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (fillComboBoxCity())
            {
                getFilesFromDocDirectory();
                comboBoxCity.SelectedIndex = 0;
            }

            AppDomain.CurrentDomain.ProcessExit += new EventHandler(ExitEvent);
            listView1.Columns[0].Width = listView1.Width - 4;

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
            dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGridView1.DataSource = dtProduct;
        }

        /// <summary>
        /// заполнение комбобокса городов из базы
        /// </summary>
        /// <returns>истино, если заполнение прошло успешно</returns>
        private bool fillComboBoxCity()
        {
            try
            {
                conn = new SqlConnection(sqlConnString);
                if (conn.State == ConnectionState.Closed) conn.Open();
                var sqlComm = new SqlCommand("select City from City", conn);
                SqlDataReader sdr = sqlComm.ExecuteReader();
                while (sdr.Read())
                {
                    comboBoxCity.Items.Add(sdr.GetString(0));
                }
                if (conn.State == ConnectionState.Open) conn.Close();
                return true;
            }
            catch { return false; }
        }

        /// <summary>
        /// Заполнение списка файлов из папки
        /// </summary>
        private void getFilesFromDocDirectory()
        {
            string md = Environment.GetFolderPath(Environment.SpecialFolder.Personal);//путь к Документам
            try
            {
                string textOfCity = "";
                if (comboBoxCity.SelectedIndex < 0) textOfCity = "Челябинск";
                else textOfCity = comboBoxCity.SelectedItem.ToString();
                if (Directory.Exists(md + "\\MetBaseFiles\\" + textOfCity) == false)
                {
                    Directory.CreateDirectory(md + "\\MetBaseFiles\\" + textOfCity);
                }
                string[] searchPatterns = "*.xls?|*.doc?|*.pdf".Split('|');
                List<string> files = new List<string>();
                foreach (string sp in searchPatterns)
                    files.AddRange(System.IO.Directory.GetFiles(md + "\\MetBaseFiles\\" + textOfCity, sp, SearchOption.TopDirectoryOnly));
                files.Sort();

                listView1.Items.Clear();
                ListViewItem lvi;
                foreach (string file in files)
                {
                    lvi = new ListViewItem(Path.GetFileName(file));
                    lvi.SubItems.Add(file);
                    listView1.Items.Add(lvi);
                }
                filesCountInDirectory = listView1.Items.Count;
            }
            catch (Exception ex) { MessageBox.Show("Ошибка №1031\n" + ex.ToString()); }
        }

        private void btn_AddBD_Click(object sender, EventArgs e)
        {
            conn = new SqlConnection(sqlConnString);
            if (textBoxOrgName.Text == "") MessageBox.Show("Необходимо заполнить название организации!");
            else
            {
                DataTable dtOrgname = new DataTable();
                if (conn.State == ConnectionState.Closed) conn.Open();

                //проверяем наличие организации в базе
                string query = @"select [datePriceList] from Organization where [Name]='" + textBoxOrgName.Text + "'";
                SqlDataAdapter adapter;
                adapter = new SqlDataAdapter(query, conn);
                adapter.Fill(dtOrgname);
                if (dtOrgname.Rows.Count > 0)
                {
                    string dt1 = dtOrgname.Rows[0]["datePriceList"].ToString().Trim();
                    string dt2 = dateTimePicker1.Value.Year.ToString() + "." +
                    dateTimePicker1.Value.Month.ToString() + "." + dateTimePicker1.Value.Day.ToString();
                    if (dt1 == dt2.Trim())
                    {
                        if (MessageBox.Show("Даты файлов совпадают\nБудут добавлены новые записи, которых еще нет в базе.", "Внимание!", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {    //
                            string name = "";
                            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                            {
                                #region добавить продукт
                                var sqlCmd = new SqlCommand("dbo.insProd", conn);
                                sqlCmd.CommandType = CommandType.StoredProcedure;
                                name = dataGridView1.Rows[i].Cells["Название"].Value.ToString();
                                if (new Regex(@"трубы", RegexOptions.IgnoreCase).IsMatch(name)) name = "Труба";
                                if (new Regex(@"листы", RegexOptions.IgnoreCase).IsMatch(name)) name = "Лист";
                                sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ name);

                                if (dataGridView1.Rows[i].Cells["Тип"].Value.ToString() != "")
                                {
                                    sqlCmd.Parameters.AddWithValue("@Type", /* Значение параметра */ dataGridView1.Rows[i].Cells["Тип"].Value);
                                }
                                else sqlCmd.Parameters.AddWithValue("@Type", /* Значение параметра */ "тип не указан");

                                string temporarary = dataGridView1.Rows[i].Cells["Диаметр (высота), мм"].Value.ToString();
                                if (temporarary != "")
                                {
                                    if (new Regex(@"^\.\d+(?:\s|$)", RegexOptions.IgnoreCase).IsMatch(temporarary)) temporarary = temporarary.Substring(1);
                                    else if (new Regex(@"^\.\d+[,\.]\d+(?:\s|$)", RegexOptions.IgnoreCase).IsMatch(temporarary)) temporarary = temporarary.Substring(1);
                                    temporarary = new Regex(@"\.").Replace(temporarary, @",");
                                    sqlCmd.Parameters.AddWithValue("@Diametr", /* Значение параметра */ Convert.ToDouble(temporarary));
                                }
                                else sqlCmd.Parameters.AddWithValue("@Diametr", /* Значение параметра */ 0);

                                temporarary = dataGridView1.Rows[i].Cells["Толщина (ширина), мм"].Value.ToString();
                                if (dataGridView1.Rows[i].Cells["Толщина (ширина), мм"].Value.ToString() != "")
                                {
                                    if (temporarary.IndexOf('-') > -1)
                                        temporarary = temporarary.Substring(temporarary.IndexOf('-') + 1);
                                    if (new Regex(@"^\.\d+(?:\s|$)", RegexOptions.IgnoreCase).IsMatch(temporarary)) temporarary = temporarary.Substring(1);
                                    else if (new Regex(@"^\.\d+[,\.]\d+(?:\s|$)", RegexOptions.IgnoreCase).IsMatch(temporarary)) temporarary = temporarary.Substring(1);
                                    temporarary = new Regex(@"\.").Replace(temporarary, @",");
                                    sqlCmd.Parameters.AddWithValue("@Tolshina", /* Значение параметра */ Convert.ToDouble(temporarary));
                                }
                                else sqlCmd.Parameters.AddWithValue("@Tolshina", /* Значение параметра */ 0);

                                if (dataGridView1.Rows[i].Cells["Класс"].Value.ToString() != "")
                                    sqlCmd.Parameters.AddWithValue("@Klass", /* Значение параметра */ dataGridView1.Rows[i].Cells["Класс"].Value);
                                else sqlCmd.Parameters.AddWithValue("@Klass", /* Значение параметра */ " ");

                                if (dataGridView1.Rows[i].Cells["Стандарт"].Value.ToString() != "")
                                    sqlCmd.Parameters.AddWithValue("@Standart", /* Значение параметра */ dataGridView1.Rows[i].Cells["Стандарт"].Value);
                                else sqlCmd.Parameters.AddWithValue("@Standart", /* Значение параметра */ " ");

                                if (dataGridView1.Rows[i].Cells["Марка"].Value.ToString() != "")
                                    sqlCmd.Parameters.AddWithValue("@Marka", /* Значение параметра */ dataGridView1.Rows[i].Cells["Марка"].Value);
                                else sqlCmd.Parameters.AddWithValue("@Marka", /* Значение параметра */ " ");

                                if (dataGridView1.Rows[i].Cells["Метраж, м (длина, мм)"].Value.ToString() != "")
                                    sqlCmd.Parameters.AddWithValue("@Dlina", /* Значение параметра */ dataGridView1.Rows[i].Cells["Метраж, м (длина, мм)"].Value);
                                else sqlCmd.Parameters.AddWithValue("@Dlina", /* Значение параметра */ " ");

                                if (dataGridView1.Rows[i].Cells["Мерность (т, м, мм)"].Value.ToString() != "")
                                    sqlCmd.Parameters.AddWithValue("@Obem", /* Значение параметра */ dataGridView1.Rows[i].Cells["Мерность (т, м, мм)"].Value);
                                else sqlCmd.Parameters.AddWithValue("@Obem", /* Значение параметра */ " ");

                                if (dataGridView1.Rows[i].Cells["Цена"].Value.ToString() != "")
                                    sqlCmd.Parameters.AddWithValue("@Price", /* Значение параметра */ dataGridView1.Rows[i].Cells["Цена"].Value);
                                else sqlCmd.Parameters.AddWithValue("@Price", /* Значение параметра */ " ");

                                if (dataGridView1.Rows[i].Cells["Примечание"].Value.ToString() != "")
                                    sqlCmd.Parameters.AddWithValue("@Primech", /* Значение параметра */ dataGridView1.Rows[i].Cells["Примечание"].Value);
                                else sqlCmd.Parameters.AddWithValue("@Primech", /* Значение параметра */ " ");

                                sqlCmd.Parameters.AddWithValue("@NameOrg", /* Значение параметра */ textBoxOrgName.Text);

                                sqlCmd.ExecuteNonQuery();
                                #endregion
                            }
                        }
                    }
                    else
                    if (MessageBox.Show("В базе уже есть такая компания\n Старая запись будет удалена...", "Внимание!", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        if (MessageBox.Show("При удалении организии произойдет удаление всей связанной\nс данной организацией продукции\nВы действительно хотите удалить \"" + textBoxOrgName.Text + "\"?", "Внимание", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.OK)
                        {
                            if (conn.State == ConnectionState.Open) conn.Close();
                            try
                            {
                                if (conn.State == ConnectionState.Closed) conn.Open();
                                var sqlCmd = new SqlCommand("dbo.delOrg", conn);
                                sqlCmd.CommandType = CommandType.StoredProcedure;

                                sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ textBoxOrgName.Text);
                                sqlCmd.ExecuteNonQuery();

                                //после удаления старых записей организации добавляем организацию и ее продукцию
                                AddToBase();
                            }
                            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                        }
                    }
                }
                else
                {
                    //добавляем организацию в базу
                    AddToBase();
                }
                if (conn.State == ConnectionState.Open) conn.Close();
            }
        }

        //функция добавления в базу значений из формы в SQL
        private void AddToBase()
        {
            try
            {
                #region добавить организацию
                if (conn.State == ConnectionState.Closed) conn.Open();
                var sqlCmd = new SqlCommand("dbo.insOrg", conn);
                sqlCmd.CommandType = CommandType.StoredProcedure;
                sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ textBoxOrgName.Text);

                if (textBoxOrgAdress.Text == "") textBoxOrgAdress.Text = " ";
                sqlCmd.Parameters.AddWithValue("@Adress", /* Значение параметра */ textBoxOrgAdress.Text);

                if (textBoxOrgTelefon.Text == "") textBoxOrgTelefon.Text = " ";
                sqlCmd.Parameters.AddWithValue("@Telefon", /* Значение параметра */ textBoxOrgTelefon.Text);

                if (textBoxOrgEmail.Text == "") textBoxOrgEmail.Text = " ";
                sqlCmd.Parameters.AddWithValue("@Email", /* Значение параметра */ textBoxOrgEmail.Text);

                if (textBoxOrgSite.Text == "") textBoxOrgSite.Text = " ";
                sqlCmd.Parameters.AddWithValue("@Site", /* Значение параметра */ textBoxOrgSite.Text);

                if (textBoxOrgINN.Text == "") textBoxOrgINN.Text = " ";
                sqlCmd.Parameters.AddWithValue("@INNKPP", /* Значение параметра */ textBoxOrgINN.Text);

                if (textBoxOrgRS.Text == "") textBoxOrgRS.Text = " ";
                sqlCmd.Parameters.AddWithValue("@RasSchet", /* Значение параметра */ textBoxOrgRS.Text);

                if (textBoxOrgKS.Text == "") textBoxOrgKS.Text = " ";
                sqlCmd.Parameters.AddWithValue("@KorSchet", /* Значение параметра */ textBoxOrgKS.Text);

                if (textBoxBIK.Text == "") textBoxBIK.Text = " ";
                sqlCmd.Parameters.AddWithValue("@BIK", /* Значение параметра */ textBoxBIK.Text);

                sqlCmd.Parameters.AddWithValue("@datePrice", /* Значение параметра */ dateTimePicker1.Value.Year.ToString() + "." +
                    dateTimePicker1.Value.Month.ToString() + "." + dateTimePicker1.Value.Day.ToString());

                sqlCmd.Parameters.AddWithValue("@CityName", /* Значение параметра */ comboBoxCity.SelectedItem);

                sqlCmd.ExecuteNonQuery();
                #endregion

                try
                {
                    string name = "";
                    for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                    {
                        #region добавить продукт
                        sqlCmd = new SqlCommand("dbo.insProd", conn);
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        name = dataGridView1.Rows[i].Cells["Название"].Value.ToString();
                        if (new Regex(@"трубы", RegexOptions.IgnoreCase).IsMatch(name)) name = "Труба";
                        if (new Regex(@"листы", RegexOptions.IgnoreCase).IsMatch(name)) name = "Лист";
                        if (new Regex(@"рельсы", RegexOptions.IgnoreCase).IsMatch(name)) name = "Рельса";
                        if (!String.IsNullOrEmpty(name))
                        {
                            if (name.Length > 2)
                                name = name.Substring(0, 1).ToUpper() + name.Substring(1, name.Length - 1).ToLower();
                        }
                        sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ name);

                        if (dataGridView1.Rows[i].Cells["Тип"].Value.ToString() != "")
                        {
                            sqlCmd.Parameters.AddWithValue("@Type", /* Значение параметра */ dataGridView1.Rows[i].Cells["Тип"].Value);
                        }
                        else sqlCmd.Parameters.AddWithValue("@Type", /* Значение параметра */ "тип не указан");

                        string temporarary = dataGridView1.Rows[i].Cells["Диаметр (высота), мм"].Value.ToString();
                        if (temporarary != "")
                        {
                            if (new Regex(@"^\.\d+(?:\s|$)", RegexOptions.IgnoreCase).IsMatch(temporarary)) temporarary = temporarary.Substring(1);
                            else if (new Regex(@"^\.\d+[,\.]\d+(?:\s|$)", RegexOptions.IgnoreCase).IsMatch(temporarary)) temporarary = temporarary.Substring(1);
                            temporarary = new Regex(@"\.").Replace(temporarary, @",");
                            sqlCmd.Parameters.AddWithValue("@Diametr", /* Значение параметра */ Convert.ToDouble(temporarary));
                        }
                        else sqlCmd.Parameters.AddWithValue("@Diametr", /* Значение параметра */ 0);

                        temporarary = dataGridView1.Rows[i].Cells["Толщина (ширина), мм"].Value.ToString();
                        if (dataGridView1.Rows[i].Cells["Толщина (ширина), мм"].Value.ToString() != "")
                        {
                            if (temporarary.IndexOf('-') > -1)
                                temporarary = temporarary.Substring(temporarary.IndexOf('-') + 1);
                            if (new Regex(@"^\.\d+(?:\s|$)", RegexOptions.IgnoreCase).IsMatch(temporarary)) temporarary = temporarary.Substring(1);
                            else if (new Regex(@"^\.\d+[,\.]\d+(?:\s|$)", RegexOptions.IgnoreCase).IsMatch(temporarary)) temporarary = temporarary.Substring(1);
                            temporarary = new Regex(@"\.").Replace(temporarary, @",");
                            sqlCmd.Parameters.AddWithValue("@Tolshina", /* Значение параметра */ Convert.ToDouble(temporarary));
                        }
                        else sqlCmd.Parameters.AddWithValue("@Tolshina", /* Значение параметра */ 0);

                        if (dataGridView1.Rows[i].Cells["Класс"].Value.ToString() != "")
                            sqlCmd.Parameters.AddWithValue("@Klass", /* Значение параметра */ dataGridView1.Rows[i].Cells["Класс"].Value);
                        else sqlCmd.Parameters.AddWithValue("@Klass", /* Значение параметра */ " ");

                        if (dataGridView1.Rows[i].Cells["Стандарт"].Value.ToString() != "")
                            sqlCmd.Parameters.AddWithValue("@Standart", /* Значение параметра */ dataGridView1.Rows[i].Cells["Стандарт"].Value);
                        else sqlCmd.Parameters.AddWithValue("@Standart", /* Значение параметра */ " ");

                        if (dataGridView1.Rows[i].Cells["Марка"].Value.ToString() != "")
                            sqlCmd.Parameters.AddWithValue("@Marka", /* Значение параметра */ dataGridView1.Rows[i].Cells["Марка"].Value);
                        else sqlCmd.Parameters.AddWithValue("@Marka", /* Значение параметра */ " ");

                        if (dataGridView1.Rows[i].Cells["Метраж, м (длина, мм)"].Value.ToString() != "")
                            sqlCmd.Parameters.AddWithValue("@Dlina", /* Значение параметра */ dataGridView1.Rows[i].Cells["Метраж, м (длина, мм)"].Value);
                        else sqlCmd.Parameters.AddWithValue("@Dlina", /* Значение параметра */ " ");

                        if (dataGridView1.Rows[i].Cells["Мерность (т, м, мм)"].Value.ToString() != "")
                            sqlCmd.Parameters.AddWithValue("@Obem", /* Значение параметра */ dataGridView1.Rows[i].Cells["Мерность (т, м, мм)"].Value);
                        else sqlCmd.Parameters.AddWithValue("@Obem", /* Значение параметра */ " ");

                        if (dataGridView1.Rows[i].Cells["Цена"].Value.ToString() != "")
                            sqlCmd.Parameters.AddWithValue("@Price", /* Значение параметра */ dataGridView1.Rows[i].Cells["Цена"].Value);
                        else sqlCmd.Parameters.AddWithValue("@Price", /* Значение параметра */ " ");

                        if (dataGridView1.Rows[i].Cells["Примечание"].Value.ToString() != "")
                            sqlCmd.Parameters.AddWithValue("@Primech", /* Значение параметра */ dataGridView1.Rows[i].Cells["Примечание"].Value);
                        else sqlCmd.Parameters.AddWithValue("@Primech", /* Значение параметра */ " ");

                        sqlCmd.Parameters.AddWithValue("@NameOrg", /* Значение параметра */ textBoxOrgName.Text);

                        sqlCmd.ExecuteNonQuery();
                        #endregion
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                try
                {
                    #region добавление менеджера
                    if (listViewManager.Items.Count > 0)
                    {

                        foreach (ListViewItem lvi in listViewManager.Items)
                        {
                            sqlCmd = new SqlCommand("dbo.insManager", conn);
                            sqlCmd.CommandType = CommandType.StoredProcedure;
                            sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ textBoxOrgName.Text);
                            sqlCmd.Parameters.AddWithValue("@NameManager", /* Значение параметра */ lvi.SubItems[0].Text);
                            sqlCmd.Parameters.AddWithValue("@TelefonManager", /* Значение параметра */ lvi.SubItems[1].Text);
                            if (lvi.SubItems.Count < 3)
                                sqlCmd.Parameters.AddWithValue("@Email", "");
                            else sqlCmd.Parameters.AddWithValue("@Email", lvi.SubItems[2].Text);
                            sqlCmd.ExecuteNonQuery();
                        }
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                try
                {
                    #region добавление склада
                    if (listViewAdrSklad.Items.Count > 0)
                    {

                        foreach (ListViewItem lvi in listViewAdrSklad.Items)
                        {
                            sqlCmd = new SqlCommand("dbo.insSklad", conn);
                            sqlCmd.CommandType = CommandType.StoredProcedure;
                            sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ textBoxOrgName.Text);
                            sqlCmd.Parameters.AddWithValue("@AdressSklad", /* Значение параметра */ lvi.SubItems[0].Text);
                            sqlCmd.ExecuteNonQuery();
                        }
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                if (ManualStringNameProd != "")
                {
                    sqlCmd = new SqlCommand("dbo.insManProd", conn);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.Parameters.AddWithValue("@NameOrg", /* Значение параметра */ orgname);
                    sqlCmd.Parameters.AddWithValue("@NameProd", /* Значение параметра */ ManualStringNameProd);
                    sqlCmd.ExecuteNonQuery();
                    ManualStringNameProd = "";
                }

                MessageBox.Show("Добавление в базу прошло успешно");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void updateBase(string idOrg)
        {
            try
            {
                #region добавить организацию
                if (conn.State == ConnectionState.Closed) conn.Open();
                var sqlCmd = new SqlCommand("dbo.updateOrganization", conn);
                sqlCmd.CommandType = CommandType.StoredProcedure;
                sqlCmd.Parameters.AddWithValue("ID_Organization", idOrg);

                sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ orgname);

                if (textBoxOrgAdress.Text == "") textBoxOrgAdress.Text = " ";
                sqlCmd.Parameters.AddWithValue("@Adress", /* Значение параметра */ textBoxOrgAdress.Text);

                if (textBoxOrgTelefon.Text == "") textBoxOrgTelefon.Text = " ";
                sqlCmd.Parameters.AddWithValue("@Telefon", /* Значение параметра */ textBoxOrgTelefon.Text);

                if (textBoxOrgEmail.Text == "") textBoxOrgEmail.Text = " ";
                sqlCmd.Parameters.AddWithValue("@Email", /* Значение параметра */ textBoxOrgEmail.Text);

                if (textBoxOrgSite.Text == "") textBoxOrgSite.Text = " ";
                sqlCmd.Parameters.AddWithValue("@Site", /* Значение параметра */ textBoxOrgSite.Text);

                if (textBoxOrgINN.Text == "") textBoxOrgINN.Text = " ";
                sqlCmd.Parameters.AddWithValue("@INNKPP", /* Значение параметра */ textBoxOrgINN.Text);

                if (textBoxOrgRS.Text == "") textBoxOrgRS.Text = " ";
                sqlCmd.Parameters.AddWithValue("@RasSchet", /* Значение параметра */ textBoxOrgRS.Text);

                if (textBoxOrgKS.Text == "") textBoxOrgKS.Text = " ";
                sqlCmd.Parameters.AddWithValue("@KorSchet", /* Значение параметра */ textBoxOrgKS.Text);

                if (textBoxBIK.Text == "") textBoxBIK.Text = " ";
                sqlCmd.Parameters.AddWithValue("@BIK", /* Значение параметра */ textBoxBIK.Text);

                sqlCmd.Parameters.AddWithValue("@datePrice", /* Значение параметра */ dateTimePicker1.Value.Year.ToString() + "." +
                    dateTimePicker1.Value.Month.ToString() + "." + dateTimePicker1.Value.Day.ToString());

                sqlCmd.ExecuteNonQuery();
                #endregion

                try
                {
                    for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                    {
                        #region добавить продукт
                        sqlCmd = new SqlCommand("dbo.insProd", conn);
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ dataGridView1.Rows[i].Cells["Название"].Value);

                        if (dataGridView1.Rows[i].Cells["Тип"].Value.ToString() != "")
                            sqlCmd.Parameters.AddWithValue("@Type", /* Значение параметра */ dataGridView1.Rows[i].Cells["Тип"].Value);
                        else sqlCmd.Parameters.AddWithValue("@Type", /* Значение параметра */ "тип не указан");

                        string temporarary = dataGridView1.Rows[i].Cells["Диаметр (высота), мм"].Value.ToString();
                        if (temporarary != "")
                        {
                            if (new Regex(@"^\.\d+(?:\s|$)", RegexOptions.IgnoreCase).IsMatch(temporarary)) temporarary = temporarary.Substring(1);
                            else if (new Regex(@"^\.\d+[,\.]\d+(?:\s|$)", RegexOptions.IgnoreCase).IsMatch(temporarary)) temporarary = temporarary.Substring(1);
                            temporarary = new Regex(@"\.").Replace(temporarary, @",");
                            sqlCmd.Parameters.AddWithValue("@Diametr", /* Значение параметра */ Convert.ToDouble(temporarary));
                        }
                        else sqlCmd.Parameters.AddWithValue("@Diametr", /* Значение параметра */ 0);

                        temporarary = dataGridView1.Rows[i].Cells["Толщина (ширина), мм"].Value.ToString();
                        if (dataGridView1.Rows[i].Cells["Толщина (ширина), мм"].Value.ToString() != "")
                        {
                            if (temporarary.IndexOf('-') > -1)
                                temporarary = temporarary.Substring(temporarary.IndexOf('-') + 1);
                            if (new Regex(@"^\.\d+(?:\s|$)", RegexOptions.IgnoreCase).IsMatch(temporarary)) temporarary = temporarary.Substring(1);
                            else if (new Regex(@"^\.\d+[,\.]\d+(?:\s|$)", RegexOptions.IgnoreCase).IsMatch(temporarary)) temporarary = temporarary.Substring(1);
                            temporarary = new Regex(@"\.").Replace(temporarary, @",");
                            sqlCmd.Parameters.AddWithValue("@Tolshina", /* Значение параметра */ Convert.ToDouble(temporarary));
                        }
                        else sqlCmd.Parameters.AddWithValue("@Tolshina", /* Значение параметра */ 0);

                        if (dataGridView1.Rows[i].Cells["Класс"].Value.ToString() != "")
                            sqlCmd.Parameters.AddWithValue("@Klass", /* Значение параметра */ dataGridView1.Rows[i].Cells["Класс"].Value);
                        else sqlCmd.Parameters.AddWithValue("@Klass", /* Значение параметра */ " ");

                        if (dataGridView1.Rows[i].Cells["Стандарт"].Value.ToString() != "")
                            sqlCmd.Parameters.AddWithValue("@Standart", /* Значение параметра */ dataGridView1.Rows[i].Cells["Стандарт"].Value);
                        else sqlCmd.Parameters.AddWithValue("@Standart", /* Значение параметра */ " ");

                        if (dataGridView1.Rows[i].Cells["Марка"].Value.ToString() != "")
                            sqlCmd.Parameters.AddWithValue("@Marka", /* Значение параметра */ dataGridView1.Rows[i].Cells["Марка"].Value);
                        else sqlCmd.Parameters.AddWithValue("@Marka", /* Значение параметра */ " ");

                        if (dataGridView1.Rows[i].Cells["Метраж, м (длина, мм)"].Value.ToString() != "")
                            sqlCmd.Parameters.AddWithValue("@Dlina", /* Значение параметра */ dataGridView1.Rows[i].Cells["Метраж, м (длина, мм)"].Value);
                        else sqlCmd.Parameters.AddWithValue("@Dlina", /* Значение параметра */ " ");

                        if (dataGridView1.Rows[i].Cells["Мерность (т, м, мм)"].Value.ToString() != "")
                            sqlCmd.Parameters.AddWithValue("@Obem", /* Значение параметра */ dataGridView1.Rows[i].Cells["Мерность (т, м, мм)"].Value);
                        else sqlCmd.Parameters.AddWithValue("@Obem", /* Значение параметра */ " ");

                        if (dataGridView1.Rows[i].Cells["Цена"].Value.ToString() != "")
                            sqlCmd.Parameters.AddWithValue("@Price", /* Значение параметра */ dataGridView1.Rows[i].Cells["Цена"].Value);
                        else sqlCmd.Parameters.AddWithValue("@Price", /* Значение параметра */ " ");

                        sqlCmd.Parameters.AddWithValue("@NameOrg", /* Значение параметра */ orgname);

                        sqlCmd.ExecuteNonQuery();
                        #endregion
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                try
                {
                    #region добавление менеджера
                    if (listViewManager.Items.Count > 0)
                    {
                        //SqlConnection connect = new SqlConnection(sqlConnString);
                        //if (connect.State == ConnectionState.Closed) connect.Open();
                        //SqlDataAdapter co = new SqlDataAdapter("select [name] from manager where id_organization='"+idOrg+"'",connect);
                        //DataTable dttt = new DataTable();
                        //co.Fill(dttt);
                        //if (connect.State == ConnectionState.Open) connect.Close();

                        foreach (ListViewItem lvi in listViewManager.Items)
                        {
                            sqlCmd = new SqlCommand("dbo.updateManager", conn);
                            sqlCmd.CommandType = CommandType.StoredProcedure;
                            sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ orgname);
                            sqlCmd.Parameters.AddWithValue("@NameManager", /* Значение параметра */ lvi.SubItems[0].Text);
                            sqlCmd.Parameters.AddWithValue("@TelefonManager", /* Значение параметра */ lvi.SubItems[1].Text);
                            sqlCmd.Parameters.AddWithValue("@Email", " ");
                            sqlCmd.ExecuteNonQuery();
                        }
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                try
                {
                    #region добавление склада
                    if (listViewAdrSklad.Items.Count > 0)
                    {

                        foreach (ListViewItem lvi in listViewAdrSklad.Items)
                        {
                            sqlCmd = new SqlCommand("dbo.insSklad", conn);
                            sqlCmd.CommandType = CommandType.StoredProcedure;
                            sqlCmd.Parameters.AddWithValue("@Name", /* Значение параметра */ orgname);
                            sqlCmd.Parameters.AddWithValue("@AdressSklad", /* Значение параметра */ lvi.SubItems[0].Text);
                            sqlCmd.ExecuteNonQuery();
                        }
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                //MessageBox.Show("Добавление в базу прошло успешно");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }


        }

        private void btnSearchPath_Click(object sender, EventArgs e)
        {

            FolderBrowserDialog fbd = new FolderBrowserDialog();
            string md = Environment.GetFolderPath(Environment.SpecialFolder.Personal);//путь к Документам
            fbd.SelectedPath = md;//System.IO.Directory.GetCurrentDirectory();
            //fbd.Filter = "All|*.xls;*.xlsx|Excel|*.xls|Excel 2010|*.xlsx";
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                string[] searchPatterns = "*.xls?|*.doc?".Split('|');
                List<string> files = new List<string>();
                foreach (string sp in searchPatterns)
                    files.AddRange(System.IO.Directory.GetFiles(fbd.SelectedPath, sp, SearchOption.TopDirectoryOnly));
                files.Sort();

                listView1.Items.Clear();
                ListViewItem lvi;
                foreach (string file in files)
                {
                    lvi = new ListViewItem(Path.GetFileName(file));
                    lvi.SubItems.Add(file);
                    listView1.Items.Add(lvi);
                }
                filesCountInDirectory = listView1.Items.Count;
            }
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {

            if (true)//MessageBox.Show("Открыть выбранный файл?", "Внимание", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                dtProduct.Clear();
                if (isOpenSqlConnection) conn.Close();
                if (isExcelOpen)
                {
                    try
                    {
                        excelappworkbook.Close(false, Type.Missing, Type.Missing);
                    }
                    catch { }
                    excelapp.Quit();
                    isExcelOpen = false;
                }
                textBoxPath.Text = listView1.SelectedItems[0].SubItems[1].Text;

                dataGridView1.DataSource = null;
                ChoosingShemaFile(listView1.SelectedItems[0].SubItems[1].Text);
            }

        }

        private void ExitEvent(object sender, EventArgs e)
        {
            if (isExcelOpen)
                excelapp.Quit();
            try
            {
                if (isOpenSqlConnection) conn.Close();
            }
            catch { }
        }

        private void btn_Refresh_Click(object sender, EventArgs e)
        {
            //progr = new progress(listView1.Items.Count);
            //progr.Show();
            refresh();
        }

        private void refresh() //обновить всё, проверка дат файлов и их обновление, если нужно
        {
            tsLabelTotalFiles.Text = listView1.Items.Count.ToString();
            try
            {
                string OrgName;
                string[] calendar;
                SqlDataAdapter da;
                DataTable dt;
                SqlCommand sqlCmd;
                conn = new SqlConnection(sqlConnString);

                int Obnovleno = 0;
                bool obnov = false;
                int Dobavleno = 0;
                int[] date; //массив для хранения даты текущего файла
                int[] dateFromBase; //массив для хранения даты по прайсу из базы

                foreach (ListViewItem lvi in listView1.Items)
                {
                    tsLabelCurFile.Text = (lvi.Index + 1).ToString();
                    lvi.Selected = true;
                    dt = new DataTable();
                    OrgName = new Regex(@".+(?=[\s_\.]\d+[\._]\d+[\._]\d+(?:г\.?)?\.[\w\d]{3,4}$)").Match(lvi.SubItems[0].Text).Value;
                    OrgName = GetFileNameForShemaFile(OrgName, regexParam.GetDateTimeFromName(lvi.SubItems[0].Text));
                    calendar = new Regex(@"(?<=.+[\s_\.])\d+[\._]\d+[\._]\d+(?=(?:г\.?)?\.[\w\d]{3,4}$)").Match(lvi.SubItems[0].Text).Value.Split('.', '_');
                    if (calendar[2].Length == 2) calendar[2] = "20" + calendar[2];
                    dateTimePicker1.Value = new DateTime(Convert.ToInt32(calendar[2]), Convert.ToInt32(calendar[1]), Convert.ToInt32(calendar[0]));
                    date = new int[] { Convert.ToInt32(dateTimePicker1.Value.Year),
                        Convert.ToInt32(dateTimePicker1.Value.Month), Convert.ToInt32(dateTimePicker1.Value.Day)};
                    DateTime dtFile = new DateTime(date[0], date[1], date[2]);
                    DateTime dtBase;
                    string idOrg = "";
                    try
                    {
                        if (conn.State == ConnectionState.Closed)
                        {
                            conn.Open();
                            isOpenSqlConnection = true;
                        }
                        //запрос даты прайса по организации из базы
                        da = new SqlDataAdapter("select datePriceList, id_organization from dbo.Organization where [Name]='" + OrgName + "';", conn);
                        da.Fill(dt);
                        if (dt.Rows.Count > 0)
                        {
                            string[] tmp = dt.Rows[0][0].ToString().Split('.');
                            dateFromBase = new int[] { Convert.ToInt32(tmp[0]), Convert.ToInt32(tmp[1]), Convert.ToInt32(tmp[2]) };
                            dtBase = new DateTime(dateFromBase[0], dateFromBase[1], dateFromBase[2]);
                            if (dtFile <= dtBase) continue;
                            da = new SqlDataAdapter("select ID_Organization from dbo.Organization where [Name]='" + OrgName + "';", conn);
                            dt.Clear();
                            da.Fill(dt);

                            if (dt.Rows.Count > 0)
                            {
                                sqlCmd = new SqlCommand("dbo.delUpdProd", conn);
                                sqlCmd.CommandType = CommandType.StoredProcedure;

                                sqlCmd.Parameters.AddWithValue("@NameOrg", /* Значение параметра */ OrgName);
                                sqlCmd.ExecuteNonQuery();
                                Obnovleno++;
                                obnov = true;
                                idOrg = dt.Rows[0][0].ToString();
                            }
                            dtProduct.Clear();
                            if (conn.State == ConnectionState.Open) conn.Close();
                            if (isExcelOpen)
                            {
                                try
                                {
                                    excelappworkbook.Close(false, Type.Missing, Type.Missing);
                                }
                                catch { }
                                excelapp.Quit();
                                isExcelOpen = false;
                            }

                            textBoxPath.Text = lvi.SubItems[0].Text;
                            dataGridView1.DataSource = null;
                            ChoosingShemaFile(lvi.SubItems[1].Text);

                            updateBase(idOrg);

                        }

                        else
                        {
                            dtProduct.Clear();
                            if (conn.State == ConnectionState.Open) conn.Close();
                            if (isExcelOpen)
                            {
                                try
                                {
                                    excelappworkbook.Close(false, Type.Missing, Type.Missing);
                                }
                                catch { }
                                excelapp.Quit();
                                isExcelOpen = false;
                            }

                            textBoxPath.Text = lvi.SubItems[0].Text;
                            dataGridView1.DataSource = null;
                            ChoosingShemaFile(lvi.SubItems[1].Text);

                            AddToBase();
                            if (!obnov)
                            {
                                Dobavleno++;
                            }
                            obnov = false;
                            try
                            {
                                if (conn.State == ConnectionState.Open) conn.Close();
                                if (isExcelOpen)
                                {
                                    try
                                    {
                                        excelappworkbook.Close(false, Type.Missing, Type.Missing);
                                    }
                                    catch { }
                                    excelapp.Quit();
                                    isExcelOpen = false;
                                }
                            }
                            catch { }
                        }
                    }
                    catch (Exception ex) { MessageBox.Show(ex.ToString()); if (isOpenSqlConnection) conn.Close(); }
                }

                //form.Abort();
                MessageBox.Show("Завершено!\nДобавлено: " + Dobavleno + "\nОбновлено: " + Obnovleno);
            }
            catch { }
        }

        private void comboBoxCity_SelectedIndexChanged(object sender, EventArgs e)
        {
            getFilesFromDocDirectory();
        }

        private void btnCancelThread_Click(object sender, EventArgs e)
        {
            try
            {
                if (thread.IsAlive) thread.Abort();
            }
            catch { }
        }
    }
}

