using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Web;
using System.Text.RegularExpressions;
using System.IO.Compression;

namespace MetalBase.Forms
{
    public partial class Form_PDF : Form
    {
        public Form_PDF()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    OpenFileDialog ofd = new OpenFileDialog();
            //    if (ofd.ShowDialog() == DialogResult.OK)
            //    {
            //        richTextBox1.Text = ofd.FileName + "\n";
            //        richTextBox1.Text = Pdf2text(ofd.FileName);
            //    }


            //}
            //catch (Exception exception)
            //{
            //    MessageBox.Show(exception.ToString());
            //}
        }


        //private string Pdf2text(string filename)
        //{
        //    // Читаем данные из pdf-файла в строку, учитываем, что файл может содержать
        //    // бинарные потоки.
        //    try
        //    {
        //        //var infile = File.Open(filename, FileMode.Open);
        //        //if (infile.Length == 0)
        //        //{
        //        //infile.Close();
        //        //return "";
        //        //}
        //        // Проход первый. Нам требуется получить все текстовые данные из файла.
        //        // В 1ом проходе мы получаем лишь "грязные" данные, с позиционированием,
        //        // с вставками hex и так далее.
        //        Dictionary<string, string> transformations = new Dictionary<string, string>();
        //        List<string> texts = new List<string>();
        //        // Для начала получим список всех объектов из pdf-файла.
        //        using (StreamReader sr = new StreamReader(filename))
        //        {
        //            List<List<string>> file = new List<List<string>>();
        //            List<string> obj_strings = new List<string>();
        //            string line = "";
        //            bool readObj = false;
        //            while (!sr.EndOfStream)
        //            {
        //                line = sr.ReadLine();
        //                if (!readObj)
        //                {
        //                    if (new Regex(@"obj", RegexOptions.IgnoreCase).IsMatch(line))
        //                    {
        //                        readObj = true;
        //                        obj_strings.Add(new Regex(@"(?<=obj).*", RegexOptions.IgnoreCase).Match(line).Value);
        //                    }
        //                }
        //                else
        //                {
        //                    if (!new Regex(@"endobj", RegexOptions.IgnoreCase).IsMatch(line))
        //                    {
        //                        obj_strings.Add(line);
        //                    }
        //                    else
        //                    {
        //                        obj_strings.Add(new Regex(@".*(?=endobj)", RegexOptions.IgnoreCase).Match(line).Value);
        //                        readObj = false;
        //                        file.Add(obj_strings);
        //                    }
        //                }

        //            }
        //            //string srr = File.ReadAllText(filename);
        //            //MatchCollection objects = new Regex(@"obj\(.*?\)endobj", RegexOptions.IgnoreCase).Matches(richTextBox1.Text);
        //            //$objects = @$objects[1];
        //            // Начнём обходить, то что нашли - помимо текста, нам может попасться
        //            // много всего интересного и не всегда "вкусного", например, те же шрифты.
        //            for (int i = 0; i < file.Count; i++)
        //            {
        //                List<string> currentObject = file[i];
        //                // Проверяем, есть ли в текущем объекте поток данных, почти всегда он
        //                // сжат с помощью gzip.
        //                List<List<string>> streams = new List<List<string>>();
        //                bool isStream = false;
        //                List<string> stream_strings = new List<string>();
        //                foreach (string str_temp in currentObject)
        //                {
        //                    if (!isStream)
        //                    {
        //                        if (new Regex(@"stream", RegexOptions.IgnoreCase).IsMatch(str_temp))
        //                        {
        //                            isStream = true;
        //                            stream_strings.Add(new Regex(@"(?<=stream).*", RegexOptions.IgnoreCase).Match(str_temp).Value);
        //                        }
        //                    }
        //                    else
        //                    {
        //                        if (!new Regex(@"stream", RegexOptions.IgnoreCase).IsMatch(line))
        //                        {
        //                            stream_strings.Add(line);
        //                        }
        //                        else
        //                        {
        //                            stream_strings.Add(new Regex(@".*(?=stream)", RegexOptions.IgnoreCase).Match(line).Value);
        //                            isStream = false;
        //                            streams.Add(stream_strings);
        //                        }
        //                    }
        //                }
        //                //MatchCollection streams = new Regex(@"(?<=stream).*?(?=endstream)", RegexOptions.IgnoreCase).Matches(currentObject);
        //                for (int j = 0; j < streams.Count; j++)
        //                    if (streams[j].Count > 0)
        //                    {
        //                        Match stream = streams[j];
        //                        // Читаем параметры данного объекта, нас интересует только текстовые
        //                        // данные, поэтому делаем минимальные отсечения, чтобы ускорить
        //                        // выполнения
        //                        var options = GetObjectOptions(currentObject);
        //                        if (options.ContainsKey("Length1") && options.ContainsKey("Type") && options.ContainsKey("Subtype"))
        //                            if (!(string.IsNullOrEmpty(options["Length1"]) && string.IsNullOrEmpty(options["Type"]) && string.IsNullOrEmpty(options["Subtype"])))
        //                                continue;
        //                        // Итак, перед нами "возможно" текст, расшифровываем его из бинарного
        //                        // представления. После этого действия мы имеем дело только с plain text.
        //                        var data = GetDecodedStream(stream, options);
        //                        if (data.Length > 0)
        //                        {
        //                            // Итак, нам нужно найти контейнер текста в текущем потоке.
        //                            // В случае успеха найденный "грязный" текст отправится к остальным
        //                            // найденным до этого
        //                            MatchCollection textContainers = new Regex(@"(?<=BT)(.*?)(?=ET)", RegexOptions.IgnoreCase).Matches(data);
        //                            for (int t = 0; t < textContainers.Count; t++)
        //                                if (textContainers[t].Length > 0)
        //                                {
        //                                    texts = getDirtyTexts(textContainers);
        //                                    // В противном случае, пытаемся найти символьные трансформации,
        //                                    // которые будем использовать во втором шаге.
        //                                }
        //                                else
        //                                    transformations = GetCharTransformations(textContainers);
        //                        }
        //                    }
        //            }
        //            // По окончанию первичного парсинга pdf-документа, начинаем разбор полученных
        //            // текстовых блоков с учётом символьных трансформаций. По окончанию, возвращаем
        //            // полученный результат.
        //            return GetTextUsingTransformations(texts, transformations);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString());
        //        return "";
        //    }
        //}

        //private Dictionary<string, string> GetObjectOptions(string mobject)
        //{
        //    // Нам нужно получить параметры текущегго объекта. Параметры
        //    // находся между ёлочек << и >>. Каждая опция начинается со слэша /.
        //    Dictionary<string, string> options = new Dictionary<string, string>();
        //    List<string> options_strings = new List<string>();
        //    foreach (Match mObject in new Regex(@"(?<=<<).*?(?=>>)", RegexOptions.IgnoreCase).Matches(mobject))
        //    {
        //        // Отделяем опции друг от друга по /. Первую пустую удаляем из массива.
        //        options_strings = mObject.Value.Split('/').ToList();
        //        //options_strings.RemoveAt(0);
        //        // Далее создадим удобный для будущего использования массив
        //        // свойств текущего объекта. Параметры вида "/Option N" запишем
        //        // в хэш, как "Option" => N, свойства типа "/Param", как
        //        // "Param" => true.
        //        Dictionary<string, string> o = new Dictionary<string, string>();
        //        for (int j = 0; j < options_strings.Count; j++)
        //        {
        //            options_strings[j] = new Regex(@"\s+").Replace(options_strings[j], " ").Trim();
        //            if (options_strings[j].IndexOf(' ') > -1)
        //            {
        //                char[] c = { ' ' };
        //                string[] parts = options_strings[j].Split(c, 2);
        //                o[parts[0]] = parts[1];
        //            }
        //            else
        //                o[options_strings[j]] = "true";
        //        }
        //        options = o;
        //        //unset(o);
        //    }
        //    // Возращаем массив найденных параметров.
        //    return options;
        //}

        //private string GetDecodedStream(Match stream, Dictionary<string, string> options)
        //{
        //    // Итак, перед нами поток, возможно кодированный каким-нибудь
        //    // методом сжатия, а то и несколькими. Попробуем расшифровать.
        //    string data = "";
        //    // Если у текущего потока есть свойство Filter, то он точно
        //    // сжат или зашифрован. Иначе, просто возвращаем содержимое
        //    // потока назад.
        //    if (options.ContainsKey("Filter"))
        //        if (options["Filter"].Length > 0)
        //            data = stream.Value;
        //        else
        //        {
        //            // Если в опциях есть длина потока данных, то нам нужно обрезать данные
        //            // по заданной длине, а не то расшифровать не сможем или ещё какая
        //            // беда случится.
        //            int length = !string.IsNullOrEmpty(options["Length"]) ? int.Parse(options["Length"]) : stream.Value.Length; //&& strpos($options["Length"], " ") == false
        //            List<string> _stream = new List<string>();
        //            for (int k = 0; k < stream.Value.Length; k += length)
        //            {
        //                _stream.Add(stream.Value.Substring(k, length));
        //            }
        //            string __stream = "";
        //            // Перебираем опции на предмет наличия указаний на сжатие данных в текущем
        //            // потоке. PDF поддерживает много всего и разного, но текст кодируется тремя
        //            // вариантами: ASCII Hex, ASCII 85-base и GZ/Deflate. Ищем соответствующие
        //            // ключи и применяем соответствующие функции для расжатия. Есть ещё вариант
        //            // Crypt, но распознавать шифрованные PDF'ки мы не будем.
        //            foreach (KeyValuePair<string, string> kvp in options)
        //            {
        //                if (kvp.Value == "ASCIIHexDecode")
        //                    __stream = DecodeAsciiHex(_stream);
        //                if (kvp.Value == "ASCII85Decode")
        //                    __stream = DecodeAscii85(_stream);
        //                if (kvp.Value == "FlateDecode")
        //                    __stream = DecodeFlate(_stream);
        //            }
        //            data = __stream;
        //        }
        //    // Возвращаем результат наших злодейств.
        //    return data;
        //}

        //private string DecodeAsciiHex(List<string> input)
        //{
        //    string output = "";
        //    bool isOdd = true;
        //    bool isComment = false;
        //    for (int i = 0, codeHigh = -1; i < input.Count && input[i].ToString() != ">"; i++)
        //    {
        //        string c = input[i].ToString();
        //        if (isComment)
        //        {
        //            if (c == "\r" || c == "\n")
        //                isComment = false;
        //            continue;
        //        }
        //        switch (c)
        //        {
        //            case "\0": case "\t": case "\r": case "\f": case "\n": case " ": break;
        //            case "%":
        //                isComment = true;
        //                break;
        //            default:
        //                int code = int.Parse(c, System.Globalization.NumberStyles.HexNumber);
        //                if (code == 0 && c != "0")
        //                    return "";
        //                if (isOdd)
        //                    codeHigh = code;
        //                else
        //                    output += (Char)(codeHigh * 16 + code);
        //                isOdd = !isOdd;
        //                break;
        //        }

        //        if (input[i].ToString() != ">")
        //            return "";
        //        if (isOdd)
        //            output += (Char)(codeHigh * 16);
        //    }
        //    return output;
        //}

        //private string DecodeAscii85(List<string> input)
        //{
        //    string output = "";
        //    bool isComment = false;
        //    List<int> ords = new List<int>();
        //    int state = 0;
        //    for (int i = 0; i < input.Count && input[i] != "~"; i++)
        //    {
        //        string c = input[i];
        //        if (isComment)
        //        {
        //            if (c == "\r" || c == "\n")
        //                isComment = false;
        //            continue;
        //        }
        //        if (c == "\0" || c == "\t" || c == "\r" || c == "\f" || c == "\n" || c == " ")
        //            continue;
        //        if (c == "%")
        //        {
        //            isComment = true;
        //            continue;
        //        }
        //        if (c == "z" && state == 0)
        //        {
        //            for (int o = 0; o < 4; o++)
        //                output += (Char)(0);
        //            continue;
        //        }
        //        if (Char.Parse(c) < '!' || Char.Parse(c) > 'u')
        //            return "";
        //        int code = Convert.ToByte(input[i]) & 0xff;
        //        ords.Add(code - Convert.ToByte('!'));
        //        state++;
        //        if (state == 5)
        //        {
        //            state = 0; int sum = 0;
        //            for (int j = 0; j < 5; j++)
        //                sum = sum * 85 + ords[j];
        //            for (int j = 3; j >= 0; j--)
        //                output += (Char)(sum >> (j * 8));
        //        }
        //    }
        //    if (state == 1)
        //        return "";
        //    else if (state > 1)
        //    {
        //        int sum = 0;
        //        for (int i = 0; i < state; i++)
        //            sum += (ords[i] + (state - 1)) * 85 ^ (4 - i);
        //        for (int i = 0; i < state - 1; i++)
        //            output += (Char)(sum >> ((3 - i) * 8));
        //    }
        //    return output;
        //}

        //private string DecodeFlate(List<string> input)
        //{
        //    // Наиболее частый тип сжатия потока данных в PDF.
        //    // Очень просто реализуется функционалом библиотек.
        //    string output = "";
        //    foreach (string str in input)
        //    {
        //        byte[] bytes = System.Text.Encoding.UTF8.GetBytes(str);
        //        string new_str = System.Text.Encoding.UTF8.GetString(GZipUncompress(bytes));
        //        output += " " + new_str;
        //    }
        //    return output;
        //}

        //private static byte[] GZipUncompress(byte[] data)
        //{
        //    using (var input = new MemoryStream(data))
        //    using (var gzip = new GZipStream(input, CompressionMode.Decompress))
        //    using (var output = new MemoryStream())
        //    {
        //        byte[] buffer = new byte[2048];
        //        int bytesRead;
        //        long totalBytes = 0;
        //        while ((bytesRead = gzip.Read(buffer, 0, buffer.Length)) > 0)
        //        {
        //            output.Write(buffer, 0, bytesRead);
        //            totalBytes += bytesRead;
        //        }
        //        return output.ToArray();
        //    }
        //}

        //private static List<string> getDirtyTexts(MatchCollection textContainers)
        //{
        //    // Итак, у нас есть массив контейнеров текста, выдранных из пары BT и ET.
        //    // Наша новая задача, найти в них текст, который отображается просмотрщиками
        //    // на экране. Вариантов много, рассмотрим пару: [...] TJ и Td (...) Tj
        //    List<string> texts = new List<string>();
        //    for (int j = 0; j < textContainers.Count; j++)
        //    {
        //        // Добавляем найденные кусочки "грязных" данных к общему массиву
        //        // текстовых объектов.
        //        string temp = textContainers[j].Value;
        //        var parts = new Regex(@"\[(.*?)\]\s*?(?=TJ)", RegexOptions.IgnoreCase).Matches(temp);
        //        if (parts.Count > 0)
        //            foreach (Match m in parts)
        //                texts.Add(parts[1].Groups[1].Value);
        //        else
        //        {
        //            parts = new Regex(@"(?<=Td)\s*(\(.*\))\s*?(?=Tj)", RegexOptions.IgnoreCase).Matches(temp);
        //            if (parts.Count > 0)
        //                foreach (Match m in parts)
        //                    texts.Add(parts[1].Groups[1].Value);
        //        }
        //    }
        //    return texts;
        //}

        //private static Dictionary<string, string> GetCharTransformations(MatchCollection textContainers)
        //{
        //    // О Мама Миа! Этого насколько я мог увидеть, никто не реализовывал на PHP, по крайней
        //    // мере в открытом доступе. Сейчас мы займёмся весёлым, начнём искать по потокам
        //    // трансформации символов. Под трансформацией я имею ввиду перевод одного символа в hex-
        //    // представлении в другой, или даже в некоторую последовательность.
        //    // Нас интересуют следующие поля, которые мы должны отыскать в текущем потоке.
        //    // Данные между beginbfchar и endbfchar преобразовывают один hex-код в другой (или 
        //    // последовательность кодов) по отдельности. Между beginbfrange и endbfrange организуется
        //    // преобразование над последовательностями данных, что сокращает  количество определений.
        //    string stream = "";
        //    Dictionary<string, string> transformations = new Dictionary<string, string>();
        //    foreach (Match m in textContainers)
        //    {
        //        stream += m.Value + "\n";
        //    }
        //    var chars = new Regex(@"([0-9]+)\s+beginbfchar(.*?)endbfchar", RegexOptions.IgnoreCase).Matches(stream);
        //    var ranges = new Regex(@"([0-9]+)\s+beginbfrange(.*?)endbfrange", RegexOptions.IgnoreCase).Matches(stream);

        //    // Вначале обрабатываем отдельные символы. Строка преобразования выглядит так:
        //    // - <0123> <abcd> -> 0123 преобразовывается в abcd;
        //    // - <0123> <abcd6789> -> 0123 преобразовывается в несколько символов (в данном случае abcd и 6789)
        //    for (int j = 0; j < chars.Count; j++)
        //    {
        //        // Перед списком данных, есть число обозначающее количество строк, которые нужно
        //        // прочитать. Мы будем брать его в рассчёт.
        //        int count = int.Parse(chars[j].Groups[1].Value);
        //        string[] current = chars[j].Groups[2].Value.Trim().Split('\n');
        //        // Читаем данные из каждой строчки.
        //        for (int k = 0; k < count && k < current.Length; k++)
        //        {
        //            // После этого записываем новую найденную трансформацию. Не забываем, что
        //            // если символов меньше четырёх, мы должны дописать нули.
        //            foreach (Match map in new Regex(@"<([0-9a-f]{2,4})>\s+<([0-9a-f]{4,512})>", RegexOptions.IgnoreCase).Matches(current[k].Trim()))
        //            {
        //                transformations.Add(map.Groups[1].Value.PadLeft(4, '0'), map.Groups[2].Value);
        //            }
        //        }
        //    }
        //    // Теперь обратимся к последовательностям. По документации последовательности бывают
        //    // двух видов, а именно:
        //    // - <0000> <0020> <0a00> -> в этом случае <0000> будет заменено на <0a00>, <0001> на <0a01> и
        //    //   так далее до <0020>, которое превратится в <0a20>.
        //    // - <0000> <0002> [<abcd> <01234567> <8900>] -> тут всё работает чуть по другому. Смотрим, сколько
        //    //   элементов находится между <0000> и <0002> (вместе с 0001 три). Потом каждому из элементов
        //    //   присваиваем значение из квадратных скобок: 0000 -> abcd, 0001 -> 0123 4567, а 0002 -> 8900.
        //    for (int j = 0; j < ranges.Count; j++)
        //    {
        //        // Опять сверяемся с количеством элементов для трансформации.
        //        int count = int.Parse(ranges[j].Groups[1].Value);
        //        string[] current = chars[j].Groups[2].Value.Trim().Split('\n');
        //        // Перебираем строчки.
        //        for (int k = 0; k < count && k < current.Length; k++)
        //        {
        //            // В данном случае последовательность первого типа.
        //            MatchCollection matchCollection = new Regex(@"<([0-9a-f]{4})>\s+<([0-9a-f]{4})>\s+<([0-9a-f]{4})>", RegexOptions.IgnoreCase).Matches(current[k].Trim());
        //            if (matchCollection.Count > 0)
        //                foreach (Match map in matchCollection)
        //                {
        //                    // Переводим данные в 10-чную систему счисления, так проще прошагать циклом.
        //                    int from = int.Parse(map.Groups[1].Value, System.Globalization.NumberStyles.HexNumber);
        //                    int to = int.Parse(map.Groups[2].Value, System.Globalization.NumberStyles.HexNumber);
        //                    int _from = int.Parse(map.Groups[3].Value, System.Globalization.NumberStyles.HexNumber);
        //                    // В массив трансформаций добавляем все элементы между началом и концом последовательности.
        //                    // По документации мы должны добавить ведущие нули, если длина hex-кода меньше четырёх символов.
        //                    for (int m = from, n = 0; m <= to; m++, n++)
        //                        transformations.Add(string.Format("{0:0000}", m.ToString("X")), string.Format("{0:0000}", (_from + n).ToString("X")));
        //                    // Второй вариант.
        //                }
        //            else
        //            {
        //                matchCollection = new Regex(@"#<([0-9a-f]{4})>\s+<([0-9a-f]{4})>\s+\[(.*)\]", RegexOptions.IgnoreCase).Matches(current[k].Trim());
        //                if (matchCollection.Count > 0)
        //                {
        //                    foreach (Match map in matchCollection)
        //                    {
        //                        // Также начало и конец последовательности. Бъём данные в квадратных скобках
        //                        // по (около)пробельным символам.
        //                        int from = int.Parse(map.Groups[1].Value, System.Globalization.NumberStyles.HexNumber);
        //                        int to = int.Parse(map.Groups[2].Value, System.Globalization.NumberStyles.HexNumber);
        //                        string[] parts = new Regex(@"\s").Split(map.Groups[3].Value);

        //                        // Обходим данные и присваиваем соответствующие данные их новым значениям.
        //                        for (int m = from, n = 0; m <= to && n < parts.Length; m++, n++)
        //                            transformations.Add(string.Format("{0:0000}", m.ToString("X")), string.Format("{0:0000}", int.Parse(parts[n], System.Globalization.NumberStyles.HexNumber).ToString("X")));
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    return transformations;
        //}

        //private string GetTextUsingTransformations(List<string> texts, Dictionary<string, string> transformations)
        //{
        //    // Начинаем второй этап - получение текста из "грязных" данных.
        //    // В PDF "грязные" текстовые строки могут выглядеть следующим образом:
        //    // - (I love)10(PHP) - в данном случае в () находятся текстовые данные,
        //    //   а 10 являет собой величину проблема.
        //    // - <01234567> - в данном случае мы имеем дело с двумя символами,
        //    //   в их HEX-представлении: 0123 и 4567. Оба символа следует проверить
        //    //   на наличие замещений в таблице трансформаций.
        //    // - (Hello, \123world!) - \123 здесь символ в 8-чной системе счисления,
        //    //   его также требуется верно обработать.
        //    // Что ж поехали. Начинаем потихоньку накапливать текстовые данные,
        //    // перебирая "грязные" текстовые кусочки"
        //    string document = "";
        //    for (int i = 0; i < texts.Count; i++)
        //    {
        //        // Нас интересуют две ситуации, когда текст находится в <> (hex) и в
        //        // () (plain-представление.
        //        bool isHex = false;
        //        bool isPlain = false;
        //        string hex = "";
        //        string plain = "";
        //        // Посимвольно сканируем текущий текстовый кусок.
        //        for (int j = 0; j < texts[i].Length; j++)
        //        {
        //            // Выбираем текущий символ
        //            Char c = texts[i][j];
        //            // ...и определяем, что нам с ним делать.
        //            switch (c)
        //            {
        //                // Перед нами начинаются 16-чные данные
        //                case '<':
        //                    hex = "";
        //                    isHex = true;
        //                    break;
        //                // Hex-данные закончились, будем их разбирать.
        //                case '>':
        //                    // Бъём строку на кусочки по 4 символа...
        //                    string[] hexs = new string[hex.Length / 4];
        //                    for (int h = 0; h < hex.Length / 4; h++)
        //                    {
        //                        hexs[i] = hex.Substring(h * 4, 4);
        //                    }

        //                    // ...и смотрим, что мы можем с каждым кусочком сделать
        //                    for (int k = 0; k < hexs.Length; k++)
        //                    {
        //                        // Возможна ситуация, что в кусочке меньше 4 символов, документация
        //                        // говорит нам дополнить кусок справа нулями.
        //                        string chex = hexs[k].PadLeft(4, '0');
        //                        // Проверяем наличие данного hex-кода в трансформациях. В случае
        //                        // успеха, заменяем кусок на требуемый.
        //                        if (transformations.ContainsKey(chex))
        //                            chex = transformations[chex];
        //                        // Пишем в выходные данные новый Unicode-символ.
        //                        document += "&#x" + chex + ";"; // WebUtility or HtmlUtility
        //                    }
        //                    // Hex-данные закончились, не забываем сказать об этом коду.
        //                    isHex = false;
        //                    break;
        //                // Начался кусок "чистого" текста
        //                case '(':
        //                    plain = "";
        //                    isPlain = true;
        //                    break;
        //                // Ну и как водится, этот кусок когда-нибудь закончится.
        //                case ')':
        //                    // Добавляем полученный текст в выходной поток.
        //                    document += plain;
        //                    isPlain = false;
        //                    break;
        //                // Символ экранирования, глянем, что стоит за ним.
        //                case '\\':
        //                    char c2 = texts[i][j + 1];
        //                    // Если это \ или одна из круглых скобок, то нужно их вывести, как есть.
        //                    if (c2 == '\\' || c2 == '(' || c2 == ')') plain += c2;
        //                    // Возможно, это пробельный символ или ещё какой перевод строки, обрабатываем.
        //                    else if (c2 == 'n') plain += '\n';
        //                    else if (c2 == 'r') plain += '\r';
        //                    else if (c2 == 't') plain += '\t';
        //                    else if (c2 == 'b') plain += '\b';
        //                    else if (c2 == 'f') plain += '\f';
        //                    // Может случится, что за \ идёт цифра. Таких цифр может быть до 3, они являются
        //                    // кодом символа в 8-чной системе счисления. Распарсим и их.
        //                    else if (c2 >= '0' && c2 <= '9')
        //                    {
        //                        // Нам нужны три цифры не более, и именно цифры.
        //                        string oct = new Regex(@"[^0-9]").Replace(texts[i].Substring(j + 1, 3), "");
        //                        // Определяем сколько символов мы откусили, чтобы сдвинуть позицию
        //                        // "текущего символа" правильно.
        //                        j += oct.Length - 1;
        //                        // В "чистый" текст пишем соответствующий символ.
        //                        plain += "&#" + Convert.ToInt32(oct, 8) + ";";//html_entity_decode("&#".octdec($oct).";");
        //                    }
        //                    // Мы сдвинули позицию "текущего символа" не меньше, чем на один, парсер
        //                    // узнай об этом.
        //                    j++;
        //                    break;
        //                // Если же перед нами что-то другое, то пишем текущий символ либо во
        //                // временную hex-строку (если до этого был символ <),
        //                default:
        //                    if (isHex)
        //                        hex += c;
        //                    // либо в "чистую" строку, если была открыта круглая скобка.
        //                    if (isPlain)
        //                        plain += c;
        //                    break;
        //            }
        //        }
        //        // Блоки текста отделяем переводами строк.
        //        document += "\n";
        //    }
        //    // Возвращаем полученные текстовые данные.
        //    return document;
        //}
    }
}
