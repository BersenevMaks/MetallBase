using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MetalBase.HelpClasses
{
    public class Class_DTM
    {
        private string diam = "", tolsh = "", metraj = "";
        private Regex dtm = new Regex(@"", RegexOptions.IgnoreCase);
        public Class_DTM(string temp, string type)
        {
            CalcDTM(temp, type);
        }
        public Class_DTM()
        { }
        /// <summary>
        /// Диаметр после CalcDTM
        /// </summary>
        /// <returns>Диаметр</returns>
        public string D()
        { return diam; }
        /// <summary>
        /// Толщина после CalcDTM
        /// </summary>
        /// <returns>Толщина</returns>
        public string T()
        { return tolsh; }
        /// <summary>
        /// Метраж после CalcDTM
        /// </summary>
        /// <returns>Метраж</returns>
        public string M()
        { return metraj; }
        /// <summary>
        /// Использованный Regex для вычисления CalcDTM
        /// </summary>
        /// <returns>Regex после CalcDTM</returns>
        public Regex DTM()
        { return dtm; }

        /// <summary>
        /// Возвращает Regex diamDxDxD
        /// </summary>
        /// <returns></returns>
        public Regex GetRegexDiamDxDxD()
        { return diamDxDxD; }

        /// <summary>
        /// Возвращает Regex diamDxD
        /// </summary>
        /// <returns></returns>
        public Regex GetRegexDiamDxD()
        { return diamDxD; }

        Regex diamDxDxD = new Regex(@"\d+(?:[,.]\d+)?\s*[xх\*×]\s*\d+(?:[,.]\d+)?\s*[xх\*×]\s*\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase);
        Regex diamDxD = new Regex(@"\d+(?:[,.]\d+)?\s*[xх\*×]\s*\d+(?:[,.]\d+)?\s*", RegexOptions.IgnoreCase);
        /// <summary>
        /// Вычислить из строки Диаметр, Толщину и Метраж
        /// </summary>
        /// <param name="temp">Исходная строка</param>
        /// <param name="type">Тип продукта</param>
        /// <param name="name">Имя продукта</param>
        /// <param name="mode">Режим: 0-все режимы, 1-только составные (D(xD)+)</param>
        public void CalcDTM(string temp, string type = "", string name = "", int mode = 0)
        {
            diam = ""; tolsh = ""; metraj = "";
            //diamDxDxD
            if (diamDxDxD.IsMatch(temp))
            {
                diam = new Regex(@"\d+(?:[,.]\d+)?(?=\s*[xх\*×]\s*\d+(?:[,.]\d+)?\s*[xх\*×]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                tolsh = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*×]\s*)\d+(?:[,.]\d+)?(?=\s*[xх\*×]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                metraj = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*×]\s*\d+(?:[,.]\d+)?\s*[xх\*×]\s*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                try
                {
                    if (name.ToLower().Contains("труб"))
                        if (double.Parse(tolsh) > double.Parse(diam)) { var max = diam; diam = tolsh; tolsh = max; }
                    if (!string.IsNullOrEmpty(type) || !string.IsNullOrEmpty(name))
                        if (new Regex(@"проф", RegexOptions.IgnoreCase).IsMatch(type) || new Regex(@"проф", RegexOptions.IgnoreCase).IsMatch(type))
                            if (double.Parse(tolsh) > double.Parse(metraj)) { var max = tolsh; tolsh = metraj; metraj = max; }
                    if (name.ToLower().Contains("угол"))
                    {
                        //string tmp = diam;
                        //diam = metraj;
                        //metraj = tmp;
                        try
                        {
                            double d = Convert.ToDouble(diam);
                            double t = Convert.ToDouble(tolsh);
                            double m = Convert.ToDouble(metraj);
                            double min = d;
                            if (t < min)
                                min = t;
                            if (m < min)
                                min = m;
                            double tmp = t;
                            if (d == min) { t = d; d = tmp; }
                            if (m == min) { t = m; m = tmp; }
                            diam = d.ToString(); tolsh = t.ToString(); metraj = m.ToString();
                        }
                        catch { }
                    }
                    dtm = new Regex(@"\d+(?:[,.]\d+)?\s*[xх\*×]\s*\d+(?:[,.]\d+)?\s*[xх\*×]\s*\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase);
                }
                catch (Exception ex) { ex.ToString(); }
            }
            //diamDxD
            else if (diamDxD.IsMatch(temp))
            {
                diam = new Regex(@"\d+(?:[,.]\d+)?(?=\s*[xх\*×]\s*\d+(?:[,.]\d+)?)", RegexOptions.IgnoreCase).Match(temp).Value;
                tolsh = new Regex(@"(?<=\d+(?:[,.]\d+)?\s*[xх\*×]\s*)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).Match(temp).Value;
                try
                {
                    if (name.ToLower().Contains("труб"))
                        if (double.Parse(tolsh) > double.Parse(diam)) { var max = diam; diam = tolsh; tolsh = max; }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка преобразования\n diam = " + diam + ", tolsh = " + tolsh + "\n" + ex.ToString());
                }
                dtm = new Regex(@"\d+(?:[,.]\d+)?\s*[xх\*×]\s*\d+(?:[,.]\d+)?\s*", RegexOptions.IgnoreCase);
            }
            else if (new Regex(@"(?<=Ø\s?)\d+\b", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                dtm = new Regex(@"(?<=Ø\s?)\d+\b", RegexOptions.IgnoreCase);
                diam = dtm.Match(temp).Value;
            }
            else if (new Regex(@"(?<=\b[дd]?)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase).IsMatch(temp))
            {
                dtm = new Regex(@"(?<=\b[дd]?)\d+(?:[,.]\d+)?", RegexOptions.IgnoreCase);
                diam = dtm.Match(temp).Value;
            }
            else if (mode == 0)
                diam = new Regex(@"\d+(?:[,.]\d+)?(?=\s|$)", RegexOptions.IgnoreCase).Match(temp).Value;
            else diam = "";
            if (string.IsNullOrEmpty(metraj) && !string.IsNullOrEmpty(tolsh))
            {
                metraj = new Regex(@"\d+(?:[,.]\d+)?(?=\s*(?:мм|м\s))", RegexOptions.IgnoreCase).Match(temp).Value;
                if (tolsh == metraj) metraj = "";
            }
        }
    }
}
