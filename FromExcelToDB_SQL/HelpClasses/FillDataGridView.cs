using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace MetalBase
{
    class FillDataGridView
    {
        /// <summary>
        /// SQLConnection
        /// </summary>
        public SqlConnection Conn { get; set; } //sqlConnection
        /// <summary>
        /// Запрос в БД
        /// </summary>
        public string Query { get; set; } //запрос в БД
        /// <summary>
        /// Количество строк в выборке (странице)
        /// </summary>
        public int CountRows { get; set; } //Количество строк в выборке (странице)
        /// <summary>
        /// Конечное значение ID_Product
        /// </summary>
        public int EndIndex { get; set; } //конечный индекс ID_Product
        /// <summary>
        /// Флаг для перевода в незанятое состояние потока чтения запроса
        /// </summary>
        public bool FutureNotBusy { get; set; }
        /// <summary>
        /// Номер выполняемого потока
        /// </summary>
        public int NumbOfThread { get; set; }
        /// <summary>
        /// Вариант запроса в зависимости от выбора в дереве
        /// </summary>
        public int Variant { get; set; }
        /// <summary>
        /// e_Node_Text
        /// </summary>
        public string e_Node_Text { get; set; }
        /// <summary>
        /// e_Node_Parent_Text
        /// </summary>
        public string e_Node_Parent_Text { get; set; }
        /// <summary>
        /// Город для выборки
        /// </summary>
        public string City { get; set; }
		/// <summary>
		/// Имя вкладки для рекадктирования
		/// </summary>
		public string TabName { get; set; }

    }
}
