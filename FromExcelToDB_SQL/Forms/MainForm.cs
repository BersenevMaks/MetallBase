using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;
using System.Runtime.InteropServices;
using MetalBase.Forms;

namespace MetalBase
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            treeView1.BeforeExpand+=new TreeViewCancelEventHandler(treeView1_BeforeExpand);
            treeView1.BeforeCollapse += new TreeViewCancelEventHandler(treeView1_BeforeCollapse);
            treeView1.MouseDown+=new MouseEventHandler(treeView1_MouseDown);
        }

        
        public delegate void FillDataGridViewDelegate(SqlConnection conn, string str);
        bool isOpenSqlConnection = false;
        bool busy = false;
        bool _cancelFillDataGrid = false;
        int NumbOfThreads = 1; //количество потоков для загрузки
        DataTable dtProduct = new DataTable();
        
        string sqlConnString = @"Server=tcp:maks-pc,1433; Database=MetalBase; User ID=metuser; Password=metuser; Pooling=true;";
        SqlConnection conn;

        List<string> set = new List<string>();

        void GetSettings(List<string> sett)
        {
            set = sett;
        }

        private void settings()
        {
            string fileName = Path.GetDirectoryName(Application.ExecutablePath)+"\\settings.set";                //пишем полный путь к файлу
            if (File.Exists(fileName) == true)
            {
                string[] allText = {"","","","","",""};
                try
                {                                  //чтение файла
                    allText = File.ReadAllLines(fileName);         //чтение всех строк файла в массив строк
                }
                catch (FileNotFoundException e)
                {
                    Console.WriteLine(e.Message);
                }
                if (allText.Length == 7)
                {
                    sqlConnString = "Server=tcp:" + allText[0] + "," + allText[2] + "; Database=" + allText[3] + "; User ID=" + allText[4] + "; Password=" + allText[5] + "; Pooling=true;";
                    if(!string.IsNullOrEmpty(allText[6]))
                        int.TryParse(allText[6], out NumbOfThreads);
                }
                else {
                    settingForm sf = new settingForm(new settingDelegate(GetSettings));
            sf.ShowDialog();
                    if (set.Count == 7)
                    { sqlConnString = "Server=tcp:" + set[0] + "," + set[2] + "; Database=" + set[3] + "; User ID=" + set[4] + "; Password=" + set[5] + "; Pooling=true;";
                        int.TryParse(set[6], out NumbOfThreads);
                    }
                    else sqlConnString = @"Server=tcp:maks-pc,1433; Database=MetalBase; User ID=metuser; Password=metuser";
                }
            }
            else sqlConnString = @"Server=tcp:maks-pc,1433; Database=MetalBase; User ID=metuser; Password=metuser";
            //Console.ReadKey();
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void добавитьЗаписьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ImportMetalBase imb = new ImportMetalBase(sqlConnString);
            imb.Show();
        }

        DataTable dt;
        BindingSource bind;

        private void connectToDBSQL(object sender, EventArgs e)
        {
            connectAndUpdate();
        }

        private void connectAndUpdate()
        {
            toolStatus.Text = "Выполняется подключение";
            try
            {
                conn = new SqlConnection(sqlConnString);
                conn.Open();
                isOpenSqlConnection = true;

                dt = new DataTable();
                SqlCommand comm;
                if(toolStripComboBoxCity.SelectedIndex>0)
                    comm = new SqlCommand("Select distinct [Name] from dbo.Product where id_organization in (select id_organization from dbo.organization where organization.[City]='" + toolStripComboBoxCity.SelectedItem + "') order by [Name]", conn);
                else comm = new SqlCommand("Select distinct [Name] from dbo.Product order by [Name]", conn);
                SqlDataReader reader = comm.ExecuteReader();
                treeView1.Nodes.Clear();
                while (reader.Read())
                {
                    treeView1.Nodes.Add(reader.GetString(0));
                }
                reader.Close();
                
                for (int i = 0; i < treeView1.Nodes.Count; i++)
                {
                    comm = new SqlCommand("Select distinct [Type] from dbo.Product where [Name]='" + treeView1.Nodes[i].Text + "' order by [Type]", conn);
                    reader = comm.ExecuteReader();
                    while (reader.Read())
                    {
                        string str = reader.GetString(0);
                        str = str.Trim();
                        if (str != "")
                            treeView1.Nodes[i].Nodes.Add(reader.GetString(0));
                    }
                    reader.Close();
                }
                treeView1.Nodes[0].Expand();
                treeView1.Nodes[0].Collapse();
                // Call Close when done reading.
                if (!reader.IsClosed)
                    reader.Close();

                //string query = "Select * from dbo.Product";
                //SqlDataAdapter adapter = new SqlDataAdapter(query, conn);
                //adapter.Fill(dt);
                dataGridView1.DataSource = dt;

                conn.Close();
                toolStatus.Text = "Подключено";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                if (isOpenSqlConnection) conn.Close();
                toolStatus.Text = "Не подключено";
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            toolStripComboBox1.Items.Add("1");
            toolStripComboBox1.Items.Add("2");
            toolStripComboBox1.Items.Add("3");
            toolStripComboBox1.SelectedIndex = 0;
            settings();
            this.WindowState = FormWindowState.Maximized;
            toolStripDropDownButton4.Image = MetalBase.Properties.Resources.connect;
            toolStripDropDownButton5.Image = MetalBase.Properties.Resources.w256h2561346685464Refresh;
            toolStripDropDownButton6.Image = MetalBase.Properties.Resources.Import_excel;
            toolStripCBdefaultCity.SelectedIndex = 1;
            toolStripComboBoxCity.SelectedIndex = toolStripCBdefaultCity.SelectedIndex;
			
		}

       

        private void readDataFromSql(TreeNodeMouseClickEventArgs e)
        {
            try
            {
                FillDataGridView fdv = new FillDataGridView();
                bind = new BindingSource();
                //dataGridView1 = new DataGridView();
                dt = new DataTable();
				dtProduct = new DataTable();
                treeView1.SelectedNode = e.Node;
                if (conn.State == ConnectionState.Closed) conn.Open();
                isOpenSqlConnection = true;
                string query;
                SqlDataAdapter adapter;
                dataGridView2.Columns.Clear();
                dataGridView2.Rows.Clear();
                //Выбираем из базы инфу по выделенной позиции в дереве
                if (e.Node.Parent == null)//Если родитель равенн Нулль, то значит родителей нет, а значит - узел корневой
                {
                    //if (e.Node.Nodes.Count < 1)
                    {

                        if (new Regex(@"труб|армату|круг|тройник", RegexOptions.IgnoreCase).IsMatch(e.Node.Text))
                        {
                            query = @"Select o.datePriceList as 'Дата', o.name as 'Организация', p.diametr as 'Диаметр', 
p.tolshina as 'Толщина', p.Marka as 'Марка стали', p.Obem as 'Объем (т.)', p.price as 'Цена', p.standart as 'Стандарт', p.Dlina as 'Метраж', 
p.Type as 'Тип', p.Primech as 'Примечание' from Organization o join  Product p on o.id_organization=p.id_organization left join Sklad s on 
o.id_organization=s.id_organization where p.Name = '" + e.Node.Text + @"'";
                            dataGridView2.Columns.Add("o.datePriceList", "Дата");
                            dataGridView2.Columns.Add("o.name", "Организация");
                            dataGridView2.Columns.Add("p.diametr", "Диаметр");
                            dataGridView2.Columns.Add("p.tolshina", "Толщина");
                            dataGridView2.Columns.Add("p.Marka", "Марка стали");
                            dataGridView2.Columns.Add("p.Obem", "Объем");
                            dataGridView2.Columns.Add("p.price", "Цена");
                            dataGridView2.Columns.Add("p.standart", "Стандарт");
                            dataGridView2.Columns.Add("p.Dlina", "Метраж");
                            dataGridView2.Columns.Add("p.Type", "Тип");
                            dataGridView2.Columns.Add("p.Primech", "Примечание");
                            dataGridView2.Rows.Add();
                            fdv.Variant = 1;
                            fdv.e_Node_Text = e.Node.Text;
                        }
                        else
                        {
                            query = @"Select o.datePriceList as 'Дата', o.name as 'Организация', p.diametr as 'Размер', 
p.tolshina as 'Ширина', p.Marka as 'Марка стали', p.Obem as 'Объем (т.)', p.price as 'Цена', p.standart as 'Стандарт', p.Dlina as 'Длина', 
p.Type as 'Тип', p.Primech as 'Примечание' from Organization o join  Product p on o.id_organization=p.id_organization left join Sklad s on 
o.id_organization=s.id_organization where p.Name = '" + e.Node.Text + @"'";
                            dataGridView2.Columns.Add("o.datePriceList", "Дата");
                            dataGridView2.Columns.Add("o.name", "Организация");
                            dataGridView2.Columns.Add("p.diametr", "Высота (толщина)");
                            dataGridView2.Columns.Add("p.tolshina", "Ширина");
                            dataGridView2.Columns.Add("p.Marka", "Марка стали");
                            dataGridView2.Columns.Add("p.Obem", "Объем");
                            dataGridView2.Columns.Add("p.price", "Цена");
                            dataGridView2.Columns.Add("p.standart", "Стандарт");
                            dataGridView2.Columns.Add("p.Dlina", "Длина");
                            dataGridView2.Columns.Add("p.Type", "Тип");
                            dataGridView2.Columns.Add("p.Primech", "Примечание");
                            dataGridView2.Rows.Add();
                            fdv.Variant = 2;
                            fdv.e_Node_Text = e.Node.Text;
                        }
                        if (toolStripComboBoxCity.SelectedItem.ToString() != "Все")
                            query += " and o.City='" + toolStripComboBoxCity.SelectedItem.ToString() + "'";
                        query += " order by p.diametr, p.Tolshina";
                        toolStatus.Text = "Выполняется запрос";
                        dataGridView2.ReadOnly= true;
                        fdv.City = toolStripComboBoxCity.SelectedItem.ToString();
                        fdv.Conn = conn;
                        fdv.Query = query;
						fdv.TabName = e.Node.Text;
						//if (busy && !_cancelFillDataGrid) { _cancelFillDataGrid = true; busy = false; System.Threading.Thread.Sleep(1000); }
						//else
						//{
						toolStripBtnStop.Visible = true;
                        System.Threading.Thread thread;
                        if (toolStripComboBox1.SelectedItem.ToString() == "1")
                        {
                            thread = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(FillDataGrid));
                            thread.Start(fdv);
                        }
                        //}

                        //adapter = new SqlDataAdapter(query, conn);
                        //adapter.Fill(dt);
                        //bind.DataSource = dt;
                        //dataGridView1.DataSource = bind;
                    }
                }
                else
                {
                    if (new Regex(@"труб|армату|круг|тройник", RegexOptions.IgnoreCase).IsMatch(e.Node.Parent.Text))
                    {
                        query = @"Select o.datePriceList as 'Дата', o.name as 'Организация', p.diametr as 'Диаметр', 
p.tolshina as 'Толщина', p.Marka as 'Марка стали', p.Obem as 'Объем (т.)', p.price as 'Цена', p.standart as 'Стандарт', p.Dlina as 'Метраж', 
p.Type as 'Тип', p.Primech as 'Примечание' from Organization o join  Product p on o.id_organization=p.id_organization left join Sklad s on 
o.id_organization=s.id_organization where p.Name = '" + e.Node.Parent.Text + "' and p.Type = '" + e.Node.Text + "'";
                        dataGridView2.Columns.Add("o.datePriceList", "Дата");
                        dataGridView2.Columns.Add("o.name", "Организация");
                        dataGridView2.Columns.Add("p.diametr", "Диаметр");
                        dataGridView2.Columns.Add("p.tolshina", "Толщина");
                        dataGridView2.Columns.Add("p.Marka", "Марка стали");
                        dataGridView2.Columns.Add("p.Obem", "Объем");
                        dataGridView2.Columns.Add("p.price", "Цена");
                        dataGridView2.Columns.Add("p.standart", "Стандарт");
                        dataGridView2.Columns.Add("p.Dlina", "Метраж");
                        dataGridView2.Columns.Add("p.Type", "Тип");
                        dataGridView2.Columns.Add("p.Primech", "Примечание");
                        dataGridView2.Rows.Add();
                        fdv.Variant = 3;
                        fdv.e_Node_Parent_Text = e.Node.Parent.Text;
                        fdv.e_Node_Text = e.Node.Text;
                    }
                    else
                    {
                        query = @"Select o.datePriceList as 'Дата', o.name as 'Организация', p.diametr as 'Размер', 
p.tolshina as 'Ширина', p.Marka as 'Марка стали', p.Obem as 'Объем (т.)', p.price as 'Цена', p.standart as 'Стандарт', p.Dlina as 'Длина', 
p.Type as 'Тип', p.Primech as 'Примечание' from Organization o join  Product p on o.id_organization=p.id_organization left join Sklad s on 
o.id_organization=s.id_organization where p.Name = '" + e.Node.Parent.Text + "' and p.Type = '" + e.Node.Text + "'";
                        dataGridView2.Columns.Add("o.datePriceList", "Дата");
                        dataGridView2.Columns.Add("o.name", "Организация");
                        dataGridView2.Columns.Add("p.diametr", "Высота (толщина)");
                        dataGridView2.Columns.Add("p.tolshina", "Ширина");
                        dataGridView2.Columns.Add("p.Marka", "Марка стали");
                        dataGridView2.Columns.Add("p.Obem", "Объем");
                        dataGridView2.Columns.Add("p.price", "Цена");
                        dataGridView2.Columns.Add("p.standart", "Стандарт");
                        dataGridView2.Columns.Add("p.Dlina", "Длина");
                        dataGridView2.Columns.Add("p.Type", "Тип");
                        dataGridView2.Columns.Add("p.Primech", "Примечание");
                        dataGridView2.Rows.Add();
                        fdv.Variant = 4;
                        fdv.e_Node_Parent_Text = e.Node.Parent.Text;
                        fdv.e_Node_Text = e.Node.Text;
                    }
                    if (toolStripComboBoxCity.SelectedItem.ToString() != "Все")
                        query += " and o.City='" + toolStripComboBoxCity.SelectedItem.ToString() + "'";
                    query += " order by p.diametr, p.Tolshina";
                    adapter = new SqlDataAdapter(query, conn);
                    toolStatus.Text = "Выполняется запрос";
                    dataGridView2.ReadOnly= true;
                    fdv.City = toolStripComboBoxCity.SelectedItem.ToString();
                    fdv.Conn = conn;
                    fdv.Query = query;
					fdv.TabName = e.Node.Parent.Text;
                    
                    //if (busy && !_cancelFillDataGrid) { _cancelFillDataGrid = true; busy = false; System.Threading.Thread.Sleep(1000); }
                    //else
                    //{
                    toolStripBtnStop.Visible = true;
                    System.Threading.Thread thread;
                    if (toolStripComboBox1.SelectedItem.ToString() == "1")
                    {
                        thread = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(FillDataGrid));
                        thread.Start(fdv);
                    }
                    else if (toolStripComboBox1.SelectedItem.ToString() == "2")
                    {
                        thread = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(FillDataGridAsync));
                        thread.Start(fdv);
                    }
                    else FillDataGridSimple(fdv);
                    
                    //}

                    //FillDataGrid(conn, query);
                    //adapter.Fill(dt);
                    //bind.DataSource = dt;
                    //dataGridView1.DataSource = bind;
                }

                if (conn.State == ConnectionState.Open) conn.Close();
                //dataGridView1.Sort(dataGridView1.Columns[0]);
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                if (isOpenSqlConnection) conn.Close();
                toolStatus.Text = "Запрос не выполнен";
                dataGridView2.ReadOnly= false;
            }

        }

        public void FillDataGrid(Object fillDataGridView)
        {
			if (!busy && !_cancelFillDataGrid)
				try
				{
                    DGV_Enabled_False();
					busy = true;
					FillDataGridView fdv = new FillDataGridView();
					fdv = (FillDataGridView)fillDataGridView;
					SqlConnection conn = fdv.Conn;
					string query = fdv.Query;
					string query1 = query.Replace("Select", "Select top 1");
					int MaxForProgressBar = 0;
					if (conn.State == ConnectionState.Closed) conn.Open();
					SqlDataAdapter sda = new SqlDataAdapter(query1, conn);
					DataTable dt = new DataTable();
					sda.Fill(dt);
					string[] ColHeaders = new string[dt.Columns.Count];
					for (int i = 0; i < dt.Columns.Count; i++)
						ColHeaders[i] = dt.Columns[i].ColumnName;
					ClearDataGridView();
					string replace = new Regex(@"Select[.'\s\d\t\r\n\w,\.\(\)=]+from", RegexOptions.IgnoreCase).Replace(query, "select count(p.Name) from");
					replace = new Regex(@"order.*", RegexOptions.IgnoreCase).Replace(replace, "");
					SqlCommand comm = new SqlCommand(replace, conn);
					if (conn.State == ConnectionState.Closed) conn.Open();
					SqlDataReader reader = comm.ExecuteReader();
					while (reader.Read()) MaxForProgressBar = reader.GetInt32(0);
					if (!reader.IsClosed) reader.Close();
					comm = null;
					AddColumnsDataGridView(ColHeaders, MaxForProgressBar);
					comm = new SqlCommand(query, conn);
					if (conn.State == ConnectionState.Closed) conn.Open();
					reader = comm.ExecuteReader();
					while (reader.Read())
					{
						if (_cancelFillDataGrid) break;
						Object[] row = new Object[11];
						reader.GetValues(row);
						AddRowDataGridView(row);
						if (reader.IsClosed) break;
					}
					Action action = () =>
					{
                        //dtProduct.DefaultView.Sort = "Диаметр ASC, Толщина ASC";
                        (tabControl1.TabPages[fdv.TabName].Controls["dgv_" + fdv.TabName] as DataGridView).DataSource = dtProduct;
                        //(tabControl1.TabPages[fdv.TabName].Controls["dgv_" + fdv.TabName] as DataGridView)

                    };
					Invoke(action);
					if (!reader.IsClosed)
						reader.Close();
					SetLabel();
					if (conn.State != ConnectionState.Closed) conn.Close();
					_cancelFillDataGrid = false;
					busy = false;
				}
				catch (Exception ex) { busy = false; _cancelFillDataGrid = false; MessageBox.Show(ex.ToString()); }
        }

        private void dataGridView1_SortCompare(object sender,
        DataGridViewSortCompareEventArgs e)
        {
            // Try to sort based on the cells in the current column.
            e.SortResult = System.String.Compare(
                dataGridView1.Rows[e.RowIndex1].Cells[3].Value.ToString(),
                    dataGridView1.Rows[e.RowIndex2].Cells[3].Value.ToString());
            //e.CellValue1.ToString(), e.CellValue2.ToString());

            // If the cells are equal, sort based on the ID column.
            if (e.SortResult == 0 && e.Column.Name != "Толщина")
            {
                e.SortResult = System.String.Compare(
                    dataGridView1.Rows[e.RowIndex1].Cells[4].Value.ToString(),
                    dataGridView1.Rows[e.RowIndex2].Cells[4].Value.ToString());
            }
            e.Handled = true;
        }

        private void FillDataGridSimple(Object fillDataGridView)
        {
            conn = new SqlConnection(sqlConnString);
            if (conn.State == ConnectionState.Closed) conn.Open();
            FillDataGridView fdv = (FillDataGridView)fillDataGridView;
            SqlCommand comm = new SqlCommand();
            switch (fdv.Variant)
            {
                case 1:
                    comm = new SqlCommand("get_this_page_1", conn);
                    break;
                case 2:
                    comm = new SqlCommand("get_this_page_2", conn);
                    break;
                case 3:
                    comm = new SqlCommand("get_this_page_3", conn);
                    break;
                case 4:
                    comm = new SqlCommand("get_this_page_4", conn);
                    break;
                default:
                    comm = new SqlCommand("get_this_page_1", conn);
                    break;
            };
            comm.CommandType = CommandType.StoredProcedure;
            comm.Parameters.AddWithValue("@rec_per_page", "500000");
            comm.Parameters.AddWithValue("@page_num", fdv.NumbOfThread);
            comm.Parameters.AddWithValue("@e_Node_Text", fdv.e_Node_Text);
            if (fdv.Variant > 2)
                comm.Parameters.AddWithValue("@e_Node_Parent_Text", fdv.e_Node_Parent_Text);
            comm.Parameters.AddWithValue("@City", fdv.City);
            SqlDataAdapter sdar = new SqlDataAdapter(comm);
            dataTable.Clear();
            sdar.Fill(dtProduct);
            dataGridView1.DataSource = dataTable;
        }

        public void FillDataGridAsync(Object fillDataGridView)
        {
            _cancelFillDataGrid = false;
            if (!busy && !_cancelFillDataGrid)
                try
                {
                    int Otnoshenie = 1; //отношение количества строк в БД к количеству потоков
                    int CountRows = 1; //начальное значение ID_Product
                    busy = true;
                    FillDataGridView fdv = new FillDataGridView();
                    fdv = (FillDataGridView)fillDataGridView;
                    SqlConnection conn = fdv.Conn;
                    string query = fdv.Query;
                    string query1 = query.Replace("Select", "Select top 1");
                    int MaxForProgressBar = 0;
                    if (conn.State == ConnectionState.Closed) conn.Open();
                    SqlDataAdapter sda = new SqlDataAdapter(query1, conn);
                    DataTable dt = new DataTable();
                    sda.Fill(dt);
                    string[] ColHeaders = new string[dt.Columns.Count];
                    for (int i = 0; i < dt.Columns.Count; i++)
                        ColHeaders[i] = dt.Columns[i].ColumnName;
                    ClearDataGridView();
                    
                    //найти количество строк в БД
                    string replace = new Regex(@"Select[.'\s\d\t\r\n\w,\.\(\)=]+from", RegexOptions.IgnoreCase).Replace(query, "select count(p.Name) from");
                    replace = new Regex(@"order.*", RegexOptions.IgnoreCase).Replace(replace, "");
                    SqlCommand comm = new SqlCommand(replace, conn);
                    if (conn.State == ConnectionState.Closed) conn.Open();
                    SqlDataReader reader = comm.ExecuteReader();
                    while (reader.Read()) {
                        MaxForProgressBar = reader.GetInt32(0);
                    }
                    if (!reader.IsClosed) reader.Close();
                    AddColumnsDataGridView(ColHeaders, MaxForProgressBar);

                    Otnoshenie = MaxForProgressBar / NumbOfThreads;
                    if (!reader.IsClosed) reader.Close();

                    //Цикл создания потоков
                    for (int countThreads = 0; countThreads < NumbOfThreads; countThreads++)
                    {
                        if (!_cancelFillDataGrid)
                        {
                            CountRows = Otnoshenie;
                            if (countThreads == NumbOfThreads - 1)
                            {
                                fdv.FutureNotBusy = true;
                            }
                            fdv.CountRows = CountRows;
                            fdv.NumbOfThread = countThreads;
                            System.Threading.Thread thread = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(ReadSQLDataAsync));
                            thread.Start(fdv);
                        }
                    }

                }
                catch (Exception ex) { busy = false; _cancelFillDataGrid = false; MessageBox.Show(ex.ToString()); }
        }

        public SqlDataReader[] dataReaders;
        public SqlConnection[] sqlConnections;
        public string[] querys;
        public SqlCommand[] sqlCommands;

        public void ReadSQLDataAsync(Object fillDataGridView)
        {
            if (!_cancelFillDataGrid)
                try
                {
                    DGV_Enabled_False();
                    FillDataGridView fdgv = new FillDataGridView();
                    fdgv = (FillDataGridView)fillDataGridView;
                    int currNumbOfThread = fdgv.NumbOfThread;
                    SqlConnection conn = fdgv.Conn;
                    conn = new SqlConnection(sqlConnString); 
                    SqlCommand comm;
                    switch (fdgv.Variant)
                    {
                        case 1:
                            comm = new SqlCommand("get_this_page_1", conn);
                            break;
                        case 2:
                            comm = new SqlCommand("get_this_page_2", conn);
                            break;
                        case 3:
                            comm = new SqlCommand("get_this_page_3", conn);
                            break;
                        case 4:
                            comm = new SqlCommand("get_this_page_4", conn);
                            break;
                        default:
                            comm = new SqlCommand("get_this_page_1", conn);
                            break;
                    };
                    comm.CommandType = CommandType.StoredProcedure;
                    comm.Parameters.AddWithValue("@rec_per_page", fdgv.CountRows);
                    comm.Parameters.AddWithValue("@page_num", fdgv.NumbOfThread);
                    comm.Parameters.AddWithValue("@e_Node_Text", fdgv.e_Node_Text);
                    if(fdgv.Variant > 2)
                        comm.Parameters.AddWithValue("@e_Node_Parent_Text", fdgv.e_Node_Parent_Text);
                    comm.Parameters.AddWithValue("@City", fdgv.City);
                    if (conn.State == ConnectionState.Closed) conn.Open();
                    SqlDataReader reader = comm.ExecuteReader();
                    int IsAdded = 0;
                    while (reader.Read())
                    {
                        if (_cancelFillDataGrid) break;
                        Object[] row = new Object[11]; //может потом использовать ColHeaders.Count
                        reader.GetValues(row);
                        IsAdded = AddRowDataGridViewAsync(row, fdgv.CountRows);
                        if (reader.IsClosed) break;
                    }
                    if (!reader.IsClosed)
                        reader.Close();

                    SetLabel();
                    //Action action = () =>
                    //    {
                    //        toolStripLabel4.Text = "Строк в таблице: " + dataGridView1.Rows.Count.ToString();
                    //    };
                    //Invoke(action);
                    if (conn.State != ConnectionState.Closed) conn.Close();
                    _cancelFillDataGrid = false;
                    if (fdgv.FutureNotBusy) { busy = false; fdgv.FutureNotBusy = false; }
                }
                catch (Exception ex) { busy = false; _cancelFillDataGrid = false; MessageBox.Show(ex.ToString()); }
        }

        private void ClearDataGridView()
        {
            Action action = () =>
            {
                dataGridView1.DataSource = null;
				dataGridView1.Columns.Clear();
				//(tabControl1.SelectedTab.Controls["dgv_" + tabControl1.SelectedTab.Name] as DataGridView).DataSource = null;
				//(tabControl1.SelectedTab.Controls["dgv_" + tabControl1.SelectedTab.Name] as DataGridView).Columns.Clear();
				dtProduct.Clear();
				dtProduct.Columns.Clear();
				toolStripProgressBar1.Value = 0;
            };
            Invoke(action);
        }
        private void AddColumnsDataGridView(string[] ColHeaders, int MaxProgressBar)
        {
            Action action = () =>
                {
                    toolStripProgressBar1.Maximum = MaxProgressBar;
					foreach (string s in ColHeaders)
						//dataGridView1.Columns.Add(s, s);
						dtProduct.Columns.Add(s);
                };
            Invoke(action);
        }
        
        private void AddRowDataGridView(Object[] row)
        {
            Action action = () =>
                {
					//dataGridView1.Rows.Add(row);
					dtProduct.Rows.Add(row);
                    if (toolStripProgressBar1.Value >= toolStripProgressBar1.Maximum) toolStripProgressBar1.Value = toolStripProgressBar1.Maximum;
                    else
                        toolStripProgressBar1.Value++;
                };
            Invoke(action);
        }
        DataTable dataTable = new DataTable();
        private int AddRowDataGridViewAsync(Object[] row, int CountRows)
        {
            Action action = () =>
            {
                dataGridView1.Rows.Add(row);
                if (toolStripProgressBar1.Value >= toolStripProgressBar1.Maximum) toolStripProgressBar1.Value = toolStripProgressBar1.Maximum;
                else
                    toolStripProgressBar1.Value++;
                //dataGridView1.DataSource = dataTable;
            };
            Invoke(action);
            return 1;
        }
        private void DGV_Enabled_False()
        {
            Action action = () =>
            {
                dataGridView2.Enabled = false;
            };
            Invoke(action);
        }
        private void SetLabel()
        {
            Action action = () =>
                {
                    toolStatus.Text = "Запрос успешно выполнен";
                    dataGridView2.Enabled= true;
                    toolStripProgressBar1.Value = 0;
                    busy = false;
                };
            Invoke(action);
        }

        private void toolStripBtnStop_Click(object sender, EventArgs e)
        {
            if (busy && !_cancelFillDataGrid) { _cancelFillDataGrid = true; busy = false; System.Threading.Thread.Sleep(1000); toolStripBtnStop.Visible = false; }
        }

        private void SearchDiam()
        {
            try
            {
                _cancelFillDataGrid = true;
                dt.Clear();
                dt.Columns.Clear();
				//dtProduct = new DataTable();
                if (conn.State == ConnectionState.Closed) conn.Open();
                isOpenSqlConnection = true;
                string query = "";
                string temp = "";
                SqlDataAdapter adapter;
                var paramList = new List<string>();
                bool isNoEmpty = false;
                bool isGetQuery = false;
                if (treeView1.SelectedNode.Parent == null)//Если родитель равенн Нулль, то значит родителей нет, а значит - узел корневой
                {
                    //if (treeView1.SelectedNode.Nodes.Count < 1)
                    {

                        if (new Regex(@"труб|армату|круг|тройник", RegexOptions.IgnoreCase).IsMatch(treeView1.SelectedNode.Text))
                        {
                            query = @"Select o.datePriceList as 'Дата', o.name as 'Организация', p.diametr as 'Диаметр', 
p.tolshina as 'Толщина', p.Marka as 'Марка стали', p.Obem as 'Объем (т.)', p.price as 'Цена', p.standart as 'Стандарт', p.Dlina as 'Метраж', 
p.Type as 'Тип', p.Primech as 'Примечание' from Organization o join  Product p on o.id_organization=p.id_organization left join Sklad s on 
o.id_organization=s.id_organization where p.Name = '" + tabControl1.SelectedTab.Name + "' and ";//+ treeView1.SelectedNode.Text + "' and ";
						}
                        else
                        {
                            query = @"Select o.datePriceList as 'Дата', o.name as 'Организация', p.diametr as 'Высота (толщина)', 
p.tolshina as 'Ширина', p.Marka as 'Марка стали', p.Obem as 'Объем (т.)', p.price as 'Цена', p.standart as 'Стандарт', p.Dlina as 'Длина', 
p.Type as 'Тип', p.Primech as 'Примечание' from Organization o join  Product p on o.id_organization=p.id_organization left join Sklad s on 
o.id_organization=s.id_organization where p.Name = '" + tabControl1.SelectedTab.Name + "' and ";
                        }
                    }
                }
                else
                {
                    if (new Regex(@"труб|армату|круг|тройник", RegexOptions.IgnoreCase).IsMatch(treeView1.SelectedNode.Parent.Text))
                    {
                        query = @"Select o.datePriceList as 'Дата', o.name as 'Организация', p.diametr as 'Диаметр', 
p.tolshina as 'Толщина', p.Marka as 'Марка стали', p.Obem as 'Объем (т.)', p.price as 'Цена', p.standart as 'Стандарт', p.Dlina as 'Метраж', 
p.Type as 'Тип', p.Primech as 'Примечание' from Organization o join  Product p on o.id_organization=p.id_organization left join Sklad s on 
o.id_organization=s.id_organization where p.Name = '" + tabControl1.SelectedTab.Name + "' and p.Type = '" + treeView1.SelectedNode.Text + "' and ";
                    }
                    else
                    {
                        query = @"Select o.datePriceList as 'Дата', o.name as 'Организация', p.diametr as 'Высота (толщина)', 
p.tolshina as 'Ширина', p.Marka as 'Марка стали', p.Obem as 'Объем (т.)', p.price as 'Цена', p.standart as 'Стандарт', p.Dlina as 'Длина', 
p.Type as 'Тип', p.Primech as 'Примечание' from Organization o join  Product p on o.id_organization=p.id_organization left join Sklad s on 
o.id_organization=s.id_organization where p.Name = '" + tabControl1.SelectedTab.Name + "' and p.Type = '" + treeView1.SelectedNode.Text + "' and ";
                    }
                }
                for (int i = 0; i < dataGridView2.Columns.Count; i++)
                {
                    if (dataGridView2.Rows.Count < 1) break;
                    if (dataGridView2.Rows[0].Cells[i].Value != null)
                        temp = dataGridView2.Rows[0].Cells[i].Value.ToString().Trim();
                    if (temp != "")
                    {
                        string colName = dataGridView2.Columns[i].Name.ToString();
                        if (colName=="o.name" || colName=="p.Primech" || colName == "p.standart" || colName == "p.Type")
                            paramList.Add( colName + " like '%" + temp + "%' and ");
                        else
                            paramList.Add(colName + " = '" + temp + "' and ");
                        isNoEmpty = true;
                        isGetQuery = true;
                    }
                    temp = "";
                    if (isNoEmpty)
                    {
                        foreach (string s in paramList)
                        {
                            query += s;
                            isNoEmpty = false;
                        }
                    }
                }
                if (isGetQuery)
                {
                    query = query.Remove(query.Length - 5, 5);
                    query += " order by p.Diametr, p.tolshina";
                    isNoEmpty = false;
                    isGetQuery = true;
                    adapter = new SqlDataAdapter(query, conn);
                    toolStatus.Text = "Выполняется запрос";
                    FillDataGridView fdv = new FillDataGridView();
                    fdv.City = toolStripComboBoxCity.SelectedItem.ToString();
                    fdv.Conn = conn;
                    fdv.Query = query;
                    fdv.TabName = tabControl1.SelectedTab.Name;
                    dtProduct = new DataTable();

                    toolStripBtnStop.Visible = true;
                    _cancelFillDataGrid = false;
                    System.Threading.Thread thread = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(FillDataGrid));
                    thread.Start(fdv);


                }



                if (conn.State == ConnectionState.Open) conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                if (isOpenSqlConnection) conn.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SearchDiam();
        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            connectAndUpdate();
        }

        private void выходToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void textBoxSortDiametr_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SearchDiam();
            }
        }

        private void MainForm_SizeChanged(object sender, EventArgs e)
        {
			tabControl1.Width = this.Width - tabControl1.Location.X - 22;
			tabControl1.Height = this.Height - tabControl1.Location.Y - 50;

			//dataGridView1.Width = this.Width - dataGridView1.Location.X - 22;
            dataGridView2.Width = this.Width - dataGridView2.Location.X - 22;
            //groupBox1.Width = this.Width - groupBox1.Location.X - 22;
            //dataGridView1.Height = this.Height - dataGridView1.Location.Y - 50;
            treeView1.Height = this.Height - treeView1.Location.Y - 50;
        }

        private void panel1_MouseClick(object sender, MouseEventArgs e)
        {
            panel1.Visible = false;
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //panel1.Location.X = Cursor.Position.X;
            //panel1.Location.Y = Cursor.Position.Y;
            textBoxOrgName.Text = "";
            textBoxTelOrg.Text = "";
            listViewAdrSklad.Items.Clear();
            listViewManager.Items.Clear();
            string norg;
            try
            {
				DataGridView dgv = (DataGridView)sender;
                norg = dgv.Rows[e.RowIndex].Cells["Организация"].Value.ToString();
                textBoxOrgName.Text = norg;
                
                if (conn.State == ConnectionState.Closed) { conn.Open(); isOpenSqlConnection = true; }
                conn.Close();
                string query;
                SqlDataAdapter adapter;
                DataTable dtdt = new DataTable();
                query = @"Select Telefon from Organization where [Name] = '"+norg+"';";
                adapter = new SqlDataAdapter(query, conn);
                toolStatus.Text = "Выполняется запрос";
                dataGridView2.ReadOnly= true;
                adapter.Fill(dtdt);
                if (dtdt.Rows.Count > 0)
                {
                    textBoxTelOrg.Text = dtdt.Rows[0][0].ToString();
                }
                dtdt.Clear();
                query = @"Select Email as 'mail' from Organization where [Name] = '" + norg + "';";
                adapter = new SqlDataAdapter(query, conn);
                toolStatus.Text = "Выполняется запрос";
                dataGridView2.ReadOnly= true;
                adapter.Fill(dtdt);
                if (dtdt.Rows.Count > 0)
                {
                    textBoxEmailOrg.Text = dtdt.Rows[0]["mail"].ToString();
                }
                dtdt.Clear();
                query = @"Select City as 'City' from Organization where [Name] = '" + norg + "';";
                adapter = new SqlDataAdapter(query, conn);
                toolStatus.Text = "Выполняется запрос";
                dataGridView2.ReadOnly= true;
                adapter.Fill(dtdt);
                if (dtdt.Rows.Count > 0)
                {
                    textBoxOrgCity.Text = dtdt.Rows[0]["City"].ToString();
                }

                dtdt.Clear();
                query = @"Select distinct m.[Name] as 'name', m.Telefon as 'tel' from Manager m where ID_Organization = (select ID_Organization from Organization where [Name] = '" + norg + "');";
                adapter = new SqlDataAdapter(query, conn);
                toolStatus.Text = "Выполняется запрос";
                dataGridView2.ReadOnly= true;
                adapter.Fill(dtdt);
                foreach (DataRow row in dtdt.Rows)
                {
                    if (row["name"].ToString() != "")
                    {
                        ListViewItem lvi = new ListViewItem(row["name"].ToString());
                        lvi.SubItems.Add(row["tel"].ToString());
                        listViewManager.Items.Add(lvi);
                    }
                }

                dtdt.Clear();
                query = @"Select distinct s.[Adress] as 'adr' from Sklad s where ID_Organization = (select ID_Organization from Organization where [Name] = '" + norg + "');";
                adapter = new SqlDataAdapter(query, conn);
                toolStatus.Text = "Выполняется запрос";
                dataGridView2.ReadOnly= true;
                adapter.Fill(dtdt);
                foreach (DataRow row in dtdt.Rows)
                {
                    if (row["adr"].ToString() != "")
                    {
                        ListViewItem lvi = new ListViewItem(row["adr"].ToString());
                        listViewAdrSklad.Items.Add(lvi);
                    }
                }

                toolStatus.Text = "Запрос выполнен";
                dataGridView2.ReadOnly= false;
                isOpenSqlConnection = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            panel1.Visible = true;
        }

        private void менеджерыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Managers managers = new Managers();
            managers.sqlConnectionString = sqlConnString;
            managers.ShowDialog();
        }

        private void складыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Sklad sk = new Sklad();
            sk.sqlConnectionString = sqlConnString;
            sk.ShowDialog();
        }

        private void организацииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Organizations org = new Organizations();
            org.sqlConnectionString = sqlConnString;
            org.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView2.Columns.Count; i++)
                dataGridView2.Rows[0].Cells[i].Value = "";

        }

        private void dataGridView2_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //    SearchDiam();
            switch (e.KeyCode)
            {
                case Keys.Enter:
                    SearchDiam();
                    break;

                case Keys.Delete:
                    dataGridView2.SelectedCells[0].Value = "";
                    break;
            }
        }

        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            settingForm sf = new settingForm(new settingDelegate(GetSettings));
            sf.ShowDialog();
            if (set.Count >= 6)
            {
                sqlConnString = "Server=tcp:" + set[0] + "," + set[2] + "; Database=" + set[3] + "; User ID=" + set[4] + "; Password=" + set[5] + "; Pooling=true;";
                if(set.Count==7)
                    int.TryParse(set[6], out NumbOfThreads);
            }
            else sqlConnString = @"Server=tcp:maks-pc,1433; Database=MetalBase; User ID=metuser; Password=metuser";
        }

        private void редакторСоответсвийToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ManualNameProdRedactor mnp = new ManualNameProdRedactor(sqlConnString);
            mnp.ShowDialog();
        }

        #region чтобы небыло сворачивания/разворачивания treeView при двойном клике
        bool    doubleClicked = false;
        
        private void treeView1_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            if (doubleClicked)
            {
                doubleClicked = false;
                e.Cancel = true;
            }
        }

        private void treeView1_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            if (doubleClicked)
            {
                doubleClicked = false;
                e.Cancel = true;
            }
        }
 
        private void treeView1_MouseDown(object sender, MouseEventArgs e)
        {
            TreeNode node;
            if (e.Button == MouseButtons.Left && e.Clicks==2)
            {
                node = treeView1.GetNodeAt(e.X, e.Y);
                if (node != null)
                {
                      doubleClicked = true;                       
                }
            }
            if (e.Button == MouseButtons.Right && e.Clicks == 1)
            {
                node = treeView1.GetNodeAt(e.X, e.Y);
                if (node != null && node.Parent == null)
                {
                    contextMenuStrip1.Tag = node.Text;
                    contextMenuStrip1.Show(e.X, e.Y+treeView1.Top);
                }
            }
        }
        #endregion

        private void button3_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
        }

        private void toolStripComboBoxCity_SelectedIndexChanged(object sender, EventArgs e)
        {
            connectAndUpdate();
        }

        private void заменитьНазваниеПродукцииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenameProductForm rpf = new RenameProductForm(sqlConnString);
            rpf.ShowDialog();

        }

        private void заменитьНаименованиеПродукцииВБазеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RenameProductForm rpf = new RenameProductForm(sqlConnString, contextMenuStrip1.Tag.ToString());
            contextMenuStrip1.Tag = "";
            rpf.ShowDialog();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void маркиНержавеющейСталиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Forms.NerjMarkForm nerjForm = new Forms.NerjMarkForm(sqlConnString);
            nerjForm.ShowDialog();
        }

		private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
		{
			toolStatus.Text = "Выполняется запрос";
            dataGridView2.ReadOnly = true;
            bool isNotEmpty = false;
			for (int i = 0; i < tabControl1.TabPages.Count; i++)
				if (tabControl1.TabPages[i].Name == e.Node.Text)
				{ isNotEmpty = true; break; }
			if (isNotEmpty)
			{
				readDataFromSql(e);
			 	tabControl1.SelectTab(e.Node.Text);
			}
			else AddTabControl(e);
            dataGridView2.ReadOnly= false;
        }

		private void AddTabControl(TreeNodeMouseClickEventArgs e)
		{
			var tabPage = new TabPage(e.Node.Text + "    ");
			tabPage.Name = e.Node.Text;// + "    ";
			tabPage.Text = e.Node.Text + "    ";
			tabControl1.TabPages.Add(tabPage);
			

			DataGridView dgv = new DataGridView();
			tabControl1.SuspendLayout();
			tabControl1.TabPages[e.Node.Text].Controls.Add(dgv);
			tabControl1.ResumeLayout();
			
			readDataFromSql(e);
			dgv.RowHeadersVisible = false;
			dgv.Name = "dgv_" + e.Node.Text;
			dgv.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
			dgv.Cursor = Cursors.Arrow;
			dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
			//dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
			dgv.RowTemplate.Resizable = DataGridViewTriState.False;
			dgv.MultiSelect = false;
			dgv.ReadOnly = true;
			dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
			dgv.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
			
			dgv.DataSource = dtProduct;
			dgv.Dock = DockStyle.Fill;
			dgv.BringToFront();

			tabControl1.SelectedIndex = tabControl1.TabPages.Count - 1;
		}

		private void tabControl1_DrawItem(object sender, DrawItemEventArgs e)
		{
			var tabPage = this.tabControl1.TabPages[e.Index];
			var tabRect = this.tabControl1.GetTabRect(e.Index);
			var closeImage = Properties.Resources.ClosePage;
			tabRect.Inflate(-2, -2);
			
			//if (e.Index == this.tabControl1.TabCount - 1)
			//{
			//	var addImage = Properties.Resources.AddPage;
			//	e.Graphics.DrawImage(addImage,
			//		tabRect.Left + (tabRect.Width - addImage.Width) / 2,
			//		tabRect.Top + (tabRect.Height - addImage.Height) / 2);
			//}
			//else
			//{

			e.Graphics.DrawImage(closeImage,
					(tabRect.Right - closeImage.Width),
					tabRect.Top + (tabRect.Height - closeImage.Height) / 2 + 2);
			//e.Graphics.ScaleTransform((float)(tabRect.Width + closeImage.Width), (float)2);
			TextRenderer.DrawText(e.Graphics, tabPage.Text, tabPage.Font,
					tabRect, tabPage.ForeColor, TextFormatFlags.Left);
			//}
		}

		private void tabControl1_MouseClick(object sender, MouseEventArgs e)
		{
			//Looping through the controls.
			var closeImage = Properties.Resources.ClosePage;
			for (int i = 0; i < this.tabControl1.TabPages.Count; i++)
			{
				Rectangle r = tabControl1.GetTabRect(i);
				//Getting the position of the "x" mark.
				Rectangle closeButton = new Rectangle(r.Right - 2 - closeImage.Width, r.Top + 2, closeImage.Width, closeImage.Height);
				if (closeButton.Contains(e.Location))
				{
					if (MessageBox.Show("Хотите закрыть эту вкладку?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
					{
						this.tabControl1.TabPages.RemoveAt(i);
						break;
					}
				}
			}
		}

        private void tsBtnPDF_Click(object sender, EventArgs e)
        {
            Form_PDF fpdf = new Form_PDF();
            fpdf.ShowDialog();
        }
    }
}
