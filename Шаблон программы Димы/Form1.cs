using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using ExcelDataReader;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace Шаблон_программы_Димы
{
    public partial class Form1 : Form
    {

        public string comp, nullValue, field, fileLine, wktGeom, geomString;
        string repl = "";
        public double NumberEl, NumberTop, startM, stepM, wellM, stopM;
        long sizefile;
        public string[][] rangeCurve; // весь диапозон кривой включая nullNamber
        public string[] pointsCurve; // точки кривой
        public string[] nullNamber; // значения -9999
        public List<string> dept = new List<string>();
        public List<string> unit = new List<string>();
        public List<string> wktGeomList = new List<string>();
        public List<string> namefile = new List<string>();
        public bool flag = false;
        

        //ID выбранной скважины
        int CurWellID = -1;
        // Лист параметров петрофизики наследованный от класса 
        public static List<ListItemPet> PetParams = new List<ListItemPet>();

        
        public string commStr = "select регион from Регионы";


        public Form1()
        {
            InitializeComponent();
            
            CreateList();

            using (OleDbConnection dbconn = new OleDbConnection(connStr))
            {
                dbconn.Open();
                DataTable dt1 = new DataTable();
                OleDbCommand dbcomm = new OleDbCommand(commStr, dbconn);
                OleDbDataAdapter da = new OleDbDataAdapter(dbcomm);
                da.Fill(dt1);
                comboBox1.DataSource = dt1;
                comboBox1.DisplayMember = "Регион";
                
                
                dbconn.Close();

            }
        }

        // шаг заполнение прогресс бара 
        void timer_tick() 
        {
            progressBar1.Value += 1;
        }

        //Метод нумерации строк в datagrid
        private void numeric_Point_DataGrid(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            if (((DataGridView)sender).Rows[e.RowIndex].HeaderCell.Value ==null)
            {
                ((DataGridView)sender).Rows[e.RowIndex].HeaderCell.Value = (e.RowIndex + 1).ToString();
            }
        }

        // Событие datagrid отвечающее за нумерацию строк 
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e) 
        {
            numeric_Point_DataGrid(sender,e);
        }

        private void dataGridView2_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            numeric_Point_DataGrid(sender, e);
        }

        //
        // Закладка "Забрать файл"
        //

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string[] dataRows = new string[dataGridView1.ColumnCount];
                string s = "";

                for (int i = 0; i < 1; i++)
                {
                    foreach (DataGridViewCell item in dataGridView1.CurrentRow.Cells)
                    {
                        s += item.Value.ToString() + " ";
                    }
                }
                dataRows = s.Split(' ');


                using (OleDbConnection dbconn = new OleDbConnection(connStr))
                {
                    dbconn.Open();
                    OleDbCommand dbcomm = new OleDbCommand($"DECLARE @g geometry = ( select Geom from Каротаж_интервалы where ID_каротажа = '{dataRows[0]}' and начало ='{dataRows[3]}' and конец = '{dataRows[4]}' and sha256 is not null) SELECT @g.BufferWithCurves(0).ToString()", dbconn);
                    OleDbCommand dbcomm2 = new OleDbCommand($"DECLARE @g geometry = ( select Geom from Каротаж_интервалы where ID_каротажа = '{dataRows[0]}' and начало ='{dataRows[3]}' and конец = '{dataRows[4]}') SELECT @g.BufferWithCurves(0).ToString()", dbconn);
                    // Выбирает из базы кривую и представляет её с типом varсhar для записи 
                    using (OleDbDataReader reader = dbcomm.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                wktGeom = reader[0].ToString();
                                if (wktGeom == "") // В таблице каротаж_интервалы есть дублированые записи с одинаковым id и интервалом , но с разным "sha256,user,data", поэтому проверяем wktgeom на наличие кривой, если пусто, то идём по второму запросу  
                                {
                                    OleDbDataReader reader1 = dbcomm2.ExecuteReader();
                                    while (reader1.Read())
                                    {
                                        wktGeom = reader1[0].ToString();
                                    }
                                }
                                wktGeom = wktGeom.Remove(0, 12); //Убирает символы с 0-12, там "LINESTRING ("
                                wktGeom = wktGeom.Remove(wktGeom.Length - 1); // Убирает последний символ, там ")"
                            }
                        }
                    }
                    dbconn.Close();
                }

                string[][] geomData = wktGeom.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(line => line.Split(" ;".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)).ToArray();

                for (int i = 0; i < geomData.Length; i++)
                {
                    geomString += $"{geomData[i][1]} {"",1} {geomData[i][0]}" + '\r';
                }

                string lasText = "~Version information"
                             + '\n' + "VERS.         2.0: CWLS LAS-VERSION 2" + '\n' +
                             "WRAP.          NO: One line per depth step" + '\n' +
                             "~Well information" + '\n' +
                             "#MNEM.UNIT DATA TYPE INFORMATION" + '\n' +
                             "#------------------------------------------------------------" + '\n' +
                             $"STRT.M  {dataRows[3],15}: First depth in file" + '\n' +
                             $"STOP.M  {dataRows[4],15}: Last depth in file" + '\n' +
                             $"STEP.M  {dataRows[2],15}: Depth increment" + '\n' +
                             $"NULL.M  {-9999,15}: Null values" + '\n' +
                             $"COMP.M  {"",15}: COMPANY" + '\n' +
                             $"WELL.   {comboBox3.Text,15}: WELL" + '\n' +
                             $"FLD.    {comboBox2.Text,15}: FIELD" + '\n' +
                             $"LOC.    {"",15}: LOCATION" + '\n' +
                             $"CNTY.   {"",15}: COUNTY" + '\n' +
                             $"STAT.   {"",15}: STATE" + '\n' +
                             $"CTRY.   {"",15}: COUNTRY" + '\n' +
                             $"SRVC.   {"",15}: SERVICE COMPANY" + '\n' +
                             $"DATE.   {"",15}: LOG DATE" + '\n' +
                             $"API.    {"",15}: API NUMBER" + '\n' +
                             "~Curve information" + '\n' +
                             "#MNEM.UNIT         API CODE         CURVE DESCRIPTION" + '\n' +
                             "#------------------------------------------------------------" + '\n' +
                             "DEPT.M  : Depth  curve" + '\n' +
                             $"{dataRows[1]}.         :" + '\n' +
                             "~Parameter information block" + '\n' +
                             "#MNEM.UNIT    VALUE    DESCRIPTION" + '\n' +
                             "#------------------------------------------------------------" + '\n' +
                             "~Other information" + '\n' +
                             "COMMENT1_.        :" + '\n' +
                             "COMMENT2_.        :" + '\n' +
                             "#----------------- REMARKS AREA -----------------------------" + '\n' +
                             "#------------------------------------------------------------" + '\n' +
                             "~ASCII Log Data" + '\n' + geomString;

                SaveFileDialog sf = new SaveFileDialog();
                sf.Filter = "Las файл(*.Las) | *.Las";
                sf.Title = "Сохранение файла";
                if (sf.ShowDialog() == DialogResult.Cancel)
                {
                    return;
                }
                File.WriteAllText(sf.FileName, lasText);
                MessageBox.Show("Сохранено!", "Сохранение");
                geomString = "";

            }
            catch (Exception)
            {
                if (comboBox1.Text == "")
                {
                    MessageBox.Show("Выберите \"Регион\"");
                }
                else if (comboBox2.Text == "")
                {
                    MessageBox.Show("Выберите \"Площадь\"");
                }
                else if (comboBox3.Text == "")
                {
                    MessageBox.Show("Выберите \"Скважину\"");
                }
                else
                {
                    MessageBox.Show("Что-то пошло не так!", "Ошибка");
                }
                return;
                throw;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            comboBox2.Text = "";
            comboBox3.Items.Clear();
            comboBox3.Text = "";
            dataGridView1.Rows.Clear();
            
            using (OleDbConnection dbconn = new OleDbConnection(connStr))
            {
                dbconn.Open();
                DataTable dt2 = new DataTable();
                OleDbCommand dbcomm = new OleDbCommand("SELECT площадь FROM Площади where ID_региона = (select ID_региона from Регионы where регион = '" + comboBox1.Text + "')", dbconn);
                OleDbDataReader reader = dbcomm.ExecuteReader();
                while (reader.Read())
                {
                    comboBox2.Items.Add(reader["площадь"].ToString());
                }
                dbconn.Close();
            }

        }


        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox3.Items.Clear();
            comboBox3.Text = "";
            dataGridView1.Rows.Clear();
            using (OleDbConnection dbconn = new OleDbConnection(connStr))
            {
                dbconn.Open();
                DataTable dt3 = new DataTable();
                OleDbCommand dbcomm = new OleDbCommand($"SELECT площадь, скважина from Площади JOIN Регионы On Площади.ID_региона = Регионы.ID_региона JOIN Скважины ON Скважины.ID_площади = Площади.ID_площади where Регионы.регион = '{comboBox1.Text}' and площадь = '{comboBox2.Text}'", dbconn);
                OleDbDataReader reader = dbcomm.ExecuteReader();
                while (reader.Read())
                {
                    comboBox3.Items.Add(reader["Скважина"].ToString());
                }
                dbconn.Close();
            }
        }

        

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

            dataGridView1.Rows.Clear();
            using (OleDbConnection dbconn = new OleDbConnection(connStr))
            {
                dbconn.Open();
                DataTable dt4 = new DataTable();
                // Получаем ID выбранной скважины
                OleDbCommand dbcomm1 = new OleDbCommand($"SELECT Скважины.ID_скважины from Скважины JOIN Площади on Площади.ID_площади = Скважины.ID_площади WHERE площадь = '{comboBox2.Text}'",dbconn);
                OleDbDataReader reader1 = dbcomm1.ExecuteReader();
                while (reader1.Read())
                {
                    CurWellID = (int)reader1["ID_скважины"];
                }

                OleDbCommand dbcomm = new OleDbCommand($"SELECT Каротаж.ID_каротажа, тип_англ, шаг_глубины, начало, конец, наименование, filename, Каротаж_интервалы.пользователь,Каротаж_интервалы.дата_ввода from Каротаж JOIN Каротаж_типы on Каротаж.ID_тип_каротажа = Каротаж_типы.ID_тип_каротажа JOIN Каротаж_интервалы on Каротаж.ID_каротажа = Каротаж_интервалы.ID_каротажа JOIN Каротаж_группы on Каротаж_типы.ID_группы_каротажа = Каротаж_группы.ID_группы_каротажа JOIN Каротаж_LAS on Каротаж_интервалы.ID_LAS = Каротаж_LAS.ID_LAS WHERE ID_скважины = (SELECT ID_скважины from GIS.dbo.Регион_Площадь_Скважина where регион = '{comboBox1.Text}' and площадь = '{comboBox2.Text}' and скважина = '{comboBox3.Text}')", dbconn);
                OleDbDataReader reader = dbcomm.ExecuteReader();
                while (reader.Read())
                {
                    DataGridViewTextBoxCell Id_каротаж = new DataGridViewTextBoxCell();
                    DataGridViewTextBoxCell type_angl = new DataGridViewTextBoxCell();
                    DataGridViewTextBoxCell stepn = new DataGridViewTextBoxCell();
                    DataGridViewTextBoxCell start = new DataGridViewTextBoxCell();
                    DataGridViewTextBoxCell end = new DataGridViewTextBoxCell();
                    DataGridViewTextBoxCell group = new DataGridViewTextBoxCell();
                    DataGridViewTextBoxCell filename = new DataGridViewTextBoxCell();
                    DataGridViewTextBoxCell user = new DataGridViewTextBoxCell();
                    DataGridViewTextBoxCell dateadd = new DataGridViewTextBoxCell();

                    Id_каротаж.Value = reader["ID_каротажа"].ToString();
                    type_angl.Value = reader["тип_англ"].ToString();
                    stepn.Value = reader["шаг_глубины"].ToString();
                    start.Value = reader["начало"].ToString();
                    end.Value = reader["конец"].ToString();
                    group.Value = reader["наименование"].ToString();
                    filename.Value = reader["filename"].ToString();
                    user.Value = reader["пользователь"].ToString();
                    dateadd.Value = reader["дата_ввода"].ToString();
                    DataGridViewRow row = new DataGridViewRow();
                    row.Cells.AddRange(Id_каротаж, type_angl, stepn, start, end, group, filename, user, dateadd);
                    dataGridView1.Rows.Add(row);
                }
                dbconn.Close();
            }

            RewritePetParams(CurWellID);
        }


        private void CreateList()
        {

            //listView1.Bounds =new Rectangle(new Point (10,10), new Size(500, 300));
            ColumnHeader header1 = new ColumnHeader();   //cоздает колонку
            header1.Text = "Field";                      // текст колонки
            header1.Width = 200;                         //размер колонки
            ColumnHeader header2 = new ColumnHeader();
            header2.Text = "Well";
            ColumnHeader header3 = new ColumnHeader();
            header3.Text = "Type";
            ColumnHeader header4 = new ColumnHeader();
            header4.Text = "Start";
            ColumnHeader header5 = new ColumnHeader();
            header5.Text = "Stop";
            ColumnHeader header6 = new ColumnHeader();
            header6.Text = "Step";
            ColumnHeader header7 = new ColumnHeader();
            header7.Width = 300;                         //размер колонки
            header7.Text = "FileName";

            listView1.Columns.AddRange(new ColumnHeader[] { header1, header2, header3, header4, header5, header6, header7 }); //добавляет несколько колонок в коллекцию

            listView1.View = View.Details;

        }
        //
        // Закладка  "Внести файл"
        //
        private void button2_Click(object sender, EventArgs e) // Открытие файла ласс
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Filter = "Las файл(*.LAS) | *.LAS";
            of.Title = "Выбор файла";
            if (of.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            FileInfo filename = new FileInfo(of.FileName);

            sizefile = filename.Length; // Размер файла 
            progressBar1.Value = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = (int)sizefile; // Задаёт максимум прогресс бара в виде размера файла 

            string encoding = string.Empty;
            using (StreamReader str = new StreamReader(of.FileName,true))
            {
                encoding = str.CurrentEncoding.ToString();
            }
            MessageBox.Show(encoding);

            bool containsfile = namefile.Contains(filename.Name);  // Проверка на уже добавленный файл
            if (containsfile == true)
            {

                MessageBox.Show("Вы уже открывали этот файл!");

            }

            else
            {

                
                using (StreamReader sr = new StreamReader(of.FileName, Encoding.Default))
                {
                    while (!sr.EndOfStream)
                    {
                       
                        comp = sr.ReadLine();

                        if (comp.IndexOf("NULL") == 0 || comp.IndexOf("  NULL.") == 0)
                        {

                            string[] data = comp.Split(new char[] { ' ', ':', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                            

                            if (data.Length > 0)
                            {
                               
                                if (data[1] == ".")
                                {
                                   
                                    nullValue = data[2];
                                }
                                if (data[0] == "NULL.M" || data[0] == "NULL.")
                                {
                                   
                                    nullValue = data[1];
                                }
                            }
                            
                        }

                        if (comp.IndexOf("STRT") == 0 || comp.IndexOf("  STRT.M") == 0) 
                        {
                            string[] data = comp.Split(new char[] { ' ', ':', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                            if (data.Length > 0)
                            {
                                if (data[1] == ".M")
                                {
                                    string qwe = data[2];
                                    startM = Convert.ToDouble(qwe);
                                }
                                if (data[0] == "STRT.M")
                                {
                                    string qwe = data[1];
                                    startM = Convert.ToDouble(qwe);
                                }
                            }


                        }

                        if (comp.IndexOf("STOP") == 0 || comp.IndexOf("  STOP.M") == 0)
                        {
                            string[] data = comp.Split(new char[] { ' ', ':', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                            if (data.Length > 0)
                            {
                                if (data[1] == ".M")
                                {
                                    string qwe = data[2];
                                    stopM = Convert.ToDouble(qwe);
                                }
                                else if (data[0] == "STOP.M")
                                {
                                    string qwe = data[1];
                                    stopM = Convert.ToDouble(qwe);
                                }



                            }
                        }

                        if (comp.IndexOf("STEP") == 0 || comp.IndexOf("  STEP.M") == 0)
                        {
                            string[] data = comp.Split(new char[] { ' ', ':', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                            if (data.Length > 0)
                            {
                                if ((data[1] == ".M"))
                                {
                                    string qwe = data[2];
                                    stepM = Convert.ToDouble(qwe);
                                }
                                if (data[0] == "STEP.M")
                                {
                                    if (data[1] == "Depth")
                                    {

                                    }
                                    else
                                    {
                                        string qwe = data[1];
                                        stepM = Convert.ToDouble(qwe);
                                    }
                                }
                            }
                        }

                        if (comp.IndexOf("WELL") == 0 || comp.IndexOf("  WELL.") == 0)
                        {
                            string[] data = comp.Split(new char[] { ' ', ':', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                            if (data.Length > 0)
                            {
                                if ((data[1] == "."))
                                {
                                    string qwe = data[2];
                                    // если встроке есть что-то кроме цифр, заменяе их на пустое значение
                                    Regex reg = new Regex(@"\D");
                                    string res = reg.Replace(qwe,repl);

                                    wellM = Convert.ToDouble(qwe);
                                    
                                }
                                if (data[0] == "WELL.")
                                {
                                    string qwe = data[1];
                                    // если встроке есть что-то кроме цифр, заменяе их на пустое значение
                                    Regex reg = new Regex(@"\D");
                                    string res = reg.Replace(qwe, repl);

                                    wellM = Convert.ToDouble(res);
                                 
                                }
                            }
                        }

                        if (comp.IndexOf("FLD") == 0 || comp.IndexOf("  FLD .") == 0)
                        {
                            string[] data = comp.Split(new char[] { ' ', ':', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                            if (data.Length > 0)
                            {
                                foreach (var item in data)
                                {
                                    MessageBox.Show(item);
                                }
                                if ((data[1] == "."))
                                {
                                    //string qwe = data[2];
                                    field = data[2];
                                }
                                if (data[0] == "FLD.")
                                {
                                    //string qwe = data[1];
                                    field = data[1];
                                }
                            }
                        }

                        if (comp.IndexOf("DEPT") == 0 || comp.IndexOf("DEPT .M") == 0)
                        {
                            while ((comp.IndexOf("~Parameter") != 0) && (comp.IndexOf("~OTHER") != 0) && (comp.IndexOf("~ASCII") !=0 ))
                            {


                                comp = sr.ReadLine();
                                string[] data = comp.Split(new char[] { ' ', '.', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                                if ((data[0] != "~Parameter") && (data[0] != "~OTHER") && (data[0] != "~ASCII"))
                                {
                                    if (data[1] == ":")
                                    {
                                        unit.Add("Null");
                                    }
                                    else
                                    {
                                        unit.Add(data[1]);
                                    }
                                    dept.Add(data[0]);
                                    
                                }
                                   
                            }

                        }

                        if (flag)
                        {
                            fileLine = comp + '\r' + '\n';
                            while (!sr.EndOfStream)
                            {

                                progressBar1.Value += (int)sizefile % 10; //Задаёт шаг прогресс бара
                                fileLine += sr.ReadLine() + '\r' + '\n';
                                
                            }

                            rangeCurve = fileLine.Split(new char[] { '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries).Select(line => line.Split(" ;".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)).ToArray();
                            flag = false;
                        }

                        if ((comp.IndexOf("~ASCII") == 0) || (comp.IndexOf("~A") == 0) || (comp.IndexOf("~Ascii") == 0))
                        {
                            flag = true;
                        }
                        
                    }
                    
                    
                    
                }
                   
                   nullNamber = new string[rangeCurve.Length];
                   pointsCurve = new string[rangeCurve.Length];
              
                   for (int j = 1; j < dept.Count + 1; j++)
                   {
                    
                       for (int i = 0; i < rangeCurve.Length; i++)
                       {
                           if (rangeCurve[i][j] == nullValue)
                           {
                               continue;
                           }
                           else
                           {
                               nullNamber[i] = rangeCurve[i][0];
                               pointsCurve[i] = rangeCurve[i][j];
                           }
                       }
              
                       nullNamber = nullNamber.Where(x => x != null).ToArray();
                       pointsCurve = pointsCurve.Where(x => x != null).ToArray();
              
                       NumberEl = pointsCurve.Length;
                       NumberEl += 1;
                       NumberTop = NumberEl / 2;
                       // MessageBox.Show("Вершин " + NumberTop + " у скважины " + dept[j - 1]);
                       NumberEl = 0;
                       NumberTop = 0;
              
                       List<string> koordArray = new List<string>(); // Лист для добавления точек кривой
                       List<string> koordArray1 = new List<string>(); // Лист для добавления точек кривой
                       List<string> txt = new List<string>(); // Лист для текстовой записи кривой перед отправкой в базу 
              
                       for (int w = 0; w < pointsCurve.Length; w++)
                       {
                           string koordX = Convert.ToString(pointsCurve[w]);
                           koordArray.Add(koordX);
              
                           string koordY = Convert.ToString(nullNamber[w]);
                           koordArray1.Add(koordY);
                       }
              
                       for (int q = 0; q < pointsCurve.Length; q++)
                       {
                        
                           if (q == 0)
                           {
                               txt.Add(Convert.ToString(koordArray[q]));
                               txt.Add(" ");
                               txt.Add(Convert.ToString(koordArray1[q]));
                           }
                           else
                           {
                               txt.Add(",");
                               txt.Add(Convert.ToString(koordArray[q]));
                               txt.Add(" ");
                               txt.Add(Convert.ToString(koordArray1[q]));
                           }
                       }
              
                       nullNamber = new string[rangeCurve.Length];
                       pointsCurve = new string[rangeCurve.Length];
                       
                       StringBuilder strb = new StringBuilder(); // Поставил вместо string так как быстрее работает
                      
                       foreach (var item in txt)
                       {
                           strb.Append(item + '\r' + '\n');
                       }
                    
                       wktGeomList.Add(strb.ToString());
              
                       strb.Remove(0,strb.Length);
                       txt.Clear();
                       koordArray.Clear();
                       koordArray1.Clear();
                       
                   }
                
              
              
                   for (int i = 0; i < dept.Count; i++)
                   {

                      ListViewItem row1Cell = new ListViewItem();                                      // создает ячейку в первой строке
                      ListViewItem.ListViewSubItem item2Row1Cell = new ListViewItem.ListViewSubItem(); // создаёт ячейку в  строке в след  колонке
                      ListViewItem.ListViewSubItem item3Row1Cell = new ListViewItem.ListViewSubItem();
                      ListViewItem.ListViewSubItem item4Row1Cell = new ListViewItem.ListViewSubItem();
                      ListViewItem.ListViewSubItem item5Row1Cell = new ListViewItem.ListViewSubItem();
                      ListViewItem.ListViewSubItem item6Row1Cell = new ListViewItem.ListViewSubItem();
                      ListViewItem.ListViewSubItem item7Row1Cell = new ListViewItem.ListViewSubItem();
                      row1Cell.Text = field;
                      item2Row1Cell.Text = wellM.ToString();
                      item3Row1Cell.Text = dept[i];
                      item4Row1Cell.Text = startM.ToString();
                      item5Row1Cell.Text = stopM.ToString();
                      item6Row1Cell.Text = stepM.ToString();
                      item7Row1Cell.Text = filename.Name;
                      
                      row1Cell.SubItems.AddRange(new ListViewItem.ListViewSubItem[] { item2Row1Cell, item3Row1Cell, item4Row1Cell, item5Row1Cell, item6Row1Cell, item7Row1Cell }); // второй вариант добавления сразу несколько ячеек по строке в колонки 
                      
                      listView1.Items.AddRange(new ListViewItem[] { row1Cell }); // добавляет несколько ячеек в коллекцию 
                      listView1.CheckBoxes = true; // Доабвляет флажок к строкам
                      //row1Cell.Checked = true; //Устанавливает начальное значение флага
                      listView1.View = View.Details;
                    
                    

                   }
                    dept.Clear();
                
                    // }

                    progressBar1.Maximum = 100; // Задаёт максимальное значение прогресс бара 100
                    progressBar1.Value = 100;   // Заполняет до макс значения
                
            }     
                    namefile.Add(filename.Name); // Добавляет имя файла в список имен файлов

            
        }



        // Занесение в базу выбранного каротажа 
        private void button3_Click(object sender, EventArgs e) 
        {
            //Строка подключения к тестовой БД, потом сдлеать к основе
            string testConnStr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=...;Data Source=...";
            

            // string asd = "";
            ListView.CheckedListViewItemCollection checkitem = listView1.CheckedItems;
            List<string> subItemsText = new List<string>(); // Лист для текста из ячеек 
            foreach (ListViewItem item in checkitem)
            {


                for (int i = 0; i < 6; i++)
                {
                    subItemsText.Add(item.SubItems[i].Text);

                    // MessageBox.Show(item.SubItems[i].Text);
                    //asd += item.SubItems[i].Text;
                    
                }
              //  MessageBox.Show("Количесвто кривых "+wktGeomList.Count.ToString());

                int a = item.Index;

                //asd = "";
                MessageBox.Show(wktGeomList.Count.ToString());
                using (OleDbConnection dbconn = new OleDbConnection(testConnStr))
                {
                    dbconn.Open();
                    OleDbCommand dbcomm = dbconn.CreateCommand();
                    dbcomm.CommandText = $"DECLARE @g geometry SET @g = geometry::STLineFromText('LINESTRING ({wktGeomList[a]})',0) INSERT INTO temp2 (Geom,Скважина,Площадь,Тип_Каротажа) VALUES(@g,{subItemsText[1]},'{subItemsText[0]}','{subItemsText[2]}')";
                    dbcomm.ExecuteNonQuery();
                    dbconn.Close();
                }
                subItemsText.Clear();
                item.Font = new Font("Calibri",12); // Изменяет размер шрифта выбранных строк 
                item.ForeColor = Color.DimGray; // Красит выбранный текст в цвет 
                item.BackColor = Color.Gainsboro; // Красит выбранные строки в цвет
                item.Checked = false; // Убирает флажки с занесеннего каротажа
            }
                    MessageBox.Show("Сохранено!");
            

        }


        private void button4_Click(object sender, EventArgs e)
        {
                       
            listView1.Items.Clear(); // Очищает таблицу
            namefile.Clear(); // Удаляет из листа имена файлов которые были открыти 
        }

        //
        // //
        // ЗАКЛАДКА "ПЕТРОФИЗИКА"
        // //
        //
        // Вызывается когда переключаешь закладки
        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (e.TabPage == tabPage3)
            {
                RequeryPet();
                
            }
        }

        private void PetTypes_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillPetBlock(PetTypes.SelectedItem == null ? -1 : ((ListItem)PetTypes.SelectedItem).id);
            FillPetGroup(PetBlock.SelectedItem == null ? -1 : ((ListItem)PetGroup.SelectedItem).id);
            FillPetParam(PetGroup.SelectedItem == null ? -1 : ((ListItem)PetGroup.SelectedItem).id);

            if (PetBlock.Items.Count > 0)
            {
                PetBlock.SelectedIndex = 0;
            }
        }

        private void PetBlock_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillPetGroup(PetBlock.SelectedItem == null ? -1 : ((ListItem)PetBlock.SelectedItem).id);


            if (PetGroup.Items.Count > 0)
            {
                PetGroup.SelectedIndex = 0;
            }
        }

        private void PetGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillPetParam(PetGroup.SelectedItem == null ? -1 : ((ListItem)PetGroup.SelectedItem).id);

            if (PetParam.Items.Count > 0)
            {
                PetParam.SelectedIndex = 0;
            }

        }


        private void PetBlock_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (PetBlock.SelectedItem == null) return;


            bool b = !(PetBlock.GetItemChecked(PetBlock.SelectedIndex));

            // Снимает все галки у подчиненых записей
            for (int i = 0; i < PetParams.Count; i++)
            {
                if (((ListItemPet)PetParams[i]).BlockID == ((ListItem)PetBlock.SelectedItem).id)
                {
                    // У параметров
                    PetParams[i].State = (b ? CheckState.Checked : CheckState.Unchecked);
                    // У группы
                    PetParams[i].GroupState = PetParams[i].State;
                    // У блока
                    PetParams[i].BlockState = PetParams[i].State;
                }
            }
        }

        private void PetGroup_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (PetGroup.SelectedItem == null) return;

            bool b = !(PetGroup.GetItemChecked(PetGroup.SelectedIndex));

            // Снимает все глаки у подчиненных записей
            for (int i = 0; i < PetParams.Count; i++)
            {
                if (((ListItemPet)PetParams[i]).GroupID == ((ListItem)PetGroup.SelectedItem).id)
                {
                    // У параметров
                    PetParams[i].State = (b ? CheckState.Checked : CheckState.Unchecked);
                    // У группы
                    PetParams[i].GroupState = PetParams[i].State;
                }
            }

            // Галку у блоков делаем серой, выбранной или снятой


            //petgroup.setitemcheckstate(petgroup.selectedindex, e.newvalue);
            //checkstate ch = petblock.getitemcheckstate(petblock.selectedindex);
            //int checkedcount = 0;
            //int uncheckedcount = 0;

            //for (int i = 0; i < petgroup.items.count; i++)
            //{
            //    if (petgroup.getitemcheckstate(i) == checkstate.checked) checkedcount++;
            //if (petgroup.getitemcheckstate(i) == checkstate.unchecked) uncheckedcount++;
            //    }
            //ch = checkstate.indeterminate;
            //if (checkedcount == petgroup.items.count) ch = checkstate.checked;
            //if (uncheckedcount == petgroup.items.count) ch = checkstate.unchecked;

            //petblock.setitemcheckstate(petblock.selectedindex, ch);


            }

        private void PetParam_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //  MessageBox.Show(PetParam.SelectedIndex.ToString());

            //PetParam.SetItemCheckState(e.Index, e.NewValue);



            //for (int i = 0; i < PetParams.Count; i++)
            //{
            //    if (PetParams[i].id == ((ListItem)PetParam.Items[e.Index]).id)
            //    {
            //        ((ListItemPet)PetParams[i]).State = PetParam.GetItemCheckState(e.Index);
            //        break;
            //    }
            //}

            SetPetGroupState();
        }

      
        // Кнопка отображение датагрида
        private void btnPetApplyFilter_Click(object sender, EventArgs e)
        {
            if (CurWellID == -1)
            {
                return;
            }
            //Собираем типы интервалов
            string IntTypes = GetIdsIntTypesOfPet();
            if (IntTypes.Length == 0) return;

            string IDParams = GetIdsParamsOfPet();

            Cursor.Current = Cursors.WaitCursor;
            DataSet ds = new DataSet();
            try
            {
                OleDbConnection dbconn = new OleDbConnection(connStr);
                dbconn.Open();
                string SQL = "EXEC spGetPet @WellID=" + CurWellID.ToString() + ", @Types='" + IntTypes + "', @IDParams=" + (IDParams.Length == 0 ? "NULL" : "'" + IDParams + "', @CutOnBlocks=" + (chkBlockPet.Checked ? "1" : "0"));
                OleDbDataAdapter dataAdapt = new OleDbDataAdapter(SQL,dbconn);
                dataAdapt.Fill(ds);

            }
            finally 
            {
                Cursor.Current = Cursors.Default;
            }

            if (ds == null) return;

            DataGridView PetGrid = new DataGridView();
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                // Настройка датагрида
                PetGrid.ColumnHeadersDefaultCellStyle.Font = new Font(new FontFamily("Arial Cyr"),(float)(9.5), FontStyle.Bold, GraphicsUnit.Point);
                PetGrid.DefaultCellStyle.Font = new Font(new FontFamily("Arial Cyr"), (float)(9.25), FontStyle.Bold, GraphicsUnit.Point);
                PetGrid.AutoGenerateColumns = true;
                PetGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
              
                PetGrid.Columns.Clear();
                PetGrid.DataSource = ds.Tables[i];
                PetGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                PetGrid.RowHeadersWidth = 60;
                PetGrid.RowPostPaint += numeric_Point_DataGrid; // Обработчик нумераций строк в гриде
               
                if(!chkBlockPet.Checked && ds.Tables.Count == 1)
                {
                    tcPet.TabPages.Add("Все блоки");
                    tcPet.Appearance = TabAppearance.FlatButtons;
                    tcPet.ItemSize = new Size(0, 1);
                    tcPet.SizeMode = TabSizeMode.Fixed;
                    tcPet.TabStop = false;

                }
                else
                {
                    tcPet.TabPages.Add(" " + ((ListItem)PetBlock.Items[i]).name + " ");
                    tcPet.SelectedTab = tcPet.TabPages[tcPet.TabPages.Count - 1];
                    tcPet.Appearance = TabAppearance.Normal;
                    tcPet.ItemSize = new Size(69, 24);
                    tcPet.SizeMode = TabSizeMode.Normal;
                    tcPet.TabStop = true;
                }

                PetGrid.Parent = tcPet.TabPages[i];
                PetGrid.Dock = DockStyle.Fill;
                PetGrid.ReadOnly = true;
                PetGrid.AllowUserToAddRows = false;

                // Искуственно  скрывает колонку ID_образца
                int h = FindColumnIndexByFieldName(PetGrid, "ID_образца");
                if (h != -1)
                {
                    PetGrid.Columns[i].Visible = false;
                }

                // Колонка динамически постороена 
                bool IsDinamicCol = false;
                CurColorIndex = -1;
                Color c = Color.Black;
                if (PetGrid.Columns.Count > 0)
                {
                    c = PetGrid.Columns[0].DefaultCellStyle.BackColor;
                }

                string prev_BlockName = "";
                string BlockName = "";

                for (int j = 0; j < PetGrid.Columns.Count; j++)
                {
                    //MessageBox.Show(PetGrid.Columns[j].ToString());
                    if (PetGrid.Columns[j].HeaderText.IndexOf(Environment.NewLine) < 0)
                    {
                        Size len = TextRenderer.MeasureText(PetGrid.Columns[j].HeaderText, PetGrid.ColumnHeadersDefaultCellStyle.Font);
                        PetGrid.Columns[j].Width = len.Width + 36;
                        PetGrid.Columns[j].MinimumWidth = PetGrid.Columns[j].Width;
                    }
                    else
                    {
                        string st = PetGrid.Columns[j].HeaderText;
                        string st1 = (st.Substring(0, st.IndexOf(Environment.NewLine)));
                        string st2 = (st.Substring(st.IndexOf(Environment.NewLine) + 2, st.Length - st1.Length - 2));

                        Size len1 = TextRenderer.MeasureText(st1, PetGrid.ColumnHeadersDefaultCellStyle.Font);
                        Size len2 = TextRenderer.MeasureText(st2, PetGrid.ColumnHeadersDefaultCellStyle.Font);
                        PetGrid.Columns[j].Width = Math.Max(len1.Width, len2.Width) + 40;
                        PetGrid.Columns[j].MinimumWidth = PetGrid.Columns[j].Width;
                        if (len1.Width == 0 && len2.Width > 0)
                        {
                            PetGrid.Columns[j].HeaderText = st2;
                        }

                        IsDinamicCol = true;

                        BlockName = "";
                        string[] words = st.Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                        if (!chkBlockPet.Checked && words.Length == 3)
                        {
                            BlockName = words[0];
                        }



                       
                    }

                    if (IsDinamicCol)
                    {
                        PetGrid.Columns[j].DefaultCellStyle.BackColor = c;
                    }

                    prev_BlockName = BlockName;
                }
                    if (PetGrid != null) PetGrid = null;
            }
                if (chkBlockPet.Checked) tcPet.SelectedIndex = 0;
        }

        public static int CurColorIndex = -1;
        public static List<Color> Colors = new List<Color>();
        public static Color GetNextColor()
        {
            CurColorIndex = (CurColorIndex == Colors.Count - 1 ? 0 : CurColorIndex + 1);
            return Colors[CurColorIndex];
        }

        /// <summary>
        /// Находит колонку грида, к которой привязано поле базы даныых с именем FieldName
        /// </summary>
        /// <param name="grid">Наш грид</param>
        /// <param name="FieldName">Поле в базе данных</param>
        /// <returns>index колонки грида или -1, если не найдено</returns>
        private static int FindColumnIndexByFieldName(DataGridView grid, string FieldName)
        {
            FieldName = FieldName.ToLower();
            int col_ind = -1;
            for (int i = 0; i < grid.Columns.Count; i++)
            {
                if (grid.Columns[i].DataPropertyName.ToLower() == FieldName)
                {
                    col_ind = i;
                    break;
                }
                
            }
            return col_ind;
        }

        private string GetIdsParamsOfPet()
        {
            string IDParams = "";

            for (int i = 0; i < PetParams.Count; i++)
            {
                if (((ListItemPet)PetParams[i]).State == CheckState.Checked)
                {
                    IDParams = IDParams + (IDParams.Length > 0 ? "," : "") + ((ListItemPet)PetParams[i]).id;
                }
            }
            return IDParams;
        }

        private string GetIdsIntTypesOfPet()
        {
            string IntTypes = "";
            // Проходит по "Типам объектов" выбранной скважины и добавляет их в виде "1,2,3" в зависимости от значения "Керн/Шельф/Шлам"
            for (int i = 0; i < PetTypes.Items.Count; i++)
            {
                if (PetTypes.GetItemChecked(i))
                {
                    IntTypes = IntTypes + (IntTypes.Length > 0 ? "," : "") + ((ListItem)PetTypes.Items[i]).id.ToString();
                }
            }
            return IntTypes;
        }

        // Кнопка очиски петрофизики
        private void btnClearFilter_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < PetTypes.Items.Count; i++)
            {
                PetTypes.SetItemChecked(i, false);
            }

            tcPet.TabPages.Clear();
            RewritePetParams(CurWellID);
            chkBlockPet.Checked = false;
            FillPetBlock(-1);// Очищаем список блоков
            FillPetGroup(-1);// Очищаем список групп
            FillPetParam(-1);// Очищаем список параметров
               
        }

        // Кнопка добавления параметров петрофизики
        private void btnAddPetParam_Click(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.Title = "Выбрать файл";
            if (op.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            Form2 newF = new Form2();
            newF.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            if (op.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }
            using (var stream = File.Open(op.FileName, FileMode.Open, FileAccess.Read)) 
            {
                
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                        {
                            int collNamb = reader.FieldCount;
                            for (int i = 0; i < collNamb-1; i++)
                            {

                            }
                            DataGridViewTextBoxCell Glybina = new DataGridViewTextBoxCell();
                            DataGridViewTextBoxCell Ugol = new DataGridViewTextBoxCell();

                            Glybina.Value = reader.GetValue(1);
                            Ugol.Value = reader.GetValue(2);

                            DataGridViewRow row = new DataGridViewRow();

                            row.Cells.AddRange(Glybina,Ugol);
                            Grid_Inklin.Rows.Add(row);
                            MessageBox.Show(collNamb.ToString());
                        }
                    } while (reader.NextResult());

                }

            }
        }
        
        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            if (op.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }
            object temppatch = op.FileName;
            var wordapp = new Word.Application();
            wordapp.Visible = false;
            wordapp.Options.Overtype = false;
            var worddoc = wordapp.Documents.Open(ref temppatch,Type.Missing);
            
            //string text = "";
            for (int i = 0; i < worddoc.Paragraphs.Count; i++)
            {
                //text += " \r\n " + worddoc.Paragraphs[i + 1].Range.Text;
                DataGridViewTextBoxCell adad = new DataGridViewTextBoxCell();

                DataGridViewTextBoxCell test = new DataGridViewTextBoxCell();
                object uni = Word.WdUnits.wdCharacter;
                object ext = Word.WdMovementType.wdMove;
                object count = 6;
                wordapp.Selection.HomeKey(Word.WdUnits.wdStory,Word.WdMovementType.wdMove);
                wordapp.Selection.MoveRight(ref uni, ref count, ref ext);
                uni = Word.WdUnits.wdWord;
                count = 1;
                ext = Word.WdMovementType.wdExtend;
                wordapp.Selection.MoveRight(ref uni, ref count, ref ext);

                adad.Value = "\r\n" + worddoc.Paragraphs[i + 1].Range.Text;
               // test.Value = wordrange;
                DataGridViewRow row1 = new DataGridViewRow();
                row1.Cells.AddRange(adad,test);
                Grid_Inklin.Rows.Add(row1);
            }



            //MessageBox.Show(text);
            worddoc.Close();
            wordapp.Quit();

        }

        private void button7_Click(object sender, EventArgs e)
        {
           string connStr1 = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ModelingBasins;Data Source=IPGG-SQL";

            using (OleDbConnection dbconn = new OleDbConnection(connStr1))
            {
                dbconn.Open();
                OleDbCommand dbcomm = new OleDbCommand($"DECLARE @g geometry = ( select Geogr from Субъекты_федерации where ID_субъекта_федерации = '{3}') SELECT @g.BufferWithCurves(0).ToString()", dbconn);
                // Выбирает из базы кривую и представляет её с типом varсhar для записи 
                using (OleDbDataReader reader = dbcomm.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            wktGeom = reader[0].ToString();
                            wktGeom = wktGeom.Remove(0, 16); //Убирает символы с 0-12, там "LINESTRING ("
                            wktGeom = wktGeom.Remove(wktGeom.Length - 3); // Убирает последний символ, там ")"
                        }
                    }
                }
                dbconn.Close();
            }

            string[][] geomData = wktGeom.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(line => line.Split(" ;".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)).ToArray();

            for (int i = 0; i < geomData.Length; i++)
            {
                geomString += $"{geomData[i][0]} {geomData[i][1]}" + '\r';
            }

            SaveFileDialog sf = new SaveFileDialog();
            sf.Filter = "Las файл(*.Las) | *.Las";
            sf.Title = "Сохранение файла";
            if (sf.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }
            File.WriteAllText(sf.FileName, geomString);
            MessageBox.Show("Сохранено!", "Сохранение");
        }




        /// <summary>
        /// Метод заполнения окна "Типы объектов"
        /// </summary>
        /// <param name="WellID"></param> ID тек. скважины или -1 если нужно очистить список
        private void FillPetTypes(int WellID)
        {
            PetTypes.Items.Clear();
           
            if (WellID == -1) return;

            using (OleDbConnection dbconn = new OleDbConnection(connStr))
            {
                dbconn.Open();
               
                OleDbCommand dbcomm = dbconn.CreateCommand();
                dbcomm.CommandText = " SELECT DISTINCT OI.Тип AS id, CASE"
                       + "   WHEN OI.Тип=1 THEN 'Керн'"
                       + "   WHEN OI.Тип=2 THEN 'Шлиф'"
                       + "   WHEN OI.Тип=3 THEN 'Шлам'"
                       + " END AS name"
                       + " FROM Образец_интервалы OI (NOLOCK)"
                       + $" WHERE OI.Объект=1 AND OI.ID_объекта={WellID} ORDER BY 1";
                OleDbDataReader reader = dbcomm.ExecuteReader();
                while (reader.Read())
                {
                    PetTypes.Items.Add(new ListItem((int)reader["id"],reader["name"].ToString()));
                }
                dbconn.Close();
            }
        }
        


        /// <summary>
        ///  Метод заполнения окна "Блоки"
        /// </summary>
        /// <param name="TypeID"></param> Выбранный тип интервала образца или -1, если нужно очистить список
        private void FillPetBlock(int TypeID)
        {
            PetBlock.Items.Clear();
            if (TypeID == -1) return;

            // Из листа PetParams заполняем по типу
            for (int i = 0; i < PetParams.Count; i++)
            {
                if (PetParams[i].TypeInt == TypeID)
                {
                  if (PetBlock.FindString(PetParams[i].BlockName) == -1)
                  {
                      PetBlock.Items.Add(new ListItem((int)PetParams[i].BlockID, PetParams[i].BlockName));
                      PetBlock.SetItemCheckState(PetBlock.Items.Count - 1, PetParams[i].BlockState);
                  }

                }
            }
        }
        
        /// <summary>
        /// Метод заполнения окна "Группы"
        /// </summary>
        /// <param name="TypeID"></param> Выбранный блок или -1
        private void FillPetGroup(int BlockID)
        {
            PetGroup.Items.Clear();
            if (BlockID == -1) return;

            // Из листа PetParams заполняем по типу
            for (int i = 0; i < PetParams.Count; i++)
            {
                if (PetParams[i].BlockID == BlockID)
                    if (PetGroup.FindString(PetParams[i].GroupName) == -1)
                    {
                        PetGroup.Items.Add(new ListItem((int)PetParams[i].GroupID, PetParams[i].GroupName));
                        PetGroup.SetItemCheckState(PetGroup.Items.Count - 1, PetParams[i].GroupState);
                    }
            }

           SetPetBlockState();
        }

        /// <summary>
        ///  Метод заполнения окна "Пареметры групп"
        /// </summary>
        /// <param name="GroupID"></param>
        private void FillPetParam(int GroupID)
        {
            PetParam.Items.Clear();
            if (GroupID == -1) return;

            for (int i = 0; i < PetParams.Count; i++)
            {
                if (PetParams[i].GroupID == GroupID)
                {
                    if (PetParam.FindString(PetParams[i].name) == -1)
                    {
                        PetParam.Items.Add(new ListItem(PetParams[i].id, PetParams[i].name), PetParams[i].State);
                    }
                }
            }
           
            SetPetGroupState();
            
        }

        // Устанавливаем нужное значение галочек
        private void SetPetGroupState()
        {
            // Если группа не выделена, то ничего не пересчитываем
            if (PetGroup.SelectedItem == null) return;

            CheckState ch = PetGroup.GetItemCheckState(PetGroup.SelectedIndex);
            int GroupID = ((ListItem)PetGroup.SelectedItem).id; // Присваивает айди группы текущей скважины
            int CheckedCount = 0; // Считает сколько отмечено
            int UncheckedCount = 0; // Считает сколько не отмечено

            // Проходит циклом по окну "Группы" и смотрит сколько галок есть
            for (int i = 0; i < PetParams.Count; i++)
            {
                if (PetParams[i].GroupID == GroupID)
                {
                    if (PetParams[i].State == CheckState.Checked) CheckedCount++;
                    if (PetParams[i].State == CheckState.Unchecked) UncheckedCount++;
                }
            }

            // Сверяем галочки
            ch = CheckState.Indeterminate;
            if (CheckedCount == PetParam.Items.Count) ch = CheckState.Checked;
            if (UncheckedCount == PetParam.Items.Count) ch = CheckState.Unchecked;

            // Ставим нужное состояние у групп
            for (int i = 0; i < PetParams.Count; i++)
            {
                if (PetParams[i].GroupID == GroupID) PetParams[i].GroupState = ch;

            }

            PetGroup.SetItemCheckState(PetGroup.SelectedIndex, ch);
        }

        private void SetPetBlockState()
        {
            // Если группа не выделена, то ничего не пересчитываем
            if (PetBlock.SelectedItem == null) return;

            CheckState ch = PetBlock.GetItemCheckState(PetBlock.SelectedIndex);
            int BlockID = ((ListItem)PetBlock.SelectedItem).id; // Присваивает айди блока текущей скважины
            int CheckedCount = 0; // Считает сколько отмечено
            int UnchekedCount = 0; // Считает сколько не отмечено

            // Проходит циклом по окну "Группы" текущего значения окна "Блоки" и смотрит сколько галок есть
            for (int i = 0; i < PetParams.Count; i++)
            {
                if (PetParams[i].BlockID == BlockID)
                {
                    if (PetParams[i].GroupState == CheckState.Checked) CheckedCount++;
                    if (PetParams[i].GroupState == CheckState.Unchecked) UnchekedCount++;
                }
            }

            // Сверяем сколько галочек проставлено
            ch = CheckState.Indeterminate;
            if (CheckedCount > 0 && UnchekedCount == 0) ch = CheckState.Checked;
            if (CheckedCount == 0 && UnchekedCount > 0) ch = CheckState.Unchecked;

            // Ставим в нужное состояние галки у блока
            for (int i = 0; i < PetParams.Count; i++)
            {
                if (PetParams[i].BlockID == BlockID) PetParams[i].BlockState = ch;
            }

            PetBlock.SetItemCheckState(PetBlock.SelectedIndex, ch);


        }

        /// <summary>
        /// Считываем из БД список параметров по выбранной скважине в PetParams
        /// </summary>
        /// <param name="WellID"></param> Текущая скважина или -1,если не выбрана текущая 
        private void RewritePetParams(int WellID)
        {
            if (PetParams != null) PetParams = null;
            if (WellID == -1) return;
            
            if (PetParams != null) PetParams.Clear();
            else PetParams = new List<ListItemPet>();

            using (OleDbConnection dbconn = new OleDbConnection(connStr))
            {
                dbconn.Open();
               
                DataTable dt = new DataTable();
                //Запрос загружает в список PetParams все параметры всех групп
               
                 string   sql= " SELECT DISTINCT OI.Тип AS ТипИнт, CASE"
                       + "   WHEN OI.Тип=1 THEN 'Керн'"
                       + "   WHEN OI.Тип=2 THEN 'Шлиф'"
                       + "   WHEN OI.Тип=3 THEN 'Шлам'"
                       + " END AS ТипНазв, "
                       + " PP.ID_петрофизич_параметра, PP.Название,"
                       + " ISNULL(PP.ID_группы_петрофизики,0) AS GroupID, "
                       + " ISNULL((SELECT Наименование FROM Петрофизика_группы (NOLOCK) WHERE ID_группы_петрофизики = PP.ID_группы_петрофизики),'Основная') AS GrName,"
                       + " ISNULL(PG.ID_блока_петрофизики,0) AS BlockID, "
                       // По этой колонке идет сортировка с указанием номера колонки 8 в ORDER BY!
                       + " ISNULL((SELECT Наименование FROM Петрофизика_блоки (NOLOCK) WHERE ID_блока_петрофизики = PG.ID_блока_петрофизики),'Основной') AS BlockName"
                       + " FROM Петрофизика_параметры PP (NOLOCK)"
                       + " JOIN Образец_петрофизика OP (NOLOCK) ON (OP.ID_петрофизич_параметра = PP.ID_петрофизич_параметра)"
                       + " JOIN Образец_интервалы OI (NOLOCK) ON (OI.ID_образца = OP.ID_образца)"
                       + " LEFT JOIN Петрофизика_группы PG (NOLOCK) ON (PG.ID_группы_петрофизики = PP.ID_группы_петрофизики)"
                       + $" WHERE OI.Объект=1 AND OI.ID_объекта={WellID}"
                       + " ORDER BY 8, ISNULL(PP.ID_группы_петрофизики,0)";
                
                OleDbDataAdapter dbadapt = new OleDbDataAdapter(sql,dbconn);
                dbadapt.Fill(dt);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ListItemPet item = new ListItemPet();
                    item.id = (int)dt.Rows[i]["ID_петрофизич_параметра"]; // Это айди параметра петрофизики к каждому названию
                    item.name = dt.Rows[i]["Название"].ToString(); // Само название параметра, та самая пористость и тому подобное
                    item.TypeInt = (int)dt.Rows[i]["ТипИнт"]; // Берёт тип обекта (1-3) данной скважины и записывает её тип 
                    item.TypeName = dt.Rows[i]["ТипНазв"].ToString(); // Это тип скважины(Керн\Шлиф\Шлам)
                    item.State = CheckState.Unchecked;
                    item.GroupID = 0; //  Это айди группы петрофизики у  айди параметра  петрофизики
                    if (dt.Rows[i]["GroupID"] != DBNull.Value) item.GroupID = Convert.ToInt32(dt.Rows[i]["GroupID"]); // DBNull проверяет есть ли значение в поле базы, если(if) не равно 0, то есть есть, то записывает его туда 
                    item.GroupName = dt.Rows[i]["GrName"].ToString(); // Это название группы петрофизики(плотность,пористость...)
                    item.BlockID = 0; // Это айди блока группы петрофизики
                    if (dt.Rows[i]["BlockID"] != DBNull.Value) item.BlockID = Convert.ToInt32(dt.Rows[i]["BlockID"]); // Точно так же как с айди группы ранее
                    item.BlockName = dt.Rows[i]["BlockName"].ToString(); // Нзвание  блока  свойств (физ-мех\физич)
                    item.BlockState = CheckState.Unchecked;
                    
                    PetParams.Add(item); // Добавляет всю инфу в лист наследованный от класса 
                }
               
                if (PetParams == null) return;

                dbconn.Close();

            }

            FillPetTypes(WellID);
        }

        /// <summary>
        /// 
        /// </summary>
        private void RequeryPet()
        {
            RewritePetParams(CurWellID);
            //Загружаем доступные типы для данной скважины
            FillPetTypes(CurWellID);
            FillPetBlock(-1); // Очищаем список блоков
            FillPetGroup(-1); // Очищаем список Групп
            FillPetParam(-1); // Очищаем список параметров

            if(CurWellID == -1)
            {
                PetTypes.Items.Clear();
                RewritePetParams(-1);
                return;
            }

            // Если есть "Керн" в списке Типов для данной скважины, то отмечаем сразу и делаем выборку остальных списков
            if (PetTypes.Items.Count>0)
            {
                PetTypes.SelectedIndex = 0;
                PetTypes.SetItemChecked(0, true);
                
            }
        }



        

        /// <summary>
        /// Класс "Элемент параметра петрофизики"
        /// </summary>
        ///  Класс для записи информации петрофизики о скважине
        public class ListItemPet  
        {
            // Это айди параметра петрофизики к каждому названию
            public int id;
            // Само название параметра, та самая пористость и тому подобное
            public string name;
            // Берёт тип обекта (1-3) данной скважины и записывает её тип 
            public int TypeInt;
            // Это тип скважины(Керн\Шлиф\Шлам)
            public string TypeName;
            public CheckState State;

            //  Это айди группы петрофизики у  айди параметра  петрофизики
            public int GroupID;
            // Это название группы петрофизики(плотность,пористость...)
            public string GroupName;
            public CheckState GroupState;
            // 3 поля ниже используются только для петрофизики (для литологии не используются)
            // Это айди блока группы петрофизики
            public int BlockID;
            // Нзвание  блока  свойств (физ-мех\физич)
            public string BlockName;
            public CheckState BlockState;
        }

        // так и не разобрался зачем нужен этот класс
        public class ListItem
        {
            public int id;
            public string name;

            public ListItem(int KeyValue, string Text)
            {
                this.id = KeyValue;
                this.name = Text;
            }
            public override string ToString()
            {
                return name;
            }
        }


    }
}
