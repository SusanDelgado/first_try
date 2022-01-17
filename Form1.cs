using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;


namespace Создатель_расчетов
{
    public partial class Form1 : Form
    {
        public DataTable application_table = new DataTable("application");
        public DataTable paper_table = new DataTable("paper");
        private Word.Application wordapp;
        private Excel.Application excelApp;
        private Word.Document worddocument;
        private Word.Paragraphs wordparagraphs;
        private Word.Paragraph wordparagraph;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView2.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView3.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView4.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView4.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            comboBox1.SelectedIndex = 0;
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox4.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.Items.Add("                    +");
            int index = 0;
            for (int i = 0; i < comboBox2.Items.Count; i++)
            {
                if (comboBox2.Items[i].ToString() == "                    +")
                    index = i;
            }
            comboBox2.SelectedIndex = index;
            comboBox3.Items.Add("Нет");
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;
            if (comboBox3.SelectedIndex == 0)
                numericUpDown1.Enabled = false;

            DataGridViewTextBoxColumn[] column_datagrid1 = new DataGridViewTextBoxColumn[3];
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Font = new Font("Times New Roman", 9);
            column_datagrid1[0] = new DataGridViewTextBoxColumn();
            column_datagrid1[1] = new DataGridViewTextBoxColumn();
            column_datagrid1[1].HeaderText = "Наименование";
            column_datagrid1[2] = new DataGridViewTextBoxColumn();
            column_datagrid1[2].HeaderText = "Количество";
            dataGridView1.Columns.AddRange(column_datagrid1);
            dataGridView1.Columns[0].Width = 26;
            dataGridView1.Columns[0].ReadOnly = true;
            dataGridView1.Columns[1].Width = 200;
            dataGridView1.Columns[2].Width = 100;

            DataGridViewTextBoxColumn[] column_datagrid2 = new DataGridViewTextBoxColumn[4];
            dataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Font = new Font("Times New Roman", 9);
            column_datagrid2[0] = new DataGridViewTextBoxColumn();
            column_datagrid2[1] = new DataGridViewTextBoxColumn();
            column_datagrid2[1].HeaderText = "Тип";
            column_datagrid2[2] = new DataGridViewTextBoxColumn();
            column_datagrid2[2].HeaderText = "Использовано";
            column_datagrid2[3] = new DataGridViewTextBoxColumn();
            column_datagrid2[3].HeaderText = "% брака";
            dataGridView2.Columns.AddRange(column_datagrid2);
            dataGridView2.Columns[0].Width = 26;
            dataGridView2.Columns[0].ReadOnly = true;
            dataGridView2.Columns[1].Width = 90;
            dataGridView2.Columns[2].Width = 70;
            dataGridView2.Columns[3].Width = 40;


            DataGridViewTextBoxColumn[] column_datagrid3 = new DataGridViewTextBoxColumn[4];
            dataGridView3.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column_datagrid3[0] = new DataGridViewTextBoxColumn();
            column_datagrid3[0].HeaderText = "Тип";
            column_datagrid3[1] = new DataGridViewTextBoxColumn();
            column_datagrid3[1].HeaderText = "К-во в пачке";
            column_datagrid3[2] = new DataGridViewTextBoxColumn();
            column_datagrid3[2].HeaderText = "Цена пачки";
            column_datagrid3[3] = new DataGridViewTextBoxColumn();
            column_datagrid3[3].HeaderText = "Цена 1 шт.";
            dataGridView3.Columns.AddRange(column_datagrid3);
            for(int i = 0; i < 4; i++)
                dataGridView3.Columns[i].Width = 80;
            dataGridView3.Font = new Font("Times New Roman", 9);

            DataGridViewTextBoxColumn[] column_datagrid4 = new DataGridViewTextBoxColumn[5];
            dataGridView4.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            column_datagrid4[0] = new DataGridViewTextBoxColumn();
            column_datagrid4[1] = new DataGridViewTextBoxColumn();
            column_datagrid4[1].HeaderText = "Тип бумаги";
            column_datagrid4[2] = new DataGridViewTextBoxColumn();
            column_datagrid4[2].HeaderText = "Ед. изм.";
            column_datagrid4[3] = new DataGridViewTextBoxColumn();
            column_datagrid4[3].HeaderText = "Цена";
            column_datagrid4[4] = new DataGridViewTextBoxColumn();
            column_datagrid4[4].HeaderText = "К-во";
            DataGridViewCheckBoxColumn dgvCmb = new DataGridViewCheckBoxColumn();
            dgvCmb.ValueType = typeof(bool);
            dgvCmb.HeaderText = "Исп. в печати";
            dataGridView4.Columns.AddRange(column_datagrid4);
            dataGridView4.Columns.Add(dgvCmb);
            dataGridView4.Columns[0].Width = 26;
            dataGridView4.Columns[0].ReadOnly = true;
            dataGridView4.Columns[1].Width = 115;
            dataGridView4.Columns[2].Width = 45;
            dataGridView4.Columns[3].Width = 45;
            dataGridView4.Columns[4].Width = 45;
            dataGridView4.Columns[5].Width = 45;
            dataGridView4.Font = new Font("Times New Roman", 9);

            foreach(DataGridViewColumn column in dataGridView1.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            foreach (DataGridViewColumn column in dataGridView2.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            foreach (DataGridViewColumn column in dataGridView3.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            foreach (DataGridViewColumn column in dataGridView4.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            try
            {
                XmlDocument new_doc = new XmlDocument();
                new_doc.Load("consumables.xml");
                XmlElement new_elem = new_doc.DocumentElement;
                foreach (XmlNode xnode in new_elem)
                {
                    if (xnode.Name == "staples_type")
                    {
                        foreach (XmlNode childnode_of_child_0 in xnode.ChildNodes)
                        {
                            if (childnode_of_child_0.Name == "staple")
                            {
                                int row_datagrid3 = dataGridView3.Rows.Add();
                                foreach (XmlNode childnode_of_child_1 in childnode_of_child_0.ChildNodes)
                                {
                                    if (childnode_of_child_1.Name == "name")
                                    {
                                        dataGridView3.Rows[row_datagrid3].Cells[0].Value = $"{childnode_of_child_1.InnerText}";
                                        comboBox3.Items.Add(childnode_of_child_1.InnerText);
                                    }
                                    if (childnode_of_child_1.Name == "count")
                                        dataGridView3.Rows[row_datagrid3].Cells[1].Value = $"{childnode_of_child_1.InnerText}";
                                    if (childnode_of_child_1.Name == "price")
                                        dataGridView3.Rows[row_datagrid3].Cells[2].Value = $"{childnode_of_child_1.InnerText}";
                                }
                                double price = Double.Parse(dataGridView3.Rows[row_datagrid3].Cells[2].Value.ToString()) / Double.Parse(dataGridView3.Rows[row_datagrid3].Cells[1].Value.ToString());
                                dataGridView3.Rows[row_datagrid3].Cells[3].Value = price;
                            }
                        }
                    }
                    string name = "";
                    if (xnode.Name == "paper_type")
                    {
                        foreach (XmlNode childnode_of_child_0 in xnode.ChildNodes)
                        {
                            if (childnode_of_child_0.Name == "paper")
                            {
                                int row_datagrid4 = dataGridView4.Rows.Add();
                                foreach (XmlNode childnode_of_child_1 in childnode_of_child_0.ChildNodes)
                                {
                                    dataGridView4.Rows[row_datagrid4].Cells[0].Value = dataGridView4.Rows.Count;
                                    if (childnode_of_child_1.Name == "name")
                                    {
                                        dataGridView4.Rows[row_datagrid4].Cells[1].Value = $"{childnode_of_child_1.InnerText}";
                                        name = childnode_of_child_1.InnerText;
                                    }
                                    if (childnode_of_child_1.Name == "unit")
                                        dataGridView4.Rows[row_datagrid4].Cells[2].Value = $"{childnode_of_child_1.InnerText}";
                                    if (childnode_of_child_1.Name == "price")
                                        dataGridView4.Rows[row_datagrid4].Cells[3].Value = $"{childnode_of_child_1.InnerText}";
                                    if (childnode_of_child_1.Name == "count")
                                        dataGridView4.Rows[row_datagrid4].Cells[4].Value = $"{childnode_of_child_1.InnerText}";
                                    if (childnode_of_child_1.Name == "used")
                                    {
                                        dataGridView4.Rows[row_datagrid4].Cells[5].Value = Int32.Parse(childnode_of_child_1.InnerText);
                                        if (Convert.ToBoolean(dataGridView4.Rows[row_datagrid4].Cells[5].Value) == true)
                                            comboBox2.Items.Add(name);
                                    }
                                }
                            }
                        }
                    }
                    if (xnode.Name == "Master")
                    {
                        foreach (XmlNode childnode_of_child_0 in xnode.ChildNodes)
                        {
                            if (childnode_of_child_0.Name == "count")
                                numericUpDown3.Value = Int32.Parse(childnode_of_child_0.InnerText);
                            if (childnode_of_child_0.Name == "price")
                                textBox4.Text = $"{childnode_of_child_0.InnerText}";
                        }
                        double count = Double.Parse(numericUpDown3.Value.ToString());
                        double price = Double.Parse(textBox4.Text) / count;
                        label12.Text = price.ToString();
                    }
                    if (xnode.Name == "Paint")
                    {
                        foreach (XmlNode childnode_of_child_0 in xnode.ChildNodes)
                        {
                            if (childnode_of_child_0.Name == "price_of_tube")
                                textBox5.Text = $"{childnode_of_child_0.InnerText}";
                            if (childnode_of_child_0.Name == "prise_of_imprint")
                                textBox6.Text = $"{childnode_of_child_0.InnerText}";
                        }
                    }
                    if (xnode.Name == "Color_paint")
                    {
                        foreach (XmlNode childnode_of_child_0 in xnode.ChildNodes)
                        {
                            if (childnode_of_child_0.Name == "Set_price")
                                textBox7.Text = $"{childnode_of_child_0.InnerText}";
                            if (childnode_of_child_0.Name == "Imprint_price")
                                textBox8.Text = $"{childnode_of_child_0.InnerText}";
                        }
                    }
                    if (xnode.Name == "Hard_leaf")
                    {
                        foreach (XmlNode childnode_of_child_0 in xnode.ChildNodes)
                        {
                            if (childnode_of_child_0.Name == "Channel_priceA3")
                                textBox12.Text = $"{childnode_of_child_0.InnerText}";
                            if (childnode_of_child_0.Name == "Cover_priceA3")
                                textBox13.Text = $"{childnode_of_child_0.InnerText}";
                            if (childnode_of_child_0.Name == "Channel_priceA4")
                                textBox16.Text = $"{childnode_of_child_0.InnerText}";
                            if (childnode_of_child_0.Name == "Cover_priceA4")
                                textBox15.Text = $"{childnode_of_child_0.InnerText}";
                            if (childnode_of_child_0.Name == "Channel_priceA5")
                                textBox18.Text = $"{childnode_of_child_0.InnerText}";
                            if (childnode_of_child_0.Name == "Cover_priceA5")
                                textBox17.Text = $"{childnode_of_child_0.InnerText}";
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Необходимые файлы отсутствуют, переустановите программу","Ошибка",MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }

            try
            {
                XmlDocument xDoc = new XmlDocument();
                xDoc.Load("application.xml");
                XmlElement xRoot = xDoc.DocumentElement;
                foreach (XmlNode xnode in xRoot)
                {
                    foreach (XmlNode childnode in xnode.ChildNodes)
                    {
                        if (childnode.Name == "application_number")
                        {
                            numericUpDown4.Value = Int32.Parse(childnode.InnerText);
                        }
                        if (childnode.Name == "applicant")
                        {
                            textBox1.Text = $"{childnode.InnerText}";
                        }
                        if (childnode.Name == "applicant_rank")
                        {
                            comboBox4.SelectedIndex = Int32.Parse(childnode.InnerText);
                        }
                        if (childnode.Name == "applicant_name")
                        {
                            textBox14.Text = $"{childnode.InnerText}";
                        }
                        if (childnode.Name == "employee_position")
                        {
                            textBox2.Text = $"{childnode.InnerText}";
                        }
                        if (childnode.Name == "employee_rank")
                        {
                            comboBox1.SelectedIndex = Int32.Parse(childnode.InnerText);
                        }
                        if (childnode.Name == "employee_name")
                        {
                            textBox3.Text = $"{childnode.InnerText}";
                        }
                    }
                }

                DataColumn PaperId = new DataColumn("Id", Type.GetType("System.Int32"));
                DataColumn paper_type_for_paper_table = new DataColumn("paper_type", Type.GetType("System.Int32"));
                DataColumn paper_count_for_paper_table = new DataColumn("paper_count", Type.GetType("System.Int32"));
                DataColumn defect_for_paper_table = new DataColumn("defect", Type.GetType("System.Double"));
                paper_table.Columns.Add(PaperId);
                paper_table.Columns.Add(paper_type_for_paper_table);
                paper_table.Columns.Add(paper_count_for_paper_table);
                paper_table.Columns.Add(defect_for_paper_table);

                DataColumn idColumn = new DataColumn("Id", Type.GetType("System.Int32"));
                DataColumn nameColumn = new DataColumn("name", Type.GetType("System.String"));
                DataColumn product_countColumn = new DataColumn("product_count", Type.GetType("System.Int32"));
                DataColumn paper_id_column = new DataColumn("paper_id", Type.GetType("System.Int32"));
                DataColumn masterColumn = new DataColumn("master", Type.GetType("System.Int32"));
                DataColumn staple_typeColumn = new DataColumn("staple_type", Type.GetType("System.Int32"));
                DataColumn staple_countColumn = new DataColumn("staple_count", Type.GetType("System.Int32"));
                DataColumn color_paintColumn = new DataColumn("color_paint", Type.GetType("System.Int32"));
                DataColumn hard_leafColumn = new DataColumn("hard_leaf", Type.GetType("System.Int32"));

                application_table.Columns.Add(idColumn);
                application_table.Columns.Add(nameColumn);
                application_table.Columns.Add(product_countColumn);
                application_table.Columns.Add(paper_id_column);
                application_table.Columns.Add(masterColumn);
                application_table.Columns.Add(staple_typeColumn);
                application_table.Columns.Add(staple_countColumn);
                application_table.Columns.Add(color_paintColumn);
                application_table.Columns.Add(hard_leafColumn);

                int counter = 1;
                string product_name = "";
                int product_count = 0;
                int master = 0;
                int staple_type = 0;
                int staple_count = 0;
                int paper_id = 0;
                int color_paint = 0;
                int hard_leaf = 0;
                int temp_paper_type = 0;
                int temp_paper_count = 0;
                double temp_defect = 0;
                foreach (XmlNode xnode in xRoot)
                {
                    foreach (XmlNode childnode in xnode.ChildNodes)
                    {
                        if (childnode.Name == "product")
                        {
                            foreach (XmlNode childofchildnode in childnode.ChildNodes)
                            {
                                DataRow row = application_table.NewRow();
                                if (childofchildnode.Name == "name")
                                {
                                    product_name = $"{childofchildnode.InnerText}";
                                }
                                if (childofchildnode.Name == "product_count")
                                {
                                    product_count = Int32.Parse(childofchildnode.InnerText);
                                }
                                if (childofchildnode.Name == "paper_id")
                                {
                                    foreach (XmlNode third_node in childofchildnode.ChildNodes)
                                    {
                                        DataRow row_for_paper_table = paper_table.NewRow();
                                        if (third_node.Name == "paper_type")
                                        {
                                            temp_paper_type = Int32.Parse(third_node.InnerText);
                                        }
                                        if (third_node.Name == "paper_count")
                                        {
                                            temp_paper_count = Int32.Parse(third_node.InnerText);
                                        }
                                        if (third_node.Name == "defect")
                                        {
                                            temp_defect = Double.Parse(third_node.InnerText);
                                            row_for_paper_table.ItemArray = new object[] { counter, temp_paper_type, temp_paper_count, temp_defect };
                                            paper_table.Rows.Add(row_for_paper_table);
                                        }
                                    }
                                }
                                if (childofchildnode.Name == "color_paint")
                                {
                                    color_paint = Int32.Parse(childofchildnode.InnerText);
                                }
                                if (childofchildnode.Name == "hard_leaf")
                                {
                                    hard_leaf = Int32.Parse(childofchildnode.InnerText);
                                }
                                if (childofchildnode.Name == "master")
                                {
                                    master = Int32.Parse(childofchildnode.InnerText);
                                }
                                if (childofchildnode.Name == "staple_type")
                                {
                                    staple_type = Int32.Parse(childofchildnode.InnerText);
                                }
                                if (childofchildnode.Name == "staple_count")
                                {
                                    paper_id = counter;
                                    staple_count = Int32.Parse(childofchildnode.InnerText);
                                    row.ItemArray = new object[] { counter, product_name, product_count, paper_id, master, staple_type, staple_count, color_paint, hard_leaf};
                                    application_table.Rows.Add(row);
                                    counter++;

                                }
                            }
                        }
                    }
                }

                for (int i = 0; i < application_table.Rows.Count; i++)
                {
                    dataGridView1.Rows.Add();
                    for (int j = 0; j < 3; j++)
                    {
                        dataGridView1.Rows[i].Cells[j].Value = application_table.Rows[i][j];
                    }
                }

                if (dataGridView1.Rows.Count >= 1 && dataGridView1.Columns.Count >= 1)
                    dataGridView1.Rows[0].Cells[1].Selected = true;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Необходимые файлы отсутствуют, переустановите программу", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
        private void comboBox2_DropDown(object sender, EventArgs e)
        {
            if (comboBox2.Items.Contains("                    +"))
                comboBox2.Items.Remove("                    +");
        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
        }

        private void comboBox2_SelectionChangeCommitted(object sender, EventArgs e)
        {
            
        }

        //Заполняется таблица по выбранной бумаге, и записывается индекс бумани в datatable
        private void comboBox2_DropDownClosed(object sender, EventArgs e)
        {
            if (comboBox2.SelectedItem != null)
            {
                int row_datagrid1 = dataGridView2.Rows.Add();
                dataGridView2.Rows[row_datagrid1].Cells[0].Value = dataGridView2.Rows.Count;
                dataGridView2.Rows[row_datagrid1].Cells[1].Value = comboBox2.SelectedItem;
                dataGridView2.Rows[row_datagrid1].Cells[2].Value = "";
                dataGridView2.Rows[row_datagrid1].Cells[3].Value = "";

                int temp_index_of_paper = 0;
                for(int i = 0; i < dataGridView4.Rows.Count; i++)
                {
                    if(comboBox2.SelectedItem.ToString() == dataGridView4.Rows[i].Cells[1].Value.ToString())
                    {
                        temp_index_of_paper = i;
                    }
                }

                DataRow row_for_paper_table = paper_table.NewRow();
                int temp_id = dataGridView1.CurrentCell.RowIndex + 1;
                int temp_paper_type = temp_index_of_paper;
                int temp_paper_count = 0;
                double temp_defect = 0;

                row_for_paper_table.ItemArray = new object[] {temp_id, temp_paper_type, temp_paper_count, temp_defect};
                paper_table.Rows.Add(row_for_paper_table);

               comboBox2.Items.Add("                    +");
                int index = 0;
                for (int i = 0; i < comboBox2.Items.Count; i++)
                {
                    if (comboBox2.Items[i].ToString() == "                    +")
                        index = i;
                }
                comboBox2.SelectedIndex = index;
            }
            else
            {
                comboBox2.Items.Add("                    +");
                int index = 0;
                for (int i = 0; i < comboBox2.Items.Count; i++)
                {
                    if (comboBox2.Items[i].ToString() == "                    +")
                        index = i;
                }
                comboBox2.SelectedIndex = index;
            }
        }

        private void comboBox3_DropDownClosed(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex == 0)
            {
                application_table.Rows[dataGridView1.SelectedCells[0].RowIndex][5] = 99;
            }
            else
            {
                application_table.Rows[dataGridView1.SelectedCells[0].RowIndex][5] = comboBox3.SelectedIndex - 1;
            }
            if (comboBox3.SelectedIndex != 0)
                numericUpDown1.Enabled = true;
            else
            {
                application_table.Rows[dataGridView1.SelectedCells[0].RowIndex][6] = 0;
                numericUpDown1.Value = 0;
                numericUpDown1.Enabled = false;
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            int row_datagrid3 = dataGridView3.Rows.Add();
            dataGridView3.Rows[row_datagrid3].Cells[0].Value = "";
            dataGridView3.Rows[row_datagrid3].Cells[1].Value = "";
            comboBox3.Items.Clear();
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                comboBox3.Items.Add(dataGridView3.Rows[i].Cells[0].Value);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (dataGridView3.Rows.Count != 0)
            {
                int selected_row = dataGridView3.CurrentCell.RowIndex;
                dataGridView3.Rows.RemoveAt(selected_row);
                dataGridView3.Refresh();
            }
            comboBox3.Items.Clear();
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                comboBox3.Items.Add(dataGridView3.Rows[i].Cells[0].Value);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            int row_datagrid4 = dataGridView4.Rows.Add();
            dataGridView4.Rows[row_datagrid4].Cells[0].Value = dataGridView4.Rows.Count;
            dataGridView4.Rows[row_datagrid4].Cells[1].Value = "";
            dataGridView4.Rows[row_datagrid4].Cells[2].Value = "";
            dataGridView4.Rows[row_datagrid4].Cells[3].Value = "";
            dataGridView4.Rows[row_datagrid4].Cells[4].Value = "";
            dataGridView4.Rows[row_datagrid4].Cells[5].Value = false;
            comboBox2.Items.Clear();
            for (int i = 0; i < dataGridView4.Rows.Count; i++)
            {
                if (Convert.ToBoolean(dataGridView4.Rows[i].Cells[5].Value) == true)
                    comboBox2.Items.Add(dataGridView4.Rows[i].Cells[1].Value);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (dataGridView4.Rows.Count != 0)
            {
                int selected_row = dataGridView4.CurrentCell.RowIndex;
                dataGridView4.Rows.RemoveAt(selected_row);
                for(int i = 0; i < dataGridView4.Rows.Count; i++)
                {
                    dataGridView4.Rows[i].Cells[0].Value = i+1;
                }
                dataGridView4.Refresh();
            }
            comboBox2.Items.Clear();
            for (int i = 0; i < dataGridView4.Rows.Count; i++)
            {
                if (Convert.ToBoolean(dataGridView4.Rows[i].Cells[5].Value) == true)
                    comboBox2.Items.Add(dataGridView4.Rows[i].Cells[1].Value);
            }
        }

        //Добавоение нового элемента
        private void button6_Click(object sender, EventArgs e)
        {
            int row_datagrid1 = dataGridView1.Rows.Add();
            dataGridView1.Rows[row_datagrid1].Cells[0].Value = dataGridView1.Rows.Count;
            dataGridView1.Rows[row_datagrid1].Cells[1].Value = "";
            dataGridView1.Rows[row_datagrid1].Cells[2].Value = "";

            DataRow row = application_table.NewRow();
            row.ItemArray = new object[] { dataGridView1.Rows.Count, "", 0, dataGridView1.Rows.Count, 0, 99, 0, 0, 0};
            application_table.Rows.Add(row);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                DataTable temp_paper_table = new DataTable("temp_paper");
                DataColumn PaperId = new DataColumn("Id", Type.GetType("System.Int32"));
                DataColumn paper_type_for_paper_table = new DataColumn("paper_type", Type.GetType("System.Int32"));
                DataColumn paper_count_for_paper_table = new DataColumn("paper_count", Type.GetType("System.Int32"));
                DataColumn defect_for_paper_table = new DataColumn("defect", Type.GetType("System.Double"));
                temp_paper_table.Columns.Add(PaperId);
                temp_paper_table.Columns.Add(paper_type_for_paper_table);
                temp_paper_table.Columns.Add(paper_count_for_paper_table);
                temp_paper_table.Columns.Add(defect_for_paper_table);

                int selected_row = dataGridView1.CurrentCell.RowIndex;
                int index = 0;
                while(index < paper_table.Rows.Count)
                {
                    if (paper_table.Rows[index][0].ToString() == application_table.Rows[selected_row][0].ToString())
                    {
                        index = index + 1;
                    }
                    else
                    {
                        DataRow temp_row = temp_paper_table.NewRow();
                        int temp_counter = Int32.Parse(paper_table.Rows[index][0].ToString());
                        if (Int32.Parse(paper_table.Rows[index][0].ToString()) < Int32.Parse(application_table.Rows[selected_row][0].ToString()))
                        {
                            temp_row.ItemArray = new object[] { temp_counter, paper_table.Rows[index][1], paper_table.Rows[index][2], paper_table.Rows[index][3] };
                        }
                        else
                        {
                            temp_counter = temp_counter - 1;
                            temp_row.ItemArray = new object[] { temp_counter, paper_table.Rows[index][1], paper_table.Rows[index][2], paper_table.Rows[index][3] };
                        }
                        temp_paper_table.Rows.Add(temp_row);
                        index = index + 1;
                    }
                }

                paper_table.Clear();
                paper_table = temp_paper_table;

                dataGridView1.Rows.RemoveAt(selected_row);
                application_table.Rows.RemoveAt(selected_row);
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.Rows[i].Cells[0].Value = i + 1;
                    application_table.Rows[i][0] = i + 1;
                }
                dataGridView1.Refresh();
                dataGridView2.Rows.Clear();
                dataGridView2.Refresh();
            }
        }

        //Добавление использованной бумаги в dtgd2 на основании какая строка выбрана в dtgd1
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < application_table.Rows.Count; i++)
            {
                try
                {
                    if (Int32.Parse(application_table.Rows[i][0].ToString()) == Int32.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString()))
                    {
                        dataGridView2.Rows.Clear();
                        numericUpDown2.Value = 0;
                        numericUpDown1.Value = 0;
                        comboBox3.SelectedIndex = 0;
                        for (int j = 0; j < paper_table.Rows.Count; j++)
                        {
                            if (application_table.Rows[i][0].ToString() == paper_table.Rows[j][0].ToString())
                            {
                                int row_datagrid2 = dataGridView2.Rows.Add();
                                dataGridView2.Rows[row_datagrid2].Cells[0].Value = dataGridView2.Rows.Count;
                                dataGridView2.Rows[row_datagrid2].Cells[1].Value = dataGridView4.Rows[Int32.Parse(paper_table.Rows[j][1].ToString())].Cells[1].Value;
                                dataGridView2.Rows[row_datagrid2].Cells[2].Value = paper_table.Rows[j][2];
                                dataGridView2.Rows[row_datagrid2].Cells[3].Value = paper_table.Rows[j][3];
                            }
                        }
                        numericUpDown2.Value = Int32.Parse(application_table.Rows[i][4].ToString());
                        if (Int32.Parse(application_table.Rows[i][5].ToString()) == 0 && dataGridView1.Rows[i].Cells[1].Value.ToString() == "Новый продукт" &&
                            Int32.Parse(dataGridView1.Rows[i].Cells[2].Value.ToString()) == 0)
                        {
                            comboBox3.SelectedIndex = Int32.Parse(application_table.Rows[i][5].ToString());
                        }
                        else
                        {
                            if (Int32.Parse(application_table.Rows[i][5].ToString()) == 99)
                            {
                                comboBox3.SelectedIndex = 0;
                            }
                            else
                            {
                                comboBox3.SelectedIndex = Int32.Parse(application_table.Rows[i][5].ToString()) + 1;
                            }
                        }
                        numericUpDown1.Value = Int32.Parse(application_table.Rows[i][6].ToString());
                        if (Int32.Parse(application_table.Rows[i][7].ToString()) == 1)
                        {
                            checkBox1.Checked = true;
                        }
                        else
                        {
                            checkBox1.Checked = false;
                        }

                        if (Int32.Parse(application_table.Rows[i][8].ToString()) == 1)
                        {
                            checkBox14.Checked = true;
                            checkBox2.Checked = true;
                        }
                        if (Int32.Parse(application_table.Rows[i][8].ToString()) == 2)
                        {
                            checkBox14.Checked = true;
                            checkBox3.Checked = true;
                        }
                        if (Int32.Parse(application_table.Rows[i][8].ToString()) == 3)
                        {
                            checkBox14.Checked = true;
                            checkBox13.Checked = true;
                        }
                        else
                        {
                            checkBox14.Checked = false;
                            checkBox2.Checked = false;
                            checkBox3.Checked = false;
                            checkBox13.Checked = false;
                        }

                    }
                }
                catch (Exception ex)
                {
                }
            }
        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void dataGridView4_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            comboBox2.Items.Clear();
            for (int i = 0; i < dataGridView4.Rows.Count; i++)
            {
                if (Convert.ToBoolean(dataGridView4.Rows[i].Cells[5].Value) == true)
                    comboBox2.Items.Add(dataGridView4.Rows[i].Cells[1].Value);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            numericUpDown2.Value = 0;
            comboBox3.SelectedIndex = 0;
            numericUpDown1.Value = 0;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count != 0)
            {
                try
                {
                    int selected_row_from_dtgrd1 = dataGridView1.CurrentCell.RowIndex;
                    int selected_row = dataGridView2.CurrentCell.RowIndex;
                    for (int i = 0; i < paper_table.Rows.Count; i++)
                    {
                        if (Int32.Parse(paper_table.Rows[i][0].ToString()) == selected_row_from_dtgrd1 + 1 &&
                        dataGridView4.Rows[Int32.Parse(paper_table.Rows[i][1].ToString())].Cells[1].Value.ToString() == dataGridView2.Rows[selected_row].Cells[1].Value.ToString())
                        {
                            paper_table.Rows.RemoveAt(i);
                        }
                    }
                    dataGridView2.Rows.RemoveAt(selected_row);
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        dataGridView2.Rows[i].Cells[0].Value = i + 1;
                    }
                    dataGridView2.Refresh();
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Добавьте хоть одно наименование продукции");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int flag = 0;
            if(dataGridView1.Rows.Count == 0)
            {
                flag = 1;
                MessageBox.Show("Считать нечего");
            }
            List<int> temp_mass1 = new List<int>();
            for (int i = 0; i < application_table.Rows.Count; i++)
            {
                temp_mass1.Add(Int32.Parse(application_table.Rows[i][0].ToString()));
            }
            var result1 = temp_mass1.Distinct().ToArray();
            List<int> temp_mass2 = new List<int>();
            for (int i = 0; i < paper_table.Rows.Count; i++)
            {
                temp_mass2.Add(Int32.Parse(paper_table.Rows[i][0].ToString()));
            }
            var result2 = temp_mass2.Distinct().ToArray();
            if (result1.Length != result2.Length)
            {
                    MessageBox.Show("Заполните данные по всем видама продукции");
                    flag = 1;
            }
            for(int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if(dataGridView1.Rows[i].Cells[2].Value.ToString() == "0")
                {
                    MessageBox.Show("Количество наименований не может равняться 0");
                    flag = 1;
                }
            }
            if (flag != 1)
            {
                Form2 form2 = new Form2();
                form2.dtgrd4 = this.dataGridView4;
                form2.dtgrd3 = this.dataGridView3;
                form2.dtgrd1 = this.dataGridView1;
                form2.master_price = this.label12;
                form2.price_of_imprint = this.textBox6;
                form2.price_of_tube = this.textBox5;
                form2.app_table = this.application_table;
                form2.ppr_table = this.paper_table;
                form2.color_paint_cost = this.textBox8;
                form2.channel_costA3 = this.textBox12;
                form2.cover_costA3 = this.textBox13;
                form2.channel_costA4 = this.textBox16;
                form2.cover_costA4 = this.textBox15;
                form2.channel_costA5 = this.textBox18;
                form2.cover_costA5 = this.textBox17;
                form2.ShowDialog();

            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            string text = textBox4.Text;
            Errors_checker cheker = new Errors_checker();
            textBox4.Text = cheker.textBox_checker(text);
            if (textBox4.Text != "")
            {
                double cost = Double.Parse(textBox4.Text) / Double.Parse(numericUpDown3.Value.ToString());
                label12.Text = Math.Round(cost, 4).ToString();
            }
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            if (textBox4.Text != "")
            {
                double cost = Double.Parse(textBox4.Text) / Double.Parse(numericUpDown3.Value.ToString());
                label12.Text = Math.Round(cost, 4).ToString();
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            for (int i = 0; i < application_table.Rows.Count; i++)
            {
                if (Int32.Parse(application_table.Rows[i][0].ToString()) == Int32.Parse(dataGridView1.CurrentCell.RowIndex.ToString()) + 1)
                {
                    application_table.Rows[i][1] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[1].Value;
                    if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value.ToString() != "")
                    {
                        application_table.Rows[i][2] = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[2].Value;
                    }
                    else
                        application_table.Rows[i][2] = 0;
                }
            }
        }

        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int selected_row = dataGridView2.CurrentCell.RowIndex;
                int second_temp_index = Int32.Parse(dataGridView1.CurrentCell.RowIndex.ToString());
                string text = dataGridView2.CurrentRow.Cells[3].Value.ToString();
                Errors_checker cheker = new Errors_checker();
                dataGridView2.CurrentRow.Cells[3].Value = cheker.textBox_checker(text);
                for(int i = 0; i < paper_table.Rows.Count; i++)
                {
                    if (Int32.Parse(paper_table.Rows[i][0].ToString()) == second_temp_index + 1 && 
                        dataGridView4.Rows[Int32.Parse(paper_table.Rows[i][1].ToString())].Cells[1].Value.ToString() == dataGridView2.Rows[selected_row].Cells[1].Value.ToString())
                    {
                        if (dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[2].Value.ToString() != "")
                            paper_table.Rows[i][2] = Int32.Parse(dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[2].Value.ToString());
                        else
                            paper_table.Rows[i][2] = 0;
                        if (dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[3].Value.ToString() != "")
                            paper_table.Rows[i][3] = dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[3].Value;
                        else
                            paper_table.Rows[i][3] = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                //dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[2].Value = 0;
                //MessageBox.Show("Количество должно быть целочисленным");
                MessageBox.Show(ex.Message);
            }
        }

        private void numericUpDown2_Leave(object sender, EventArgs e)
        {
            if(numericUpDown2.Text == "")
            {
                numericUpDown2.Value = 0;
            }
            application_table.Rows[dataGridView1.SelectedCells[0].RowIndex][4] = numericUpDown2.Value;
        }

        private void numericUpDown1_Leave(object sender, EventArgs e)
        {
            application_table.Rows[dataGridView1.SelectedCells[0].RowIndex][6] = numericUpDown1.Value;
        }

        //ОТМЕНИТЬ ВСЕ ИЗМЕНЕНИЯ, ПРИ НАЖАТИИ НА КНОПКУ ИДЕТ ОЧИСКА ВСЕХ ТАБЛИЦ, А ПОТОМ ЗАГРУЗКА В НИХ ДАННЫХ ИЗ ФАЙЛА
        private void button4_Click(object sender, EventArgs e)
        {
            application_table.Rows.Clear();
            paper_table.Rows.Clear();
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            dataGridView4.Rows.Clear();
            comboBox2.Items.Clear();
            comboBox3.Items.Clear();
            comboBox3.Items.Add("Нет");
            numericUpDown1.Value = 0;
            numericUpDown2.Value = 0;

            XmlDocument new_doc = new XmlDocument();
            new_doc.Load("consumables.xml");
            XmlElement new_elem = new_doc.DocumentElement;
            foreach (XmlNode xnode in new_elem)
            {
                if (xnode.Name == "staples_type")
                {
                    foreach (XmlNode childnode_of_child_0 in xnode.ChildNodes)
                    {
                        if (childnode_of_child_0.Name == "staple")
                        {
                            int row_datagrid3 = dataGridView3.Rows.Add();
                            foreach (XmlNode childnode_of_child_1 in childnode_of_child_0.ChildNodes)
                            {
                                if (childnode_of_child_1.Name == "name")
                                {
                                    dataGridView3.Rows[row_datagrid3].Cells[0].Value = $"{childnode_of_child_1.InnerText}";
                                    comboBox3.Items.Add(childnode_of_child_1.InnerText);
                                }
                                if (childnode_of_child_1.Name == "count")
                                    dataGridView3.Rows[row_datagrid3].Cells[1].Value = $"{childnode_of_child_1.InnerText}";
                                if (childnode_of_child_1.Name == "price")
                                    dataGridView3.Rows[row_datagrid3].Cells[2].Value = $"{childnode_of_child_1.InnerText}";
                            }
                            double price = Double.Parse(dataGridView3.Rows[row_datagrid3].Cells[2].Value.ToString()) / Double.Parse(dataGridView3.Rows[row_datagrid3].Cells[1].Value.ToString());
                            dataGridView3.Rows[row_datagrid3].Cells[3].Value = price;
                        }
                    }
                }
                string name = "";
                if (xnode.Name == "paper_type")
                {
                    foreach (XmlNode childnode_of_child_0 in xnode.ChildNodes)
                    {
                        if (childnode_of_child_0.Name == "paper")
                        {
                            int row_datagrid4 = dataGridView4.Rows.Add();
                            foreach (XmlNode childnode_of_child_1 in childnode_of_child_0.ChildNodes)
                            {
                                dataGridView4.Rows[row_datagrid4].Cells[0].Value = dataGridView4.Rows.Count;
                                if (childnode_of_child_1.Name == "name")
                                {
                                    dataGridView4.Rows[row_datagrid4].Cells[1].Value = $"{childnode_of_child_1.InnerText}";
                                    name = childnode_of_child_1.InnerText;
                                }
                                if (childnode_of_child_1.Name == "unit")
                                    dataGridView4.Rows[row_datagrid4].Cells[2].Value = $"{childnode_of_child_1.InnerText}";
                                if (childnode_of_child_1.Name == "price")
                                    dataGridView4.Rows[row_datagrid4].Cells[3].Value = $"{childnode_of_child_1.InnerText}";
                                if (childnode_of_child_1.Name == "count")
                                    dataGridView4.Rows[row_datagrid4].Cells[4].Value = $"{childnode_of_child_1.InnerText}";
                                if (childnode_of_child_1.Name == "used")
                                {
                                    dataGridView4.Rows[row_datagrid4].Cells[5].Value = Int32.Parse(childnode_of_child_1.InnerText);
                                    if (Convert.ToBoolean(dataGridView4.Rows[row_datagrid4].Cells[5].Value) == true)
                                        comboBox2.Items.Add(name);
                                }
                            }
                        }
                    }
                }
                if (xnode.Name == "Master")
                {
                    foreach (XmlNode childnode_of_child_0 in xnode.ChildNodes)
                    {
                        if (childnode_of_child_0.Name == "count")
                            numericUpDown3.Value = Int32.Parse(childnode_of_child_0.InnerText);
                        if (childnode_of_child_0.Name == "price")
                            textBox4.Text = $"{childnode_of_child_0.InnerText}";
                    }
                    double count = Double.Parse(numericUpDown3.Value.ToString());
                    double price = Double.Parse(textBox4.Text) / count;
                    label12.Text = price.ToString();
                }
                if (xnode.Name == "Paint")
                {
                    foreach (XmlNode childnode_of_child_0 in xnode.ChildNodes)
                    {
                        if (childnode_of_child_0.Name == "price_of_tube")
                            textBox5.Text = $"{childnode_of_child_0.InnerText}";
                        if (childnode_of_child_0.Name == "prise_of_imprint")
                            textBox6.Text = $"{childnode_of_child_0.InnerText}";
                    }
                }
                if (xnode.Name == "Color_paint")
                {
                    foreach (XmlNode childnode_of_child_0 in xnode.ChildNodes)
                    {
                        if (childnode_of_child_0.Name == "Set_price")
                            textBox7.Text = $"{childnode_of_child_0.InnerText}";
                        if (childnode_of_child_0.Name == "Imprint_price")
                            textBox8.Text = $"{childnode_of_child_0.InnerText}";
                    }
                }
                if (xnode.Name == "Hard_leaf")
                {
                    foreach (XmlNode childnode_of_child_0 in xnode.ChildNodes)
                    {
                        if (childnode_of_child_0.Name == "Channel_priceA3")
                            textBox12.Text = $"{childnode_of_child_0.InnerText}";
                        if (childnode_of_child_0.Name == "Cover_priceA3")
                            textBox13.Text = $"{childnode_of_child_0.InnerText}";
                        if (childnode_of_child_0.Name == "Channel_priceA4")
                            textBox16.Text = $"{childnode_of_child_0.InnerText}";
                        if (childnode_of_child_0.Name == "Cover_priceA4")
                            textBox15.Text = $"{childnode_of_child_0.InnerText}";
                        if (childnode_of_child_0.Name == "Channel_priceA5")
                            textBox18.Text = $"{childnode_of_child_0.InnerText}";
                        if (childnode_of_child_0.Name == "Cover_priceA5")
                            textBox17.Text = $"{childnode_of_child_0.InnerText}";
                    }
                }
            }

            XmlDocument xDoc = new XmlDocument();
            xDoc.Load("application.xml");
            XmlElement xRoot = xDoc.DocumentElement;
            foreach (XmlNode xnode in xRoot)
            {
                foreach (XmlNode childnode in xnode.ChildNodes)
                {
                    if (childnode.Name == "application_number")
                    {
                        numericUpDown4.Value = Int32.Parse(childnode.InnerText);
                    }
                    if (childnode.Name == "applicant")
                    {
                        textBox1.Text = $"{childnode.InnerText}";
                    }
                    if (childnode.Name == "applicant_rank")
                    {
                        comboBox4.SelectedIndex = Int32.Parse(childnode.InnerText);
                    }
                    if (childnode.Name == "applicant_name")
                    {
                        textBox14.Text = $"{childnode.InnerText}";
                    }
                    if (childnode.Name == "employee_position")
                    {
                        textBox2.Text = $"{childnode.InnerText}";
                    }
                    if (childnode.Name == "employee_rank")
                    {
                        comboBox1.SelectedIndex = Int32.Parse(childnode.InnerText);
                    }
                    if (childnode.Name == "employee_name")
                    {
                        textBox3.Text = $"{childnode.InnerText}";
                    }
                }
            }

            int counter = 1;
            string product_name = "";
            int product_count = 0;
            int master = 0;
            int staple_type = 0;
            int staple_count = 0;
            int paper_id = 0;
            int color_paint = 0;
            int hard_leaf = 0;
            int temp_paper_type = 0;
            int temp_paper_count = 0;
            double temp_defect = 0;
            foreach (XmlNode xnode in xRoot)
            {
                foreach (XmlNode childnode in xnode.ChildNodes)
                {
                    if (childnode.Name == "product")
                    {
                        foreach (XmlNode childofchildnode in childnode.ChildNodes)
                        {
                            DataRow row = application_table.NewRow();
                            if (childofchildnode.Name == "name")
                            {
                                product_name = $"{childofchildnode.InnerText}";
                            }
                            if (childofchildnode.Name == "product_count")
                            {
                                product_count = Int32.Parse(childofchildnode.InnerText);
                            }
                            if (childofchildnode.Name == "paper_id")
                            {
                                foreach (XmlNode third_node in childofchildnode.ChildNodes)
                                {
                                    DataRow row_for_paper_table = paper_table.NewRow();
                                    if (third_node.Name == "paper_type")
                                    {
                                        temp_paper_type = Int32.Parse(third_node.InnerText);
                                    }
                                    if (third_node.Name == "paper_count")
                                    {
                                        temp_paper_count = Int32.Parse(third_node.InnerText);
                                    }
                                    if (third_node.Name == "defect")
                                    {
                                        temp_defect = Double.Parse(third_node.InnerText);
                                        row_for_paper_table.ItemArray = new object[] { counter, temp_paper_type, temp_paper_count, temp_defect };
                                        paper_table.Rows.Add(row_for_paper_table);
                                    }
                                }
                            }
                            if (childofchildnode.Name == "color_paint")
                            {
                                color_paint = Int32.Parse(childofchildnode.InnerText);
                            }
                            if (childofchildnode.Name == "hard_leaf")
                            {
                                hard_leaf = Int32.Parse(childofchildnode.InnerText);
                            }
                            if (childofchildnode.Name == "master")
                            {
                                master = Int32.Parse(childofchildnode.InnerText);
                            }
                            if (childofchildnode.Name == "staple_type")
                            {
                                staple_type = Int32.Parse(childofchildnode.InnerText);
                            }
                            if (childofchildnode.Name == "staple_count")
                            {
                                paper_id = counter;
                                staple_count = Int32.Parse(childofchildnode.InnerText);
                                row.ItemArray = new object[] { counter, product_name, product_count, paper_id, master, staple_type, staple_count, color_paint, hard_leaf };
                                application_table.Rows.Add(row);
                                counter++;

                            }
                        }
                    }
                }
            }

            for (int i = 0; i < application_table.Rows.Count; i++)
            {
                dataGridView1.Rows.Add();
                for (int j = 0; j < 3; j++)
                {
                    dataGridView1.Rows[i].Cells[j].Value = application_table.Rows[i][j];
                }
            }

            if (dataGridView1.Rows.Count >= 1 && dataGridView1.Columns.Count >= 1)
                dataGridView1.Rows[0].Cells[1].Selected = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int flag = 0;
            if (dataGridView1.Rows.Count == 0)
            {
                flag = 1;
                MessageBox.Show("Считать нечего");
            }
            for (int i = 0; i < application_table.Rows.Count; i++)
            {
                if (application_table.Rows[i][3].ToString() == "")
                {
                    MessageBox.Show("Заполните данные по продукции в " + application_table.Rows[i][1].ToString());
                    flag = 1;
                }
            }
            if (flag == 0)
            {
                XmlDocument save_Doc = new XmlDocument();
                save_Doc.Load("application.xml");
                
                
                XmlElement xRoot = save_Doc.DocumentElement;
                xRoot.RemoveAll();

                XmlElement application_tab = save_Doc.CreateElement("application_tab");
                xRoot.AppendChild(application_tab);

                XmlElement application_number = save_Doc.CreateElement("application_number");
                XmlText application_number_text = save_Doc.CreateTextNode(numericUpDown4.Value.ToString());
                application_number.AppendChild(application_number_text);

                XmlElement applicant = save_Doc.CreateElement("applicant");
                XmlText applicant_text = save_Doc.CreateTextNode(textBox1.Text);
                applicant.AppendChild(applicant_text);

                XmlElement applicant_rank = save_Doc.CreateElement("applicant_rank");
                XmlText applicant_rank_text = save_Doc.CreateTextNode(comboBox4.SelectedIndex.ToString());
                applicant_rank.AppendChild(applicant_rank_text);

                XmlElement applicant_number = save_Doc.CreateElement("applicant_name");
                XmlText applicant_number_text = save_Doc.CreateTextNode(textBox14.Text);
                applicant_number.AppendChild(applicant_number_text);

                XmlElement employee_position = save_Doc.CreateElement("employee_position");
                XmlText employee_position_text = save_Doc.CreateTextNode(textBox2.Text);
                employee_position.AppendChild(employee_position_text);

                XmlElement employee_rank = save_Doc.CreateElement("employee_rank");
                XmlText employee_rank_text = save_Doc.CreateTextNode(comboBox1.SelectedIndex.ToString());
                employee_rank.AppendChild(employee_rank_text);

                XmlElement employee_name = save_Doc.CreateElement("employee_name");
                XmlText employee_name_text = save_Doc.CreateTextNode(textBox3.Text);
                employee_name.AppendChild(employee_name_text);

                application_tab.AppendChild(application_number);
                application_tab.AppendChild(applicant);
                application_tab.AppendChild(applicant_rank);
                application_tab.AppendChild(applicant_number);
                application_tab.AppendChild(employee_position);
                application_tab.AppendChild(employee_rank);
                application_tab.AppendChild(employee_name);

                XmlElement production_tab = save_Doc.CreateElement("production_tab");
                xRoot.AppendChild(production_tab);

                for (int i = 0; i < application_table.Rows.Count; i++)
                {
                    XmlElement xmlroot = save_Doc.DocumentElement;
                    XmlElement product = save_Doc.CreateElement("product");
                    xmlroot.LastChild.AppendChild(product);

                    XmlElement name = save_Doc.CreateElement("name");
                    xmlroot.LastChild.LastChild.AppendChild(name);
                    XmlText name_text = save_Doc.CreateTextNode(application_table.Rows[i][1].ToString());
                    xmlroot.LastChild.LastChild.LastChild.AppendChild(name_text);

                    XmlElement product_count = save_Doc.CreateElement("product_count");
                    xmlroot.LastChild.LastChild.AppendChild(product_count);
                    XmlText product_count_text = save_Doc.CreateTextNode(application_table.Rows[i][2].ToString());
                    xmlroot.LastChild.LastChild.LastChild.AppendChild(product_count_text);

                    for (int j = 0; j < paper_table.Rows.Count; j++)
                    {
                        if (application_table.Rows[i][0].ToString() == paper_table.Rows[j][0].ToString())
                        {
                            XmlElement paper_id = save_Doc.CreateElement("paper_id");
                            xmlroot.LastChild.LastChild.AppendChild(paper_id);

                            XmlElement paper_type = save_Doc.CreateElement("paper_type");
                            xmlroot.LastChild.LastChild.LastChild.AppendChild(paper_type);
                            XmlText paper_type_text = save_Doc.CreateTextNode(paper_table.Rows[j][1].ToString());
                            xmlroot.LastChild.LastChild.LastChild.LastChild.AppendChild(paper_type_text);

                            XmlElement paper_count = save_Doc.CreateElement("paper_count");
                            xmlroot.LastChild.LastChild.LastChild.AppendChild(paper_count);
                            XmlText paper_count_text = save_Doc.CreateTextNode(paper_table.Rows[j][2].ToString());
                            xmlroot.LastChild.LastChild.LastChild.LastChild.AppendChild(paper_count_text);

                            XmlElement defect = save_Doc.CreateElement("defect");
                            xmlroot.LastChild.LastChild.LastChild.AppendChild(defect);
                            XmlText defect_text = save_Doc.CreateTextNode(paper_table.Rows[i][3].ToString());
                            xmlroot.LastChild.LastChild.LastChild.LastChild.AppendChild(defect_text);
                        }
                    }
                    XmlElement color_paint = save_Doc.CreateElement("color_paint");
                    xmlroot.LastChild.LastChild.AppendChild(color_paint);
                    XmlText color_paint_text = save_Doc.CreateTextNode(application_table.Rows[i][7].ToString());
                    xmlroot.LastChild.LastChild.LastChild.AppendChild(color_paint_text);

                    XmlElement hard_leaf = save_Doc.CreateElement("hard_leaf");
                    xmlroot.LastChild.LastChild.AppendChild(hard_leaf);
                    XmlText hard_leaf_text = save_Doc.CreateTextNode(application_table.Rows[i][8].ToString());
                    xmlroot.LastChild.LastChild.LastChild.AppendChild(hard_leaf_text);

                    XmlElement master = save_Doc.CreateElement("master");
                    xmlroot.LastChild.LastChild.AppendChild(master);
                    XmlText master_text = save_Doc.CreateTextNode(application_table.Rows[i][4].ToString());
                    xmlroot.LastChild.LastChild.LastChild.AppendChild(master_text);

                    XmlElement staple_type = save_Doc.CreateElement("staple_type");
                    xmlroot.LastChild.LastChild.AppendChild(staple_type);
                    XmlText staple_type_text = save_Doc.CreateTextNode(application_table.Rows[i][5].ToString());
                    xmlroot.LastChild.LastChild.LastChild.AppendChild(staple_type_text);

                    XmlElement staple_count = save_Doc.CreateElement("staple_count");
                    xmlroot.LastChild.LastChild.AppendChild(staple_count);
                    XmlText staple_count_text = save_Doc.CreateTextNode(application_table.Rows[i][6].ToString());
                    xmlroot.LastChild.LastChild.LastChild.AppendChild(staple_count_text);


                    save_Doc.Save("application.xml");
                }

                XmlDocument new_doc = new XmlDocument();
                new_doc.Load("consumables.xml");
                XmlElement new_elem = new_doc.DocumentElement;
                new_elem.RemoveAll();

                XmlElement staples_type = new_doc.CreateElement("staples_type");
                new_elem.AppendChild(staples_type);

                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    XmlElement staple = new_doc.CreateElement("staple");
                    new_elem.LastChild.AppendChild(staple);

                    XmlElement name = new_doc.CreateElement("name");
                    new_elem.LastChild.LastChild.AppendChild(name);
                    XmlText name_text = new_doc.CreateTextNode(dataGridView3.Rows[i].Cells[0].Value.ToString());
                    new_elem.LastChild.LastChild.LastChild.AppendChild(name_text);

                    XmlElement count = new_doc.CreateElement("count");
                    new_elem.LastChild.LastChild.AppendChild(count);
                    XmlText count_text = new_doc.CreateTextNode(dataGridView3.Rows[i].Cells[1].Value.ToString());
                    new_elem.LastChild.LastChild.LastChild.AppendChild(count_text);

                    XmlElement price = new_doc.CreateElement("price");
                    new_elem.LastChild.LastChild.AppendChild(price);
                    XmlText price_text = new_doc.CreateTextNode(dataGridView3.Rows[i].Cells[2].Value.ToString());
                    new_elem.LastChild.LastChild.LastChild.AppendChild(price_text);

                    new_doc.Save("consumables.xml");
                }

                XmlElement paper_type_second = new_doc.CreateElement("paper_type");
                new_elem.AppendChild(paper_type_second);

                for (int i = 0; i < dataGridView4.Rows.Count; i++)
                {
                    XmlElement paper = new_doc.CreateElement("paper");
                    new_elem.LastChild.AppendChild(paper);

                    XmlElement name = new_doc.CreateElement("name");
                    new_elem.LastChild.LastChild.AppendChild(name);
                    XmlText name_text = new_doc.CreateTextNode(dataGridView4.Rows[i].Cells[1].Value.ToString());
                    new_elem.LastChild.LastChild.LastChild.AppendChild(name_text);

                    XmlElement unit = new_doc.CreateElement("unit");
                    new_elem.LastChild.LastChild.AppendChild(unit);
                    XmlText unit_text = new_doc.CreateTextNode(dataGridView4.Rows[i].Cells[2].Value.ToString());
                    new_elem.LastChild.LastChild.LastChild.AppendChild(unit_text);

                    XmlElement price = new_doc.CreateElement("price");
                    new_elem.LastChild.LastChild.AppendChild(price);
                    XmlText price_text = new_doc.CreateTextNode(dataGridView4.Rows[i].Cells[3].Value.ToString());
                    new_elem.LastChild.LastChild.LastChild.AppendChild(price_text);

                    XmlElement count = new_doc.CreateElement("count");
                    new_elem.LastChild.LastChild.AppendChild(count);
                    XmlText count_text = new_doc.CreateTextNode(dataGridView4.Rows[i].Cells[4].Value.ToString());
                    new_elem.LastChild.LastChild.LastChild.AppendChild(count_text);

                    XmlElement used = new_doc.CreateElement("used");
                    new_elem.LastChild.LastChild.AppendChild(used);
                    XmlText used_text = new_doc.CreateTextNode("1");
                    if (Convert.ToBoolean(dataGridView4.Rows[i].Cells[5].Value) == true)
                        used_text = new_doc.CreateTextNode("1");
                    else
                        used_text = new_doc.CreateTextNode("0");
                    new_elem.LastChild.LastChild.LastChild.AppendChild(used_text);

                    new_doc.Save("consumables.xml");
                }

                XmlElement master_tab = new_doc.CreateElement("Master");
                new_elem.AppendChild(master_tab);

                XmlElement count_of_master = new_doc.CreateElement("count");
                XmlText count_of_master_text = new_doc.CreateTextNode(numericUpDown3.Value.ToString());
                count_of_master.AppendChild(count_of_master_text);
                master_tab.AppendChild(count_of_master);

                XmlElement price_of_master = new_doc.CreateElement("price");
                XmlText price_of_master_text = new_doc.CreateTextNode(textBox4.Text);
                price_of_master.AppendChild(price_of_master_text);
                master_tab.AppendChild(price_of_master);

                XmlElement paint = new_doc.CreateElement("Paint");
                new_elem.AppendChild(paint);

                XmlElement price_of_tube = new_doc.CreateElement("price_of_tube");
                XmlText price_of_tube_text = new_doc.CreateTextNode(textBox5.Text);
                price_of_tube.AppendChild(price_of_tube_text);
                paint.AppendChild(price_of_tube);

                XmlElement prise_of_imprint = new_doc.CreateElement("prise_of_imprint");
                XmlText prise_of_imprint_text = new_doc.CreateTextNode(textBox6.Text);
                prise_of_imprint.AppendChild(prise_of_imprint_text);
                paint.AppendChild(prise_of_imprint);

                XmlElement Color_paint = new_doc.CreateElement("Color_paint");
                new_elem.AppendChild(Color_paint);

                XmlElement Set_price = new_doc.CreateElement("Set_price");
                XmlText Set_price_text = new_doc.CreateTextNode(textBox7.Text);
                Set_price.AppendChild(Set_price_text);
                Color_paint.AppendChild(Set_price);

                XmlElement Imprint_price = new_doc.CreateElement("Imprint_price");
                XmlText Imprint_price_text = new_doc.CreateTextNode(textBox8.Text);
                Imprint_price.AppendChild(Imprint_price_text);
                Color_paint.AppendChild(Imprint_price);

                XmlElement Hard_leaf = new_doc.CreateElement("Hard_leaf");
                new_elem.AppendChild(Hard_leaf);

                XmlElement Channel_priceA3 = new_doc.CreateElement("Channel_priceA3");
                XmlText Channel_priceA3_text = new_doc.CreateTextNode(textBox12.Text);
                Channel_priceA3.AppendChild(Channel_priceA3_text);
                Hard_leaf.AppendChild(Channel_priceA3);

                XmlElement Cover_priceA3 = new_doc.CreateElement("Cover_priceA3");
                XmlText Cover_priceA3_text = new_doc.CreateTextNode(textBox13.Text);
                Cover_priceA3.AppendChild(Cover_priceA3_text);
                Hard_leaf.AppendChild(Cover_priceA3);

                XmlElement Channel_priceA4 = new_doc.CreateElement("Channel_priceA4");
                XmlText Channel_priceA4_text = new_doc.CreateTextNode(textBox16.Text);
                Channel_priceA4.AppendChild(Channel_priceA4_text);
                Hard_leaf.AppendChild(Channel_priceA4);

                XmlElement Cover_priceA4 = new_doc.CreateElement("Cover_priceA4");
                XmlText Cover_priceA4_text = new_doc.CreateTextNode(textBox15.Text);
                Cover_priceA4.AppendChild(Cover_priceA4_text);
                Hard_leaf.AppendChild(Cover_priceA4);

                XmlElement Channel_priceA5 = new_doc.CreateElement("Channel_priceA5");
                XmlText Channel_priceA5_text = new_doc.CreateTextNode(textBox18.Text);
                Channel_priceA5.AppendChild(Channel_priceA5_text);
                Hard_leaf.AppendChild(Channel_priceA5);

                XmlElement Cover_priceA5 = new_doc.CreateElement("Cover_priceA5");
                XmlText Cover_priceA5_text = new_doc.CreateTextNode(textBox17.Text);
                Cover_priceA5.AppendChild(Cover_priceA5_text);
                Hard_leaf.AppendChild(Cover_priceA5);

                new_doc.Save("consumables.xml");
            }
        }
        //НАКЛАДНАЯ
        private void button5_Click(object sender, EventArgs e)
        {
            int flag = 0;
            if (dataGridView1.Rows.Count == 0)
            {
                flag = 1;
                MessageBox.Show("Считать нечего");
            }
            List<int> temp_mass1 = new List<int>();
            for (int i = 0; i < application_table.Rows.Count; i++)
            {
                temp_mass1.Add(Int32.Parse(application_table.Rows[i][0].ToString()));
            }
            var result1 = temp_mass1.Distinct().ToArray();
            List<int> temp_mass2 = new List<int>();
            for (int i = 0; i < paper_table.Rows.Count; i++)
            {
                temp_mass2.Add(Int32.Parse(paper_table.Rows[i][0].ToString()));
            }
            var result2 = temp_mass2.Distinct().ToArray();
            if (result1.Length != result2.Length)
            {
                MessageBox.Show("Заполните данные по всем видама продукции");
                flag = 1;
            }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells[2].Value.ToString() == "0")
                {
                    MessageBox.Show("Количество наименований не может равняться 0");
                    flag = 1;
                }
            }
            if (dataGridView1.Rows.Count == 0)
            {
                flag = 1;
                MessageBox.Show("Считать нечего");
            }
            for (int i = 0; i < application_table.Rows.Count; i++)
            {
                if (application_table.Rows[i][3].ToString() == "")
                {
                    MessageBox.Show("Заполните данные по продукции в " + application_table.Rows[i][1].ToString());
                    flag = 1;
                }
            }
            try
            {
                if (flag == 0)
                {
                    wordapp = new Word.Application();
                    Object template = Type.Missing;
                    Object newTemplate = false;
                    Object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
                    Object visible = true;
                    worddocument = wordapp.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);

                    Object begin = Type.Missing;
                    Object end = Type.Missing;
                    Word.Range wordrange = worddocument.Range(ref begin, ref end);
                    wordrange.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
                    wordrange.PageSetup.TextColumns.SetCount(2);
                    wordrange.PageSetup.LeftMargin = float.Parse(28.3.ToString());
                    wordrange.PageSetup.RightMargin = float.Parse(28.3.ToString());
                    wordrange.PageSetup.TopMargin = float.Parse(28.3.ToString());
                    wordrange.PageSetup.BottomMargin = float.Parse(28.3.ToString());
                    wordrange.Font.Size = 12;
                    wordrange.Font.Name = "Times New Roman";
                    wordparagraphs = worddocument.Paragraphs;

                    wordrange = worddocument.Range(0, 0);
                    wordrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    wordparagraph = (Word.Paragraph)wordparagraphs[1];
                    wordparagraph.Range.Font.Size = 10;
                    wordrange.Text = "Форма 10";

                    object oMissing = System.Reflection.Missing.Value;
                    worddocument.Paragraphs.Add(ref oMissing);
                    wordparagraph = (Word.Paragraph)wordparagraphs[2];
                    wordparagraph.Range.Font.Size = 12;
                    wordparagraph.Range.Font.Bold = 1;
                    wordparagraph.Range.Text = "НАКЛАДНАЯ №";
                    wordparagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    String[] Months = new string[] {"января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"};
                    oMissing = System.Reflection.Missing.Value;
                    worddocument.Paragraphs.Add(ref oMissing);
                    wordparagraph = (Word.Paragraph)wordparagraphs[3];
                    wordparagraph.Range.Font.Size = 12;
                    wordparagraph.Range.Font.Bold = 1;
                    wordparagraph.Range.Text = DateTime.Now.Day.ToString() + ' ' + Months[Int32.Parse(DateTime.Now.Month.ToString()) - 1] + ' ' + DateTime.Now.Year.ToString();
                    wordparagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

                    worddocument.Paragraphs.Add(ref oMissing);
                    wordparagraph = (Word.Paragraph)wordparagraphs[4];
                    wordparagraph.Range.Font.Size = 12;
                    wordparagraph.Range.Font.Bold = 0;
                    wordrange = wordparagraph.Range;
                    Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                    Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;
                    Word.Table wordtable = worddocument.Tables.Add(wordrange, 8, 2, ref defaultTableBehavior, ref autoFitBehavior);
                    wordtable.Columns[1].Width = 40;
                    wordtable.Columns[2].Width = 350;

                    object begCell = wordtable.Cell(1, 1).Range.Start;
                    object endCell = wordtable.Cell(2, 1).Range.End;
                    Word.Range wordcellrange = worddocument.Range(ref begCell, ref endCell);
                    wordcellrange.Select();
                    wordapp.Selection.Cells.Merge();

                    begCell = wordtable.Cell(3, 1).Range.Start;
                    endCell = wordtable.Cell(4, 1).Range.End;
                    wordcellrange = worddocument.Range(ref begCell, ref endCell);
                    wordcellrange.Select();
                    wordapp.Selection.Cells.Merge();

                    begCell = wordtable.Cell(5, 1).Range.Start;
                    endCell = wordtable.Cell(8, 1).Range.End;
                    wordcellrange = worddocument.Range(ref begCell, ref endCell);
                    wordcellrange.Select();
                    wordapp.Selection.Cells.Merge();

                    begCell = wordtable.Cell(5, 2).Range.Start;
                    endCell = wordtable.Cell(7, 2).Range.End;
                    wordcellrange = worddocument.Range(ref begCell, ref endCell);
                    wordcellrange.Select();
                    wordapp.Selection.Cells.Merge();

                    wordcellrange = worddocument.Tables[1].Cell(1, 1).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Text = "Выдать (принять)";

                    wordcellrange = worddocument.Tables[1].Cell(3, 1).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Text = "Основание";

                    wordcellrange = worddocument.Tables[1].Cell(5, 1).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Text = "Получатель (сдатчик)";

                    wordcellrange = worddocument.Tables[1].Cell(1, 2).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Text = "из полиграфического центра";

                    wordcellrange = worddocument.Tables[1].Cell(2, 2).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Text = "(кому, от кого)";

                    wordcellrange = worddocument.Tables[1].Cell(3, 2).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Text = "заявка № " + numericUpDown4.Value.ToString();

                    wordcellrange = worddocument.Tables[1].Cell(4, 2).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Text = "( наименование и № документа) ";

                    wordcellrange = worddocument.Tables[1].Cell(5, 2).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Text = "хранилище вещевого имущества группы материального обеспечения внутренних войск батальона обеспечения войсковой части 5448";

                    wordcellrange = worddocument.Tables[1].Cell(8, 2).Range;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    wordcellrange.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;

                    wordparagraph = (Word.Paragraph)wordparagraphs[22];
                    wordparagraph.Range.Font.Size = 1;
                    wordrange.Paragraphs.SpaceAfter = 0;

                    worddocument.Paragraphs.Add(ref oMissing);
                    wordparagraph = (Word.Paragraph)wordparagraphs[23];
                    wordparagraph.Range.Font.Size = 9;
                    wordparagraph.Range.Font.Bold = 0;
                    wordrange = wordparagraph.Range;
                    Word.Table wordtable1 = worddocument.Tables.Add(wordrange, application_table.Rows.Count + 3, 7, ref defaultTableBehavior, ref autoFitBehavior);
                    wordtable1.Columns[1].Width = 30;
                    wordtable1.Columns[2].Width = 150;
                    wordtable1.Columns[3].Width = 20;
                    wordtable1.Columns[4].Width = 50;
                    wordtable1.Columns[5].Width = 40;
                    wordtable1.Columns[6].Width = 50;
                    wordtable1.Columns[7].Width = 50;

                    wordcellrange = worddocument.Tables[2].Cell(1, 1).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Font.Bold = 1;
                    wordcellrange.Text = "№ п/п";

                    wordcellrange = worddocument.Tables[2].Cell(1, 2).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Font.Bold = 1;
                    wordcellrange.Text = "Наименование имущества";

                    wordcellrange = worddocument.Tables[2].Cell(1, 3).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Font.Bold = 1;
                    wordcellrange.Text = "Ед. учета";

                    wordcellrange = worddocument.Tables[2].Cell(1, 4).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Font.Bold = 1;
                    wordcellrange.Text = "Выдать";

                    wordcellrange = worddocument.Tables[2].Cell(1, 5).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Font.Bold = 1;
                    wordcellrange.Text = "Выдано";

                    wordcellrange = worddocument.Tables[2].Cell(1, 6).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Font.Bold = 1;
                    wordcellrange.Text = "Стоимость 1 экз. (руб.)";

                    wordcellrange = worddocument.Tables[2].Cell(1, 7).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Font.Bold = 1;
                    wordcellrange.Text = "Общая стоимость (руб.)";

                    wordcellrange = worddocument.Tables[2].Cell(2, 1).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Text = "1";

                    wordcellrange = worddocument.Tables[2].Cell(2, 2).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Text = "2";

                    wordcellrange = worddocument.Tables[2].Cell(2, 3).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Text = "3";

                    wordcellrange = worddocument.Tables[2].Cell(2, 4).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Text = "4";

                    wordcellrange = worddocument.Tables[2].Cell(2, 5).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Text = "5";

                    wordcellrange = worddocument.Tables[2].Cell(2, 6).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Text = "6";

                    wordcellrange = worddocument.Tables[2].Cell(2, 7).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordcellrange.Text = "7";

                    for (int i = 0; i < application_table.Rows.Count; i++)
                    {
                        double paper_cost = 0;
                        int total_paper_count = 0;
                        int paper_count = 0;
                        int paper_type = 0;
                        double defect = 0;
                        for (int j = 0; j < paper_table.Rows.Count; j++)
                        {
                            if(application_table.Rows[i][0].ToString() == paper_table.Rows[j][0].ToString())
                            {
                                paper_count = Int32.Parse(paper_table.Rows[j][2].ToString());
                                total_paper_count = total_paper_count + paper_count;
                                paper_type = Int32.Parse(paper_table.Rows[j][1].ToString());
                                defect = Double.Parse(paper_table.Rows[j][3].ToString()) / 100;
                                paper_cost = Math.Round((((paper_count * defect + paper_count) /
                                    (Double.Parse(dataGridView4.Rows[paper_type].Cells[4].Value.ToString()))) *
                                    Double.Parse(dataGridView4.Rows[paper_type].Cells[3].Value.ToString())), 2) + paper_cost;
                            }
                        }
                        int product_count = Int32.Parse(application_table.Rows[i][2].ToString());
                        int master_count = Int32.Parse(application_table.Rows[i][4].ToString());
                        int staple_type = 99;
                        if (Int32.Parse(application_table.Rows[i][5].ToString()) != 99) 
                        {
                            staple_type = Int32.Parse(application_table.Rows[i][5].ToString());
                        }
                        int staple_count = Int32.Parse(application_table.Rows[i][6].ToString());
                        double paint_cost = 0;
                        if(application_table.Rows[i][7].ToString() == "1")
                        {
                            paint_cost = Math.Round((total_paper_count * (Double.Parse(textBox8.Text) / Double.Parse(textBox7.Text))) * Double.Parse(textBox7.Text), 4);
                        }
                        else
                        {
                            paint_cost = Math.Round((total_paper_count * (Double.Parse(textBox6.Text) / Double.Parse(textBox5.Text))) * Double.Parse(textBox5.Text), 4);
                        }

                        double hard_leaf_cost = 0;
                        if(application_table.Rows[i][8].ToString() == "1")
                        {
                            hard_leaf_cost = Double.Parse(textBox12.Text) * product_count + Double.Parse(textBox13.Text) * product_count;
                        }
                        if (application_table.Rows[i][8].ToString() == "2")
                        {
                            hard_leaf_cost = Double.Parse(textBox16.Text) * product_count + Double.Parse(textBox15.Text) * product_count;
                        }
                        if (application_table.Rows[i][8].ToString() == "3")
                        {
                            hard_leaf_cost = Double.Parse(textBox18.Text) * product_count + Double.Parse(textBox17.Text) * product_count;
                        }
                        double total_cost = 0;
                        double cost_of_1 = 0;
                        if (staple_type == 99) 
                        {
                            total_cost = Math.Round((paper_cost) + ((master_count * Double.Parse(label12.Text))) + (paint_cost) + hard_leaf_cost, 2);
                            cost_of_1 = Math.Round(((paper_cost) +
                                (paint_cost) + ((master_count * Double.Parse(label12.Text)) + hard_leaf_cost)) / product_count, 2);
                            if (cost_of_1 < double.Parse("0,01"))
                                cost_of_1 = double.Parse("0,01");
                            total_cost = product_count * cost_of_1;

                        }
                        else
                        {
                            total_cost = Math.Round((paper_cost) + ((master_count * Double.Parse(label12.Text)) + staple_count *
                                Double.Parse(dataGridView3.Rows[staple_type].Cells[3].Value.ToString())) +
                                (paint_cost) + hard_leaf_cost, 2);
                            cost_of_1 = Math.Round(((paper_cost) +
                                (paint_cost) + ((master_count * Double.Parse(label12.Text)) + hard_leaf_cost + staple_count *
                                Double.Parse(dataGridView3.Rows[staple_type].Cells[3].Value.ToString()))) / product_count, 2);
                            if (cost_of_1 < double.Parse("0,01"))
                                cost_of_1 = double.Parse("0,01");
                            total_cost = product_count * cost_of_1;
                        }

                        wordcellrange = worddocument.Tables[2].Cell(i + 3, 1).Range;
                        wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wordcellrange.Text = application_table.Rows[i][0].ToString();

                        wordcellrange = worddocument.Tables[2].Cell(i + 3, 2).Range;
                        wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wordcellrange.Text = application_table.Rows[i][1].ToString();

                        wordcellrange = worddocument.Tables[2].Cell(i + 3, 3).Range;
                        wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wordcellrange.Text = "шт.";

                        wordcellrange = worddocument.Tables[2].Cell(i + 3, 4).Range;
                        wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wordcellrange.Text = application_table.Rows[i][2].ToString();

                        wordcellrange = worddocument.Tables[2].Cell(i + 3, 7).Range;
                        wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wordcellrange.Text = total_cost.ToString("0.00");

                        wordcellrange = worddocument.Tables[2].Cell(i + 3, 6).Range;
                        wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wordcellrange.Text = cost_of_1.ToString("0.00");

                        wordtable1.Range.Rows[i + 1].Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    }

                    String[] number_as_word = new string[] {"одно", "два", "три", "четыре", "пять", "шеть", "семь", "восемь", "девять", "десять", "одинадцать"
                    , "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать", "двадцать", "двадцать одно"
                    , "двадцать два", "двадцать три", "двадцать четыре", "двадцать пять", "двадцать шесть", "двадцать семь", "двадцать восемь", "двадцать девять"
                    , "тридцать", "тридцать одно", "тридцать два", "тридцать три", "тридцать четыре", "тридцать пять", "тридцать шесть", "тридцать семи", "тридцать восемь"
                    , "тридцать девять", "сорок", "сорок одно", "сорок два", "сорок три", "сорок четыре", "сорок пять", "сорок шесть", "сорок семь", "сорок восемь"
                    , "сорок девять", "пятьдесят", "пятьдесят одно", "пятьдесят два", "пятьдесят три" , "пятьдесят четыре" , "пятьдесят пять", "пятьдесят шесть", "пятьдесят семь"
                    , "пятьдесят восемь" , "пятьдесят девять" , "шестьдесят", "шестьдесят одно", "шестьдесят два", "шестьдесят три", "шестьдесят четыре", "шестьдесят пять"
                    , "шестьдесят шесть", "шестьдесят семь" , "шестьдесят восемь", "шестьдесят девять", "семьдесят", "семьдесят ондо", "семьдесят два" , "семьдесят три"
                    , "семьдесят четыре", "семьдесят пять", "семьдесят шесть" , "семьдесят семь" , "семьдесят восемь", "семьдесят девять", "восемьдесят", "восемьдесят одно"
                    , "восемьдесят два" , "восемьдесят три" , "восемьдесят четыре", "восемьдесят пять" , "восемьдесят шесть" , "восемьдесят семь" , "восемьдесят восемь"
                    , "восемьдесят девять" , "девяносто", "девяносто ондо", "девяносто два" , "девяносто три", "девяносто четыре", "девяносто пять" , "девяносто шесть"
                    , "девяносто семь" , "девяносто восемь", "девяносто девять", "сто"};
                    int number = application_table.Rows.Count - 1;
                    string name = " наименований";
                    if (number == 0 || number == 20 || number == 30 || number == 40 || number == 50 || number == 60 || number == 70 || number == 80 || number == 90)
                    {
                        name = " наименование";
                    }
                    if (number == 1 || number == 2 || number == 3 || number == 21 || number == 22 || number == 23 || number == 31 || number == 32 || number == 33
                         || number == 41 || number == 42 || number == 43 || number == 51 || number == 52 || number == 53 || number == 61 || number == 62
                          || number == 63 || number == 71 || number == 72 || number == 73 || number == 81 || number == 82 || number == 83 || number == 91
                           || number == 92 || number == 93)
                    {
                        name = " наименования";
                    }

                    worddocument.Tables[2].Rows[application_table.Rows.Count + 3].Cells[1].Merge(worddocument.Tables[2].Rows[application_table.Rows.Count + 3].Cells[7]);
                    wordcellrange = worddocument.Tables[2].Cell(application_table.Rows.Count + 3, 1).Range;
                    wordcellrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    wordcellrange.Font.Bold = 1;
                    wordcellrange.Font.Size = 12;
                    if(application_table.Rows.Count >= 1 && application_table.Rows.Count <= 100)
                        wordcellrange.Text = "Итого: " + application_table.Rows.Count + " (" + number_as_word[application_table.Rows.Count - 1] + ") " + name;
                    else
                        wordcellrange.Text = "Итого: " + application_table.Rows.Count + name;

                    wordparagraph = (Word.Paragraph)wordparagraphs[worddocument.Paragraphs.Count];
                    wordrange = wordparagraph.Range;
                    wordrange.Font.Size = 1;

                    string temp_name = "";
                    if (comboBox1.SelectedItem.ToString() == "прапорщик")
                    {
                        temp_name = "пр-к";
                    }
                    if (comboBox1.SelectedItem.ToString() == "старший прапорщик")
                    {
                        temp_name = "ст. пр-к";
                    }
                    if (comboBox1.SelectedItem.ToString() == "младший лейтенант")
                    {
                        temp_name = "мл. л-т.";
                    }
                    if (comboBox1.SelectedItem.ToString() == "лейтенант")
                    {
                        temp_name = "л-т.";
                    }
                    if (comboBox1.SelectedItem.ToString() == "старший лейтенант")
                    {
                        temp_name = "ст. л-т.";
                    }
                    if (comboBox1.SelectedItem.ToString() == "капитан")
                    {
                        temp_name = "к-н";
                    }
                    if (comboBox1.SelectedItem.ToString() == "майор")
                    {
                        temp_name = "м-р";
                    }
                    if (comboBox1.SelectedItem.ToString() == "подполковник")
                    {
                        temp_name = "п/п-к";
                    }
                    if (comboBox1.SelectedItem.ToString() == "полковник")
                    {
                        temp_name = "п-к";
                    }

                    worddocument.Paragraphs.Add(ref oMissing);
                    wordparagraph = (Word.Paragraph)wordparagraphs[worddocument.Paragraphs.Count];
                    wordparagraph.Range.Font.Size = 9;
                    wordparagraph.Range.Font.Bold = 0;

                    String[] rank_name = new string[] { "мл. с-нт", "с-нт", "ст. с-нт", "ст-на", "пр-к", "ст. пр-к", "мл. лт.", "лт.", "ст. лт.", "к-н", "м-р", "п/п-к", "п-к"};
                    worddocument.Paragraphs.Add(ref oMissing);
                    wordparagraph = (Word.Paragraph)wordparagraphs[worddocument.Paragraphs.Count];
                    wordrange = wordparagraph.Range;
                    wordrange.Font.Size = 10;
                    wordrange.Paragraphs.SpaceAfter = 0;
                    wordrange.Paragraphs.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                    wordrange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    wordrange.Text = "Выдал: " + temp_name + "________" + textBox3.Text + "       Получил: " + rank_name[comboBox4.SelectedIndex] + "________" + textBox14.Text;

                    worddocument.Paragraphs.Add(ref oMissing);
                    wordparagraph = (Word.Paragraph)wordparagraphs[worddocument.Paragraphs.Count];
                    wordrange = wordparagraph.Range;
                    wordrange.Font.Size = 10;
                    wordrange.Text = "                       (подпись)                                                            (подпись)";

                    
                    worddocument.Save();
                    wordapp.Documents.Close();
                    wordapp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worddocument);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordapp);
                    System.GC.Collect();

                }
            }
            catch (System.Runtime.InteropServices.COMException exception)
            {
                wordapp.Documents.Close(false);
                wordapp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worddocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordapp);
                System.GC.Collect();
            }
            catch (Exception ex)
            {
                wordapp.Documents.Close(false);
                wordapp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worddocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordapp);
                System.GC.Collect();
                MessageBox.Show(ex.ToString());
                
            }

            try
            {
                if (flag == 0)
                {
                    List<int> temp_mass = new List<int>();
                    for (int i = 0; i < paper_table.Rows.Count; i++)
                    {
                        temp_mass.Add(Int32.Parse(paper_table.Rows[i][1].ToString()));
                    }
                    var result = temp_mass.Distinct().ToArray();

                    List<int> mass_for_staples = new List<int>();
                    for(int i = 0; i < application_table.Rows.Count; i++)
                    {
                        if (Int32.Parse(application_table.Rows[i][5].ToString()) != 99)
                        {
                            mass_for_staples.Add(Int32.Parse(application_table.Rows[i][5].ToString()));
                        }
                    }
                    var staples_mass = mass_for_staples.Distinct().ToArray();
                    Array.Sort(staples_mass);
                    Array.Reverse(staples_mass);

                    int width = 0;
                    int color_paint_used = 0;
                    int hard_leaf_used = 0;
                    for (int i = 0; i < application_table.Rows.Count; i++)
                    {
                        if (application_table.Rows[i][7].ToString() == "1")
                            color_paint_used = 1;
                        if (application_table.Rows[i][8].ToString() == "1")
                            hard_leaf_used = 2;
                    }

                    width = 10 + color_paint_used + hard_leaf_used + result.Length * 3 + staples_mass.Length * 2;

                    excelApp = new Excel.Application();
                    Excel.Workbook workBook;
                    Excel.Worksheet workSheet;
                    workBook = excelApp.Workbooks.Add();
                    workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);

                    Excel.Range range = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, width]];
                    range.Merge();
                    range.WrapText = true;
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 1;
                    range.Font.Size = 13;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "Расчет затрат";

                    range = workSheet.Range[workSheet.Cells[2, 1], workSheet.Cells[2, width]];
                    range.Merge();
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 0;
                    range.Font.Size = 13;
                    range.RowHeight = 25;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "на изготовление печатной продукции к заявке № " + numericUpDown4.Value.ToString() + " (" + textBox1.Text + ")";

                    range = workSheet.Range["A3:A5"];
                    range.Merge();
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 1;
                    range.Font.Size = 7;
                    range.Columns.ColumnWidth = 4;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "№ п/п";

                    range = workSheet.Range["B3:B5"];
                    range.Merge();
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 1;
                    range.Font.Size = 7;
                    range.Columns.ColumnWidth = 20;
                    range.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "Наименование продукции";

                    range = workSheet.Range["C3:C5"];
                    range.Merge();
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 1;
                    range.Font.Size = 7;
                    range.Columns.ColumnWidth = 10;
                    range.RowHeight = 50;
                    range.Orientation = 90;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "Количество экземпляров";

                    range = workSheet.Range[workSheet.Cells[3, 4], workSheet.Cells[3, 3 + result.Length * 3]];
                    range.Merge();
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 1;
                    range.Font.Size = 7;
                    range.Columns.ColumnWidth = 10;
                    range.RowHeight = 20;
                    range.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "Используемая бумага";

                    range = workSheet.Range[workSheet.Cells[3, 4 + result.Length * 3], workSheet.Cells[3, width - 3]];
                    range.Merge();
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 1;
                    range.Font.Size = 7;
                    range.Columns.ColumnWidth = 10;
                    range.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "Расходные материалы";

                    range = workSheet.Range[workSheet.Cells[3, width - 2], workSheet.Cells[3, width - 1]];
                    range.Merge();
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 1;
                    range.Font.Size = 7;
                    range.Columns.ColumnWidth = 10;
                    range.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "Стоимость";

                    range = workSheet.Range[workSheet.Cells[3, width], workSheet.Cells[5, width]];
                    range.Merge();
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 1;
                    range.Font.Size = 7;
                    range.Columns.ColumnWidth = 5;
                    range.Orientation = 90;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "примечание";

                    for (int i = 0; i < result.Length; i++)
                    {
                        range = workSheet.Range[workSheet.Cells[4, 3 * (i + 1) + 1], workSheet.Cells[4, 3 + (i + 1) * 3]];
                        range.Merge();
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.Font.Size = 7;
                        range.RowHeight = 30;
                        range.WrapText = true;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = dataGridView4.Rows[Int32.Parse(result[i].ToString())].Cells[1].Value;

                        range = workSheet.Cells[5, 3 * (i + 1) + 1];
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.Font.Size = 7;
                        range.RowHeight = 120;
                        range.WrapText = true;
                        range.Orientation = 90;
                        range.Columns.ColumnWidth = 7;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = "количество";

                        range = workSheet.Cells[5, 3 * (i + 1) + 2];
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.Font.Size = 7;
                        range.RowHeight = 120;
                        range.Orientation = 90;
                        range.WrapText = true;
                        range.Columns.ColumnWidth = 7;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = "из них производственные отходы";

                        range = workSheet.Cells[5, 3 * (i + 1) + 3];
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.Font.Size = 7;
                        range.RowHeight = 120;
                        range.WrapText = true;
                        range.Orientation = 90;
                        range.Columns.ColumnWidth = 10;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = "стоимость";
                    }

                    for (int i = staples_mass.Length - 1; i >= 0; i--)
                    {
                        
                        range = workSheet.Range[workSheet.Cells[4, width - 2 - hard_leaf_used - (i + 1) * 2], workSheet.Cells[4, width - 1 - hard_leaf_used - (i + 1) * 2]];
                        range.Merge();
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.Font.Size = 7;
                        range.RowHeight = 30;
                        range.WrapText = true;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = dataGridView3.Rows[staples_mass[i]].Cells[0].Value.ToString();

                        range = workSheet.Cells[5, width - 2 - hard_leaf_used - (i + 1) * 2];
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.Font.Size = 7;
                        range.RowHeight = 120;
                        range.Orientation = 90;
                        range.WrapText = true;
                        range.Columns.ColumnWidth = 10;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = "количество (штук)";

                        range = workSheet.Cells[5, width - 1 - hard_leaf_used - (i + 1) * 2];
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.Font.Size = 7;
                        range.RowHeight = 120;
                        range.Orientation = 90;
                        range.Columns.ColumnWidth = 10;
                        range.WrapText = true;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = "стоимость";
                    }

                    range = workSheet.Range[workSheet.Cells[4, 4 + result.Length * 3], workSheet.Cells[4, 5 + result.Length * 3]];
                    range.Merge();
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 0;
                    range.Font.Size = 7;
                    range.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "Мастер-пленка";

                    range = workSheet.Range[workSheet.Cells[4, 6 + result.Length * 3], workSheet.Cells[4, 7 + result.Length * 3]];
                    range.Merge();
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 0;
                    range.Font.Size = 7;
                    range.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "Краска";

                    if (color_paint_used > 0)
                    {
                        range = workSheet.Range[workSheet.Cells[4, 8 + result.Length * 3], workSheet.Cells[4, 8 + result.Length * 3]];
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.Font.Size = 7;
                        range.WrapText = true;
                        range.Columns.ColumnWidth = 20;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = "Цветная печать";

                        range = workSheet.Cells[5, 8 + result.Length * 3];
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.Font.Size = 7;
                        range.Orientation = 90;
                        range.WrapText = true;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = "стоимость";
                    }

                    range = workSheet.Cells[5, 4 + result.Length * 3];
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 0;
                    range.Font.Size = 7;
                    range.Orientation = 90;
                    range.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "количество (мастеров)";

                    range = workSheet.Cells[5, 5 + result.Length * 3];
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 0;
                    range.Font.Size = 7;
                    range.Orientation = 90;
                    range.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "стоимость";

                    range = workSheet.Cells[5, 6 + result.Length * 3];
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 0;
                    range.Font.Size = 7;
                    range.Orientation = 90;
                    range.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "количество (туб)";

                    range = workSheet.Cells[5, 7 + result.Length * 3];
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 0;
                    range.Font.Size = 7;
                    range.Orientation = 90;
                    range.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "стоимость";

                    if (hard_leaf_used > 0)
                    {
                        range = workSheet.Range[workSheet.Cells[4, width - 4], workSheet.Cells[4, width - 3]];
                        range.Merge();
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.Font.Size = 7;
                        range.WrapText = true;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = "Твердый переплет";

                        range = workSheet.Cells[5, width - 4];
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.Font.Size = 7;
                        range.Orientation = 90;
                        range.WrapText = true;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = "стоимость обложки";

                        range = workSheet.Cells[5, width - 3];
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.Font.Size = 7;
                        range.Orientation = 90;
                        range.WrapText = true;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = "стоимость канала";
                    }

                    range = workSheet.Range[workSheet.Cells[4, width - 2], workSheet.Cells[5, width - 2]];
                    range.Merge();
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 0;
                    range.Font.Size = 7;
                    range.Orientation = 90;
                    range.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "1 экземпляр";

                    range = workSheet.Range[workSheet.Cells[4, width - 1], workSheet.Cells[5, width - 1]];
                    range.Merge();
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 0;
                    range.Font.Size = 7;
                    range.Orientation = 90;
                    range.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "весь тираж";

                    range = workSheet.Range[workSheet.Cells[6, 1], workSheet.Cells[application_table.Rows.Count + 6, width]];
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    range = workSheet.Range[workSheet.Cells[6, 1], workSheet.Cells[application_table.Rows.Count + 6,width - 1]];
                    range.NumberFormat = "0.000";
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 0;
                    range.Font.Size = 7;
                    range.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = 0;

                    for (int i = 0; i < application_table.Rows.Count; i++)
                    {
                        int product_count = Int32.Parse(application_table.Rows[i][2].ToString());
                        int master_count = Int32.Parse(application_table.Rows[i][4].ToString());
                        int staple_type = 0;
                        if (Int32.Parse(application_table.Rows[i][5].ToString()) == 99)
                        {
                            staple_type = 99;
                        }
                        else
                        {
                            staple_type = Int32.Parse(application_table.Rows[i][5].ToString());
                        }
                        int staple_count = Int32.Parse(application_table.Rows[i][6].ToString());
                        int total_paper_count = 0;
                        double paper_cost = 0;

                        range = workSheet.Cells[i + 6, 1];
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.NumberFormat = "0";
                        range.Font.Size = 7;
                        range.WrapText = true;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = application_table.Rows[i][0].ToString();

                        range = workSheet.Cells[i + 6, 2];
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.Font.Size = 7;
                        range.WrapText = true;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = application_table.Rows[i][1].ToString();

                        range = workSheet.Cells[i + 6, 3];
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.Font.Size = 7;
                        range.WrapText = true;
                        range.NumberFormat = "0.000";
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = application_table.Rows[i][2].ToString();

                        for (int j = 0; j < paper_table.Rows.Count; j++)
                        {
                            int index = 0;
                            if (application_table.Rows[i][0].ToString() == paper_table.Rows[j][0].ToString())
                            {
                                for(int z = 0; z < result.Length; z++)
                                {
                                    if (result[z].ToString() == paper_table.Rows[j][1].ToString())
                                        index = z;
                                }
                                int paper_type = Int32.Parse(paper_table.Rows[j][1].ToString());
                                int paper_count = Int32.Parse(paper_table.Rows[j][2].ToString());
                                double defect = Double.Parse(paper_table.Rows[j][3].ToString()) / 100;
                                total_paper_count = total_paper_count + paper_count;

                                range = workSheet.Cells[i + 6, 4 + index * 3];
                                range.Font.Name = "Times New Roman";
                                range.Font.Bold = 0;
                                range.Font.Size = 7;
                                range.WrapText = true;
                                range.NumberFormat = "0.000";
                                range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.HorizontalAlignment = Excel.Constants.xlCenter;
                                range.VerticalAlignment = Excel.Constants.xlCenter;
                                range.Value = Math.Round((((paper_count * defect + paper_count) / (Double.Parse(dataGridView4.Rows[paper_type].Cells[4].Value.ToString())))), 3);

                                range = workSheet.Cells[i + 6, 5 + index * 3];
                                range.Font.Name = "Times New Roman";
                                range.Font.Bold = 0;
                                range.Font.Size = 7;
                                range.WrapText = true;
                                range.NumberFormat = "0.000";
                                range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.HorizontalAlignment = Excel.Constants.xlCenter;
                                range.VerticalAlignment = Excel.Constants.xlCenter;
                                range.Value = Math.Round((((paper_count * defect) / (Double.Parse(dataGridView4.Rows[paper_type].Cells[4].Value.ToString())))), 3);
                                paper_cost = Math.Round((((paper_count * defect + paper_count) /
                                    (Double.Parse(dataGridView4.Rows[paper_type].Cells[4].Value.ToString())) *
                                    Double.Parse(dataGridView4.Rows[paper_type].Cells[3].Value.ToString()))), 3) + paper_cost;

                                range = workSheet.Cells[i + 6, 6 + index * 3];
                                range.Font.Name = "Times New Roman";
                                range.Font.Bold = 0;
                                range.Font.Size = 7;
                                range.WrapText = true;
                                range.NumberFormat = "0.000";
                                range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.HorizontalAlignment = Excel.Constants.xlCenter;
                                range.VerticalAlignment = Excel.Constants.xlCenter;
                                range.Value = Math.Round((((paper_count * defect + paper_count) /
                                    (Double.Parse(dataGridView4.Rows[paper_type].Cells[4].Value.ToString())) *
                                    Double.Parse(dataGridView4.Rows[paper_type].Cells[3].Value.ToString()))), 3);
                            }
                        }

                        for (int j = staples_mass.Length - 1; j >= 0; j--)
                        {
                            if (application_table.Rows[i][5].ToString() == staples_mass[j].ToString())
                            {
                                range = workSheet.Cells[i + 6, width - 2 - hard_leaf_used - (j + 1) * 2];
                                range.Font.Name = "Times New Roman";
                                range.Font.Bold = 0;
                                range.Font.Size = 7;
                                range.WrapText = true;
                                range.NumberFormat = "0.000";
                                range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.HorizontalAlignment = Excel.Constants.xlCenter;
                                range.VerticalAlignment = Excel.Constants.xlCenter;
                                range.Value = staple_count;

                                range = workSheet.Cells[i + 6, width - 1 - hard_leaf_used - (j + 1) * 2];
                                range.Font.Name = "Times New Roman";
                                range.Font.Bold = 0;
                                range.Font.Size = 7;
                                range.WrapText = true;
                                range.NumberFormat = "0.000";
                                range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                range.HorizontalAlignment = Excel.Constants.xlCenter;
                                range.VerticalAlignment = Excel.Constants.xlCenter;
                                range.Value = Math.Round(staple_count * Double.Parse(dataGridView3.Rows[staples_mass[j]].Cells[3].Value.ToString()),3);
                            }
                        }

                        range = workSheet.Cells[i + 6, 4 + result.Length * 3];
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.Font.Size = 7;
                        range.NumberFormat = "0.000";
                        range.WrapText = true;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = master_count;

                        range = workSheet.Cells[i + 6, 5 + result.Length * 3];
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 0;
                        range.Font.Size = 7;
                        range.WrapText = true;
                        range.NumberFormat = "0.000";
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = Math.Round((master_count * Double.Parse(label12.Text)), 3);

                        double paint_cost = 0;
                        if (application_table.Rows[i][7].ToString() == "1")
                        {
                            paint_cost = Math.Round((total_paper_count * (Double.Parse(textBox8.Text) / Double.Parse(textBox7.Text))) * Double.Parse(textBox7.Text), 4);
                            
                            range = workSheet.Cells[i + 6, 8 + result.Length * 3];
                            range.Font.Name = "Times New Roman";
                            range.Font.Bold = 0;
                            range.Font.Size = 7;
                            range.WrapText = true;
                            range.NumberFormat = "0.000";
                            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.HorizontalAlignment = Excel.Constants.xlCenter;
                            range.VerticalAlignment = Excel.Constants.xlCenter;
                            range.Value = Math.Round(paint_cost, 3);
                        }
                        else
                        {
                            paint_cost = Math.Round((total_paper_count * (Double.Parse(textBox6.Text) / Double.Parse(textBox5.Text))) * Double.Parse(textBox5.Text), 4);

                            double temp_price_of_tube = Math.Round((total_paper_count * (Double.Parse(textBox6.Text) / Double.Parse(textBox5.Text))), 3);
                            range = workSheet.Cells[i + 6, 6 + result.Length * 3];
                            range.Font.Name = "Times New Roman";
                            range.Font.Bold = 0;
                            range.Font.Size = 7;
                            range.WrapText = true;
                            range.NumberFormat = "0.000";
                            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.HorizontalAlignment = Excel.Constants.xlCenter;
                            range.VerticalAlignment = Excel.Constants.xlCenter;
                            range.Value = temp_price_of_tube;

                            range = workSheet.Cells[i + 6, 7 + result.Length * 3];
                            range.Font.Name = "Times New Roman";
                            range.Font.Bold = 0;
                            range.Font.Size = 7;
                            range.WrapText = true;
                            range.NumberFormat = "0.000";
                            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.HorizontalAlignment = Excel.Constants.xlCenter;
                            range.VerticalAlignment = Excel.Constants.xlCenter;
                            range.Value = Math.Round(temp_price_of_tube * Double.Parse(textBox5.Text), 3);
                        }

                        double hard_leaf_cost = 0;
                        if (application_table.Rows[i][8].ToString() == "1")
                        {
                            hard_leaf_cost = Double.Parse(textBox12.Text) * product_count + Double.Parse(textBox13.Text) * product_count;
                            //Стоимость обложки
                            range = workSheet.Cells[i + 6, width - 4];
                            range.Font.Name = "Times New Roman";
                            range.Font.Bold = 0;
                            range.Font.Size = 7;
                            range.WrapText = true;
                            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.HorizontalAlignment = Excel.Constants.xlCenter;
                            range.VerticalAlignment = Excel.Constants.xlCenter;
                            range.Value = Double.Parse(textBox13.Text) * product_count;
                            //Стоимость канала
                            range = workSheet.Cells[i + 6, width - 3];
                            range.Font.Name = "Times New Roman";
                            range.Font.Bold = 0;
                            range.Font.Size = 7;
                            range.WrapText = true;
                            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.HorizontalAlignment = Excel.Constants.xlCenter;
                            range.VerticalAlignment = Excel.Constants.xlCenter;
                            range.Value = Double.Parse(textBox12.Text) * product_count;
                        }

                        if (application_table.Rows[i][8].ToString() == "2")
                        {
                            hard_leaf_cost = Double.Parse(textBox16.Text) * product_count + Double.Parse(textBox15.Text) * product_count;
                            //Стоимость обложки
                            range = workSheet.Cells[i + 6, width - 4];
                            range.Font.Name = "Times New Roman";
                            range.Font.Bold = 0;
                            range.Font.Size = 7;
                            range.WrapText = true;
                            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.HorizontalAlignment = Excel.Constants.xlCenter;
                            range.VerticalAlignment = Excel.Constants.xlCenter;
                            range.Value = Double.Parse(textBox15.Text) * product_count;
                            //Стоимость канала
                            range = workSheet.Cells[i + 6, width - 3];
                            range.Font.Name = "Times New Roman";
                            range.Font.Bold = 0;
                            range.Font.Size = 7;
                            range.WrapText = true;
                            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.HorizontalAlignment = Excel.Constants.xlCenter;
                            range.VerticalAlignment = Excel.Constants.xlCenter;
                            range.Value = Double.Parse(textBox16.Text) * product_count;
                        }

                        if (application_table.Rows[i][8].ToString() == "3")
                        {
                            hard_leaf_cost = Double.Parse(textBox18.Text) * product_count + Double.Parse(textBox17.Text) * product_count;
                            //Стоимость обложки
                            range = workSheet.Cells[i + 6, width - 4];
                            range.Font.Name = "Times New Roman";
                            range.Font.Bold = 0;
                            range.Font.Size = 7;
                            range.WrapText = true;
                            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.HorizontalAlignment = Excel.Constants.xlCenter;
                            range.VerticalAlignment = Excel.Constants.xlCenter;
                            range.Value = Double.Parse(textBox17.Text) * product_count;
                            //Стоимость канала
                            range = workSheet.Cells[i + 6, width - 3];
                            range.Font.Name = "Times New Roman";
                            range.Font.Bold = 0;
                            range.Font.Size = 7;
                            range.WrapText = true;
                            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.HorizontalAlignment = Excel.Constants.xlCenter;
                            range.VerticalAlignment = Excel.Constants.xlCenter;
                            range.Value = Double.Parse(textBox18.Text) * product_count;
                        }

                        if (staple_type == 99)
                        {
                            range = workSheet.Cells[i + 6, width - 2];
                            range.Font.Name = "Times New Roman";
                            range.Font.Bold = 0;
                            range.Font.Size = 7;
                            range.WrapText = true;
                            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.HorizontalAlignment = Excel.Constants.xlCenter;
                            range.VerticalAlignment = Excel.Constants.xlCenter;
                            range.Value = Math.Round(((paper_cost) + hard_leaf_cost + paint_cost + ((master_count * Double.Parse(label12.Text)))) / product_count, 3);

                            range = workSheet.Cells[i + 6, width - 1];
                            range.Font.Name = "Times New Roman";
                            range.Font.Bold = 0;
                            range.Font.Size = 7;
                            range.WrapText = true;
                            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.HorizontalAlignment = Excel.Constants.xlCenter;
                            range.VerticalAlignment = Excel.Constants.xlCenter;
                            range.Value = Math.Round((paper_cost) + ((master_count * Double.Parse(label12.Text))) + paint_cost + hard_leaf_cost, 3);
                        }
                        else
                        {
                            range = workSheet.Cells[i + 6, width - 2];
                            range.Font.Name = "Times New Roman";
                            range.Font.Bold = 0;
                            range.Font.Size = 7;
                            range.WrapText = true;
                            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.HorizontalAlignment = Excel.Constants.xlCenter;
                            range.VerticalAlignment = Excel.Constants.xlCenter;
                            range.Value = Math.Round(((paper_cost) + hard_leaf_cost + paint_cost + ((master_count * Double.Parse(label12.Text)) + staple_count *
                                        Double.Parse(dataGridView3.Rows[staple_type].Cells[3].Value.ToString()))) / product_count, 3);

                            range = workSheet.Cells[i + 6, width - 1];
                            range.Font.Name = "Times New Roman";
                            range.Font.Bold = 0;
                            range.Font.Size = 7;
                            range.WrapText = true;
                            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.HorizontalAlignment = Excel.Constants.xlCenter;
                            range.VerticalAlignment = Excel.Constants.xlCenter;
                            range.Value = Math.Round((paper_cost) +
                                        ((master_count * Double.Parse(label12.Text)) + staple_count *
                                        Double.Parse(dataGridView3.Rows[staple_type].Cells[3].Value.ToString())) + paint_cost + hard_leaf_cost, 3);
                        }
                    }

                    range = workSheet.Cells[application_table.Rows.Count + 6, 1];
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 1;
                    range.Font.Size = 7;
                    range.WrapText = true;
                    range.HorizontalAlignment = Excel.Constants.xlCenter;
                    range.VerticalAlignment = Excel.Constants.xlCenter;
                    range.Value = "Итого";

                    for (int i = 2; i < width; i++)
                    {
                        Excel.Range range1 = workSheet.Range[workSheet.Cells[6, i], workSheet.Cells[application_table.Rows.Count + 6, i]];
                        double sum = excelApp.WorksheetFunction.Sum(range1);

                        range = workSheet.Cells[application_table.Rows.Count + 6, i];
                        range.Font.Name = "Times New Roman";
                        range.Font.Bold = 1;
                        range.Font.Size = 7;
                        range.WrapText = true;
                        range.NumberFormat = "0.000";
                        range.HorizontalAlignment = Excel.Constants.xlCenter;
                        range.VerticalAlignment = Excel.Constants.xlCenter;
                        range.Value = sum;
                    }

                    range = workSheet.Cells[application_table.Rows.Count + 6, 2];
                    range.Value = "";

                    range = workSheet.Cells[application_table.Rows.Count + 8, 1];
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 0;
                    range.Font.Size = 11;
                    range.Value = textBox2.Text;

                    range = workSheet.Cells[application_table.Rows.Count + 9, 1];
                    range.Font.Name = "Times New Roman";
                    range.Font.Bold = 0;
                    range.Font.Size = 11;
                    range.Value = comboBox1.SelectedItem.ToString() + "                    " + textBox3.Text;

                    excelApp.Visible = false;
                    excelApp.UserControl = true;
                    workBook.Saved = false;
                    workBook.Close(SaveChanges: true);
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    System.GC.Collect();
                }
            }
            catch(Exception ex)
            {
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                System.GC.Collect();
            }

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            string text = textBox5.Text;
            Errors_checker cheker = new Errors_checker();
            textBox5.Text = cheker.textBox_checker(text);

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            string text = textBox6.Text;
            Errors_checker cheker = new Errors_checker();
            textBox6.Text = cheker.textBox_checker(text);
        }

        private void dataGridView4_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView4.Rows.Count > 1)
            {
                string text = dataGridView4.CurrentRow.Cells[3].Value.ToString();
                Errors_checker cheker = new Errors_checker();
                dataGridView4.CurrentRow.Cells[3].Value = cheker.textBox_checker(text);
            }
        }

        private void dataGridView3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            int current_row = dataGridView1.CurrentCell.RowIndex;
            if (checkBox1.Checked == true)
                application_table.Rows[current_row][7] = 1;
            else
                application_table.Rows[current_row][7] = 0;
        }

        private void checkBox14_CheckedChanged(object sender, EventArgs e)
        {
            int current_row = dataGridView1.CurrentCell.RowIndex;
            if (checkBox14.Checked == false)
            {
                application_table.Rows[current_row][8] = 0;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox13.Checked = false;
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            string text = textBox7.Text;
            Errors_checker cheker = new Errors_checker();
            textBox7.Text = cheker.textBox_checker(text);
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            string text = textBox8.Text;
            Errors_checker cheker = new Errors_checker();
            textBox8.Text = cheker.textBox_checker(text);
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            string text = textBox12.Text;
            Errors_checker cheker = new Errors_checker();
            textBox12.Text = cheker.textBox_checker(text);
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            string text = textBox13.Text;
            Errors_checker cheker = new Errors_checker();
            textBox13.Text = cheker.textBox_checker(text);
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox3.SelectedIndex == 0)
            {
                numericUpDown1.Enabled = false;
                numericUpDown1.Value = 0;
            }
            else
            {
                numericUpDown1.Enabled = true;
            }
        }

        private void dataGridView3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView3.Rows.Count > 1)
            {
                if (dataGridView3.CurrentRow.Cells[2].Value != null)
                {
                    string text = dataGridView3.CurrentRow.Cells[2].Value.ToString();
                    Errors_checker cheker = new Errors_checker();
                    dataGridView3.CurrentRow.Cells[2].Value = cheker.textBox_checker(text);

                    if (double.Parse(dataGridView3.CurrentRow.Cells[1].Value.ToString()) > 1 && double.Parse(dataGridView3.CurrentRow.Cells[2].Value.ToString()) > 1)
                    {
                        double temp_price_of_one = double.Parse(dataGridView3.CurrentRow.Cells[2].Value.ToString()) / double.Parse(dataGridView3.CurrentRow.Cells[1].Value.ToString());
                        dataGridView3.CurrentRow.Cells[3].Value = temp_price_of_one;
                    }

                    text = dataGridView3.CurrentRow.Cells[3].Value.ToString();
                    dataGridView3.CurrentRow.Cells[3].Value = cheker.textBox_checker(text);
                }
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            //A3 = 1
            int current_row = dataGridView1.CurrentCell.RowIndex;
            if (checkBox2.Checked == true)
            {
                checkBox14.Checked = true;
                checkBox3.Checked = false;
                checkBox13.Checked = false;
                application_table.Rows[current_row][8] = 1;
            }
            else
            {
                application_table.Rows[current_row][8] = 0;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            //A4 = 2
            int current_row = dataGridView1.CurrentCell.RowIndex;
            if (checkBox3.Checked == true)
            {
                checkBox14.Checked = true;
                checkBox2.Checked = false;
                checkBox13.Checked = false;
                application_table.Rows[current_row][8] = 2;
            }
            else
            {
                application_table.Rows[current_row][8] = 0;
            }
        }

        private void checkBox13_CheckedChanged(object sender, EventArgs e)
        {
            //A5 = 3
            int current_row = dataGridView1.CurrentCell.RowIndex;
            if (checkBox13.Checked == true)
            {
                checkBox14.Checked = true;
                checkBox3.Checked = false;
                checkBox2.Checked = false;
                application_table.Rows[current_row][8] = 3;
            }
            else
            {
                application_table.Rows[current_row][8] = 0;
            }
        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {
            string text = textBox16.Text;
            Errors_checker cheker = new Errors_checker();
            textBox16.Text = cheker.textBox_checker(text);
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            string text = textBox15.Text;
            Errors_checker cheker = new Errors_checker();
            textBox15.Text = cheker.textBox_checker(text);
        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {
            string text = textBox18.Text;
            Errors_checker cheker = new Errors_checker();
            textBox18.Text = cheker.textBox_checker(text);
        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {
            string text = textBox17.Text;
            Errors_checker cheker = new Errors_checker();
            textBox17.Text = cheker.textBox_checker(text);
        }
    }
}
