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

namespace Создатель_расчетов
{
    public partial class Form2 : Form
    {
        public DataGridView dtgrd4;
        public DataGridView dtgrd3;
        public DataGridView dtgrd1;
        public Label master_price;
        public TextBox price_of_tube;
        public TextBox price_of_imprint;
        public DataTable app_table;
        public DataTable ppr_table;
        public TextBox color_paint_cost;
        public TextBox cover_cost;
        public TextBox channel_cost;

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            //проверка используется ли бумага 
            int used_paper_counter = 0;
            for (int i = 0; i < dtgrd4.Rows.Count; i++)
            {
                if (Convert.ToBoolean(dtgrd4.Rows[i].Cells[5].Value) == true)
                    used_paper_counter++;
            }
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.GridColor = Color.Black;
            int columns_count = dtgrd4.Rows.Count*6 + dtgrd3.Rows.Count*2+13;
            int rows_count = 5 + dtgrd1.Rows.Count;

            DataGridViewTextBoxColumn[] column = new DataGridViewTextBoxColumn[columns_count];
            for(int i = 0; i < columns_count; i++)
            {
                column[i] = new DataGridViewTextBoxColumn();
                if(i != 0)
                    column[i].Width = 90;
                else 
                    column[i].Width = 30;
            }
            dataGridView1.Columns.AddRange(column);
            for(int i = 0; i < rows_count; i++)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[i].Height = 70;
            }
            dataGridView1.Rows[0].Height = 40;
            dataGridView1.AutoGenerateColumns = false;

            //заполнение таблицы данными
            int temp_index = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 3; j < (dtgrd4.RowCount * 6) + 2; j = j + 6)
                {
                    dataGridView1.Rows[2].Cells[0].Value = "№ п/п";
                    dataGridView1.Rows[2].Cells[1].Value = "Наименование продукции";
                    dataGridView1.Rows[2].Cells[2].Value = "Количество экземпляров";
                    dataGridView1.Rows[0].Cells[3].Value = "Используемая бумага";
                    dataGridView1.Rows[3].Cells[j].Value = "Количество (пачка) с браком";
                    dataGridView1.Rows[3].Cells[j + 1].Value = "Количество (пачка) без брака";
                    if(temp_index < dtgrd4.Rows.Count)
                        dataGridView1.Rows[3].Cells[j + 2].Value = "Листов в " + dtgrd4.Rows[temp_index].Cells[2].Value.ToString();
                    dataGridView1.Rows[4].Cells[j + 2].Value = "Количество листов";
                    dataGridView1.Rows[4].Cells[j + 3].Value = "Процент брака";
                    dataGridView1.Rows[4].Cells[j + 4].Value = "Из низ производ. отходы";
                    dataGridView1.Rows[4].Cells[j + 5].Value = "Стоимость";
                    temp_index++;
                }
                for (int j = 3 + dtgrd4.RowCount * 6; j < ((3 + dtgrd4.RowCount * 6) + (4 + dtgrd3.RowCount * 2) - 1); j = j + 2)
                {
                    dataGridView1.Rows[0].Cells[(dtgrd4.RowCount * 6) + 3].Value = "Расходные материалы";
                    dataGridView1.Rows[1].Cells[(dtgrd4.RowCount * 6) + 3].Value = "Мастер-пленка";
                    dataGridView1.Rows[1].Cells[(dtgrd4.RowCount * 6) + 5].Value = "Краска";
                    dataGridView1.Rows[3].Cells[(dtgrd4.RowCount * 6) + 3].Value = "Количество (мастеров)";
                    dataGridView1.Rows[3].Cells[(dtgrd4.RowCount * 6) + 5].Value = "Количество (туб)";
                    dataGridView1.Rows[4].Cells[j + 1].Value = "Стоимость";
                    if (j > (dtgrd4.RowCount * 6) + 5)
                        dataGridView1.Rows[3].Cells[j].Value = "Количество (штук)";
                }
                for (int j = (3 + dtgrd4.RowCount * 6) + (4 + dtgrd3.RowCount * 2); j <= (3 + dtgrd4.RowCount * 6) + (4 + dtgrd3.RowCount * 2) + 2; j++)
                {
                    dataGridView1.Rows[0].Cells[(3 + dtgrd4.RowCount * 6) + (7 + dtgrd3.RowCount * 2)].Value = "Стоимость";
                    dataGridView1.Rows[2].Cells[(3 + dtgrd4.RowCount * 6) + (7 + dtgrd3.RowCount * 2)].Value = "1 экземпляр";
                    dataGridView1.Rows[2].Cells[(3 + dtgrd4.RowCount * 6) + (7 + dtgrd3.RowCount * 2) + 1].Value = "весь тираж";
                    dataGridView1.Rows[2].Cells[(3 + dtgrd4.RowCount * 6) + (7 + dtgrd3.RowCount * 2) + 2].Value = "примечание";
                }
            }
            //заполнение таблицы данными из предыдущей формы
            string master = master_price.Text;
            string paint = price_of_tube.Text;
            dataGridView1.Rows[3].Cells[4 + dtgrd4.RowCount * 6].Value = master;
            dataGridView1.Rows[3].Cells[6 + dtgrd4.RowCount * 6].Value = paint;
            int column_count_for_paper = 3;
            for (int i = 0; i < dtgrd4.RowCount; i++)
            {
                dataGridView1.Rows[1].Cells[column_count_for_paper].Value = dtgrd4.Rows[i].Cells[1].Value;
                column_count_for_paper = column_count_for_paper + 6;
            }

            column_count_for_paper = 7 + dtgrd4.RowCount * 6;
            for (int i = 0; i < dtgrd3.RowCount; i++)
            {
                dataGridView1.Rows[1].Cells[column_count_for_paper].Value = dtgrd3.Rows[i].Cells[0].Value;
                dataGridView1.Rows[3].Cells[column_count_for_paper + 1].Value = dtgrd3.Rows[i].Cells[3].Value;
                column_count_for_paper = column_count_for_paper + 2;
            }

            //расчеты в таблицу 

            int index = 5;
            int product_count = 0;
            int paper_type = 0;
            int paper_count = 0;
            double defect = 0;
            int master_count = 0;
            int staple_type = 99;
            int staple_count = 0;
            try
            {
                for (int i = 0; i < app_table.Rows.Count; i++)
                {
                    dataGridView1.Rows[index].Cells[0].Value = index - 4;
                    dataGridView1.Rows[index].Cells[1].Value = app_table.Rows[i][1];
                    product_count = Int32.Parse(app_table.Rows[i][2].ToString());
                    dataGridView1.Rows[index].Cells[2].Value = app_table.Rows[i][2];
                    master_count = Int32.Parse(app_table.Rows[i][4].ToString());
                    if (Int32.Parse(app_table.Rows[i][5].ToString()) != 99)
                    {
                        staple_type = Int32.Parse(app_table.Rows[i][5].ToString());
                        staple_count = Int32.Parse(app_table.Rows[i][6].ToString());
                    }
                    else
                    {
                        staple_type = 0;
                        staple_count = 0;
                    }
                    int total_paper_count = 0;
                    double paper_cost = 0;
                    for (int j = 0; j < ppr_table.Rows.Count; j++)
                    {
                        if (app_table.Rows[i][0].ToString() == ppr_table.Rows[j][0].ToString())
                        {
                            paper_type = Int32.Parse(ppr_table.Rows[j][1].ToString());
                            paper_count = Int32.Parse(ppr_table.Rows[j][2].ToString());
                            defect = Double.Parse(ppr_table.Rows[j][3].ToString()) / 100;

                            //Пачка с браком
                            dataGridView1.Rows[index].Cells[3 + (paper_type) * 6].Value = Math.Round(((paper_count * defect + paper_count) / Double.Parse(dtgrd4.Rows[paper_type].Cells[4].Value.ToString())), 4);
                            //Пачка без брака
                            dataGridView1.Rows[index].Cells[4 + (paper_type) * 6].Value = Math.Round((paper_count / Double.Parse(dtgrd4.Rows[paper_type].Cells[4].Value.ToString())), 4);
                            //Количество листов
                            dataGridView1.Rows[index].Cells[5 + (paper_type) * 6].Value = paper_count.ToString();
                            //Процент брака
                            dataGridView1.Rows[index].Cells[6 + (paper_type) * 6].Value = (defect * 100).ToString();
                            //Из них производственные отходы
                            dataGridView1.Rows[index].Cells[7 + (paper_type) * 6].Value = Math.Round(((paper_count * defect) / Double.Parse(dtgrd4.Rows[paper_type].Cells[4].Value.ToString())), 4);
                            //Стоимость
                            dataGridView1.Rows[index].Cells[8 + (paper_type) * 6].Value = Math.Round((((paper_count * defect + paper_count) / (Double.Parse(dtgrd4.Rows[paper_type].Cells[4].Value.ToString()))) *
                                Double.Parse(dtgrd4.Rows[paper_type].Cells[3].Value.ToString())), 4);

                            total_paper_count = total_paper_count + paper_count;

                            paper_cost = Math.Round((((paper_count * defect + paper_count) /
                                    (Double.Parse(dtgrd4.Rows[paper_type].Cells[4].Value.ToString())) *
                                    Double.Parse(dtgrd4.Rows[paper_type].Cells[3].Value.ToString()))), 3) + paper_cost;
                        }
                    }

                    //Количество мастеров
                    dataGridView1.Rows[index].Cells[3 + (dtgrd4.RowCount * 6)].Value = master_count;
                    //Стоимость мастера
                    dataGridView1.Rows[index].Cells[4 + (dtgrd4.RowCount * 6)].Value = Math.Round((master_count * Double.Parse(master_price.Text)), 4);
                    //Количество туб
                    dataGridView1.Rows[index].Cells[5 + (dtgrd4.RowCount * 6)].Value = Math.Round((total_paper_count * (Double.Parse(price_of_imprint.Text) / Double.Parse(price_of_tube.Text))), 4);
                    //Стоимость краски
                    dataGridView1.Rows[index].Cells[6 + (dtgrd4.RowCount * 6)].Value = Math.Round((total_paper_count * (Double.Parse(price_of_imprint.Text) / Double.Parse(price_of_tube.Text))) * Double.Parse(price_of_tube.Text), 4);
                    //Количество скоб
                    dataGridView1.Rows[index].Cells[7 + (dtgrd4.RowCount * 6) + staple_type * 2].Value = staple_count;
                    //Стоимость скоб
                    dataGridView1.Rows[index].Cells[8 + (dtgrd4.RowCount * 6) + staple_type * 2].Value = Math.Round((staple_count * Double.Parse(dtgrd3.Rows[staple_type].Cells[3].Value.ToString())), 4);
                    //Стоимость 1 экземпляра
                    double temp_paint_cost = 0;
                    if (app_table.Rows[i][7].ToString() == "1")
                    {
                        temp_paint_cost = total_paper_count * Double.Parse(color_paint_cost.Text);
                    }
                    else
                        temp_paint_cost = Math.Round((total_paper_count * (Double.Parse(price_of_imprint.Text) / Double.Parse(price_of_tube.Text))) * Double.Parse(price_of_tube.Text), 4);
                    double temp_channel_cost = 0;
                    double temp_cover_cost = 0;
                    double hard_leaf = 0;
                    if (app_table.Rows[i][8].ToString() == "1")
                    {
                        temp_channel_cost = product_count * Double.Parse(channel_cost.Text);
                        temp_cover_cost = product_count * Double.Parse(cover_cost.Text);
                        hard_leaf = temp_cover_cost + temp_channel_cost;
                    }
                    dataGridView1.Rows[index].Cells[dataGridView1.ColumnCount - 2].Value = Math.Round((paper_cost) +
                        ((master_count * Double.Parse(master_price.Text)) + hard_leaf + temp_paint_cost + staple_count *
                        Double.Parse(dtgrd3.Rows[staple_type].Cells[3].Value.ToString())), 4);
                    //Стоимость всего тиража
                    dataGridView1.Rows[index].Cells[dataGridView1.ColumnCount - 3].Value = Math.Round(((paper_cost) + +hard_leaf + temp_paint_cost + 
                        ((master_count * Double.Parse(master_price.Text)) + staple_count *
                        Double.Parse(dtgrd3.Rows[staple_type].Cells[3].Value.ToString()))) / product_count, 4);

                    index = index + 1;

                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            index = 7;
            for(int i = 0; i < dtgrd4.Rows.Count; i++)
            {
                dataGridView1.Rows[3].Cells[index + 1].Value = dtgrd4.Rows[i].Cells[3].Value;
                dataGridView1.Rows[3].Cells[index].Value = dtgrd4.Rows[i].Cells[4].Value;
                index = index + 6;
            }

            for(int i = 5; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount - 1; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value == null)
                        dataGridView1.Rows[i].Cells[j].Value = 0;
                }
            }
            dataGridView1.Rows[1].Cells[columns_count - 6].Value = "Цветная печать";
            dataGridView1.Rows[4].Cells[columns_count - 6].Value = "Стоимость";
            dataGridView1.Rows[3].Cells[columns_count - 6].Value = color_paint_cost.Text;
            dataGridView1.Rows[1].Cells[columns_count - 5].Value = "Твердый переплет";
            dataGridView1.Rows[4].Cells[columns_count - 5].Value = "Стоимость обложки";
            dataGridView1.Rows[3].Cells[columns_count - 5].Value = cover_cost.Text;
            dataGridView1.Rows[4].Cells[columns_count - 4].Value = "Стоимость канала";
            dataGridView1.Rows[3].Cells[columns_count - 4].Value = channel_cost.Text;

            for(int i = 0; i < app_table.Rows.Count; i++)
            {
                int temp_product_count = Int32.Parse(app_table.Rows[i][2].ToString());
                if (app_table.Rows[i][7].ToString() == "1")
                {
                    int total_paper_count = 0;
                    for(int j = 0; j < ppr_table.Rows.Count; j++)
                    {
                        total_paper_count = Int32.Parse(ppr_table.Rows[j][2].ToString());
                    }
                    dataGridView1.Rows[i + 5].Cells[columns_count - 6].Value = total_paper_count * Double.Parse(color_paint_cost.Text);
                }
                if(app_table.Rows[i][8].ToString() == "1")
                {
                    dataGridView1.Rows[i + 5].Cells[columns_count - 4].Value = temp_product_count * Double.Parse(channel_cost.Text);
                    dataGridView1.Rows[i + 5].Cells[columns_count - 5].Value = temp_product_count * Double.Parse(cover_cost.Text);
                }
            }
        }

        //прорисовка ячеек, так как нету конструктора таблицы, рисуем вручную 
        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            int columns_count = dtgrd4.Rows.Count * 6 + dtgrd3.Rows.Count * 2 + 13;
            //Твердый переплет
            if (e.RowIndex == 1 && e.ColumnIndex == columns_count - 5)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.Pink;
            }
            if (e.RowIndex == 2 && e.ColumnIndex == columns_count - 5)
            {
                e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.Pink;
            }
            if (e.RowIndex == 1 && e.ColumnIndex == columns_count - 4)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.Pink;
            }
            if (e.RowIndex == 2 && e.ColumnIndex == columns_count - 4)
                e.CellStyle.BackColor = System.Drawing.Color.Pink;
            if (e.RowIndex == 3 && e.ColumnIndex == columns_count - 4)
                e.CellStyle.BackColor = System.Drawing.Color.Red;
            if (e.RowIndex == 3 && e.ColumnIndex == columns_count - 5)
                e.CellStyle.BackColor = System.Drawing.Color.Red;
            if (e.RowIndex == 4 && e.ColumnIndex == columns_count - 5)
                e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
            if (e.RowIndex == 4 && e.ColumnIndex == columns_count - 4)
                e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
            for(int i = 5; i < app_table.Rows.Count + 5; i++)
            {
                if(e.RowIndex == i && e.ColumnIndex == columns_count - 4)
                    e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
                if (e.RowIndex == i && e.ColumnIndex == columns_count - 5)
                    e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
                if (e.RowIndex == i && e.ColumnIndex == columns_count - 6)
                    e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
            }
            //Цветная печать
            if (e.RowIndex == 1 && e.ColumnIndex == columns_count - 6)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.Pink;
            }
            if (e.RowIndex == 2 && e.ColumnIndex == columns_count - 6)
                e.CellStyle.BackColor = System.Drawing.Color.Pink;
            if (e.RowIndex == 3 && e.ColumnIndex == columns_count - 6)
                e.CellStyle.BackColor = System.Drawing.Color.Red;
            if (e.RowIndex == 4 && e.ColumnIndex == columns_count - 6)
                e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;

            if (e.RowIndex == 0 && e.ColumnIndex == 0)
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
            if (e.RowIndex == 1 && e.ColumnIndex == 0)
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
            if (e.RowIndex == 2 && e.ColumnIndex == 0)
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
            if (e.RowIndex == 3 && e.ColumnIndex == 0)
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
            if (e.RowIndex == 0 && e.ColumnIndex == 1)
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
            if (e.RowIndex == 1 && e.ColumnIndex == 1)
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
            if (e.RowIndex == 2 && e.ColumnIndex == 1)
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
            if (e.RowIndex == 3 && e.ColumnIndex == 1)
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
            if (e.RowIndex == 0 && e.ColumnIndex == 2)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.Khaki;
            }
            if (e.RowIndex == 1 && e.ColumnIndex == 2)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.Khaki;
            }
            if (e.RowIndex == 2 && e.ColumnIndex == 2)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.Khaki;
            }
            if (e.RowIndex == 3 && e.ColumnIndex == 2)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.Khaki;
            }
            if (e.RowIndex == 4 && e.ColumnIndex == 2)
            {
                e.CellStyle.BackColor = System.Drawing.Color.Khaki;
            }

            if (e.RowIndex == 0 && e.ColumnIndex == columns_count - 1)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
            }
            if (e.RowIndex == 1 && e.ColumnIndex == columns_count - 1)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
            }
            if (e.RowIndex == 2 && e.ColumnIndex == columns_count - 1)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
            }
            if (e.RowIndex == 3 && e.ColumnIndex == columns_count - 1)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
            }
            if (e.RowIndex == 1 && e.ColumnIndex == columns_count - 2)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
            }
            if (e.RowIndex == 2 && e.ColumnIndex == columns_count - 2)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
            }
            if (e.RowIndex == 3 && e.ColumnIndex == columns_count - 2)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
            }
            if (e.RowIndex == 1 && e.ColumnIndex == columns_count - 3)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
            }
            if (e.RowIndex == 2 && e.ColumnIndex == columns_count - 3)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
            }
            if (e.RowIndex == 3 && e.ColumnIndex == columns_count - 3)
            {
                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
            }
            if (e.RowIndex == 4 && e.ColumnIndex == columns_count - 3)
            {
                e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
            }
            if (e.RowIndex == 4 && e.ColumnIndex == columns_count - 2)
            {
                e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
            }
            if (e.RowIndex == 4 && e.ColumnIndex == columns_count - 1)
            {
                e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
            }

            for (int i = 3; i < (dtgrd4.RowCount*6) + 2; i++)
            {
                if (e.RowIndex == 0 && e.ColumnIndex == i)
                {
                    e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
                    e.CellStyle.BackColor = System.Drawing.Color.AliceBlue;
                }
                if (e.RowIndex == 0 && e.ColumnIndex == i+1)
                    e.CellStyle.BackColor = System.Drawing.Color.AliceBlue;
            }
            for(int i = 3+dtgrd4.RowCount*6; i < ((3+dtgrd4.RowCount * 6)+(7+dtgrd3.RowCount*2)-1); i++)
            {
                if (e.RowIndex == 0 && e.ColumnIndex == i)
                {
                    e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
                    e.CellStyle.BackColor = System.Drawing.Color.AliceBlue;
                }
                if (e.RowIndex == 0 && e.ColumnIndex == i + 1)
                    e.CellStyle.BackColor = System.Drawing.Color.AliceBlue;
            }
            for (int i = (3 + dtgrd4.RowCount * 6) + (7 + dtgrd3.RowCount * 2); i <= (3 + dtgrd4.RowCount * 6) + (7 + dtgrd3.RowCount * 2); i++)
            {
                if (e.RowIndex == 0 && e.ColumnIndex == i)
                {
                    e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
                    e.CellStyle.BackColor = System.Drawing.Color.AliceBlue;
                }
                if (e.RowIndex == 0 && e.ColumnIndex == i+1)
                    e.CellStyle.BackColor = System.Drawing.Color.AliceBlue;
            }
            for(int i = 0; i < dtgrd4.RowCount; i++)
            {
                for(int j = 3+(i*6); j < (3 + (i * 6)) + 5; j++)
                {
                    if (e.RowIndex == 1 && e.ColumnIndex == j)
                    {
                        e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
                        e.CellStyle.BackColor = System.Drawing.Color.Pink;
                    }
                    if (e.RowIndex == 2 && e.ColumnIndex == j)
                    {
                        e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
                        e.CellStyle.BackColor = System.Drawing.Color.Pink;
                    }
                    if (e.RowIndex == 2 && e.ColumnIndex == j + 1)
                        e.CellStyle.BackColor = System.Drawing.Color.Pink;
                    if (e.RowIndex == 1 && e.ColumnIndex == j + 1)
                        e.CellStyle.BackColor = System.Drawing.Color.Pink;
                }
            }
            for (int i = 0; i < dtgrd4.RowCount; i++)
            {
                for (int j = 3 + (i * 6); j < (3 + (i * 6)) + 6; j++)
                {
                    if (e.RowIndex == 1 && e.ColumnIndex == j)
                    {
                        e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                    }
                }
            }

            for (int i = 0; i < dtgrd3.RowCount+2; i++)
            {
                for (int j = ((3 + (dtgrd4.RowCount * 6))) + (i * 2); j < (3 + (dtgrd4.RowCount * 6)) + (i * 2) + 1; j++)
                {
                    if (e.RowIndex == 1 && e.ColumnIndex == j)
                    {
                        e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
                        e.CellStyle.BackColor = System.Drawing.Color.Pink;
                    }
                    if (e.RowIndex == 2 && e.ColumnIndex == j)
                    {
                        e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
                        e.CellStyle.BackColor = System.Drawing.Color.Pink;
                    }
                    if (e.RowIndex == 1 && e.ColumnIndex == j + 1)
                        e.CellStyle.BackColor = System.Drawing.Color.Pink;
                    if (e.RowIndex == 2 && e.ColumnIndex == j + 1)
                        e.CellStyle.BackColor = System.Drawing.Color.Pink;
                }

            }

            for (int i = 0; i < dtgrd3.RowCount + 2; i++)
            {
                for (int j = ((3 + (dtgrd4.RowCount * 6))) + (i * 2); j < (3 + (dtgrd4.RowCount * 6)) + (i * 2) + 2; j++)
                {
                    if (e.RowIndex == 1 && e.ColumnIndex == j)
                        e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                }

            }

            for (int i = 0; i < dtgrd3.RowCount + 2; i++)
            {
                for (int j = ((3 + (dtgrd4.RowCount * 6))) + (i * 2); j < (3 + (dtgrd4.RowCount * 6)) + (i * 2) + 1; j = j+2)
                {
                    if (e.RowIndex == 3 && e.ColumnIndex == j)
                        e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                }

            }

            for (int i = 0; i < dtgrd4.RowCount; i++)
            {
                for (int j = 3 + (i * 6); j < (3 + (i * 6)) + 5; j=j+6)
                {
                    if (e.RowIndex == 3 && e.ColumnIndex == j)
                        e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                    if (e.RowIndex == 3 && e.ColumnIndex == j+1)
                        e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                    if (e.RowIndex == 3 && e.ColumnIndex == j + 2)
                        e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;
                }
            }
            // закрашывание таблицы начиная с 3 строки в разделе используемая бумага
            for (int i = 5; i < (dtgrd4.RowCount * 6) + 2; i = i+6)
            {
                for (int j = 4; j < dataGridView1.RowCount; j++)
                {
                    if (e.RowIndex == j && e.ColumnIndex == i)
                        e.CellStyle.BackColor = System.Drawing.Color.Khaki;
                    if (e.RowIndex == j && e.ColumnIndex == i + 1)
                        e.CellStyle.BackColor = System.Drawing.Color.Khaki;
                    if (e.RowIndex == 3 && e.ColumnIndex == i)
                        e.CellStyle.BackColor = System.Drawing.Color.GreenYellow;
                    if (e.RowIndex == 3 && e.ColumnIndex == i + 1)
                        e.CellStyle.BackColor = System.Drawing.Color.GreenYellow;
                    if (e.RowIndex == 3 && e.ColumnIndex == i + 2)
                        e.CellStyle.BackColor = System.Drawing.Color.GreenYellow;
                    if (e.RowIndex == 3 && e.ColumnIndex == i + 3)
                        e.CellStyle.BackColor = System.Drawing.Color.Red;
                    if (e.RowIndex == j - 1 && e.ColumnIndex == i - 2)
                        e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
                    if (e.RowIndex == j - 1 && e.ColumnIndex == i - 1)
                        e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
                    if (e.RowIndex == j && e.ColumnIndex == i - 2)
                        e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
                    if (e.RowIndex == j && e.ColumnIndex == i - 1)
                        e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
                    if (e.RowIndex == j && e.ColumnIndex == i + 2)
                        e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
                    if (e.RowIndex == j && e.ColumnIndex == i + 3)
                        e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
                    if (e.RowIndex == j + 1 && e.ColumnIndex == 2)
                        e.CellStyle.BackColor = System.Drawing.Color.Khaki;
                }
            }

            // закрашывание таблицы начиная с 3 строки в разделе расходные материалы
            for (int i = 3 + dtgrd4.RowCount * 6; i < ((3 + dtgrd4.RowCount * 6) + (4 + dtgrd3.RowCount * 2) - 1); i = i + 2)
            {
                for(int j = 3; j < dataGridView1.RowCount; j++)
                {
                    if (e.RowIndex == j && e.ColumnIndex == i)
                        e.CellStyle.BackColor = System.Drawing.Color.Khaki;
                    if (e.RowIndex == 3 && e.ColumnIndex == i + 1)
                        e.CellStyle.BackColor = System.Drawing.Color.Red;
                    if (e.RowIndex == j + 1 && e.ColumnIndex == i + 1)
                        e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
                }
            }

            // закрашывание таблицы начиная с 3 строки в разделе расходные стоимость и примечание
            for (int i = (3 + dtgrd4.RowCount * 6) + (7 + dtgrd3.RowCount * 2); i <= (3 + dtgrd4.RowCount * 6) + (7 + dtgrd3.RowCount * 2) + 2; i++)
            {
                for (int j = 5; j < dataGridView1.RowCount; j++)
                {
                    if (e.RowIndex == j && e.ColumnIndex == i)
                        e.CellStyle.BackColor = System.Drawing.Color.LightSteelBlue;
                }
            }

        }
    }
}
