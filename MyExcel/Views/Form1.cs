using System;
using System.Activities.Statements;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using MyExcel.Controllers;

namespace MyExcel
{
    public partial class Form1 : Form
    {
        Controller controller = new Controller();
        //int rowD, colD;       // код для события при двойном щелчке
        //bool isDoub = false;
        bool IsEndDraw = false;
        public Form1()
        {
            InitializeComponent();
            DataGridInit();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                //if (isDoub == true)           // код для события при двойном щелчке
                //{
                //    dataGridView1.Rows[rowD].Cells[colD].Value = viewM[rowD, colD];
                //    isDoub = false;
                //    MessageBox.Show("kjgv");
                //}

                address.Text = $"{dataGridView1.Columns[e.ColumnIndex].HeaderText}{e.RowIndex + 1}";
                function.Text = controller.GetProgItem(e.RowIndex, e.ColumnIndex);

                dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
            }
        }

        private void address_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    dataGridView1.CurrentCell = 
                        dataGridView1.Rows[Convert.ToInt32(address.Text.Remove(0, 1)) - 1].Cells["C" + 
                        address.Text[0]];
                    dataGridView1.BeginEdit(true);
                }
                catch
                {
                    MessageBox.Show("Введите корректный адрес!");
                }
            }
        }

        private void DataGridInit()
        {
            func.BackColor = Color.White;
            dataGridView1.DataSource = controller.viewDatas;
            dataGridView1.RowHeadersVisible = true;
            IsEndDraw = true;
            address.Text = "A1";
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (IsEndDraw == true)
            {
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                {
                    dataGridView1.CellValueChanged -= dataGridView1_CellValueChanged;
                    controller.Evaluate(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), 
                        e.RowIndex, e.ColumnIndex, dataGridView1.Columns[e.ColumnIndex].HeaderText);
                    if (controller.IsUpdate == true)
                    {
                        dataGridView1.Refresh();
                        controller.IsUpdate = false;
                    }
                    Thread myThread = new Thread(new ThreadStart(Ref));
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            controller.Upload();
            dataGridView1.CellValueChanged -= dataGridView1_CellValueChanged;
            dataGridView1.Refresh();
            dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //isDoub = true;             // код для события при двойном щелчке
            //dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = controller.progDatas[e.RowIndex][e.ColumnIndex];
            //dataGridView1.CurrentCell = dataGridView1.Rows[e.RowIndex - 1].Cells[e.ColumnIndex];
            //dataGridView1.CurrentCell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
            //dataGridView1.BeginEdit(true);
            //dataGridView1.Refresh();
            //rowD = e.RowIndex;
            //colD = e.ColumnIndex;
        }

        private void dataGridView1_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            e.Row.HeaderCell.Value = (e.Row.Index + 1).ToString();
        }

        public void Ref(/*int row, int col*/)
        {
            dataGridView1.Refresh();
            dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
        }

        private void function_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                controller.Evaluate(func.Text, 
                    dataGridView1.CurrentCell.RowIndex, 
                    dataGridView1.CurrentCell.ColumnIndex, 
                    dataGridView1.Columns[dataGridView1.CurrentCell.ColumnIndex].HeaderText);
                dataGridView1.CellValueChanged -= dataGridView1_CellValueChanged;
                Thread myThread = new Thread(new ThreadStart(Ref));
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "progDataDataSet.View". При необходимости она может быть перемещена или удалена.

        }
    }
}
