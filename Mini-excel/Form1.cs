using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Mini-excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.RowHeadersWidth = 60;
        }

        private string GetColumnName(int number)
        {
            string result = "";

            while (number > 0)
            {
                number--;

                result = (char)('A' +number %26) + result;

                number /= 26;
            }

            return result;
        }

        private void UpdateRowNumber()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1.Rows[i].HeaderCell.Value = (i+1).ToString();
            }
        }

        private void AddColumn()
        {
            int number = dataGridView1.Columns.Count + 1;

            dataGridView1.Columns.Add("Column" + number, GetColumnName(number));
        }

        private void AddRow()
        {
            dataGridView1.Rows.Add();

            UpdateRowNumber();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void добавитьСтолбецToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddColumn();
        }

        private void добавитьСтрокуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddRow();
        }

        private void LoadCsv(string path)
        {
            string[] lines = File.ReadAllLines(path);

            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            if (lines.Length == 0)
            {
                return;
            }

            List<string[]> data = new List<string[]>();

            int maxColumns = 0;

            foreach (string line in lines)
            {
                string[] cells = line.Split(';');

                data.Add(cells);

                if(cells.Length > maxColumns)
                    maxColumns = cells.Length;
            }

            for (int i = 0; i < maxColumns; i++)
            {
                AddColumn();
            }

            dataGridView1.Rows.Add(data.Count);

            for(int row = 0; row< data.Count; row++)
            {
                for (int column = 0; column < data[row].Length; column++)
                {
                    dataGridView1.Rows[row].Cells[column].Value = data[row][column];
                }
            }

            UpdateRowNumber();
        }

        private void SaveCsv(string path)
        {
            List<string> lines = new List<string>();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                List<string> cells = new List<string>();

                foreach (DataGridViewCell cell in row.Cells)
                {
                    string value = cell.Value?.ToString() ?? "";

                    cells.Add(value);
                }    

                lines.Add(string.Join(";", cells));
            }

            File.WriteAllLines(path, lines);
        }
        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();

            dialog.Filter = "CSV файл (*.csv|*.csv";

            if (dialog.ShowDialog() == DialogResult.OK )
            {
                LoadCsv(dialog.FileName);
            }
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();

            dialog.Filter = "CSV файл (*.csv|*.csv";

            dialog.DefaultExt = "csv";

            if(dialog.ShowDialog() == DialogResult.OK)
            {
                SaveCsv(dialog.FileName);
            }
        }
    }
}
