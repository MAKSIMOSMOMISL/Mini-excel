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

namespace Mini_excel
{
    public partial class Form1 : Form
    {
        private string currentFilePath = null;//путь к файлу null значит новый документ 
        private bool suppressSelectionUpdate = false;//предохранитель для того чтобы событие SelectionChanged не вызывало мигание и лишних вызовов

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.RowHeadersWidth = 60;//размер хедера в ширину

            for (int i = 0; i < 10; i++) AddColumn();//задал здесь 10 столбцов по умолчанию
            for (int i = 0; i < 20; i++) AddRow();//задал здесь 20 строк по умолчанию
            //циклы потому что использованы буквенные заголовки, их дает AddColumn() через GetColumnName()
            this.Text = "Mini-Excel — Новый документ";//заголовок
        }

        private string GetColumnName(int number)
        {
            string result = "";

            while (number > 0)
            {
                number--;

                result = (char)('A' + number % 26) + result;

                number /= 26;
            }

            return result;//алгоритм переводит число в bijective-26, получает младший разряд, делит на 26 и переходит к старшему.
        }

        private void UpdateRowNumber()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                dataGridView1.Rows[i].HeaderCell.Value = (i + 1).ToString();//DataGridView не нумерует строки по дефолту, можно вручную через HeaderCell.value, где номер i + 1. вызывается после каждого AddRow()
            }
        }

        private void AddColumn()
        {
            int number = dataGridView1.Columns.Count + 1;//поле number следующий номер столбца. номер столбца + 1

            dataGridView1.Columns.Add("Column" + number, GetColumnName(number));//создает столбец. первый параметр-внутреннее имя, второй-видимый заголовок
        }

        private void AddRow()
        {
            dataGridView1.Rows.Add();//пустая строка

            UpdateRowNumber();//пересчет номеров
        }

        private string GetCellAddress(int rowIndex, int colIndex)
        {
            return GetColumnName(colIndex + 1) + (rowIndex + 1).ToString();//геттер индекса column, склеивает букву столбца и номер строки col + 1 and row + 1
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)//срабатывает когда выделенная ячейка изменяется пользователем
        {
            if (suppressSelectionUpdate) return;//если флаг supressSelectionUpdate, то выходит
            if (dataGridView1.CurrentCell == null) return;//если нет ячейки, тоже выходит

            int row = dataGridView1.CurrentCell.RowIndex;
            int col = dataGridView1.CurrentCell.ColumnIndex;

            tsAddress.Text = GetCellAddress(row, col);
            //обновляем поле Адрес
            object value = dataGridView1.CurrentCell.Value;
            tsFormula.Text = value?.ToString() ?? "";//обновляет поле формула, если ячейка пустая, то выводится пустая строка
        }

        private void tsFormula_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter) return;
            if (dataGridView1.CurrentCell == null) return;
            //реагирует на enter, игнорируя другие клавиши
            e.SuppressKeyPress = true;
            
            string newValue = tsFormula.Text;
            dataGridView1.CurrentCell.Value = newValue;
            dataGridView1.EndEdit();//записывает текст из поля в текущую ячейку, и завершает редактирование ячейки, чтобы применить значение
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)//синхронизация
        {
            if (dataGridView1.CurrentCell != null &&
                dataGridView1.CurrentCell.RowIndex == e.RowIndex &&
                dataGridView1.CurrentCell.ColumnIndex == e.ColumnIndex)//когда пользователь редактирует ячейку в таблице, то вызывается cellendedit
            {//обноляется только если редактируемая ячейка совпадает с текущей
                tsFormula.Text = dataGridView1.CurrentCell.Value?.ToString() ?? "";
            }
        }

        private void добавитьСтолбецToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddColumn();
        }

        private void добавитьСтрокуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddRow();
        }

        private void создатьToolStripMenuItem_Click(object sender, EventArgs e)//меню файл
        {
            if (dataGridView1.Rows.Count > 0 && HasAnyData())//если в таблице есть данные, то спрашивает разрешение
            {
                var result = MessageBox.Show(
                    "В таблице есть данные. Создать новый документ и потерять изменения?",
                    "Подтверждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result != DialogResult.Yes) return;
            }

            suppressSelectionUpdate = true;//предохранитель события на время
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();//очищаем строки и столбцы

            for (int i = 0; i < 10; i++) AddColumn();
            for (int i = 0; i < 20; i++) AddRow();//создаем заново 10х20

            currentFilePath = null;//сбрасывает путь
            tsAddress.Text = "";
            tsFormula.Text = "";//сбрасывает поля адреса и формулы
            suppressSelectionUpdate = false;

            this.Text = "Mini-Excel — Новый документ";//заголовок
        }

        private void сохранитьКакToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("«Сохранить как...» скоро будет реализовано.");//кнопка сохранить как
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)//выход
        {
            this.Close();
        }

        private bool HasAnyData()//циклы и методы с пары
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null && !string.IsNullOrWhiteSpace(cell.Value.ToString()))
                        return true;
                }
            }
            return false;
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

                if (cells.Length > maxColumns)
                    maxColumns = cells.Length;
            }

            for (int i = 0; i < maxColumns; i++)
            {
                AddColumn();
            }

            dataGridView1.Rows.Add(data.Count);

            for (int row = 0; row < data.Count; row++)
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

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                LoadCsv(dialog.FileName);
            }
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();

            dialog.Filter = "CSV файл (*.csv|*.csv";

            dialog.DefaultExt = "csv";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                SaveCsv(dialog.FileName);
            }
        }
    }
}