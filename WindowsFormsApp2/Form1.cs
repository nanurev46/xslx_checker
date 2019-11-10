using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            
        }

        public DataTable openTable()
        {
            DataTable result = new DataTable();

            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFile.InitialDirectory = "C:\\POS\\EORM\\17 release";
            if (openFile.ShowDialog() == DialogResult.Cancel)
                return result;


            // получаем выбранный файл
            string filename = openFile.FileName;
            label1.Text = filename;

            String name = "Sheet1";
            String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                            openFile.FileName +
                            ";Extended Properties='Excel 12.0 XML;HDR=YES;';";
            OleDbConnection con = new OleDbConnection(constr);
            OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
            con.Open();

            OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            sda.Fill(result);

            return result;
        }

        public static void findRecWithHands(DataGridView dsSource, DataGridView dgToFind)
        {
            int colCount = dsSource.ColumnCount;
            int rowCount = dgToFind.RowCount - 1;//- строка с шапкой

            if (checkColCount(dsSource, dgToFind))
            {
                return;
            }

            bool flag = true;
            List<int> listNum = new List<int>();
            for (int i = 0; i < rowCount; i++)
            {
                flag = true;
                for (int j = 0; j < colCount; j++)
                {
                    if (!dsSource.SelectedRows[0].Cells[j].Value.ToString().Equals(dgToFind.Rows[i].Cells[j].Value.ToString()))
                    {
                        flag = false;
                        break;
                    }
                }

                if (flag)
                {
                    listNum.Add(i);
                }

            }
            if (listNum.Count == 1)
            {
                dgToFind.ClearSelection();
                dgToFind.Rows[listNum[0]].Selected = true;
                dgToFind.CurrentCell = dgToFind.SelectedRows[0].Cells[0];
            }

            if (listNum.Count > 1)
            {
                String message = "Количество совпадений: " + listNum.Count + ". /r/n;" +
                                "Номера строк: ";
                foreach (int l in listNum)
                {
                    message = String.Concat(message + l + ", ");
                }
                MessageBox.Show("Количество столбцов в гридах не совпадает");
            }


            //логирование

        }

        public void findRecsAuto (DataGridView dgSource, DataGridView dgToFind)
        {
            int colCount = dgSource.ColumnCount;
            int rowCountToFind = dgToFind.RowCount - 1;//- строка с шапкой
            int rowCountSource = dgSource.RowCount - 1;//- строка с шапкой

            if (checkColCount(dgSource, dgToFind))
            {
                return;
            }

            List<int> zeroCoincidence = new List<int>();//список ненайденных строк
            List<int> oneCoincidence = new List<int>();//список найденных строк (1 совпадение)
            List<int> manyCoincidence = new List<int>();//список найденных строк (больше 1 совпадения)

            for (int i = 0; i < rowCountSource; i++)
            {
                switch (findRec(dgSource.Rows[i], dgToFind))
                {
                    case 0:
                        zeroCoincidence.Add(i);
                        break;
                    case 1:
                        oneCoincidence.Add(i);
                        break;
                    default:
                        manyCoincidence.Add(i);
                        break;
                }
            }

            //Логи
            richTextBox1.Text = "Дата проверки: " + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + Environment.NewLine + Environment.NewLine +
                                "Ненайденные строки (" + zeroCoincidence.Count() + "): ";
            foreach (int z in zeroCoincidence)
            {
                richTextBox1.Text = String.Concat(richTextBox1.Text, z, " ");
            }

            richTextBox1.Text = String.Concat(richTextBox1.Text, Environment.NewLine, "Найденные строки (" + oneCoincidence.Count() + "): ");
            foreach (int o in oneCoincidence)
            {
                richTextBox1.Text = String.Concat(richTextBox1.Text, o, " ");
            }

            richTextBox1.Text = String.Concat(richTextBox1.Text, Environment.NewLine, "Строки найдены боль 1 раза (" + manyCoincidence.Count() + "): ");
            foreach (int m in manyCoincidence)
            {
                richTextBox1.Text = String.Concat(richTextBox1.Text, m, " ");
            }

            richTextBox1.Text = String.Concat(richTextBox1.Text, Environment.NewLine, "=========================", Environment.NewLine);


        }

        //Считаем кол-во вхождений искомой строки в НСИ
        public static int findRec(DataGridViewRow row, DataGridView dgToFind)
        {
            int colCount = dgToFind.ColumnCount;
            int rowCount = dgToFind.RowCount - 1;
            /*
            List<int> zeroCoincidence = new List<int>();//список ненайденных строк
            List<int> oneCoincidence = new List<int>();//список найденных строк (1 совпадение)
            List<int> manyCoincidence = new List<int>();//список найденных строк (больше 1 совпадения)
            List<List<int>> result = new List<List<int>>();
            */
            int result = 0;

            bool flag = false;
            int recNum = 0; //для Логов
            for (int i = 0; i < rowCount; i++)//по строкам НСИ
            {
                flag = false;
                //counter = 0;
                for (int j = 0; j < colCount; j++)//по ячейкам
                {
                    if (!row.Cells[j].Value.ToString().Equals(dgToFind.Rows[i].Cells[j].Value.ToString()))//если ячейки не совпали, переходим к следующей строке
                    {
                        break;
                    }
                    if (j == colCount - 1)//если все ячейки совпали
                    {
                        flag = true;
                    }
                }

                if (flag)
                {
                    result++;
                }

            }

            return result;
        }

        public static bool checkColCount(DataGridView dsSource, DataGridView dgToFind)
        {
            int colCount = dsSource.ColumnCount;
            int rowCount = dgToFind.RowCount - 1;//- строка с шапкой

            if (colCount != dgToFind.ColumnCount)
            {
                MessageBox.Show("Количество столбцов в гридах не совпадает");
                return true;
            }
            return false;
        }

            private void button1_Click(object sender, EventArgs e)
        {

            dataGridView1.DataSource = openTable();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = openTable();

        }

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            findRecWithHands(dataGridView1, dataGridView2);
/*
            int colCount = dataGridView1.ColumnCount;
            int rowCount = dataGridView2.RowCount-1;

            if (colCount != dataGridView2.ColumnCount){
                MessageBox.Show("Количество столбцов в гридах не совпадает");
                return;
            }

                        bool flag = true;
                        List<int> listNum = new List<int>();
                        for (int i = 0; i < rowCount; i++)
                        {
                            flag = true;
                            for (int j = 0; j < colCount; j++)
                            {
                                if (!dataGridView1.SelectedRows[0].Cells[j].Value.ToString().Equals(dataGridView2.Rows[i].Cells[j].Value.ToString())){
                                    flag = false;
                                    break;
                                }
                            }

                            if (flag)
                            {
                                listNum.Add(i);
                            }

                        }
                        if (listNum.Count == 1)
                        {
                            dataGridView2.ClearSelection();
                            dataGridView2.Rows[listNum[0]].Selected = true;
                            dataGridView2.CurrentCell = dataGridView2.SelectedRows[0].Cells[0];
                        }

                        if (listNum.Count > 1)
                        {
                            String message = "Количество совпадений: " + listNum.Count + ". /r/n;" +
                                            "Номера строк: ";
                            foreach (int l in listNum){
                                message = String.Concat(message + l + ", ");
                            }        
                            MessageBox.Show("Количество столбцов в гридах не совпадает");
                        }
                        */
        }

        private void dataGridView2_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            findRecWithHands(dataGridView2, dataGridView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                findRecsAuto(dataGridView1, dataGridView2);
            }
            else
            {
                findRecsAuto(dataGridView2, dataGridView1);
            }
        }
    }
}
