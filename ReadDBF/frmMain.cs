using IranSystemConvertor;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReadDBF
{
    public partial class frmMain : Form
    {
        private DataTable DT;
        public frmMain()
        {
            InitializeComponent();
        }

        public DataTable GetYourData(string fullPath)
        {
            DataTable YourResultSet = new DataTable();

            OleDbConnection yourConnectionHandler = new OleDbConnection($"Provider=VFPOLEDB.1;Data Source={new FileInfo(fullPath)};");
            yourConnectionHandler.Open();

            if (yourConnectionHandler.State == ConnectionState.Open)
            {
                string mySQL = $"select * from {new FileInfo(fullPath).Name}";

                OleDbCommand MyQuery = new OleDbCommand(mySQL, yourConnectionHandler);
                OleDbDataReader reader = MyQuery.ExecuteReader();
                DataTable schemaTable = reader.GetSchemaTable();
                foreach (DataRow row in schemaTable.Rows)
                {
                    foreach (DataColumn column in schemaTable.Columns)
                    {
                        if (column.ColumnName=="ColumnName")
                        {
                            YourResultSet.Columns.Add(new DataColumn(row[column].ToString()));
                        }
                    }
                }
                while (reader.Read())
                {
                    DataRow dataRow = YourResultSet.NewRow();
                    for (int i = 0; i < YourResultSet.Columns.Count; i++)
                    {
                        try
                        {
                            if (((DataColumn)YourResultSet.Columns[i]).DataType == typeof(System.String))
                            {
                                dataRow[((DataColumn)YourResultSet.Columns[i])] = ConvertTo.UnicodeFrom(TextEncoding.Arabic1256, Convert.ToString(reader[i]));
                            }
                        }
                        catch (Exception)
                        {
                            dataRow[((DataColumn)YourResultSet.Columns[i])] = String.Empty;
                        }
                    }
                    YourResultSet.Rows.Add(dataRow);
                    Application.DoEvents();
                }
                reader.Close();
                MyQuery.Clone();
                yourConnectionHandler.Close();

                MyQuery.Dispose();
                yourConnectionHandler.Dispose();
            }
            return YourResultSet;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "DBF|*.dbf";
            if (openFileDialog.ShowDialog()==DialogResult.OK)
            {
                DT = GetYourData(openFileDialog.FileName);
                dataGridView1.DataSource = DT;
            }
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (DT == null || DT.Rows.Count == 0) return;
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = "CSV|*.csv";
            saveFileDialog.AddExtension = true;
            saveFileDialog.Filter = "CSV|*.csv";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                ConvertTo.ToCSV(DT, saveFileDialog.FileName);
            }
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void SaveConvertedFile(string fullPathName)
        {
            String line;
            StreamReader sr = new StreamReader(fullPathName);
            line = sr.ReadLine();
            //Continue to read until you reach end of file
            while (line != null)
            {
                //write the lie to console window
                Debug.WriteLine(ConvertTo.UnicodeFrom(TextEncoding.Arabic1256, line));
                //Read the next line
                line = sr.ReadLine();
            }
            //close the file
            sr.Close();
        }
    }
}
