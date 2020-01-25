using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SqlQueryInsert.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace SqlQueryInsert
{
    public partial class SqlQueryInsertForm : Form
    {
        public SqlQueryInsertForm()
        {
            InitializeComponent();
        }

        List<SQLQuery> queries = new List<SQLQuery>();
        
        private void OpenButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Exel files (*.xlsx)|*.xlsx"
            };

            label1.Text = "Ожидайте!";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = openFileDialog.FileName;
                Excel.Application excelApp = new Excel.Application();
                excelApp.Workbooks.Open(fileName);
                Excel.Worksheet currentSheet = (Excel.Worksheet)excelApp.Workbooks[1].Worksheets[1];
                for (int i=1; currentSheet.get_Range("A" + i).Value2 != null; i++)
                {
                    Excel.Range cellA = currentSheet.get_Range("A" + i);
                    Excel.Range cellB = currentSheet.get_Range("B" + i);
                    Excel.Range cellC = currentSheet.get_Range("C" + i);
                    queries.Add(new SQLQuery() { code = cellA.Text, checkNum = cellB.Text, name = cellC.Text });
                }
                excelApp.Quit();
                label1.Text = fileName;
            }
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                DefaultExt="sql",
                Filter = "SQL files (*.sql)|*sql",
                AddExtension = true
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fileName = saveFileDialog.FileName;
                try
                {
                    using (StreamWriter stream = new StreamWriter(fileName, true, Encoding.Default))
                    {
                        foreach (SQLQuery sQL in queries)
                        {
                            if (sQL.code.Length == 4)
                            {
                                string query = "INSERT INTO [dbo].[Salary_Okz] ([Code],[CheckNum],[Name])\n\t VALUES\n\t\t   ('" + sQL.code + "', " + sQL.checkNum + ", '" + sQL.name + "');\n";
                                stream.WriteLine(query);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            label2.Text = "ВЫПОЛНЕНО";
        }
    }
}
