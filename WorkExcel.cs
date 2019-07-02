using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp12Design
{
    class WorkExcel
    {
        static private string[] tableHeaders;
        static private string[][] table;

        public static void ExcelRecording(ListView lv, string text, string fileName, int index)
        {
            InfoInListView(lv);

            Excel.Application exApp = new Excel.Application();
            exApp.Visible = true;
            string path = Application.StartupPath.ToString() + @"\" + fileName + ".xlsx";
            exApp.Workbooks.Open(path);

            Excel.Worksheet workSheet = (Excel.Worksheet)exApp.ActiveSheet;

            workSheet.Cells[1, index] = text;

            for (int i = 0; i < tableHeaders.Length; i++)
                workSheet.Cells[3, i + 1] = tableHeaders[i];


            for (int i = 0, row = 4; i < table.Length; i++)
            {
                for (int j = 0; j < table[0].Length; j++)
                {
                    workSheet.Cells[row, j + 1] = table[i][j];
                }
                row++;
            }

            //workSheet.SaveAs(fileName + ".xls");
            //exApp.Quit();
        }

        static public void InfoInListView(ListView lv)
        {
            tableHeaders = new string[lv.Columns.Count];
            table = new string[lv.Items.Count][];
            for (int i = 0; i < lv.Columns.Count; i++)
                tableHeaders[i] = lv.Columns[i].Text.ToString();


            for (int i = 0; i < lv.Items.Count; i++)
            {
                table[i] = new string[lv.Columns.Count];

                for (int j = 0; j < lv.Columns.Count; j++)
                {
                    table[i][j] = lv.Items[i].SubItems[j].Text;
                }
            }
        }
    }
}
