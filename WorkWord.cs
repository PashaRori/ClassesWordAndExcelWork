using System;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApp12Design
{
    class WorkWord
    {
        static private Word.Document worddocument;
        static private string[] tableHeaders;
        static private string[][] table;

        static public void OpenDocument(String fileName)
        {
            Word.Application wordapp = new Word.Application();
            wordapp.Visible = true;
            Object path = Application.StartupPath.ToString() + @"\" + fileName + ".docx";
            WorkWord.worddocument = wordapp.Documents.Open(ref path, true, false, true, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing);
        }

        static public void WordRecordingRange(String textWord, Object start, object end)
        {
            Word.Paragraphs wordparagraphs = worddocument.Paragraphs;
            Word.Range wordrange = worddocument.Range(ref start, ref end);
            Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
            Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;
            wordrange.Text = textWord;
        }

        static public void WordRecordingTable(ListView lv, Object start, object end)
        {
            InfoInListView(lv);

            Word.Range wordrange = worddocument.Range(ref start, ref end);
            Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
            Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;

            //Добавляем таблицу и получаем объект wordtable 
            Word.Table wordtable = worddocument.Tables.Add(wordrange, table.Length + 1, table[0].Length, ref defaultTableBehavior, ref autoFitBehavior);
            Word.Range wordcellrange;

            for (int i = 0, j = 1; i < tableHeaders.Length; i++, j++)
            {
                wordcellrange = worddocument.Tables[1].Cell(1, j).Range;
                wordcellrange.Text = tableHeaders[i];
            }

            for (int i = 0; i < wordtable.Rows.Count - 1; i++)
            {
                for (int j = 0; j < wordtable.Columns.Count; j++)
                {
                    wordcellrange = worddocument.Tables[1].Cell(i + 2, j + 1).Range;
                    wordcellrange.Text = table[i][j];
                }
            }
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
