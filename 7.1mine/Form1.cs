using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace _7._1mine
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            dataGridView1.RowCount = 4;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Word.Application app = new Word.Application();
            
            Word.Document doc = app.Documents.Add();                     
            Word.Paragraph p = doc.Content.Paragraphs.Add();
            p.Range.Text = "Отчет № " + textBox1.Text + " от " + Convert.ToString(dateTimePicker1.Value);
            p.Range.Font.Bold = 1;
            p.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            
            p.Format.SpaceAfter = 20; 
            p.Range.InsertParagraphAfter(); 

            Word.Table tab = doc.Tables.Add(p.Range, dataGridView1.RowCount, 3);
            tab.Borders.Enable = 1;

            tab.Range.Bold = 1;
            tab.Cell(1, 1).Range.Text = "Район";
            tab.Cell(1, 2).Range.Text = "Адрес";
            tab.Cell(1, 3).Range.Text = "Количество жителей";
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {

                tab.Cell(i + 2, 1).Range.Text = dataGridView1.Rows[i].Cells[0].Value.ToString();
                tab.Cell(i + 2, 2).Range.Text = dataGridView1.Rows[i].Cells[1].Value.ToString();
                tab.Cell(i + 2, 3).Range.Text = dataGridView1.Rows[i].Cells[2].Value.ToString();
            }

            p = doc.Content.Paragraphs.Add();
            p.Range.Font.Bold = 0;           
            doc.Save();
            app.Visible = true;
        }
    }
}
