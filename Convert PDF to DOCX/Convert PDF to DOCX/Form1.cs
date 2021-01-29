using System;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Convert_PDF_to_DOCX
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void TextBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
             
                e.Effect = DragDropEffects.None;
        }

        private void TextBox1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = e.Data.GetData(DataFormats.FileDrop) as string[]; // get all files droppeds  
            if (files != null && files.Any())
                textBox1.Text = files.First(); //select the first one  
        }

        private void TextBox1_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Link;
            else
                e.Effect = DragDropEffects.None;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            string pdfFile = textBox1.Text.ToString();
            string wordFile = pdfFile.ToString().Replace(".pdf", ".docx");

            if (pdfFile.ToString().Contains(".pdf"))
            {
                var wordApp = new Word.Application();
                wordApp.Visible = true;
                var txt = wordApp.Documents.Open(pdfFile);

                wordApp.Documents[1].SaveAs(wordFile);
                wordApp.Documents[1].Close();
            }
        }
    }
}
