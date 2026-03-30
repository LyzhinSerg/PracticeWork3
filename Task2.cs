using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace Task2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void createButton_Click(object sender, EventArgs e)
        {
            var wordApp = new Word.Application();
            wordApp.Visible = true;
            var document = wordApp.Documents.Add();
            var range = document.Content;
            // это
            range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;

            var paragraph = document.Paragraphs.Add(range);
            range = document.Range(1, 1);
            range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            var table = document.Tables.Add(range, Convert.ToInt32(taskCountTextBox.Text) + 1, 2);
            table.Borders.Enable = 1;

            range = table.Cell(1, 1).Range;
            range.Text = "№";
            range = table.Cell(1, 2).Range;
            range.Text = "Текст";
            for (int i = 2; i < Convert.ToInt32(taskCountTextBox.Text) + 2; i++)
            {
                range = table.Cell(i, 1).Range;
                range.Text = $"{i - 1}";
            }
            // это
            paragraph = document.Paragraphs.Add();
            range = paragraph.Range;
            range.InsertDateTime();

            range = document.Range(0, 0);
            Word.InlineShape imageShape = range.InlineShapes.AddPicture(Path.Combine(Environment.CurrentDirectory,"photo2.png"));
            imageShape.Height = 420;
            imageShape.Width = 204;

        }
    }
}
