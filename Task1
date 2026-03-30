using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace Task1
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
            object template = Path.Combine(Environment.CurrentDirectory, "Шаблон.docx");
            var document = wordApp.Documents.Add(template);
            var range = document.Content;
            range.Find.Execute(FindText: "text",
                ReplaceWith: inputTextBox.Text.Replace("\n",""),
                Replace: WdReplace.wdReplaceAll
            );
          
            range.Find.Execute(FindText: "data",
                ReplaceWith: DateTime.Now,
                Replace: WdReplace.wdReplaceAll
            );

            range.Find.Execute(FindText: "table");
            if (range.Find.Found)
            {
                range.Text = "";
                var paragraph = document.Paragraphs.Add(range);
                range = paragraph.Range;
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
            }
        }
    }
}
