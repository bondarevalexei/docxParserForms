using docxParserForms.Db;
using Microsoft.EntityFrameworkCore;

namespace docxParserForms
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();

            button1.DoubleClick += new EventHandler(button1_Click);

            richTextBox1.AllowDrop = true;
            richTextBox1.DragDrop += new DragEventHandler(richTextBox1_DragDrop);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dr = this.openFileDialog1.ShowDialog();
            var filenames = new List<string>();

            if (dr == DialogResult.OK)
            {
                richTextBox1.Clear();
                foreach(var file in openFileDialog1.FileNames)
                {
                    richTextBox1.Text += file;
                    richTextBox1.Text += "\n";
                }

                checkForEmptyTextBox();
            }
        }

        private void richTextBox1_DragDrop(object sender, DragEventArgs e)
        {
            var filename = e.Data.GetData("FileDrop");
            if(filename != null)
            {
                var list = filename as string[];

                if(list != null)
                {
                    richTextBox1.Clear();
                    foreach (var item in list)
                    {
                        if(item.EndsWith(".docx"))
                            richTextBox1.Text += item + "\n";
                    }

                    checkForEmptyTextBox();
                }
            }
        }

        private void checkForEmptyTextBox()
        {
            if (richTextBox1.Text.Trim().Length == 0)
            {
                richTextBox1.Text = "Перетащите файлы или воспользуйтесь кнопкой";
            }
        }
    }
}