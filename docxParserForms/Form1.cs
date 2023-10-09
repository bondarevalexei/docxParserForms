using docxParserForms.DocxHandler;

namespace docxParserForms
{
    public partial class Form1 : Form
    {
        MainHandler _handlerDocx;
        private string _connection = @"Data Source=ALEXEI;Initial Catalog=DocxParserDB;Integrated Security=True";

        public Form1()
        {
            InitializeComponent();
            _handlerDocx = new MainHandler(_connection);

            button1.DoubleClick += new EventHandler(button1_Click);

            richTextBox1.AllowDrop = true;
            richTextBox1.DragDrop += new DragEventHandler(richTextBox1_DragDrop);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dr = this.openFileDialog1.ShowDialog();

            if (dr == DialogResult.OK)
            {
                richTextBox1.Clear();
                foreach(string file in openFileDialog1.FileNames)
                {
                    _handlerDocx.HandleFile(file);
                }

                CheckForEmptyTextBox();
            }
        }

        private void richTextBox1_DragDrop(object sender, DragEventArgs e)
        {
            object filename = e.Data.GetData("FileDrop");
            if(filename != null)
            {
                if (filename is string[] list)
                {
                    richTextBox1.Clear();
                    foreach (var item in list)
                    {
                        if (item.EndsWith(".docx"))
                            _handlerDocx.HandleFile(item);
                    }

                    CheckForEmptyTextBox();
                }
            }
        }

        private void CheckForEmptyTextBox()
        {
            if (richTextBox1.Text.Trim().Length == 0)
            {
                richTextBox1.Text = "Перетащите файлы или воспользуйтесь кнопкой";
            }
        }
    }
}