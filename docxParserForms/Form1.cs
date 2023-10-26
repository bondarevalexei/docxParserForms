using docxParserForms.DocxHandler;
using System.Text;

namespace docxParserForms
{
    public partial class Form1 : Form
    {
        MainHandler _handlerDocx;
        private int _count = 0;
        private int _index = 0;
        private List<Model> _modelList = new();

        public Form1()
        {
            InitializeComponent();
            InitializingElements();
        }

        private void InitializingElements()
        {
            _handlerDocx = new MainHandler();
            button1.DoubleClick += new EventHandler(button1_Click);
            richTextBox1.AllowDrop = true;
            richTextBox1.DragDrop += new DragEventHandler(richTextBox1_DragDrop);
            CheckButtons();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dr = this.openFileDialog1.ShowDialog();

            if (dr == DialogResult.OK)
            {
                richTextBox1.Clear();
                foreach (string file in openFileDialog1.FileNames)
                    _modelList.AddRange(_handlerDocx.HandleFile(file));

                _count = _modelList.Count;
                if (_count > 0)
                {
                    label1.Text = (_index + 1).ToString() + " / " + _count.ToString();
                    ShowModel(_modelList[0]);
                }

                CheckButtons();
                CheckForEmptyTextBox();
            }
        }

        private void richTextBox1_DragDrop(object sender, DragEventArgs e)
        {
            object filename = e.Data.GetData("FileDrop");
            if (filename != null)
            {
                if (filename is string[] list)
                {
                    richTextBox1.Clear();
                    foreach (var item in list)
                        if (item.EndsWith(".docx"))
                            _handlerDocx.HandleFile(item);

                    CheckForEmptyTextBox();
                }
            }
        }

        private void CheckForEmptyTextBox()
        {
            if (richTextBox1.Text.Trim().Length == 0)
                richTextBox1.Text = "Перетащите файлы или воспользуйтесь кнопкой.";
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e) =>
            richTextBox1.Text = "Перетащите файлы или воспользуйтесь кнопкой.";

        private void prevButton_Click(object sender, EventArgs e)
        {
            if (prevButton.Enabled)
                _index--;
            label1.Text = (_index + 1).ToString() + " / " + _modelList.Count.ToString();
            CheckButtons();

            if (_count > 0) ShowModel(_modelList[_index]);
        }

        private void nextButton_Click(object sender, EventArgs e)
        {
            if (nextButton.Enabled)
                _index++;
            label1.Text = (_index + 1).ToString() + " / " + _modelList.Count.ToString();
            CheckButtons();

            if (_count > 0) ShowModel(_modelList[_index]);
        }

        private void CheckButtons()
        {
            if (_count > 0)
            {
                if (_index == 0 && _index == _count - 1)
                    (prevButton.Enabled, nextButton.Enabled) = (false, false);
                else if (_index == 0)
                    (prevButton.Enabled, nextButton.Enabled) = (false, true);
                else if (_index == _count - 1)
                    (prevButton.Enabled, nextButton.Enabled) = (true, false);
                else
                    (prevButton.Enabled, nextButton.Enabled) = (true, true);
            }
            else
                (prevButton.Enabled, nextButton.Enabled) = (false, false);
        }

        private void ShowModel(Model model)
        {
            StringBuilder sb = new();
            sb.Append(model.Description);
            sb.Append(Environment.NewLine);
            sb.Append("Filename: " + model.Filename);
            sb.Append(Environment.NewLine);
            sb.Append("Size: " + model.Width + "x" + model.Height);

            descriptionBox.Text = sb.ToString().Trim();
            pictureBox.Image = model.Image;
        }
    }
}