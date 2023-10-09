namespace docxParserForms
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            button1 = new Button();
            openFileDialog1 = new OpenFileDialog();
            richTextBox1 = new RichTextBox();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(70, 168);
            button1.Name = "button1";
            button1.Size = new Size(186, 36);
            button1.TabIndex = 0;
            button1.Text = "Открыть файл";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            openFileDialog1.Filter = "docx files (*.docx)|*.docx";
            openFileDialog1.Multiselect = true;
            // 
            // richTextBox1
            // 
            richTextBox1.EnableAutoDragDrop = true;
            richTextBox1.Location = new Point(21, 12);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(300, 150);
            richTextBox1.TabIndex = 1;
            richTextBox1.Text = "Перетащите файлы или воспользуйтесь кнопкой.";
            richTextBox1.TextChanged += richTextBox1_TextChanged;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(344, 211);
            Controls.Add(richTextBox1);
            Controls.Add(button1);
            Name = "Form1";
            Text = "DocxParser";
            ResumeLayout(false);
        }

        #endregion

        private Button button1;
        private OpenFileDialog openFileDialog1;
        private RichTextBox richTextBox1;
    }
}