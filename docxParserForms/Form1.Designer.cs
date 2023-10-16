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
            pictureBox = new PictureBox();
            descriptionBox = new TextBox();
            prevButton = new Button();
            nextButton = new Button();
            ((System.ComponentModel.ISupportInitialize)pictureBox).BeginInit();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(95, 280);
            button1.Name = "button1";
            button1.Size = new Size(137, 24);
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
            richTextBox1.Location = new Point(26, 123);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(288, 151);
            richTextBox1.TabIndex = 1;
            richTextBox1.Text = "Перетащите файлы или воспользуйтесь кнопкой.";
            richTextBox1.TextChanged += richTextBox1_TextChanged;
            // 
            // pictureBox
            // 
            pictureBox.Location = new Point(342, 12);
            pictureBox.Name = "pictureBox";
            pictureBox.Size = new Size(319, 262);
            pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox.TabIndex = 2;
            pictureBox.TabStop = false;
            // 
            // descriptionBox
            // 
            descriptionBox.Location = new Point(342, 280);
            descriptionBox.Multiline = true;
            descriptionBox.Name = "descriptionBox";
            descriptionBox.Size = new Size(319, 75);
            descriptionBox.TabIndex = 3;
            // 
            // prevButton
            // 
            prevButton.Location = new Point(342, 361);
            prevButton.Name = "prevButton";
            prevButton.Size = new Size(109, 28);
            prevButton.TabIndex = 4;
            prevButton.Text = "Предыдущый";
            prevButton.UseVisualStyleBackColor = true;
            prevButton.Click += prevButton_Click;
            // 
            // nextButton
            // 
            nextButton.Location = new Point(552, 362);
            nextButton.Name = "nextButton";
            nextButton.Size = new Size(109, 27);
            nextButton.TabIndex = 5;
            nextButton.Text = "Следующий";
            nextButton.UseVisualStyleBackColor = true;
            nextButton.Click += nextButton_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(684, 411);
            Controls.Add(nextButton);
            Controls.Add(prevButton);
            Controls.Add(descriptionBox);
            Controls.Add(pictureBox);
            Controls.Add(richTextBox1);
            Controls.Add(button1);
            Name = "Form1";
            Text = "DocxParser";
            ((System.ComponentModel.ISupportInitialize)pictureBox).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private OpenFileDialog openFileDialog1;
        private RichTextBox richTextBox1;
        private PictureBox pictureBox;
        private TextBox descriptionBox;
        private Button prevButton;
        private Button nextButton;
    }
}