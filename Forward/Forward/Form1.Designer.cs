namespace Forward
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.OpenExcel = new System.Windows.Forms.Button();
            this.tbExcel = new System.Windows.Forms.TextBox();
            this.result = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.tbWord = new System.Windows.Forms.TextBox();
            this.OpenWord = new System.Windows.Forms.Button();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // OpenExcel
            // 
            this.OpenExcel.Location = new System.Drawing.Point(21, 21);
            this.OpenExcel.Name = "OpenExcel";
            this.OpenExcel.Size = new System.Drawing.Size(75, 36);
            this.OpenExcel.TabIndex = 0;
            this.OpenExcel.Text = "Открыть Excel";
            this.OpenExcel.UseVisualStyleBackColor = true;
            this.OpenExcel.Click += new System.EventHandler(this.OpenExcel_Click);
            // 
            // tbExcel
            // 
            this.tbExcel.Location = new System.Drawing.Point(146, 30);
            this.tbExcel.Name = "tbExcel";
            this.tbExcel.Size = new System.Drawing.Size(305, 20);
            this.tbExcel.TabIndex = 1;
            // 
            // result
            // 
            this.result.Location = new System.Drawing.Point(21, 116);
            this.result.Name = "result";
            this.result.Size = new System.Drawing.Size(75, 23);
            this.result.TabIndex = 2;
            this.result.Text = "результат";
            this.result.UseVisualStyleBackColor = true;
            this.result.Click += new System.EventHandler(this.result_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(146, 116);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(228, 134);
            this.listBox1.TabIndex = 3;
            // 
            // tbWord
            // 
            this.tbWord.Location = new System.Drawing.Point(146, 72);
            this.tbWord.Name = "tbWord";
            this.tbWord.Size = new System.Drawing.Size(305, 20);
            this.tbWord.TabIndex = 5;
            // 
            // OpenWord
            // 
            this.OpenWord.Location = new System.Drawing.Point(21, 63);
            this.OpenWord.Name = "OpenWord";
            this.OpenWord.Size = new System.Drawing.Size(75, 36);
            this.OpenWord.TabIndex = 4;
            this.OpenWord.Text = "Открыть Word";
            this.OpenWord.UseVisualStyleBackColor = true;
            this.OpenWord.Click += new System.EventHandler(this.OpenWord_Click_1);
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog2";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(652, 427);
            this.Controls.Add(this.tbWord);
            this.Controls.Add(this.OpenWord);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.result);
            this.Controls.Add(this.tbExcel);
            this.Controls.Add(this.OpenExcel);
            this.Name = "Form1";
            this.Text = "Forward1.0";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button OpenExcel;
        private System.Windows.Forms.TextBox tbExcel;
        private System.Windows.Forms.Button result;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.TextBox tbWord;
        private System.Windows.Forms.Button OpenWord;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
    }
}

