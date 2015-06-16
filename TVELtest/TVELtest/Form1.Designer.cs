namespace TVELtest
{
    partial class Form1
    {
        /// <summary>
        /// Требуется переменная конструктора.
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
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.getOrpoButton = new System.Windows.Forms.Button();
            this.resultTextBox = new System.Windows.Forms.TextBox();
            this.resultLabel = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.getIbpoButton = new System.Windows.Forms.Button();
            this.testTextBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.testlabel = new System.Windows.Forms.Label();
            this.reslab = new System.Windows.Forms.Label();
            this.testlab = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // getOrpoButton
            // 
            this.getOrpoButton.Location = new System.Drawing.Point(103, 173);
            this.getOrpoButton.Name = "getOrpoButton";
            this.getOrpoButton.Size = new System.Drawing.Size(75, 23);
            this.getOrpoButton.TabIndex = 0;
            this.getOrpoButton.Text = "ОРПО";
            this.getOrpoButton.UseVisualStyleBackColor = true;
            this.getOrpoButton.Click += new System.EventHandler(this.getOrpoButton_Click);
            // 
            // resultTextBox
            // 
            this.resultTextBox.Location = new System.Drawing.Point(103, 114);
            this.resultTextBox.Name = "resultTextBox";
            this.resultTextBox.Size = new System.Drawing.Size(100, 20);
            this.resultTextBox.TabIndex = 1;
            // 
            // resultLabel
            // 
            this.resultLabel.AutoSize = true;
            this.resultLabel.Location = new System.Drawing.Point(65, 117);
            this.resultLabel.Name = "resultLabel";
            this.resultLabel.Size = new System.Drawing.Size(32, 13);
            this.resultLabel.TabIndex = 4;
            this.resultLabel.Text = "result";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(103, 62);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(100, 20);
            this.textBox1.TabIndex = 5;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // getIbpoButton
            // 
            this.getIbpoButton.Location = new System.Drawing.Point(103, 144);
            this.getIbpoButton.Name = "getIbpoButton";
            this.getIbpoButton.Size = new System.Drawing.Size(75, 23);
            this.getIbpoButton.TabIndex = 6;
            this.getIbpoButton.Text = "ИБПО";
            this.getIbpoButton.UseVisualStyleBackColor = true;
            this.getIbpoButton.Click += new System.EventHandler(this.getIbpoButton_Click);
            // 
            // testTextBox
            // 
            this.testTextBox.Location = new System.Drawing.Point(103, 88);
            this.testTextBox.Name = "testTextBox";
            this.testTextBox.Size = new System.Drawing.Size(100, 20);
            this.testTextBox.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(65, 65);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(24, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "text";
            // 
            // testlabel
            // 
            this.testlabel.AutoSize = true;
            this.testlabel.Location = new System.Drawing.Point(65, 91);
            this.testlabel.Name = "testlabel";
            this.testlabel.Size = new System.Drawing.Size(24, 13);
            this.testlabel.TabIndex = 9;
            this.testlabel.Text = "test";
            // 
            // reslab
            // 
            this.reslab.AutoSize = true;
            this.reslab.Location = new System.Drawing.Point(209, 117);
            this.reslab.Name = "reslab";
            this.reslab.Size = new System.Drawing.Size(24, 13);
            this.reslab.TabIndex = 10;
            this.reslab.Text = "text";
            // 
            // testlab
            // 
            this.testlab.AutoSize = true;
            this.testlab.Location = new System.Drawing.Point(209, 91);
            this.testlab.Name = "testlab";
            this.testlab.Size = new System.Drawing.Size(24, 13);
            this.testlab.TabIndex = 11;
            this.testlab.Text = "text";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(292, 273);
            this.Controls.Add(this.testlab);
            this.Controls.Add(this.reslab);
            this.Controls.Add(this.testlabel);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.getIbpoButton);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.resultLabel);
            this.Controls.Add(this.testTextBox);
            this.Controls.Add(this.resultTextBox);
            this.Controls.Add(this.getOrpoButton);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button getOrpoButton;
        private System.Windows.Forms.TextBox resultTextBox;
        private System.Windows.Forms.Label resultLabel;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button getIbpoButton;
        private System.Windows.Forms.TextBox testTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label testlabel;
        private System.Windows.Forms.Label reslab;
        private System.Windows.Forms.Label testlab;


    }
}

