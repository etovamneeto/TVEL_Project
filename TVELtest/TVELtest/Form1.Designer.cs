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
            this.testTextBox = new System.Windows.Forms.TextBox();
            this.testLabel = new System.Windows.Forms.Label();
            this.resultLabel = new System.Windows.Forms.Label();
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
            this.resultTextBox.TextChanged += new System.EventHandler(this.resultTextBox_TextChanged);
            // 
            // testTextBox
            // 
            this.testTextBox.Location = new System.Drawing.Point(103, 88);
            this.testTextBox.Name = "testTextBox";
            this.testTextBox.Size = new System.Drawing.Size(100, 20);
            this.testTextBox.TabIndex = 2;
            this.testTextBox.TextChanged += new System.EventHandler(this.testTextBox_TextChanged);
            // 
            // testLabel
            // 
            this.testLabel.AutoSize = true;
            this.testLabel.Location = new System.Drawing.Point(65, 91);
            this.testLabel.Name = "testLabel";
            this.testLabel.Size = new System.Drawing.Size(24, 13);
            this.testLabel.TabIndex = 3;
            this.testLabel.Text = "test";
            this.testLabel.Click += new System.EventHandler(this.testLabel_Click);
            // 
            // resultLabel
            // 
            this.resultLabel.AutoSize = true;
            this.resultLabel.Location = new System.Drawing.Point(65, 117);
            this.resultLabel.Name = "resultLabel";
            this.resultLabel.Size = new System.Drawing.Size(32, 13);
            this.resultLabel.TabIndex = 4;
            this.resultLabel.Text = "result";
            this.resultLabel.Click += new System.EventHandler(this.resultLabel_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(292, 273);
            this.Controls.Add(this.resultLabel);
            this.Controls.Add(this.testLabel);
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
        private System.Windows.Forms.TextBox testTextBox;
        private System.Windows.Forms.Label testLabel;
        private System.Windows.Forms.Label resultLabel;


    }
}

