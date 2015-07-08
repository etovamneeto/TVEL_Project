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
            this.openFileButton = new System.Windows.Forms.Button();
            this.getOrpoAverDoseButton = new System.Windows.Forms.Button();
            this.getIbpoButton = new System.Windows.Forms.Button();
            this.orpoGroupBox = new System.Windows.Forms.GroupBox();
            this.womanIntOrpoBox95 = new System.Windows.Forms.TextBox();
            this.womanExtOrpoBox95 = new System.Windows.Forms.TextBox();
            this.orpoBoxWomanIntLabel95 = new System.Windows.Forms.Label();
            this.orpoBoxWomanExtLabel95 = new System.Windows.Forms.Label();
            this.womanIntOrpoBox = new System.Windows.Forms.TextBox();
            this.womanExtOrpoBox = new System.Windows.Forms.TextBox();
            this.orpoBoxWomanIntLabel = new System.Windows.Forms.Label();
            this.orpoBoxWomanExtLabel = new System.Windows.Forms.Label();
            this.manIntOrpoBox95 = new System.Windows.Forms.TextBox();
            this.manExtOrpoBox95 = new System.Windows.Forms.TextBox();
            this.orpoGroupManInt95Label = new System.Windows.Forms.Label();
            this.orpoGroupManExt95Label = new System.Windows.Forms.Label();
            this.manIntOrpoBox = new System.Windows.Forms.TextBox();
            this.manExtOrpoBox = new System.Windows.Forms.TextBox();
            this.orpoBoxManIntLabel = new System.Windows.Forms.Label();
            this.orpoBoxManExtLabel = new System.Windows.Forms.Label();
            this.ibpoGroupBox = new System.Windows.Forms.GroupBox();
            this.womanIntIbpoBox95 = new System.Windows.Forms.TextBox();
            this.womanExtIbpoBox95 = new System.Windows.Forms.TextBox();
            this.ibpoBoxWomanIntLabel95 = new System.Windows.Forms.Label();
            this.ibpoBoxWomanExtLabel95 = new System.Windows.Forms.Label();
            this.womanIntIbpoBox = new System.Windows.Forms.TextBox();
            this.womanExtIbpoBox = new System.Windows.Forms.TextBox();
            this.ibpoBoxWomanIntLabel = new System.Windows.Forms.Label();
            this.ibpoBoxWomanExtLabel = new System.Windows.Forms.Label();
            this.manIntIbpoBox95 = new System.Windows.Forms.TextBox();
            this.manExtIbpoBox95 = new System.Windows.Forms.TextBox();
            this.ibpoBoxManIntLabel95 = new System.Windows.Forms.Label();
            this.ibpoBoxManExtLabel95 = new System.Windows.Forms.Label();
            this.manIntIbpoBox = new System.Windows.Forms.TextBox();
            this.manExtIbpoBox = new System.Windows.Forms.TextBox();
            this.ibpoBoxManIntLabel = new System.Windows.Forms.Label();
            this.ibpoBoxManExtLabel = new System.Windows.Forms.Label();
            this.getOrpoAverLarButton = new System.Windows.Forms.Button();
            this.orpoGroupBox.SuspendLayout();
            this.ibpoGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileButton
            // 
            this.openFileButton.Location = new System.Drawing.Point(12, 12);
            this.openFileButton.Name = "openFileButton";
            this.openFileButton.Size = new System.Drawing.Size(468, 43);
            this.openFileButton.TabIndex = 12;
            this.openFileButton.Text = "Шаг 1: Выберите базу данных";
            this.openFileButton.UseVisualStyleBackColor = true;
            this.openFileButton.Click += new System.EventHandler(this.openFileButton_Click);
            // 
            // getOrpoAverDoseButton
            // 
            this.getOrpoAverDoseButton.Location = new System.Drawing.Point(12, 61);
            this.getOrpoAverDoseButton.Name = "getOrpoAverDoseButton";
            this.getOrpoAverDoseButton.Size = new System.Drawing.Size(234, 43);
            this.getOrpoAverDoseButton.TabIndex = 0;
            this.getOrpoAverDoseButton.Text = "Шаг 2-а: Рассчитать ОРПО (Ср. доза)";
            this.getOrpoAverDoseButton.UseVisualStyleBackColor = true;
            this.getOrpoAverDoseButton.Click += new System.EventHandler(this.getOrpoAverDoseButton_Click);
            // 
            // getIbpoButton
            // 
            this.getIbpoButton.Location = new System.Drawing.Point(12, 110);
            this.getIbpoButton.Name = "getIbpoButton";
            this.getIbpoButton.Size = new System.Drawing.Size(468, 43);
            this.getIbpoButton.TabIndex = 6;
            this.getIbpoButton.Text = "Шаг 3: Рассчитать ИБПО";
            this.getIbpoButton.UseVisualStyleBackColor = true;
            this.getIbpoButton.Click += new System.EventHandler(this.getIbpoButton_Click);
            // 
            // orpoGroupBox
            // 
            this.orpoGroupBox.Controls.Add(this.womanIntOrpoBox95);
            this.orpoGroupBox.Controls.Add(this.womanExtOrpoBox95);
            this.orpoGroupBox.Controls.Add(this.orpoBoxWomanIntLabel95);
            this.orpoGroupBox.Controls.Add(this.orpoBoxWomanExtLabel95);
            this.orpoGroupBox.Controls.Add(this.womanIntOrpoBox);
            this.orpoGroupBox.Controls.Add(this.womanExtOrpoBox);
            this.orpoGroupBox.Controls.Add(this.orpoBoxWomanIntLabel);
            this.orpoGroupBox.Controls.Add(this.orpoBoxWomanExtLabel);
            this.orpoGroupBox.Controls.Add(this.manIntOrpoBox95);
            this.orpoGroupBox.Controls.Add(this.manExtOrpoBox95);
            this.orpoGroupBox.Controls.Add(this.orpoGroupManInt95Label);
            this.orpoGroupBox.Controls.Add(this.orpoGroupManExt95Label);
            this.orpoGroupBox.Controls.Add(this.manIntOrpoBox);
            this.orpoGroupBox.Controls.Add(this.manExtOrpoBox);
            this.orpoGroupBox.Controls.Add(this.orpoBoxManIntLabel);
            this.orpoGroupBox.Controls.Add(this.orpoBoxManExtLabel);
            this.orpoGroupBox.Location = new System.Drawing.Point(12, 161);
            this.orpoGroupBox.Name = "orpoGroupBox";
            this.orpoGroupBox.Size = new System.Drawing.Size(468, 150);
            this.orpoGroupBox.TabIndex = 13;
            this.orpoGroupBox.TabStop = false;
            this.orpoGroupBox.Text = "Среднее ОРПО";
            // 
            // womanIntOrpoBox95
            // 
            this.womanIntOrpoBox95.Location = new System.Drawing.Point(362, 105);
            this.womanIntOrpoBox95.Name = "womanIntOrpoBox95";
            this.womanIntOrpoBox95.Size = new System.Drawing.Size(100, 20);
            this.womanIntOrpoBox95.TabIndex = 15;
            // 
            // womanExtOrpoBox95
            // 
            this.womanExtOrpoBox95.Location = new System.Drawing.Point(362, 79);
            this.womanExtOrpoBox95.Name = "womanExtOrpoBox95";
            this.womanExtOrpoBox95.Size = new System.Drawing.Size(100, 20);
            this.womanExtOrpoBox95.TabIndex = 14;
            // 
            // orpoBoxWomanIntLabel95
            // 
            this.orpoBoxWomanIntLabel95.AutoSize = true;
            this.orpoBoxWomanIntLabel95.Location = new System.Drawing.Point(235, 108);
            this.orpoBoxWomanIntLabel95.Name = "orpoBoxWomanIntLabel95";
            this.orpoBoxWomanIntLabel95.Size = new System.Drawing.Size(125, 13);
            this.orpoBoxWomanIntLabel95.TabIndex = 13;
            this.orpoBoxWomanIntLabel95.Text = "Женщины, внутр. (95%)";
            // 
            // orpoBoxWomanExtLabel95
            // 
            this.orpoBoxWomanExtLabel95.AutoSize = true;
            this.orpoBoxWomanExtLabel95.Location = new System.Drawing.Point(235, 82);
            this.orpoBoxWomanExtLabel95.Name = "orpoBoxWomanExtLabel95";
            this.orpoBoxWomanExtLabel95.Size = new System.Drawing.Size(123, 13);
            this.orpoBoxWomanExtLabel95.TabIndex = 12;
            this.orpoBoxWomanExtLabel95.Text = "Женщины, внеш. (95%)";
            // 
            // womanIntOrpoBox
            // 
            this.womanIntOrpoBox.Location = new System.Drawing.Point(102, 105);
            this.womanIntOrpoBox.Name = "womanIntOrpoBox";
            this.womanIntOrpoBox.Size = new System.Drawing.Size(100, 20);
            this.womanIntOrpoBox.TabIndex = 11;
            // 
            // womanExtOrpoBox
            // 
            this.womanExtOrpoBox.Location = new System.Drawing.Point(102, 79);
            this.womanExtOrpoBox.Name = "womanExtOrpoBox";
            this.womanExtOrpoBox.Size = new System.Drawing.Size(100, 20);
            this.womanExtOrpoBox.TabIndex = 10;
            // 
            // orpoBoxWomanIntLabel
            // 
            this.orpoBoxWomanIntLabel.AutoSize = true;
            this.orpoBoxWomanIntLabel.Location = new System.Drawing.Point(5, 108);
            this.orpoBoxWomanIntLabel.Name = "orpoBoxWomanIntLabel";
            this.orpoBoxWomanIntLabel.Size = new System.Drawing.Size(96, 13);
            this.orpoBoxWomanIntLabel.TabIndex = 9;
            this.orpoBoxWomanIntLabel.Text = "Женщины, внутр.";
            // 
            // orpoBoxWomanExtLabel
            // 
            this.orpoBoxWomanExtLabel.AutoSize = true;
            this.orpoBoxWomanExtLabel.Location = new System.Drawing.Point(5, 82);
            this.orpoBoxWomanExtLabel.Name = "orpoBoxWomanExtLabel";
            this.orpoBoxWomanExtLabel.Size = new System.Drawing.Size(94, 13);
            this.orpoBoxWomanExtLabel.TabIndex = 8;
            this.orpoBoxWomanExtLabel.Text = "Женщины, внеш.";
            // 
            // manIntOrpoBox95
            // 
            this.manIntOrpoBox95.Location = new System.Drawing.Point(362, 53);
            this.manIntOrpoBox95.Name = "manIntOrpoBox95";
            this.manIntOrpoBox95.Size = new System.Drawing.Size(100, 20);
            this.manIntOrpoBox95.TabIndex = 7;
            // 
            // manExtOrpoBox95
            // 
            this.manExtOrpoBox95.Location = new System.Drawing.Point(362, 30);
            this.manExtOrpoBox95.Name = "manExtOrpoBox95";
            this.manExtOrpoBox95.Size = new System.Drawing.Size(100, 20);
            this.manExtOrpoBox95.TabIndex = 6;
            // 
            // orpoGroupManInt95Label
            // 
            this.orpoGroupManInt95Label.AutoSize = true;
            this.orpoGroupManInt95Label.Location = new System.Drawing.Point(235, 56);
            this.orpoGroupManInt95Label.Name = "orpoGroupManInt95Label";
            this.orpoGroupManInt95Label.Size = new System.Drawing.Size(120, 13);
            this.orpoGroupManInt95Label.TabIndex = 5;
            this.orpoGroupManInt95Label.Text = "Мужчины, внутр. (95%)";
            // 
            // orpoGroupManExt95Label
            // 
            this.orpoGroupManExt95Label.AutoSize = true;
            this.orpoGroupManExt95Label.Location = new System.Drawing.Point(235, 30);
            this.orpoGroupManExt95Label.Name = "orpoGroupManExt95Label";
            this.orpoGroupManExt95Label.Size = new System.Drawing.Size(118, 13);
            this.orpoGroupManExt95Label.TabIndex = 4;
            this.orpoGroupManExt95Label.Text = "Мужчины, внеш. (95%)";
            // 
            // manIntOrpoBox
            // 
            this.manIntOrpoBox.Location = new System.Drawing.Point(102, 53);
            this.manIntOrpoBox.Name = "manIntOrpoBox";
            this.manIntOrpoBox.Size = new System.Drawing.Size(100, 20);
            this.manIntOrpoBox.TabIndex = 3;
            // 
            // manExtOrpoBox
            // 
            this.manExtOrpoBox.Location = new System.Drawing.Point(102, 27);
            this.manExtOrpoBox.Name = "manExtOrpoBox";
            this.manExtOrpoBox.Size = new System.Drawing.Size(100, 20);
            this.manExtOrpoBox.TabIndex = 2;
            // 
            // orpoBoxManIntLabel
            // 
            this.orpoBoxManIntLabel.AutoSize = true;
            this.orpoBoxManIntLabel.Location = new System.Drawing.Point(5, 56);
            this.orpoBoxManIntLabel.Name = "orpoBoxManIntLabel";
            this.orpoBoxManIntLabel.Size = new System.Drawing.Size(91, 13);
            this.orpoBoxManIntLabel.TabIndex = 1;
            this.orpoBoxManIntLabel.Text = "Мужчины, внутр.";
            // 
            // orpoBoxManExtLabel
            // 
            this.orpoBoxManExtLabel.AutoSize = true;
            this.orpoBoxManExtLabel.Location = new System.Drawing.Point(5, 30);
            this.orpoBoxManExtLabel.Name = "orpoBoxManExtLabel";
            this.orpoBoxManExtLabel.Size = new System.Drawing.Size(89, 13);
            this.orpoBoxManExtLabel.TabIndex = 0;
            this.orpoBoxManExtLabel.Text = "Мужчины, внеш.";
            // 
            // ibpoGroupBox
            // 
            this.ibpoGroupBox.Controls.Add(this.womanIntIbpoBox95);
            this.ibpoGroupBox.Controls.Add(this.womanExtIbpoBox95);
            this.ibpoGroupBox.Controls.Add(this.ibpoBoxWomanIntLabel95);
            this.ibpoGroupBox.Controls.Add(this.ibpoBoxWomanExtLabel95);
            this.ibpoGroupBox.Controls.Add(this.womanIntIbpoBox);
            this.ibpoGroupBox.Controls.Add(this.womanExtIbpoBox);
            this.ibpoGroupBox.Controls.Add(this.ibpoBoxWomanIntLabel);
            this.ibpoGroupBox.Controls.Add(this.ibpoBoxWomanExtLabel);
            this.ibpoGroupBox.Controls.Add(this.manIntIbpoBox95);
            this.ibpoGroupBox.Controls.Add(this.manExtIbpoBox95);
            this.ibpoGroupBox.Controls.Add(this.ibpoBoxManIntLabel95);
            this.ibpoGroupBox.Controls.Add(this.ibpoBoxManExtLabel95);
            this.ibpoGroupBox.Controls.Add(this.manIntIbpoBox);
            this.ibpoGroupBox.Controls.Add(this.manExtIbpoBox);
            this.ibpoGroupBox.Controls.Add(this.ibpoBoxManIntLabel);
            this.ibpoGroupBox.Controls.Add(this.ibpoBoxManExtLabel);
            this.ibpoGroupBox.Location = new System.Drawing.Point(12, 317);
            this.ibpoGroupBox.Name = "ibpoGroupBox";
            this.ibpoGroupBox.Size = new System.Drawing.Size(468, 150);
            this.ibpoGroupBox.TabIndex = 14;
            this.ibpoGroupBox.TabStop = false;
            this.ibpoGroupBox.Text = "Среднее ИБПО";
            // 
            // womanIntIbpoBox95
            // 
            this.womanIntIbpoBox95.Location = new System.Drawing.Point(362, 105);
            this.womanIntIbpoBox95.Name = "womanIntIbpoBox95";
            this.womanIntIbpoBox95.Size = new System.Drawing.Size(100, 20);
            this.womanIntIbpoBox95.TabIndex = 15;
            // 
            // womanExtIbpoBox95
            // 
            this.womanExtIbpoBox95.Location = new System.Drawing.Point(362, 79);
            this.womanExtIbpoBox95.Name = "womanExtIbpoBox95";
            this.womanExtIbpoBox95.Size = new System.Drawing.Size(100, 20);
            this.womanExtIbpoBox95.TabIndex = 14;
            // 
            // ibpoBoxWomanIntLabel95
            // 
            this.ibpoBoxWomanIntLabel95.AutoSize = true;
            this.ibpoBoxWomanIntLabel95.Location = new System.Drawing.Point(235, 108);
            this.ibpoBoxWomanIntLabel95.Name = "ibpoBoxWomanIntLabel95";
            this.ibpoBoxWomanIntLabel95.Size = new System.Drawing.Size(125, 13);
            this.ibpoBoxWomanIntLabel95.TabIndex = 13;
            this.ibpoBoxWomanIntLabel95.Text = "Женщины, внутр. (95%)";
            // 
            // ibpoBoxWomanExtLabel95
            // 
            this.ibpoBoxWomanExtLabel95.AutoSize = true;
            this.ibpoBoxWomanExtLabel95.Location = new System.Drawing.Point(235, 82);
            this.ibpoBoxWomanExtLabel95.Name = "ibpoBoxWomanExtLabel95";
            this.ibpoBoxWomanExtLabel95.Size = new System.Drawing.Size(123, 13);
            this.ibpoBoxWomanExtLabel95.TabIndex = 12;
            this.ibpoBoxWomanExtLabel95.Text = "Женщины, внеш. (95%)";
            // 
            // womanIntIbpoBox
            // 
            this.womanIntIbpoBox.Location = new System.Drawing.Point(102, 105);
            this.womanIntIbpoBox.Name = "womanIntIbpoBox";
            this.womanIntIbpoBox.Size = new System.Drawing.Size(100, 20);
            this.womanIntIbpoBox.TabIndex = 11;
            // 
            // womanExtIbpoBox
            // 
            this.womanExtIbpoBox.Location = new System.Drawing.Point(102, 79);
            this.womanExtIbpoBox.Name = "womanExtIbpoBox";
            this.womanExtIbpoBox.Size = new System.Drawing.Size(100, 20);
            this.womanExtIbpoBox.TabIndex = 10;
            // 
            // ibpoBoxWomanIntLabel
            // 
            this.ibpoBoxWomanIntLabel.AutoSize = true;
            this.ibpoBoxWomanIntLabel.Location = new System.Drawing.Point(5, 108);
            this.ibpoBoxWomanIntLabel.Name = "ibpoBoxWomanIntLabel";
            this.ibpoBoxWomanIntLabel.Size = new System.Drawing.Size(96, 13);
            this.ibpoBoxWomanIntLabel.TabIndex = 9;
            this.ibpoBoxWomanIntLabel.Text = "Женщины, внутр.";
            // 
            // ibpoBoxWomanExtLabel
            // 
            this.ibpoBoxWomanExtLabel.AutoSize = true;
            this.ibpoBoxWomanExtLabel.Location = new System.Drawing.Point(5, 82);
            this.ibpoBoxWomanExtLabel.Name = "ibpoBoxWomanExtLabel";
            this.ibpoBoxWomanExtLabel.Size = new System.Drawing.Size(94, 13);
            this.ibpoBoxWomanExtLabel.TabIndex = 8;
            this.ibpoBoxWomanExtLabel.Text = "Женщины, внеш.";
            // 
            // manIntIbpoBox95
            // 
            this.manIntIbpoBox95.Location = new System.Drawing.Point(362, 53);
            this.manIntIbpoBox95.Name = "manIntIbpoBox95";
            this.manIntIbpoBox95.Size = new System.Drawing.Size(100, 20);
            this.manIntIbpoBox95.TabIndex = 7;
            // 
            // manExtIbpoBox95
            // 
            this.manExtIbpoBox95.Location = new System.Drawing.Point(362, 27);
            this.manExtIbpoBox95.Name = "manExtIbpoBox95";
            this.manExtIbpoBox95.Size = new System.Drawing.Size(100, 20);
            this.manExtIbpoBox95.TabIndex = 6;
            // 
            // ibpoBoxManIntLabel95
            // 
            this.ibpoBoxManIntLabel95.AutoSize = true;
            this.ibpoBoxManIntLabel95.Location = new System.Drawing.Point(235, 56);
            this.ibpoBoxManIntLabel95.Name = "ibpoBoxManIntLabel95";
            this.ibpoBoxManIntLabel95.Size = new System.Drawing.Size(120, 13);
            this.ibpoBoxManIntLabel95.TabIndex = 5;
            this.ibpoBoxManIntLabel95.Text = "Мужчины, внутр. (95%)";
            // 
            // ibpoBoxManExtLabel95
            // 
            this.ibpoBoxManExtLabel95.AutoSize = true;
            this.ibpoBoxManExtLabel95.Location = new System.Drawing.Point(235, 30);
            this.ibpoBoxManExtLabel95.Name = "ibpoBoxManExtLabel95";
            this.ibpoBoxManExtLabel95.Size = new System.Drawing.Size(118, 13);
            this.ibpoBoxManExtLabel95.TabIndex = 4;
            this.ibpoBoxManExtLabel95.Text = "Мужчины, внеш. (95%)";
            // 
            // manIntIbpoBox
            // 
            this.manIntIbpoBox.Location = new System.Drawing.Point(102, 53);
            this.manIntIbpoBox.Name = "manIntIbpoBox";
            this.manIntIbpoBox.Size = new System.Drawing.Size(100, 20);
            this.manIntIbpoBox.TabIndex = 3;
            // 
            // manExtIbpoBox
            // 
            this.manExtIbpoBox.Location = new System.Drawing.Point(102, 27);
            this.manExtIbpoBox.Name = "manExtIbpoBox";
            this.manExtIbpoBox.Size = new System.Drawing.Size(100, 20);
            this.manExtIbpoBox.TabIndex = 2;
            // 
            // ibpoBoxManIntLabel
            // 
            this.ibpoBoxManIntLabel.AutoSize = true;
            this.ibpoBoxManIntLabel.Location = new System.Drawing.Point(5, 56);
            this.ibpoBoxManIntLabel.Name = "ibpoBoxManIntLabel";
            this.ibpoBoxManIntLabel.Size = new System.Drawing.Size(91, 13);
            this.ibpoBoxManIntLabel.TabIndex = 1;
            this.ibpoBoxManIntLabel.Text = "Мужчины, внутр.";
            // 
            // ibpoBoxManExtLabel
            // 
            this.ibpoBoxManExtLabel.AutoSize = true;
            this.ibpoBoxManExtLabel.Location = new System.Drawing.Point(5, 30);
            this.ibpoBoxManExtLabel.Name = "ibpoBoxManExtLabel";
            this.ibpoBoxManExtLabel.Size = new System.Drawing.Size(89, 13);
            this.ibpoBoxManExtLabel.TabIndex = 0;
            this.ibpoBoxManExtLabel.Text = "Мужчины, внеш.";
            // 
            // getOrpoAverLarButton
            // 
            this.getOrpoAverLarButton.Location = new System.Drawing.Point(246, 61);
            this.getOrpoAverLarButton.Name = "getOrpoAverLarButton";
            this.getOrpoAverLarButton.Size = new System.Drawing.Size(234, 43);
            this.getOrpoAverLarButton.TabIndex = 15;
            this.getOrpoAverLarButton.Text = "Шаг 2-б: Рассчитать ОРПО (Ср. LAR)";
            this.getOrpoAverLarButton.UseVisualStyleBackColor = true;
            this.getOrpoAverLarButton.Click += new System.EventHandler(this.getOrpoAverLarButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(492, 473);
            this.Controls.Add(this.getOrpoAverLarButton);
            this.Controls.Add(this.ibpoGroupBox);
            this.Controls.Add(this.orpoGroupBox);
            this.Controls.Add(this.openFileButton);
            this.Controls.Add(this.getIbpoButton);
            this.Controls.Add(this.getOrpoAverDoseButton);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.orpoGroupBox.ResumeLayout(false);
            this.orpoGroupBox.PerformLayout();
            this.ibpoGroupBox.ResumeLayout(false);
            this.ibpoGroupBox.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button openFileButton;
        private System.Windows.Forms.Button getOrpoAverDoseButton;
        private System.Windows.Forms.Button getIbpoButton;
        private System.Windows.Forms.GroupBox orpoGroupBox;
        private System.Windows.Forms.TextBox womanIntOrpoBox95;
        private System.Windows.Forms.TextBox womanExtOrpoBox95;
        private System.Windows.Forms.Label orpoBoxWomanIntLabel95;
        private System.Windows.Forms.Label orpoBoxWomanExtLabel95;
        private System.Windows.Forms.TextBox womanIntOrpoBox;
        private System.Windows.Forms.TextBox womanExtOrpoBox;
        private System.Windows.Forms.Label orpoBoxWomanIntLabel;
        private System.Windows.Forms.Label orpoBoxWomanExtLabel;
        private System.Windows.Forms.TextBox manIntOrpoBox95;
        private System.Windows.Forms.TextBox manExtOrpoBox95;
        private System.Windows.Forms.Label orpoGroupManInt95Label;
        private System.Windows.Forms.Label orpoGroupManExt95Label;
        private System.Windows.Forms.TextBox manIntOrpoBox;
        private System.Windows.Forms.TextBox manExtOrpoBox;
        private System.Windows.Forms.Label orpoBoxManIntLabel;
        private System.Windows.Forms.Label orpoBoxManExtLabel;
        private System.Windows.Forms.GroupBox ibpoGroupBox;
        private System.Windows.Forms.TextBox womanIntIbpoBox95;
        private System.Windows.Forms.TextBox womanExtIbpoBox95;
        private System.Windows.Forms.Label ibpoBoxWomanIntLabel95;
        private System.Windows.Forms.Label ibpoBoxWomanExtLabel95;
        private System.Windows.Forms.TextBox womanIntIbpoBox;
        private System.Windows.Forms.TextBox womanExtIbpoBox;
        private System.Windows.Forms.Label ibpoBoxWomanIntLabel;
        private System.Windows.Forms.Label ibpoBoxWomanExtLabel;
        private System.Windows.Forms.TextBox manIntIbpoBox95;
        private System.Windows.Forms.TextBox manExtIbpoBox95;
        private System.Windows.Forms.Label ibpoBoxManIntLabel95;
        private System.Windows.Forms.Label ibpoBoxManExtLabel95;
        private System.Windows.Forms.TextBox manIntIbpoBox;
        private System.Windows.Forms.TextBox manExtIbpoBox;
        private System.Windows.Forms.Label ibpoBoxManIntLabel;
        private System.Windows.Forms.Label ibpoBoxManExtLabel;
        private System.Windows.Forms.Button getOrpoAverLarButton;


    }
}

