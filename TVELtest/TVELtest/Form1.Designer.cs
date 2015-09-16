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
            this.getOrpoButton = new System.Windows.Forms.Button();
            this.getIbpoButton = new System.Windows.Forms.Button();
            this.shopComboBox = new System.Windows.Forms.ComboBox();
            this.shopNameLabel = new System.Windows.Forms.Label();
            this.tabControl = new System.Windows.Forms.TabControl();
            this.manOrpoPage = new System.Windows.Forms.TabPage();
            this.manOrpoGridView = new System.Windows.Forms.DataGridView();
            this.womanOrpoPage = new System.Windows.Forms.TabPage();
            this.womanOrpoGridView = new System.Windows.Forms.DataGridView();
            this.manIbpoPage = new System.Windows.Forms.TabPage();
            this.manIbpoGridView = new System.Windows.Forms.DataGridView();
            this.womanIbpoPage = new System.Windows.Forms.TabPage();
            this.larOrDetGroup = new System.Windows.Forms.GroupBox();
            this.larRB = new System.Windows.Forms.RadioButton();
            this.detRB = new System.Windows.Forms.RadioButton();
            this.womanIbpoGridView = new System.Windows.Forms.DataGridView();
            this.aMethodRB = new System.Windows.Forms.RadioButton();
            this.bMethodRB = new System.Windows.Forms.RadioButton();
            this.methodGroup = new System.Windows.Forms.GroupBox();
            this.tabControl.SuspendLayout();
            this.manOrpoPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.manOrpoGridView)).BeginInit();
            this.womanOrpoPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.womanOrpoGridView)).BeginInit();
            this.manIbpoPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.manIbpoGridView)).BeginInit();
            this.womanIbpoPage.SuspendLayout();
            this.larOrDetGroup.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.womanIbpoGridView)).BeginInit();
            this.methodGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileButton
            // 
            this.openFileButton.Location = new System.Drawing.Point(12, 12);
            this.openFileButton.Name = "openFileButton";
            this.openFileButton.Size = new System.Drawing.Size(468, 43);
            this.openFileButton.TabIndex = 12;
            this.openFileButton.Text = "Обновление базы данных";
            this.openFileButton.UseVisualStyleBackColor = true;
            this.openFileButton.Click += new System.EventHandler(this.openFileButton_Click);
            // 
            // getOrpoButton
            // 
            this.getOrpoButton.Location = new System.Drawing.Point(12, 61);
            this.getOrpoButton.Name = "getOrpoButton";
            this.getOrpoButton.Size = new System.Drawing.Size(468, 43);
            this.getOrpoButton.TabIndex = 0;
            this.getOrpoButton.Text = "Рассчитать ОРПО";
            this.getOrpoButton.UseVisualStyleBackColor = true;
            this.getOrpoButton.Click += new System.EventHandler(this.getOrpoButton_Click);
            // 
            // getIbpoButton
            // 
            this.getIbpoButton.Location = new System.Drawing.Point(12, 110);
            this.getIbpoButton.Name = "getIbpoButton";
            this.getIbpoButton.Size = new System.Drawing.Size(468, 43);
            this.getIbpoButton.TabIndex = 6;
            this.getIbpoButton.Text = "Рассчитать ИБПО";
            this.getIbpoButton.UseVisualStyleBackColor = true;
            this.getIbpoButton.Click += new System.EventHandler(this.getIbpoButton_Click);
            // 
            // shopComboBox
            // 
            this.shopComboBox.FormattingEnabled = true;
            this.shopComboBox.Items.AddRange(new object[] {
            "СХК",
            "АЭХК",
            "МСЗ",
            "УЭХК",
            "ПО ЭХЗ",
            "ЧМЗ",
            "ВСЕ ПРЕДПРИЯТИЯ"});
            this.shopComboBox.Location = new System.Drawing.Point(489, 34);
            this.shopComboBox.Name = "shopComboBox";
            this.shopComboBox.Size = new System.Drawing.Size(175, 21);
            this.shopComboBox.TabIndex = 18;
            this.shopComboBox.SelectedIndexChanged += new System.EventHandler(this.shopComboBox_SelectedIndexChanged);
            // 
            // shopNameLabel
            // 
            this.shopNameLabel.AutoSize = true;
            this.shopNameLabel.Location = new System.Drawing.Point(486, 18);
            this.shopNameLabel.Name = "shopNameLabel";
            this.shopNameLabel.Size = new System.Drawing.Size(147, 13);
            this.shopNameLabel.TabIndex = 19;
            this.shopNameLabel.Text = "Топливная компания ТВЭЛ";
            // 
            // tabControl
            // 
            this.tabControl.Controls.Add(this.manOrpoPage);
            this.tabControl.Controls.Add(this.womanOrpoPage);
            this.tabControl.Controls.Add(this.manIbpoPage);
            this.tabControl.Controls.Add(this.womanIbpoPage);
            this.tabControl.Location = new System.Drawing.Point(12, 159);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(718, 302);
            this.tabControl.TabIndex = 20;
            // 
            // manOrpoPage
            // 
            this.manOrpoPage.Controls.Add(this.manOrpoGridView);
            this.manOrpoPage.Location = new System.Drawing.Point(4, 22);
            this.manOrpoPage.Name = "manOrpoPage";
            this.manOrpoPage.Padding = new System.Windows.Forms.Padding(3);
            this.manOrpoPage.Size = new System.Drawing.Size(710, 276);
            this.manOrpoPage.TabIndex = 0;
            this.manOrpoPage.Text = "ОРПО, Мужчины";
            this.manOrpoPage.UseVisualStyleBackColor = true;
            // 
            // manOrpoGridView
            // 
            this.manOrpoGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.manOrpoGridView.Location = new System.Drawing.Point(0, 0);
            this.manOrpoGridView.Name = "manOrpoGridView";
            this.manOrpoGridView.Size = new System.Drawing.Size(710, 276);
            this.manOrpoGridView.TabIndex = 0;
            // 
            // womanOrpoPage
            // 
            this.womanOrpoPage.Controls.Add(this.womanOrpoGridView);
            this.womanOrpoPage.Location = new System.Drawing.Point(4, 22);
            this.womanOrpoPage.Name = "womanOrpoPage";
            this.womanOrpoPage.Padding = new System.Windows.Forms.Padding(3);
            this.womanOrpoPage.Size = new System.Drawing.Size(710, 276);
            this.womanOrpoPage.TabIndex = 1;
            this.womanOrpoPage.Text = "ОРПО, Женщины";
            this.womanOrpoPage.UseVisualStyleBackColor = true;
            // 
            // womanOrpoGridView
            // 
            this.womanOrpoGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.womanOrpoGridView.Location = new System.Drawing.Point(0, 0);
            this.womanOrpoGridView.Name = "womanOrpoGridView";
            this.womanOrpoGridView.Size = new System.Drawing.Size(710, 276);
            this.womanOrpoGridView.TabIndex = 0;
            // 
            // manIbpoPage
            // 
            this.manIbpoPage.Controls.Add(this.manIbpoGridView);
            this.manIbpoPage.Location = new System.Drawing.Point(4, 22);
            this.manIbpoPage.Name = "manIbpoPage";
            this.manIbpoPage.Padding = new System.Windows.Forms.Padding(3);
            this.manIbpoPage.Size = new System.Drawing.Size(710, 276);
            this.manIbpoPage.TabIndex = 2;
            this.manIbpoPage.Text = "ИБПО, Мужчины";
            this.manIbpoPage.UseVisualStyleBackColor = true;
            // 
            // manIbpoGridView
            // 
            this.manIbpoGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.manIbpoGridView.Location = new System.Drawing.Point(0, 0);
            this.manIbpoGridView.Name = "manIbpoGridView";
            this.manIbpoGridView.Size = new System.Drawing.Size(710, 276);
            this.manIbpoGridView.TabIndex = 0;
            // 
            // womanIbpoPage
            // 
            this.womanIbpoPage.Controls.Add(this.womanIbpoGridView);
            this.womanIbpoPage.Location = new System.Drawing.Point(4, 22);
            this.womanIbpoPage.Name = "womanIbpoPage";
            this.womanIbpoPage.Padding = new System.Windows.Forms.Padding(3);
            this.womanIbpoPage.Size = new System.Drawing.Size(710, 276);
            this.womanIbpoPage.TabIndex = 3;
            this.womanIbpoPage.Text = "ИБПО, Женщины";
            this.womanIbpoPage.UseVisualStyleBackColor = true;
            // 
            // larOrDetGroup
            // 
            this.larOrDetGroup.Controls.Add(this.larRB);
            this.larOrDetGroup.Controls.Add(this.detRB);
            this.larOrDetGroup.Location = new System.Drawing.Point(651, 61);
            this.larOrDetGroup.Name = "larOrDetGroup";
            this.larOrDetGroup.Size = new System.Drawing.Size(86, 67);
            this.larOrDetGroup.TabIndex = 25;
            this.larOrDetGroup.TabStop = false;
            this.larOrDetGroup.Text = "LAR (Det)";
            // 
            // larRB
            // 
            this.larRB.AutoSize = true;
            this.larRB.Location = new System.Drawing.Point(6, 19);
            this.larRB.Name = "larRB";
            this.larRB.Size = new System.Drawing.Size(46, 17);
            this.larRB.TabIndex = 16;
            this.larRB.TabStop = true;
            this.larRB.Text = "LAR";
            this.larRB.UseVisualStyleBackColor = true;
            // 
            // detRB
            // 
            this.detRB.AutoSize = true;
            this.detRB.Location = new System.Drawing.Point(6, 42);
            this.detRB.Name = "detRB";
            this.detRB.Size = new System.Drawing.Size(42, 17);
            this.detRB.TabIndex = 17;
            this.detRB.TabStop = true;
            this.detRB.Text = "Det";
            this.detRB.UseVisualStyleBackColor = true;
            // 
            // womanIbpoGridView
            // 
            this.womanIbpoGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.womanIbpoGridView.Location = new System.Drawing.Point(0, 0);
            this.womanIbpoGridView.Name = "womanIbpoGridView";
            this.womanIbpoGridView.Size = new System.Drawing.Size(710, 276);
            this.womanIbpoGridView.TabIndex = 0;
            // 
            // aMethodRB
            // 
            this.aMethodRB.AutoSize = true;
            this.aMethodRB.Location = new System.Drawing.Point(6, 19);
            this.aMethodRB.Name = "aMethodRB";
            this.aMethodRB.Size = new System.Drawing.Size(140, 17);
            this.aMethodRB.TabIndex = 21;
            this.aMethodRB.TabStop = true;
            this.aMethodRB.Text = "Метод А (упрощенный)";
            this.aMethodRB.UseVisualStyleBackColor = true;
            this.aMethodRB.CheckedChanged += new System.EventHandler(this.aMethodRB_CheckedChanged);
            // 
            // bMethodRB
            // 
            this.bMethodRB.AutoSize = true;
            this.bMethodRB.Location = new System.Drawing.Point(6, 42);
            this.bMethodRB.Name = "bMethodRB";
            this.bMethodRB.Size = new System.Drawing.Size(149, 17);
            this.bMethodRB.TabIndex = 22;
            this.bMethodRB.TabStop = true;
            this.bMethodRB.Text = "Метод Б (точная оценка)";
            this.bMethodRB.UseVisualStyleBackColor = true;
            this.bMethodRB.CheckedChanged += new System.EventHandler(this.bMethodRB_CheckedChanged);
            // 
            // methodGroup
            // 
            this.methodGroup.Controls.Add(this.aMethodRB);
            this.methodGroup.Controls.Add(this.bMethodRB);
            this.methodGroup.Location = new System.Drawing.Point(486, 61);
            this.methodGroup.Name = "methodGroup";
            this.methodGroup.Size = new System.Drawing.Size(159, 67);
            this.methodGroup.TabIndex = 23;
            this.methodGroup.TabStop = false;
            this.methodGroup.Text = "Методы";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(742, 473);
            this.Controls.Add(this.larOrDetGroup);
            this.Controls.Add(this.methodGroup);
            this.Controls.Add(this.shopNameLabel);
            this.Controls.Add(this.shopComboBox);
            this.Controls.Add(this.openFileButton);
            this.Controls.Add(this.getIbpoButton);
            this.Controls.Add(this.getOrpoButton);
            this.Controls.Add(this.tabControl);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl.ResumeLayout(false);
            this.manOrpoPage.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.manOrpoGridView)).EndInit();
            this.womanOrpoPage.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.womanOrpoGridView)).EndInit();
            this.manIbpoPage.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.manIbpoGridView)).EndInit();
            this.womanIbpoPage.ResumeLayout(false);
            this.larOrDetGroup.ResumeLayout(false);
            this.larOrDetGroup.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.womanIbpoGridView)).EndInit();
            this.methodGroup.ResumeLayout(false);
            this.methodGroup.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button openFileButton;
        private System.Windows.Forms.Button getOrpoButton;
        private System.Windows.Forms.Button getIbpoButton;
        private System.Windows.Forms.ComboBox shopComboBox;
        private System.Windows.Forms.Label shopNameLabel;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage manOrpoPage;
        private System.Windows.Forms.TabPage womanOrpoPage;
        private System.Windows.Forms.RadioButton aMethodRB;
        private System.Windows.Forms.RadioButton bMethodRB;
        private System.Windows.Forms.TabPage manIbpoPage;
        private System.Windows.Forms.TabPage womanIbpoPage;
        private System.Windows.Forms.DataGridView manOrpoGridView;
        private System.Windows.Forms.DataGridView womanOrpoGridView;
        private System.Windows.Forms.DataGridView manIbpoGridView;
        private System.Windows.Forms.DataGridView womanIbpoGridView;
        private System.Windows.Forms.GroupBox methodGroup;
        private System.Windows.Forms.GroupBox larOrDetGroup;
        private System.Windows.Forms.RadioButton larRB;
        private System.Windows.Forms.RadioButton detRB;


    }
}

