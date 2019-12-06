namespace Aviasale
{
    partial class GetInfoForm
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtKids = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtAdult = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.coBoxTo = new System.Windows.Forms.ComboBox();
            this.coBoxFrom = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnNext = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtKids);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.txtAdult);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(151, 103);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Людей";
            // 
            // txtKids
            // 
            this.txtKids.Location = new System.Drawing.Point(79, 62);
            this.txtKids.Name = "txtKids";
            this.txtKids.Size = new System.Drawing.Size(58, 29);
            this.txtKids.TabIndex = 4;
            this.txtKids.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtKids.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValue_KeyPress);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 21);
            this.label2.TabIndex = 3;
            this.label2.Text = "Детей:";
            // 
            // txtAdult
            // 
            this.txtAdult.Location = new System.Drawing.Point(79, 22);
            this.txtAdult.Name = "txtAdult";
            this.txtAdult.Size = new System.Drawing.Size(58, 29);
            this.txtAdult.TabIndex = 2;
            this.txtAdult.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtAdult.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValue_KeyPress);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(74, 21);
            this.label1.TabIndex = 1;
            this.label1.Text = "Взролых:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.coBoxTo);
            this.groupBox2.Controls.Add(this.coBoxFrom);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Location = new System.Drawing.Point(169, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(204, 103);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Назначение";
            // 
            // coBoxTo
            // 
            this.coBoxTo.FormattingEnabled = true;
            this.coBoxTo.Items.AddRange(new object[] {
            "Москва",
            "Адлер",
            "Томск",
            "Астрахань"});
            this.coBoxTo.Location = new System.Drawing.Point(71, 62);
            this.coBoxTo.Name = "coBoxTo";
            this.coBoxTo.Size = new System.Drawing.Size(121, 29);
            this.coBoxTo.TabIndex = 7;
            this.coBoxTo.TextChanged += new System.EventHandler(this.coBoxTo_TextChanged);
            // 
            // coBoxFrom
            // 
            this.coBoxFrom.FormattingEnabled = true;
            this.coBoxFrom.Items.AddRange(new object[] {
            "Пермь"});
            this.coBoxFrom.Location = new System.Drawing.Point(71, 22);
            this.coBoxFrom.Name = "coBoxFrom";
            this.coBoxFrom.Size = new System.Drawing.Size(121, 29);
            this.coBoxFrom.TabIndex = 6;
            this.coBoxFrom.Text = "Пермь";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(20, 65);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(45, 21);
            this.label3.TabIndex = 3;
            this.label3.Text = "Куда:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 25);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(62, 21);
            this.label4.TabIndex = 1;
            this.label4.Text = "Откуда:";
            // 
            // btnNext
            // 
            this.btnNext.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnNext.Location = new System.Drawing.Point(12, 121);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(361, 28);
            this.btnNext.TabIndex = 6;
            this.btnNext.Text = "Далее";
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // GetInfoForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(383, 157);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Segoe UI Semilight", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "GetInfoForm";
            this.Text = "Aviasale";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtKids;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtAdult;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ComboBox coBoxTo;
        private System.Windows.Forms.ComboBox coBoxFrom;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnNext;
    }
}

