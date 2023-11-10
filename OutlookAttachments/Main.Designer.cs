namespace OutlookAttachments
{
    partial class Main
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
            groupBox1 = new GroupBox();
            label2 = new Label();
            label1 = new Label();
            dtpEndDate = new DateTimePicker();
            dtpStartDate = new DateTimePicker();
            btnSave = new Button();
            setting = new Button();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(label2);
            groupBox1.Controls.Add(label1);
            groupBox1.Controls.Add(dtpEndDate);
            groupBox1.Controls.Add(dtpStartDate);
            groupBox1.Font = new Font("Times New Roman", 12F, FontStyle.Regular, GraphicsUnit.Point);
            groupBox1.Location = new Point(12, 27);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(362, 105);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "Период выборки писем";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(6, 66);
            label2.Name = "label2";
            label2.Size = new Size(116, 19);
            label2.TabIndex = 4;
            label2.Text = "Дата окончания";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(6, 28);
            label1.Name = "label1";
            label1.Size = new Size(90, 19);
            label1.TabIndex = 3;
            label1.Text = "Дата начала";
            // 
            // dtpEndDate
            // 
            dtpEndDate.Location = new Point(127, 60);
            dtpEndDate.Name = "dtpEndDate";
            dtpEndDate.Size = new Size(200, 26);
            dtpEndDate.TabIndex = 1;
            // 
            // dtpStartDate
            // 
            dtpStartDate.Location = new Point(127, 22);
            dtpStartDate.Name = "dtpStartDate";
            dtpStartDate.Size = new Size(200, 26);
            dtpStartDate.TabIndex = 0;
            // 
            // btnSave
            // 
            btnSave.Font = new Font("Times New Roman", 12F, FontStyle.Regular, GraphicsUnit.Point);
            btnSave.Location = new Point(116, 138);
            btnSave.Name = "btnSave";
            btnSave.Size = new Size(140, 42);
            btnSave.TabIndex = 2;
            btnSave.Text = "Выполнить";
            btnSave.UseVisualStyleBackColor = true;
            btnSave.Click += btnSave_Click;
            // 
            // setting
            // 
            setting.BackColor = Color.Transparent;
            setting.FlatStyle = FlatStyle.Flat;
            setting.Image = Properties.Resources.settings__1_;
            setting.Location = new Point(349, 3);
            setting.Name = "setting";
            setting.Size = new Size(25, 27);
            setting.TabIndex = 3;
            setting.UseVisualStyleBackColor = false;
            setting.Click += setting_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(388, 196);
            Controls.Add(setting);
            Controls.Add(groupBox1);
            Controls.Add(btnSave);
            Name = "Form1";
            Text = "Сохранение вложений из писем";
            Load += Form1_Load;
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private GroupBox groupBox1;
        private DateTimePicker dtpEndDate;
        private Button btnSave;
        private DateTimePicker dtpStartDate;
        private Label label2;
        private Label label1;
        private Button setting;
    }
}