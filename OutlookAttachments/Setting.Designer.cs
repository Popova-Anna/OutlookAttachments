namespace OutlookAttachments
{
    partial class Setting
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            tBPath = new TextBox();
            btnSetPath = new Button();
            groupBox1 = new GroupBox();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // tBPath
            // 
            tBPath.Enabled = false;
            tBPath.Location = new Point(7, 26);
            tBPath.Margin = new Padding(4);
            tBPath.Multiline = true;
            tBPath.Name = "tBPath";
            tBPath.Size = new Size(226, 58);
            tBPath.TabIndex = 1;
            // 
            // btnSetPath
            // 
            btnSetPath.Location = new Point(7, 92);
            btnSetPath.Margin = new Padding(4);
            btnSetPath.Name = "btnSetPath";
            btnSetPath.Size = new Size(226, 44);
            btnSetPath.TabIndex = 2;
            btnSetPath.Text = "Изменить папку сохранения";
            btnSetPath.UseVisualStyleBackColor = true;
            btnSetPath.Click += btnSetPath_Click;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(tBPath);
            groupBox1.Controls.Add(btnSetPath);
            groupBox1.Location = new Point(12, 12);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(242, 147);
            groupBox1.TabIndex = 3;
            groupBox1.TabStop = false;
            groupBox1.Text = "Путь для сохранения вложений:";
            // 
            // Setting
            // 
            AutoScaleDimensions = new SizeF(9F, 19F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(265, 173);
            Controls.Add(groupBox1);
            Font = new Font("Times New Roman", 12F, FontStyle.Regular, GraphicsUnit.Point);
            Margin = new Padding(4);
            Name = "Setting";
            Text = "Setting";
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
        }

        #endregion
        private TextBox tBPath;
        private Button btnSetPath;
        private GroupBox groupBox1;
    }
}