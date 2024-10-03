namespace ExcelToXA
{
    partial class Configuration
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Configuration));
            vendorBox = new ComboBox();
            label1 = new Label();
            label2 = new Label();
            environmentBox = new ComboBox();
            okBtn = new Button();
            hlpBtn = new Button();
            checkBox1 = new CheckBox();
            label3 = new Label();
            pass = new TextBox();
            label4 = new Label();
            user = new TextBox();
            SuspendLayout();
            // 
            // vendorBox
            // 
            vendorBox.FormattingEnabled = true;
            vendorBox.Location = new Point(93, 9);
            vendorBox.Name = "vendorBox";
            vendorBox.Size = new Size(154, 23);
            vendorBox.TabIndex = 0;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(43, 12);
            label1.Name = "label1";
            label1.Size = new Size(44, 15);
            label1.TabIndex = 1;
            label1.Text = "Vendor";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(12, 99);
            label2.Name = "label2";
            label2.Size = new Size(75, 15);
            label2.TabIndex = 3;
            label2.Text = "Environment";
            // 
            // environmentBox
            // 
            environmentBox.FormattingEnabled = true;
            environmentBox.Location = new Point(93, 96);
            environmentBox.Name = "environmentBox";
            environmentBox.Size = new Size(154, 23);
            environmentBox.TabIndex = 2;
            // 
            // okBtn
            // 
            okBtn.Location = new Point(12, 163);
            okBtn.Name = "okBtn";
            okBtn.Size = new Size(75, 23);
            okBtn.TabIndex = 4;
            okBtn.Text = "Confirm";
            okBtn.UseVisualStyleBackColor = true;
            okBtn.Click += okBtn_Click;
            // 
            // hlpBtn
            // 
            hlpBtn.Location = new Point(172, 163);
            hlpBtn.Name = "hlpBtn";
            hlpBtn.Size = new Size(75, 23);
            hlpBtn.TabIndex = 5;
            hlpBtn.Text = "Help";
            hlpBtn.UseVisualStyleBackColor = true;
            hlpBtn.Click += button1_Click;
            // 
            // checkBox1
            // 
            checkBox1.AutoSize = true;
            checkBox1.Location = new Point(93, 125);
            checkBox1.Name = "checkBox1";
            checkBox1.Size = new Size(129, 19);
            checkBox1.TabIndex = 12;
            checkBox1.Text = "Remember Settings";
            checkBox1.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(30, 70);
            label3.Name = "label3";
            label3.Size = new Size(57, 15);
            label3.TabIndex = 10;
            label3.Text = "Password";
            // 
            // pass
            // 
            pass.Location = new Point(93, 67);
            pass.Name = "pass";
            pass.PasswordChar = '*';
            pass.Size = new Size(154, 23);
            pass.TabIndex = 9;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(43, 41);
            label4.Name = "label4";
            label4.Size = new Size(44, 15);
            label4.TabIndex = 8;
            label4.Text = "User ID";
            // 
            // user
            // 
            user.Location = new Point(93, 38);
            user.Name = "user";
            user.Size = new Size(154, 23);
            user.TabIndex = 7;
            // 
            // Configuration
            // 
            AcceptButton = okBtn;
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(259, 194);
            Controls.Add(checkBox1);
            Controls.Add(label3);
            Controls.Add(pass);
            Controls.Add(label4);
            Controls.Add(user);
            Controls.Add(hlpBtn);
            Controls.Add(okBtn);
            Controls.Add(label2);
            Controls.Add(environmentBox);
            Controls.Add(label1);
            Controls.Add(vendorBox);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Configuration";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Configuration";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private ComboBox vendorBox;
        private Label label1;
        private Label label2;
        private ComboBox environmentBox;
        private Button okBtn;
        private Button hlpBtn;
        private CheckBox checkBox1;
        private Label label3;
        private TextBox pass;
        private Label label4;
        private TextBox user;
    }
}