namespace ExcelToXA
{
    partial class Main
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
            components = new System.ComponentModel.Container();
            ListViewItem listViewItem1 = new ListViewItem(new string[] { "Test", "Test" }, -1);
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            textBox1 = new TextBox();
            label1 = new Label();
            browseBtn = new Button();
            excelPathDiag = new OpenFileDialog();
            label2 = new Label();
            excelDataPreview = new ListView();
            poPreview = new ListView();
            label3 = new Label();
            uploadBtn = new Button();
            delItem = new Button();
            progressBar1 = new ProgressBar();
            button2 = new Button();
            refreshBtn = new Button();
            refreshTT = new ToolTip(components);
            SuspendLayout();
            // 
            // textBox1
            // 
            textBox1.Location = new Point(12, 27);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(159, 23);
            textBox1.TabIndex = 0;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 9);
            label1.Name = "label1";
            label1.Size = new Size(93, 15);
            label1.TabIndex = 1;
            label1.Text = "Excel Document";
            // 
            // browseBtn
            // 
            browseBtn.Location = new Point(177, 27);
            browseBtn.Name = "browseBtn";
            browseBtn.Size = new Size(75, 23);
            browseBtn.TabIndex = 2;
            browseBtn.Text = "Browse";
            browseBtn.UseVisualStyleBackColor = true;
            browseBtn.Click += browseBtn_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(12, 71);
            label2.Name = "label2";
            label2.Size = new Size(75, 15);
            label2.TabIndex = 4;
            label2.Text = "Data Preview";
            // 
            // excelDataPreview
            // 
            excelDataPreview.Items.AddRange(new ListViewItem[] { listViewItem1 });
            excelDataPreview.Location = new Point(12, 89);
            excelDataPreview.Name = "excelDataPreview";
            excelDataPreview.Size = new Size(240, 297);
            excelDataPreview.TabIndex = 5;
            excelDataPreview.UseCompatibleStateImageBehavior = false;
            excelDataPreview.View = View.Details;
            excelDataPreview.SelectedIndexChanged += excelDataPreview_SelectedIndexChanged;
            // 
            // poPreview
            // 
            poPreview.Location = new Point(288, 89);
            poPreview.Name = "poPreview";
            poPreview.Size = new Size(240, 268);
            poPreview.TabIndex = 7;
            poPreview.UseCompatibleStateImageBehavior = false;
            poPreview.View = View.Details;
            poPreview.SelectedIndexChanged += poPreview_SelectedIndexChanged;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(288, 71);
            label3.Name = "label3";
            label3.Size = new Size(137, 15);
            label3.TabIndex = 8;
            label3.Text = "Purchase Orders Preview";
            // 
            // uploadBtn
            // 
            uploadBtn.Location = new Point(371, 392);
            uploadBtn.Name = "uploadBtn";
            uploadBtn.Size = new Size(75, 23);
            uploadBtn.TabIndex = 9;
            uploadBtn.Text = "Upload";
            uploadBtn.UseVisualStyleBackColor = true;
            uploadBtn.Click += uploadBtn_Click;
            // 
            // delItem
            // 
            delItem.Location = new Point(95, 392);
            delItem.Name = "delItem";
            delItem.Size = new Size(75, 23);
            delItem.TabIndex = 10;
            delItem.Text = "Delete Item";
            delItem.UseVisualStyleBackColor = true;
            delItem.Click += delItem_Click;
            // 
            // progressBar1
            // 
            progressBar1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            progressBar1.BackColor = SystemColors.Control;
            progressBar1.Location = new Point(288, 363);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(240, 23);
            progressBar1.Style = ProgressBarStyle.Continuous;
            progressBar1.TabIndex = 11;
            // 
            // button2
            // 
            button2.Location = new Point(439, 27);
            button2.Name = "button2";
            button2.Size = new Size(89, 23);
            button2.TabIndex = 12;
            button2.Text = "Configuration";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // refreshBtn
            // 
            refreshBtn.Image = Properties.Resources.refresh__2_;
            refreshBtn.Location = new Point(503, 61);
            refreshBtn.Name = "refreshBtn";
            refreshBtn.Size = new Size(25, 25);
            refreshBtn.TabIndex = 13;
            refreshBtn.UseVisualStyleBackColor = true;
            refreshBtn.Click += refreshBtn_Click;
            // 
            // Main
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(540, 423);
            Controls.Add(refreshBtn);
            Controls.Add(button2);
            Controls.Add(progressBar1);
            Controls.Add(delItem);
            Controls.Add(uploadBtn);
            Controls.Add(label3);
            Controls.Add(poPreview);
            Controls.Add(excelDataPreview);
            Controls.Add(label2);
            Controls.Add(browseBtn);
            Controls.Add(label1);
            Controls.Add(textBox1);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Main";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Excel PO Loading";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private TextBox textBox1;
        private Label label1;
        private Button browseBtn;
        private OpenFileDialog excelPathDiag;
        private Label label2;
        private ListView excelDataPreview;
        private ListView poPreview;
        private Label label3;
        private Button uploadBtn;
        private Button delItem;
        private ProgressBar progressBar1;
        private Button button2;
        private Button refreshBtn;
        private ToolTip refreshTT;
    }
}