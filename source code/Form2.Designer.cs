namespace ExcelDocConverter
{
    partial class Form2
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2));
            this.btSelectFiles = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.cbFileType = new System.Windows.Forms.ComboBox();
            this.btOutputFolder = new System.Windows.Forms.Button();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.btStartConvert = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.colnSelect = new System.Windows.Forms.DataGridViewLinkColumn();
            this.colnFiles = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colnStatus = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btRemoveAll = new System.Windows.Forms.Button();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.lbError = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // btSelectFiles
            // 
            this.btSelectFiles.Location = new System.Drawing.Point(12, 90);
            this.btSelectFiles.Name = "btSelectFiles";
            this.btSelectFiles.Size = new System.Drawing.Size(124, 23);
            this.btSelectFiles.TabIndex = 0;
            this.btSelectFiles.Text = "Select Files";
            this.btSelectFiles.UseVisualStyleBackColor = true;
            this.btSelectFiles.Click += new System.EventHandler(this.btSelectFiles_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(140, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Convert Selected File(s) To";
            // 
            // cbFileType
            // 
            this.cbFileType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbFileType.FormattingEnabled = true;
            this.cbFileType.Items.AddRange(new object[] {
            "xls - Microsoft Excel 2003",
            "xlsx - Microsoft Excel 2007 & Above",
            "pdf - Portable Document Format",
            "xps - XML Paper Specification",
            "csv - Comma Separated Value"});
            this.cbFileType.Location = new System.Drawing.Point(158, 6);
            this.cbFileType.Name = "cbFileType";
            this.cbFileType.Size = new System.Drawing.Size(279, 21);
            this.cbFileType.TabIndex = 2;
            // 
            // btOutputFolder
            // 
            this.btOutputFolder.Location = new System.Drawing.Point(12, 33);
            this.btOutputFolder.Name = "btOutputFolder";
            this.btOutputFolder.Size = new System.Drawing.Size(262, 23);
            this.btOutputFolder.TabIndex = 3;
            this.btOutputFolder.Text = "Select a folder to save the converted files";
            this.btOutputFolder.UseVisualStyleBackColor = true;
            this.btOutputFolder.Click += new System.EventHandler(this.btOutputFolder_Click);
            // 
            // txtOutput
            // 
            this.txtOutput.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtOutput.Location = new System.Drawing.Point(12, 62);
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.Size = new System.Drawing.Size(710, 22);
            this.txtOutput.TabIndex = 4;
            // 
            // btStartConvert
            // 
            this.btStartConvert.Location = new System.Drawing.Point(272, 90);
            this.btStartConvert.Name = "btStartConvert";
            this.btStartConvert.Size = new System.Drawing.Size(124, 23);
            this.btStartConvert.TabIndex = 5;
            this.btStartConvert.Text = "Start Convert";
            this.btStartConvert.UseVisualStyleBackColor = true;
            this.btStartConvert.Click += new System.EventHandler(this.btStartConvert_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.BackgroundColor = System.Drawing.Color.White;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colnSelect,
            this.colnFiles,
            this.colnStatus});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.Location = new System.Drawing.Point(12, 119);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.Size = new System.Drawing.Size(710, 368);
            this.dataGridView1.TabIndex = 6;
            // 
            // colnSelect
            // 
            this.colnSelect.HeaderText = "";
            this.colnSelect.Name = "colnSelect";
            this.colnSelect.ReadOnly = true;
            this.colnSelect.Width = 80;
            // 
            // colnFiles
            // 
            this.colnFiles.HeaderText = "File(s)";
            this.colnFiles.Name = "colnFiles";
            this.colnFiles.ReadOnly = true;
            this.colnFiles.Width = 500;
            // 
            // colnStatus
            // 
            this.colnStatus.HeaderText = "Status";
            this.colnStatus.Name = "colnStatus";
            this.colnStatus.ReadOnly = true;
            // 
            // btRemoveAll
            // 
            this.btRemoveAll.Location = new System.Drawing.Point(142, 90);
            this.btRemoveAll.Name = "btRemoveAll";
            this.btRemoveAll.Size = new System.Drawing.Size(124, 23);
            this.btRemoveAll.TabIndex = 7;
            this.btRemoveAll.Text = "Remove All";
            this.btRemoveAll.UseVisualStyleBackColor = true;
            this.btRemoveAll.Click += new System.EventHandler(this.btRemoveAll_Click);
            // 
            // linkLabel1
            // 
            this.linkLabel1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(664, 492);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(58, 13);
            this.linkLabel1.TabIndex = 9;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "More Info";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // lbError
            // 
            this.lbError.AutoSize = true;
            this.lbError.Location = new System.Drawing.Point(590, 89);
            this.lbError.Name = "lbError";
            this.lbError.Size = new System.Drawing.Size(122, 26);
            this.lbError.TabIndex = 10;
            this.lbError.Text = "Double click \"Error\"\r\nto view error message.";
            this.lbError.Visible = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::ExcelDocConverter.Properties.Resources.p;
            this.pictureBox1.InitialImage = null;
            this.pictureBox1.Location = new System.Drawing.Point(401, 92);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(220, 20);
            this.pictureBox1.TabIndex = 11;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Visible = false;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(734, 512);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.lbError);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.btRemoveAll);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btStartConvert);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.btOutputFolder);
            this.Controls.Add(this.cbFileType);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btSelectFiles);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(650, 350);
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Simple MS Excel Document Converter 2.1";
            this.Load += new System.EventHandler(this.Form2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btSelectFiles;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cbFileType;
        private System.Windows.Forms.Button btOutputFolder;
        private System.Windows.Forms.TextBox txtOutput;
        private System.Windows.Forms.Button btStartConvert;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewLinkColumn colnSelect;
        private System.Windows.Forms.DataGridViewTextBoxColumn colnFiles;
        private System.Windows.Forms.DataGridViewTextBoxColumn colnStatus;
        private System.Windows.Forms.Button btRemoveAll;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.Label lbError;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}