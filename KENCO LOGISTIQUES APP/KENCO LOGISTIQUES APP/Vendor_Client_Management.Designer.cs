namespace KENCO_LOGISTIQUES_APP
{
    partial class Vendor_Client_Management
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.PhoneNumberBox = new System.Windows.Forms.TextBox();
            this.PDBox = new System.Windows.Forms.TextBox();
            this.FirstNameBox = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.AddressBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SCManagement = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.SaveButton = new System.Windows.Forms.Button();
            this.Reset = new System.Windows.Forms.Button();
            this.Delete = new System.Windows.Forms.Button();
            this.AddNew = new System.Windows.Forms.Button();
            this.MainMenu = new System.Windows.Forms.Button();
            this.Search = new System.Windows.Forms.Button();
            this.SearchBox = new System.Windows.Forms.ComboBox();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column13 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column10 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column13,
            this.Column10,
            this.Column2,
            this.Column3});
            this.dataGridView1.Location = new System.Drawing.Point(56, 340);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 29;
            this.dataGridView1.Size = new System.Drawing.Size(803, 195);
            this.dataGridView1.TabIndex = 270;
            // 
            // PhoneNumberBox
            // 
            this.PhoneNumberBox.Location = new System.Drawing.Point(632, 147);
            this.PhoneNumberBox.Name = "PhoneNumberBox";
            this.PhoneNumberBox.Size = new System.Drawing.Size(177, 27);
            this.PhoneNumberBox.TabIndex = 278;
            // 
            // PDBox
            // 
            this.PDBox.Location = new System.Drawing.Point(632, 67);
            this.PDBox.Name = "PDBox";
            this.PDBox.Size = new System.Drawing.Size(177, 27);
            this.PDBox.TabIndex = 277;
            // 
            // FirstNameBox
            // 
            this.FirstNameBox.Location = new System.Drawing.Point(189, 67);
            this.FirstNameBox.Name = "FirstNameBox";
            this.FirstNameBox.Size = new System.Drawing.Size(177, 27);
            this.FirstNameBox.TabIndex = 275;
            this.FirstNameBox.TextChanged += new System.EventHandler(this.FirstNameBox_TextChanged);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label10.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.label10.Location = new System.Drawing.Point(523, 146);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(67, 40);
            this.label10.TabIndex = 274;
            this.label10.Text = "Phone:\r\nNumber";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label6.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.label6.Location = new System.Drawing.Point(523, 67);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(89, 40);
            this.label6.TabIndex = 273;
            this.label6.Text = "Product:\r\nDescription";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.label3.Location = new System.Drawing.Point(58, 146);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(125, 20);
            this.label3.TabIndex = 272;
            this.label3.Text = "Supplier / Client:";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.label2.Location = new System.Drawing.Point(56, 70);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(127, 20);
            this.label2.TabIndex = 271;
            this.label2.Text = "Customer Name:";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // AddressBox
            // 
            this.AddressBox.Location = new System.Drawing.Point(189, 233);
            this.AddressBox.Name = "AddressBox";
            this.AddressBox.Size = new System.Drawing.Size(177, 27);
            this.AddressBox.TabIndex = 280;
            this.AddressBox.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.label1.Location = new System.Drawing.Point(56, 236);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 20);
            this.label1.TabIndex = 279;
            this.label1.Text = "Address:";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // SCManagement
            // 
            this.SCManagement.FormattingEnabled = true;
            this.SCManagement.Items.AddRange(new object[] {
            "Supplier",
            "Client"});
            this.SCManagement.Location = new System.Drawing.Point(189, 146);
            this.SCManagement.Name = "SCManagement";
            this.SCManagement.Size = new System.Drawing.Size(177, 28);
            this.SCManagement.TabIndex = 281;
            this.SCManagement.SelectedIndexChanged += new System.EventHandler(this.IncomeDescription_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Showcard Gothic", 10.2F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point);
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.label4.Location = new System.Drawing.Point(12, 23);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(300, 21);
            this.label4.TabIndex = 282;
            this.label4.Text = "SUPPLIER / CLIENT MANAGEMENT";
            // 
            // SaveButton
            // 
            this.SaveButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.SaveButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SaveButton.Font = new System.Drawing.Font("Showcard Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.SaveButton.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.SaveButton.Location = new System.Drawing.Point(658, 553);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new System.Drawing.Size(94, 29);
            this.SaveButton.TabIndex = 287;
            this.SaveButton.Text = "Save";
            this.SaveButton.UseVisualStyleBackColor = false;
            this.SaveButton.Click += new System.EventHandler(this.SaveButton_Click);
            // 
            // Reset
            // 
            this.Reset.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.Reset.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Reset.Font = new System.Drawing.Font("Showcard Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.Reset.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.Reset.Location = new System.Drawing.Point(527, 553);
            this.Reset.Name = "Reset";
            this.Reset.Size = new System.Drawing.Size(94, 29);
            this.Reset.TabIndex = 286;
            this.Reset.Text = "Reset";
            this.Reset.UseVisualStyleBackColor = false;
            this.Reset.Click += new System.EventHandler(this.Reset_Click);
            // 
            // Delete
            // 
            this.Delete.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.Delete.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Delete.Font = new System.Drawing.Font("Showcard Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.Delete.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.Delete.Location = new System.Drawing.Point(399, 553);
            this.Delete.Name = "Delete";
            this.Delete.Size = new System.Drawing.Size(94, 29);
            this.Delete.TabIndex = 285;
            this.Delete.Text = "Delete";
            this.Delete.UseVisualStyleBackColor = false;
            this.Delete.Click += new System.EventHandler(this.Delete_Click_1);
            // 
            // AddNew
            // 
            this.AddNew.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.AddNew.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.AddNew.Font = new System.Drawing.Font("Showcard Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.AddNew.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.AddNew.Location = new System.Drawing.Point(268, 553);
            this.AddNew.Name = "AddNew";
            this.AddNew.Size = new System.Drawing.Size(94, 29);
            this.AddNew.TabIndex = 284;
            this.AddNew.Text = "Add New";
            this.AddNew.UseVisualStyleBackColor = false;
            this.AddNew.Click += new System.EventHandler(this.AddNew_Click_1);
            // 
            // MainMenu
            // 
            this.MainMenu.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.MainMenu.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.MainMenu.Font = new System.Drawing.Font("Showcard Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.MainMenu.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.MainMenu.Location = new System.Drawing.Point(778, 553);
            this.MainMenu.Name = "MainMenu";
            this.MainMenu.Size = new System.Drawing.Size(118, 29);
            this.MainMenu.TabIndex = 283;
            this.MainMenu.Text = "Main Menu";
            this.MainMenu.UseVisualStyleBackColor = false;
            this.MainMenu.Click += new System.EventHandler(this.MainMenu_Click);
            // 
            // Search
            // 
            this.Search.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.Search.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Search.Font = new System.Drawing.Font("Showcard Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.Search.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.Search.Location = new System.Drawing.Point(758, 286);
            this.Search.Name = "Search";
            this.Search.Size = new System.Drawing.Size(138, 29);
            this.Search.TabIndex = 289;
            this.Search.Text = "Search";
            this.Search.UseVisualStyleBackColor = false;
            this.Search.Click += new System.EventHandler(this.Search_Click);
            // 
            // SearchBox
            // 
            this.SearchBox.FormattingEnabled = true;
            this.SearchBox.Items.AddRange(new object[] {
            "Supplier",
            "Client"});
            this.SearchBox.Location = new System.Drawing.Point(557, 287);
            this.SearchBox.Name = "SearchBox";
            this.SearchBox.Size = new System.Drawing.Size(177, 28);
            this.SearchBox.TabIndex = 290;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Customer Name";
            this.Column1.MinimumWidth = 6;
            this.Column1.Name = "Column1";
            this.Column1.ReadOnly = true;
            this.Column1.Width = 150;
            // 
            // Column13
            // 
            this.Column13.HeaderText = "Supplier / Client";
            this.Column13.MinimumWidth = 6;
            this.Column13.Name = "Column13";
            this.Column13.ReadOnly = true;
            this.Column13.Width = 150;
            // 
            // Column10
            // 
            this.Column10.HeaderText = "Product Description";
            this.Column10.MinimumWidth = 6;
            this.Column10.Name = "Column10";
            this.Column10.ReadOnly = true;
            this.Column10.Width = 150;
            // 
            // Column2
            // 
            this.Column2.HeaderText = "Phone Number";
            this.Column2.MinimumWidth = 6;
            this.Column2.Name = "Column2";
            this.Column2.ReadOnly = true;
            this.Column2.Width = 150;
            // 
            // Column3
            // 
            this.Column3.HeaderText = "Address";
            this.Column3.MinimumWidth = 6;
            this.Column3.Name = "Column3";
            this.Column3.ReadOnly = true;
            this.Column3.Width = 150;
            // 
            // Vendor_Client_Management
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(908, 594);
            this.Controls.Add(this.SearchBox);
            this.Controls.Add(this.Search);
            this.Controls.Add(this.SaveButton);
            this.Controls.Add(this.Reset);
            this.Controls.Add(this.Delete);
            this.Controls.Add(this.AddNew);
            this.Controls.Add(this.MainMenu);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.SCManagement);
            this.Controls.Add(this.AddressBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.PhoneNumberBox);
            this.Controls.Add(this.PDBox);
            this.Controls.Add(this.FirstNameBox);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Vendor_Client_Management";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Vendor_Client_Management";
            this.Load += new System.EventHandler(this.Vendor_Client_Management_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DataGridView dataGridView1;
        private TextBox PhoneNumberBox;
        private TextBox PDBox;
        private TextBox FirstNameBox;
        private Label label10;
        private Label label6;
        private Label label3;
        private Label label2;
        private TextBox AddressBox;
        private Label label1;
        private ComboBox SCManagement;
        private string SCMString;
        private Label label4;
        private Button SaveButton;
        private Button Reset;
        private Button Delete;
        private Button AddNew;
        private Button MainMenu;
        string newFilePath = Path.Combine(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName,
"bin\\Debug\\net6.0-windows\\Resources\\Files\\KENCO_DB.xlsx");
        private Button Search;
        private ComboBox SearchBox;
        private DataGridViewTextBoxColumn Column1;
        private DataGridViewTextBoxColumn Column13;
        private DataGridViewTextBoxColumn Column10;
        private DataGridViewTextBoxColumn Column2;
        private DataGridViewTextBoxColumn Column3;
    }
}