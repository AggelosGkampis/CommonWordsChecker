namespace DesktopAppCommonWords
{
    partial class Form1
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
            this.label1 = new System.Windows.Forms.Label();
            this.btnchoose = new System.Windows.Forms.Button();
            this.txtpath = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btnExport = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label1.Location = new System.Drawing.Point(75, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "File Path";
            // 
            // btnchoose
            // 
            this.btnchoose.FlatAppearance.BorderColor = System.Drawing.Color.Blue;
            this.btnchoose.FlatAppearance.BorderSize = 2;
            this.btnchoose.FlatAppearance.MouseDownBackColor = System.Drawing.Color.RosyBrown;
            this.btnchoose.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Lime;
            this.btnchoose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnchoose.Font = new System.Drawing.Font("Gill Sans MT", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.btnchoose.Location = new System.Drawing.Point(247, 98);
            this.btnchoose.Name = "btnchoose";
            this.btnchoose.Size = new System.Drawing.Size(121, 31);
            this.btnchoose.TabIndex = 1;
            this.btnchoose.Text = "Choose File";
            this.btnchoose.UseVisualStyleBackColor = true;
            this.btnchoose.Click += new System.EventHandler(this.btnchoose_Click);
            // 
            // txtpath
            // 
            this.txtpath.Location = new System.Drawing.Point(145, 49);
            this.txtpath.Name = "txtpath";
            this.txtpath.Size = new System.Drawing.Size(571, 23);
            this.txtpath.TabIndex = 2;
            this.txtpath.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.dataGridView1);
            this.panel1.Location = new System.Drawing.Point(1, 157);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(799, 294);
            this.panel1.TabIndex = 3;
            // 
            // dataGridView1
            // 
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.MenuText;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 25;
            this.dataGridView1.Size = new System.Drawing.Size(799, 294);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // btnExport
            // 
            this.btnExport.FlatAppearance.BorderColor = System.Drawing.Color.Blue;
            this.btnExport.FlatAppearance.BorderSize = 2;
            this.btnExport.FlatAppearance.MouseDownBackColor = System.Drawing.Color.RosyBrown;
            this.btnExport.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Lime;
            this.btnExport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExport.Font = new System.Drawing.Font("Gill Sans MT", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.btnExport.Location = new System.Drawing.Point(455, 98);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(121, 31);
            this.btnExport.TabIndex = 5;
            this.btnExport.Text = "Export File";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.txtpath);
            this.Controls.Add(this.btnchoose);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "CommonWordsChecker";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Label label1;
        private Button btnchoose;
        private TextBox txtpath;
        private Panel panel1;
        private DataGridView dataGridView1;
        private Button btnExport;
    }
}