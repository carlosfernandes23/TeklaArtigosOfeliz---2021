namespace TeklaArtigosOfeliz
{
    partial class Frm_Nest
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
            this.components = new System.ComponentModel.Container();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource1 = new Microsoft.Reporting.WinForms.ReportDataSource();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource2 = new Microsoft.Reporting.WinForms.ReportDataSource();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource3 = new Microsoft.Reporting.WinForms.ReportDataSource();
            this.dataTable1BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.DataSet1 = new TeklaArtigosOfeliz.DataSet1();
            this.dataTable2BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.dataTable3BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.REFERECIA = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PERFIL = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label2 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.cARREGARPEÇASTEKLAToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.sTOCKToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aTRIBUIRSTOKAUTOMATICOToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.nESTINGToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cALCULARToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            ((System.ComponentModel.ISupportInitialize)(this.dataTable1BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DataSet1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataTable2BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataTable3BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataTable1BindingSource
            // 
            this.dataTable1BindingSource.DataMember = "DataTable1";
            this.dataTable1BindingSource.DataSource = this.DataSet1;
            // 
            // DataSet1
            // 
            this.DataSet1.DataSetName = "DataSet1";
            this.DataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // dataTable2BindingSource
            // 
            this.dataTable2BindingSource.DataMember = "DataTable2";
            this.dataTable2BindingSource.DataSource = this.DataSet1;
            // 
            // dataTable3BindingSource
            // 
            this.dataTable3BindingSource.DataMember = "DataTable3";
            this.dataTable3BindingSource.DataSource = this.DataSet1;
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.REFERECIA,
            this.PERFIL});
            this.dataGridView1.Location = new System.Drawing.Point(2, 40);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(447, 731);
            this.dataGridView1.TabIndex = 0;
            // 
            // Column1
            // 
            this.Column1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Column1.HeaderText = "QUANT.";
            this.Column1.Name = "Column1";
            this.Column1.Width = 73;
            // 
            // Column2
            // 
            this.Column2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Column2.HeaderText = "COMP.";
            this.Column2.Name = "Column2";
            this.Column2.Width = 66;
            // 
            // REFERECIA
            // 
            this.REFERECIA.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.REFERECIA.HeaderText = "REFERECIA";
            this.REFERECIA.Name = "REFERECIA";
            // 
            // PERFIL
            // 
            this.PERFIL.HeaderText = "PERFIL";
            this.PERFIL.Name = "PERFIL";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(42, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "PEÇAS";
            // 
            // dataGridView2
            // 
            this.dataGridView2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column3,
            this.Column4,
            this.Column5,
            this.Column6});
            this.dataGridView2.Location = new System.Drawing.Point(455, 40);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(368, 731);
            this.dataGridView2.TabIndex = 2;
            // 
            // Column3
            // 
            this.Column3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Column3.HeaderText = "QUANT.";
            this.Column3.Name = "Column3";
            this.Column3.Width = 73;
            // 
            // Column4
            // 
            this.Column4.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Column4.HeaderText = "COMP.";
            this.Column4.Name = "Column4";
            this.Column4.Width = 66;
            // 
            // Column5
            // 
            this.Column5.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Column5.HeaderText = "PRIORIDADE";
            this.Column5.Name = "Column5";
            this.Column5.Width = 99;
            // 
            // Column6
            // 
            this.Column6.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Column6.HeaderText = "PERFIL";
            this.Column6.Name = "Column6";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(452, 24);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "STOCK";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1,
            this.sTOCKToolStripMenuItem,
            this.nESTINGToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1482, 24);
            this.menuStrip1.TabIndex = 4;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cARREGARPEÇASTEKLAToolStripMenuItem});
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(70, 20);
            this.toolStripMenuItem1.Text = "FICHEIRO";
            // 
            // cARREGARPEÇASTEKLAToolStripMenuItem
            // 
            this.cARREGARPEÇASTEKLAToolStripMenuItem.Name = "cARREGARPEÇASTEKLAToolStripMenuItem";
            this.cARREGARPEÇASTEKLAToolStripMenuItem.Size = new System.Drawing.Size(207, 22);
            this.cARREGARPEÇASTEKLAToolStripMenuItem.Text = "CARREGAR PEÇAS TEKLA";
            this.cARREGARPEÇASTEKLAToolStripMenuItem.Click += new System.EventHandler(this.cARREGARPEÇASTEKLAToolStripMenuItem_Click);
            // 
            // sTOCKToolStripMenuItem
            // 
            this.sTOCKToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aTRIBUIRSTOKAUTOMATICOToolStripMenuItem});
            this.sTOCKToolStripMenuItem.Name = "sTOCKToolStripMenuItem";
            this.sTOCKToolStripMenuItem.Size = new System.Drawing.Size(54, 20);
            this.sTOCKToolStripMenuItem.Text = "STOCK";
            // 
            // aTRIBUIRSTOKAUTOMATICOToolStripMenuItem
            // 
            this.aTRIBUIRSTOKAUTOMATICOToolStripMenuItem.Name = "aTRIBUIRSTOKAUTOMATICOToolStripMenuItem";
            this.aTRIBUIRSTOKAUTOMATICOToolStripMenuItem.Size = new System.Drawing.Size(275, 22);
            this.aTRIBUIRSTOKAUTOMATICOToolStripMenuItem.Text = "ATRIBUIR STOCK AUTOMATICO PERFIS";
            this.aTRIBUIRSTOKAUTOMATICOToolStripMenuItem.Click += new System.EventHandler(this.aTRIBUIRSTOKAUTOMATICOToolStripMenuItem_Click);
            // 
            // nESTINGToolStripMenuItem
            // 
            this.nESTINGToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cALCULARToolStripMenuItem});
            this.nESTINGToolStripMenuItem.Name = "nESTINGToolStripMenuItem";
            this.nESTINGToolStripMenuItem.Size = new System.Drawing.Size(66, 20);
            this.nESTINGToolStripMenuItem.Text = "NESTING";
            // 
            // cALCULARToolStripMenuItem
            // 
            this.cALCULARToolStripMenuItem.Name = "cALCULARToolStripMenuItem";
            this.cALCULARToolStripMenuItem.Size = new System.Drawing.Size(133, 22);
            this.cALCULARToolStripMenuItem.Text = "CALCULAR";
            this.cALCULARToolStripMenuItem.Click += new System.EventHandler(this.cALCULARToolStripMenuItem_Click);
            // 
            // reportViewer1
            // 
            this.reportViewer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            reportDataSource1.Name = "DataSet1";
            reportDataSource1.Value = this.dataTable1BindingSource;
            reportDataSource2.Name = "DataSet2";
            reportDataSource2.Value = this.dataTable2BindingSource;
            reportDataSource3.Name = "DataSet3";
            reportDataSource3.Value = this.dataTable3BindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource2);
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource3);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "TeklaArtigosOfeliz.Report1.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(829, 40);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.ServerReport.BearerToken = null;
            this.reportViewer1.Size = new System.Drawing.Size(653, 731);
            this.reportViewer1.TabIndex = 5;
            // 
            // FrmNest
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1482, 770);
            this.Controls.Add(this.reportViewer1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "FrmNest";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FrmNest";
            this.Load += new System.EventHandler(this.FrmNest_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataTable1BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DataSet1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataTable2BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataTable3BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn REFERECIA;
        private System.Windows.Forms.DataGridViewTextBoxColumn PERFIL;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column6;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem cARREGARPEÇASTEKLAToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem nESTINGToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cALCULARToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem sTOCKToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aTRIBUIRSTOKAUTOMATICOToolStripMenuItem;
        private DataSet1 DataSet1;
        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource dataTable1BindingSource;
        private System.Windows.Forms.BindingSource dataTable2BindingSource;
        private System.Windows.Forms.BindingSource dataTable3BindingSource;
    }
}