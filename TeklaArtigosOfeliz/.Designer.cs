using System.Windows.Forms;

namespace TeklaArtigosOfeliz
{
    partial class Frm_Inico
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
            System.Windows.Forms.ToolStripMenuItem alterarBaseDeDadosCPEDapToolStripMenuItem;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Frm_Inico));
            this.Chb_alw_top = new System.Windows.Forms.CheckBox();
            this.label5 = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.LBLestado = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.button54 = new System.Windows.Forms.Button();
            this.PASTAEXPORTACAO = new System.Windows.Forms.TextBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.parametrosToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.parametrosToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.quantificaçãoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.macrosToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.alteraFaseToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lotesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.desenhoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.criarFasesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ferramentasToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.soldaduraToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.desenhoToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.utilitariosToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.abrirObraToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.testeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.limparTodasAsUDADaPeçaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pDFSoldaduraToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.nESTDESENVOLVIMENTOToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.exportarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exportarNC1ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.EmailStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.enviarParaFabricoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.enviarEmailParaLotearToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.enviarEmialAprovToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.revisãoPeçasEConjuntosToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.button9 = new Guna.UI2.WinForms.Guna2Button();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            alterarBaseDeDadosCPEDapToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // alterarBaseDeDadosCPEDapToolStripMenuItem
            // 
            alterarBaseDeDadosCPEDapToolStripMenuItem.Name = "alterarBaseDeDadosCPEDapToolStripMenuItem";
            alterarBaseDeDadosCPEDapToolStripMenuItem.Size = new System.Drawing.Size(285, 22);
            alterarBaseDeDadosCPEDapToolStripMenuItem.Text = "Alterar Base de Dados CP e DAP";
            alterarBaseDeDadosCPEDapToolStripMenuItem.Click += new System.EventHandler(this.alterarBaseDeDadosCPEDapToolStripMenuItem_Click);
            // 
            // Chb_alw_top
            // 
            this.Chb_alw_top.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.Chb_alw_top.AutoSize = true;
            this.Chb_alw_top.Checked = true;
            this.Chb_alw_top.CheckState = System.Windows.Forms.CheckState.Checked;
            this.Chb_alw_top.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.Chb_alw_top.Location = new System.Drawing.Point(380, 90);
            this.Chb_alw_top.Name = "Chb_alw_top";
            this.Chb_alw_top.Size = new System.Drawing.Size(15, 14);
            this.Chb_alw_top.TabIndex = 0;
            this.Chb_alw_top.UseVisualStyleBackColor = true;
            this.Chb_alw_top.CheckedChanged += new System.EventHandler(this.Chb_alw_top_CheckedChanged);
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label5.Location = new System.Drawing.Point(10, 85);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(18, 17);
            this.label5.TabIndex = 12;
            this.label5.Text = "--";
            this.label5.MouseClick += new System.Windows.Forms.MouseEventHandler(this.label5_MouseClick);
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // LBLestado
            // 
            this.LBLestado.AutoSize = true;
            this.LBLestado.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
            this.LBLestado.Location = new System.Drawing.Point(10, 65);
            this.LBLestado.Name = "LBLestado";
            this.LBLestado.Size = new System.Drawing.Size(20, 17);
            this.LBLestado.TabIndex = 17;
            this.LBLestado.Text = "--";
            this.LBLestado.Visible = false;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(110, 40);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(89, 20);
            this.label11.TabIndex = 52;
            this.label11.Text = "Sem Obra";
            this.label11.Click += new System.EventHandler(this.label11_Click);
            // 
            // button54
            // 
            this.button54.Location = new System.Drawing.Point(0, 0);
            this.button54.Name = "button54";
            this.button54.Size = new System.Drawing.Size(75, 23);
            this.button54.TabIndex = 0;
            // 
            // PASTAEXPORTACAO
            // 
            this.PASTAEXPORTACAO.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.PASTAEXPORTACAO.Location = new System.Drawing.Point(340, 65);
            this.PASTAEXPORTACAO.Multiline = true;
            this.PASTAEXPORTACAO.Name = "PASTAEXPORTACAO";
            this.PASTAEXPORTACAO.Size = new System.Drawing.Size(53, 16);
            this.PASTAEXPORTACAO.TabIndex = 62;
            this.PASTAEXPORTACAO.Text = "C:\\R\\";
            this.PASTAEXPORTACAO.Visible = false;
            this.PASTAEXPORTACAO.TextChanged += new System.EventHandler(this.PASTAEXPORTACAO_TextChanged);
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.parametrosToolStripMenuItem,
            this.desenhoToolStripMenuItem,
            this.utilitariosToolStripMenuItem,
            this.exportarToolStripMenuItem,
            this.EmailStripMenuItem1});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(409, 25);
            this.menuStrip1.TabIndex = 54;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // parametrosToolStripMenuItem
            // 
            this.parametrosToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.parametrosToolStripMenuItem1,
            this.quantificaçãoToolStripMenuItem,
            this.macrosToolStripMenuItem,
            this.alteraFaseToolStripMenuItem,
            this.lotesToolStripMenuItem});
            this.parametrosToolStripMenuItem.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.parametrosToolStripMenuItem.Name = "parametrosToolStripMenuItem";
            this.parametrosToolStripMenuItem.Size = new System.Drawing.Size(66, 21);
            this.parametrosToolStripMenuItem.Text = "Modelo";
            // 
            // parametrosToolStripMenuItem1
            // 
            this.parametrosToolStripMenuItem1.Name = "parametrosToolStripMenuItem1";
            this.parametrosToolStripMenuItem1.Size = new System.Drawing.Size(163, 22);
            this.parametrosToolStripMenuItem1.Text = "Parametros";
            this.parametrosToolStripMenuItem1.Click += new System.EventHandler(this.parametrosToolStripMenuItem1_Click);
            // 
            // quantificaçãoToolStripMenuItem
            // 
            this.quantificaçãoToolStripMenuItem.Name = "quantificaçãoToolStripMenuItem";
            this.quantificaçãoToolStripMenuItem.Size = new System.Drawing.Size(163, 22);
            this.quantificaçãoToolStripMenuItem.Text = "Quantificação";
            this.quantificaçãoToolStripMenuItem.Click += new System.EventHandler(this.quantificaçãoToolStripMenuItem_Click);
            // 
            // macrosToolStripMenuItem
            // 
            this.macrosToolStripMenuItem.Name = "macrosToolStripMenuItem";
            this.macrosToolStripMenuItem.Size = new System.Drawing.Size(163, 22);
            this.macrosToolStripMenuItem.Text = "Macros";
            this.macrosToolStripMenuItem.Click += new System.EventHandler(this.macrosToolStripMenuItem_Click);
            // 
            // alteraFaseToolStripMenuItem
            // 
            this.alteraFaseToolStripMenuItem.Name = "alteraFaseToolStripMenuItem";
            this.alteraFaseToolStripMenuItem.Size = new System.Drawing.Size(163, 22);
            this.alteraFaseToolStripMenuItem.Text = "Altera Fase";
            this.alteraFaseToolStripMenuItem.Click += new System.EventHandler(this.alteraFaseToolStripMenuItem_Click);
            // 
            // lotesToolStripMenuItem
            // 
            this.lotesToolStripMenuItem.Name = "lotesToolStripMenuItem";
            this.lotesToolStripMenuItem.Size = new System.Drawing.Size(163, 22);
            this.lotesToolStripMenuItem.Text = "Lotes";
            this.lotesToolStripMenuItem.Click += new System.EventHandler(this.lotesToolStripMenuItem_Click);
            // 
            // desenhoToolStripMenuItem
            // 
            this.desenhoToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.criarFasesToolStripMenuItem,
            this.ferramentasToolStripMenuItem,
            this.soldaduraToolStripMenuItem,
            this.desenhoToolStripMenuItem1});
            this.desenhoToolStripMenuItem.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.desenhoToolStripMenuItem.Name = "desenhoToolStripMenuItem";
            this.desenhoToolStripMenuItem.Size = new System.Drawing.Size(77, 21);
            this.desenhoToolStripMenuItem.Text = "Desenho";
            // 
            // criarFasesToolStripMenuItem
            // 
            this.criarFasesToolStripMenuItem.Name = "criarFasesToolStripMenuItem";
            this.criarFasesToolStripMenuItem.Size = new System.Drawing.Size(189, 22);
            this.criarFasesToolStripMenuItem.Text = "Criar Fases";
            this.criarFasesToolStripMenuItem.Click += new System.EventHandler(this.criarFasesToolStripMenuItem_Click);
            // 
            // ferramentasToolStripMenuItem
            // 
            this.ferramentasToolStripMenuItem.Name = "ferramentasToolStripMenuItem";
            this.ferramentasToolStripMenuItem.Size = new System.Drawing.Size(189, 22);
            this.ferramentasToolStripMenuItem.Text = "Ferramentas";
            this.ferramentasToolStripMenuItem.Click += new System.EventHandler(this.ferramentasToolStripMenuItem_Click);
            // 
            // soldaduraToolStripMenuItem
            // 
            this.soldaduraToolStripMenuItem.Name = "soldaduraToolStripMenuItem";
            this.soldaduraToolStripMenuItem.Size = new System.Drawing.Size(189, 22);
            this.soldaduraToolStripMenuItem.Text = "Soldadura";
            this.soldaduraToolStripMenuItem.Click += new System.EventHandler(this.soldaduraToolStripMenuItem_Click);
            // 
            // desenhoToolStripMenuItem1
            // 
            this.desenhoToolStripMenuItem1.Name = "desenhoToolStripMenuItem1";
            this.desenhoToolStripMenuItem1.Size = new System.Drawing.Size(189, 22);
            this.desenhoToolStripMenuItem1.Text = "Verificar Desenho";
            this.desenhoToolStripMenuItem1.Click += new System.EventHandler(this.desenhoToolStripMenuItem1_Click);
            // 
            // utilitariosToolStripMenuItem
            // 
            this.utilitariosToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.abrirObraToolStripMenuItem,
            this.testeToolStripMenuItem,
            this.limparTodasAsUDADaPeçaToolStripMenuItem,
            this.pDFSoldaduraToolStripMenuItem,
            alterarBaseDeDadosCPEDapToolStripMenuItem,
            this.nESTDESENVOLVIMENTOToolStripMenuItem1});
            this.utilitariosToolStripMenuItem.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.utilitariosToolStripMenuItem.Name = "utilitariosToolStripMenuItem";
            this.utilitariosToolStripMenuItem.Size = new System.Drawing.Size(78, 21);
            this.utilitariosToolStripMenuItem.Text = "Utilitarios";
            // 
            // abrirObraToolStripMenuItem
            // 
            this.abrirObraToolStripMenuItem.Name = "abrirObraToolStripMenuItem";
            this.abrirObraToolStripMenuItem.Size = new System.Drawing.Size(285, 22);
            this.abrirObraToolStripMenuItem.Text = "Abrir Obra";
            this.abrirObraToolStripMenuItem.Click += new System.EventHandler(this.abrirObraToolStripMenuItem_Click);
            // 
            // testeToolStripMenuItem
            // 
            this.testeToolStripMenuItem.Name = "testeToolStripMenuItem";
            this.testeToolStripMenuItem.Size = new System.Drawing.Size(285, 22);
            this.testeToolStripMenuItem.Text = "Imprimri Desenhos";
            this.testeToolStripMenuItem.Click += new System.EventHandler(this.testeToolStripMenuItem_Click);
            // 
            // limparTodasAsUDADaPeçaToolStripMenuItem
            // 
            this.limparTodasAsUDADaPeçaToolStripMenuItem.Name = "limparTodasAsUDADaPeçaToolStripMenuItem";
            this.limparTodasAsUDADaPeçaToolStripMenuItem.Size = new System.Drawing.Size(285, 22);
            this.limparTodasAsUDADaPeçaToolStripMenuItem.Text = "Limpar todas as UDA\'s da Peça";
            this.limparTodasAsUDADaPeçaToolStripMenuItem.Click += new System.EventHandler(this.limparTodasAsUDADaPeçaToolStripMenuItem_Click);
            // 
            // pDFSoldaduraToolStripMenuItem
            // 
            this.pDFSoldaduraToolStripMenuItem.Name = "pDFSoldaduraToolStripMenuItem";
            this.pDFSoldaduraToolStripMenuItem.Size = new System.Drawing.Size(285, 22);
            this.pDFSoldaduraToolStripMenuItem.Text = "PDF Soldadura";
            this.pDFSoldaduraToolStripMenuItem.Click += new System.EventHandler(this.pDFSoldaduraToolStripMenuItem_Click);
            // 
            // nESTDESENVOLVIMENTOToolStripMenuItem1
            // 
            this.nESTDESENVOLVIMENTOToolStripMenuItem1.Name = "nESTDESENVOLVIMENTOToolStripMenuItem1";
            this.nESTDESENVOLVIMENTOToolStripMenuItem1.Size = new System.Drawing.Size(285, 22);
            this.nESTDESENVOLVIMENTOToolStripMenuItem1.Text = "NEST(DESENVOLVIMENTO)";
            this.nESTDESENVOLVIMENTOToolStripMenuItem1.Click += new System.EventHandler(this.nESTDESENVOLVIMENTOToolStripMenuItem1_Click);
            // 
            // exportarToolStripMenuItem
            // 
            this.exportarToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exportarNC1ToolStripMenuItem});
            this.exportarToolStripMenuItem.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.exportarToolStripMenuItem.Name = "exportarToolStripMenuItem";
            this.exportarToolStripMenuItem.Size = new System.Drawing.Size(73, 21);
            this.exportarToolStripMenuItem.Text = "Exportar";
            // 
            // exportarNC1ToolStripMenuItem
            // 
            this.exportarNC1ToolStripMenuItem.Name = "exportarNC1ToolStripMenuItem";
            this.exportarNC1ToolStripMenuItem.Size = new System.Drawing.Size(259, 22);
            this.exportarNC1ToolStripMenuItem.Text = "Exportar e conversor de NC1";
            this.exportarNC1ToolStripMenuItem.Click += new System.EventHandler(this.exportarNC1ToolStripMenuItem_Click);
            // 
            // EmailStripMenuItem1
            // 
            this.EmailStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.enviarParaFabricoToolStripMenuItem,
            this.enviarEmailParaLotearToolStripMenuItem,
            this.enviarEmialAprovToolStripMenuItem,
            this.revisãoPeçasEConjuntosToolStripMenuItem});
            this.EmailStripMenuItem1.Name = "EmailStripMenuItem1";
            this.EmailStripMenuItem1.Size = new System.Drawing.Size(48, 21);
            this.EmailStripMenuItem1.Text = "Email";
            // 
            // enviarParaFabricoToolStripMenuItem
            // 
            this.enviarParaFabricoToolStripMenuItem.Name = "enviarParaFabricoToolStripMenuItem";
            this.enviarParaFabricoToolStripMenuItem.Size = new System.Drawing.Size(458, 22);
            this.enviarParaFabricoToolStripMenuItem.Text = "Enviar Email para Fabrico";
            this.enviarParaFabricoToolStripMenuItem.Click += new System.EventHandler(this.enviarParaFabricoToolStripMenuItem_Click);
            // 
            // enviarEmailParaLotearToolStripMenuItem
            // 
            this.enviarEmailParaLotearToolStripMenuItem.Name = "enviarEmailParaLotearToolStripMenuItem";
            this.enviarEmailParaLotearToolStripMenuItem.Size = new System.Drawing.Size(458, 22);
            this.enviarEmailParaLotearToolStripMenuItem.Text = "Enviar Email para Lotear";
            this.enviarEmailParaLotearToolStripMenuItem.Click += new System.EventHandler(this.enviarEmailParaLotearToolStripMenuItem_Click);
            // 
            // enviarEmialAprovToolStripMenuItem
            // 
            this.enviarEmialAprovToolStripMenuItem.Name = "enviarEmialAprovToolStripMenuItem";
            this.enviarEmialAprovToolStripMenuItem.Size = new System.Drawing.Size(458, 22);
            this.enviarEmialAprovToolStripMenuItem.Text = "Enviar Email de Aprovisionamentos";
            this.enviarEmialAprovToolStripMenuItem.Click += new System.EventHandler(this.enviarEmialAprovToolStripMenuItem_Click);
            // 
            // revisãoPeçasEConjuntosToolStripMenuItem
            // 
            this.revisãoPeçasEConjuntosToolStripMenuItem.Name = "revisãoPeçasEConjuntosToolStripMenuItem";
            this.revisãoPeçasEConjuntosToolStripMenuItem.Size = new System.Drawing.Size(458, 22);
            this.revisãoPeçasEConjuntosToolStripMenuItem.Text = "Enviar Email ( Estudo Previo /Projeto de execução/ Desenho montagem )";
            this.revisãoPeçasEConjuntosToolStripMenuItem.Click += new System.EventHandler(this.revisãoPeçasEConjuntosToolStripMenuItem_Click);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label1.Location = new System.Drawing.Point(10, 40);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(102, 17);
            this.label1.TabIndex = 64;
            this.label1.Text = "Numero Obra :";
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(285, 85);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(88, 15);
            this.label2.TabIndex = 65;
            this.label2.Text = "Sempre Visivel";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold);
            this.label3.Location = new System.Drawing.Point(10, 42);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(285, 20);
            this.label3.TabIndex = 66;
            this.label3.Text = "Abra o Tekla ou Verifique a versão";
            this.label3.Visible = false;
            // 
            // button9
            // 
            this.button9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button9.BorderColor = System.Drawing.Color.Silver;
            this.button9.BorderRadius = 8;
            this.button9.BorderThickness = 1;
            this.button9.DisabledState.BorderColor = System.Drawing.Color.DarkGray;
            this.button9.DisabledState.CustomBorderColor = System.Drawing.Color.DarkGray;
            this.button9.DisabledState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(169)))), ((int)(((byte)(169)))), ((int)(((byte)(169)))));
            this.button9.DisabledState.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(141)))), ((int)(((byte)(141)))), ((int)(((byte)(141)))));
            this.button9.FillColor = System.Drawing.SystemColors.ButtonHighlight;
            this.button9.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.button9.ForeColor = System.Drawing.Color.Black;
            this.button9.Location = new System.Drawing.Point(300, 25);
            this.button9.Margin = new System.Windows.Forms.Padding(0);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(110, 35);
            this.button9.TabIndex = 90;
            this.button9.Text = "Carregar Obra";
            this.button9.TextOffset = new System.Drawing.Point(1, -2);
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // webBrowser1
            // 
            this.webBrowser1.Location = new System.Drawing.Point(10, 90);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.ScrollBarsEnabled = false;
            this.webBrowser1.Size = new System.Drawing.Size(380, 185);
            this.webBrowser1.TabIndex = 91;
            this.webBrowser1.Visible = false;
            // 
            // Frm_Inico
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLight;
            this.ClientSize = new System.Drawing.Size(409, 111);
            this.Controls.Add(this.webBrowser1);
            this.Controls.Add(this.button9);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.PASTAEXPORTACAO);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.LBLestado);
            this.Controls.Add(this.Chb_alw_top);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.MinimumSize = new System.Drawing.Size(425, 150);
            this.Name = "Frm_Inico";
            this.Text = "   Versão 2021     Processo de Fabrico ";
            this.TopMost = true;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Frm_Inico_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Timer timer1;
        public System.Windows.Forms.Label LBLestado;
        private Button button54;
        private TextBox PASTAEXPORTACAO;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem exportarToolStripMenuItem;
        private ToolStripMenuItem exportarNC1ToolStripMenuItem;
        private ToolStripMenuItem parametrosToolStripMenuItem;
        private ToolStripMenuItem parametrosToolStripMenuItem1;
        private ToolStripMenuItem quantificaçãoToolStripMenuItem;
        private ToolStripMenuItem desenhoToolStripMenuItem;
        private ToolStripMenuItem criarFasesToolStripMenuItem;
        private ToolStripMenuItem ferramentasToolStripMenuItem;
        public Label label11;
        public CheckBox Chb_alw_top;
        private ToolStripMenuItem soldaduraToolStripMenuItem;
        private ToolStripMenuItem macrosToolStripMenuItem;
        private ToolStripMenuItem utilitariosToolStripMenuItem;
        private ToolStripMenuItem abrirObraToolStripMenuItem;
        private ToolStripMenuItem alteraFaseToolStripMenuItem;
        private ToolStripMenuItem testeToolStripMenuItem;
        private ContextMenuStrip contextMenuStrip1;
        private ToolStripMenuItem pDFSoldaduraToolStripMenuItem;
        private ToolStripMenuItem lotesToolStripMenuItem;
        private Label label1;
        public Label label2;
        public Label label3;
        private ToolStripMenuItem nESTDESENVOLVIMENTOToolStripMenuItem1;
        private Guna.UI2.WinForms.Guna2Button button9;
        private ToolStripMenuItem EmailStripMenuItem1;
        private ToolStripMenuItem revisãoPeçasEConjuntosToolStripMenuItem;
        private ToolStripMenuItem enviarParaFabricoToolStripMenuItem;
        private ToolStripMenuItem enviarEmailParaLotearToolStripMenuItem;
        private ToolStripMenuItem enviarEmialAprovToolStripMenuItem;
        private ToolStripMenuItem limparTodasAsUDADaPeçaToolStripMenuItem;
        private ToolStripMenuItem desenhoToolStripMenuItem1;
        private WebBrowser webBrowser1;
    }
}

