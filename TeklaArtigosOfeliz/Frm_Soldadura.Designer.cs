namespace TeklaArtigosOfeliz
{
    partial class Frm_Soldadura
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Frm_Soldadura));
            this.checkBox6 = new System.Windows.Forms.CheckBox();
            this.label18 = new System.Windows.Forms.Label();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.button41 = new System.Windows.Forms.Button();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // checkBox6
            // 
            this.checkBox6.AutoSize = true;
            this.checkBox6.Checked = true;
            this.checkBox6.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox6.Location = new System.Drawing.Point(6, 2);
            this.checkBox6.Name = "checkBox6";
            this.checkBox6.Size = new System.Drawing.Size(192, 17);
            this.checkBox6.TabIndex = 11;
            this.checkBox6.Text = "Colocar marcas em todas as vistas.";
            this.checkBox6.UseVisualStyleBackColor = true;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(7, 67);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(37, 13);
            this.label18.TabIndex = 10;
            this.label18.Text = "FASE:";
            // 
            // textBox6
            // 
            this.textBox6.Location = new System.Drawing.Point(54, 60);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(49, 20);
            this.textBox6.TabIndex = 9;
            this.textBox6.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox6_KeyPress);
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(2, 33);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(46, 13);
            this.label17.TabIndex = 8;
            this.label17.Text = "TEXTO:";
            // 
            // button41
            // 
            this.button41.Location = new System.Drawing.Point(444, 52);
            this.button41.Name = "button41";
            this.button41.Size = new System.Drawing.Size(91, 59);
            this.button41.TabIndex = 7;
            this.button41.Text = "Criar Desenhos ";
            this.button41.UseVisualStyleBackColor = true;
            this.button41.Click += new System.EventHandler(this.button41_Click);
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "NOTA: Soldar segundo os nossos procedimentos habituais",
            "NOTA: Soldar com um cordão de 0,7xmenor espessura a soldar, por indicação do clie" +
                "nte"});
            this.comboBox2.Location = new System.Drawing.Point(54, 25);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(481, 21);
            this.comboBox2.TabIndex = 6;
            // 
            // FrmSoldadura
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(547, 121);
            this.Controls.Add(this.checkBox6);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.textBox6);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.button41);
            this.Controls.Add(this.comboBox2);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(563, 160);
            this.Name = "FrmSoldadura";
            this.Text = "FrmSoldadura";
            this.Load += new System.EventHandler(this.FrmSoldadura_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox checkBox6;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.TextBox textBox6;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Button button41;
        private System.Windows.Forms.ComboBox comboBox2;
    }
}