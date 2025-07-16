namespace TeklaArtigosOfeliz
{
    partial class Frm_EnviarEmailparaFabrico
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Frm_EnviarEmailparaFabrico));
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.Chb_alw_top = new System.Windows.Forms.CheckBox();
            this.button9 = new Guna.UI2.WinForms.Guna2Button();
            this.guna2CustomGradientPanel1 = new Guna.UI2.WinForms.Guna2CustomGradientPanel();
            this.guna2CustomGradientPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Location = new System.Drawing.Point(5, 9);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(271, 13);
            this.textBox1.TabIndex = 0;
            this.textBox1.Tag = "";
            this.textBox1.Text = "Nome do Material enviar para Fabrico";
            // 
            // Chb_alw_top
            // 
            this.Chb_alw_top.AutoSize = true;
            this.Chb_alw_top.Checked = true;
            this.Chb_alw_top.CheckState = System.Windows.Forms.CheckState.Checked;
            this.Chb_alw_top.Location = new System.Drawing.Point(384, -7);
            this.Chb_alw_top.Name = "Chb_alw_top";
            this.Chb_alw_top.Size = new System.Drawing.Size(15, 14);
            this.Chb_alw_top.TabIndex = 2;
            this.Chb_alw_top.UseVisualStyleBackColor = true;
            this.Chb_alw_top.Visible = false;
            this.Chb_alw_top.CheckedChanged += new System.EventHandler(this.Chb_alw_top_CheckedChanged);
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
            this.button9.Location = new System.Drawing.Point(302, 10);
            this.button9.Margin = new System.Windows.Forms.Padding(0);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(97, 35);
            this.button9.TabIndex = 89;
            this.button9.Text = "Abrir Email";
            this.button9.TextOffset = new System.Drawing.Point(1, -2);
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // guna2CustomGradientPanel1
            // 
            this.guna2CustomGradientPanel1.BackColor = System.Drawing.Color.Transparent;
            this.guna2CustomGradientPanel1.BorderColor = System.Drawing.Color.Silver;
            this.guna2CustomGradientPanel1.BorderRadius = 10;
            this.guna2CustomGradientPanel1.BorderThickness = 1;
            this.guna2CustomGradientPanel1.Controls.Add(this.textBox1);
            this.guna2CustomGradientPanel1.Location = new System.Drawing.Point(10, 10);
            this.guna2CustomGradientPanel1.Name = "guna2CustomGradientPanel1";
            this.guna2CustomGradientPanel1.Size = new System.Drawing.Size(282, 30);
            this.guna2CustomGradientPanel1.TabIndex = 92;
            // 
            // Frm_FabricoEmail
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLight;
            this.ClientSize = new System.Drawing.Size(409, 51);
            this.Controls.Add(this.guna2CustomGradientPanel1);
            this.Controls.Add(this.button9);
            this.Controls.Add(this.Chb_alw_top);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(425, 90);
            this.MinimumSize = new System.Drawing.Size(425, 90);
            this.Name = "Frm_FabricoEmail";
            this.Text = "Enviar Email para Fabrico";
            this.Load += new System.EventHandler(this.FabricoEmail_Load);
            this.guna2CustomGradientPanel1.ResumeLayout(false);
            this.guna2CustomGradientPanel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        public System.Windows.Forms.CheckBox Chb_alw_top;
        private Guna.UI2.WinForms.Guna2Button button9;
        private Guna.UI2.WinForms.Guna2CustomGradientPanel guna2CustomGradientPanel1;
    }
}