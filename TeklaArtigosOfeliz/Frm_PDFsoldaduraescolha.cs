using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TSM = Tekla.Structures.Model;
using Tekla.Structures;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model.UI;
using Tekla.Structures.Filtering;
using Tekla.Structures.Filtering.Categories;
using System.Collections;
using Tekla.Structures.Model;
using iTextSharp.text.pdf;
using iTextSharp.text;


namespace TeklaArtigosOfeliz
{
    public partial class Frm_PDFsoldaduraescolha : Form
    {
              
        public Frm_PDFsoldaduraescolha()
        {
            InitializeComponent();
        }

        string path = null;
        string obra = null;
        string cliente = null;
        string designacao = null;
        string classe = null;

        private void PDFsoldaduraescolha_Load(object sender, EventArgs e)
        {

        }

        void refreshprog()
        {

            TopMost = true;

            TSM.Model M = new Model();

            try
            {
                string Exc = string.Empty;
                string teste = M.GetInfo().ModelName.Replace(".db1", "");
                path = M.GetInfo().ModelPath.Replace(teste, "");
                obra = M.GetProjectInfo().ProjectNumber;
                cliente = M.GetProjectInfo().Name;
                designacao = M.GetProjectInfo().Builder;

                string EXC = null;
                new Model().GetProjectInfo().GetUserProperty("PROJECT_USERFIELD_2", ref EXC);
                if (EXC.Trim() == "")
                    EXC = "2";

                string classe = EXC;

                label2.Text = obra;
                label3.Text = cliente;
                label4.Text = classe;
                label5.Text = designacao;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "ERRO P.F. ABRA O TEKLA SE ESTIVER ABERTO P.F. REINICIE ESTE PROGRAMA" + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
        }

        private void textBoxObra_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            refreshprog();
        }

        private void buttonEnviarPDF_Click(object sender, EventArgs e)
        {
            PDF();
        }

        public void PDF()
        {
            string obra = label2.Text;
            string fase = textBoxFase.Text;
            string ano = string.Empty;
            string designacao = label5.Text;
            string cliente = label3.Text;
            string classe = label4.Text;

            if (obra.Contains("PT"))
            {
                ano = "20" + obra.Substring(2, 2);
            }
            else
            {
                ano = "20" + obra.Substring(0, 2);
            }

            if (int.TryParse(fase, out int faseNum))
            {
                if (faseNum >= 1 && faseNum <= 99)
                {
                    fase = faseNum.ToString().PadLeft(3, '0');
                }
            }

            string ficheironome = "Plano_Soldadura_Fase" + fase + ".pdf";

            string caminhoPDFLimpo = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Plano_Soldadura_Fase.pdf";

            string pdfPath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras"
                 + "\\" + ano + "\\" + obra
                 + @"\1.9 Gestão de fabrico\"
                 + fase + @"\20005\"
                 + ficheironome;

            FecharFicheiroPDF(pdfPath);


            if (File.Exists(pdfPath))
            {
                if (File.Exists(pdfPath))
                {
                    File.Delete(pdfPath);
                }

                File.Copy(caminhoPDFLimpo, pdfPath);

                string textoParaPDF138_1 = ("SAFDUAL 206A - 1,2mm (EN ISO 17632 - A: T42 2 M M 1 H5");
                string textoParaPDF138_2 = ("MX - 100T - 1,2mm(EN ISO 17632 - A: T42 2 M M / C 1 H5)");
                string textoParaPDF121 = ("AIR LIQUIDE: OE - S2 / AIR LIQUIDE: F.OP 139");
                string Anotacoes = textBox1.Text;

                try
                    {
                        using (FileStream fs = new FileStream(pdfPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                        {
                            PdfReader reader = new PdfReader(fs);

                            PdfStamper stamper = new PdfStamper(reader, new FileStream(pdfPath, FileMode.Create));

                            PdfContentByte pbover = stamper.GetOverContent(1);

                            var baseFont = BaseFont.CreateFont(@"C:\Windows\Fonts\arialbd.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                            var baseFont2 = BaseFont.CreateFont(@"C:\Windows\Fonts\Calibri.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

                        if (checkBox138.Checked)
                        {
                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(textoParaPDF138_1.ToString(), new iTextSharp.text.Font(baseFont2, 10)), 215, 297, 0);
                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(textoParaPDF138_2.ToString(), new iTextSharp.text.Font(baseFont2, 10)), 215, 282, 0);

                        } else{}

                        if (checkBox121.Checked)
                        {
                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(textoParaPDF121.ToString(), new iTextSharp.text.Font(baseFont2, 10)), 215, 297, 0);
                        } else { }

                        if (checkBoxAmbas.Checked)
                        {
                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(textoParaPDF138_1.ToString(), new iTextSharp.text.Font(baseFont2, 10)), 215, 297, 0);
                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(textoParaPDF138_2.ToString(), new iTextSharp.text.Font(baseFont2, 10)), 215, 282, 0);
                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(textoParaPDF121.ToString(), new iTextSharp.text.Font(baseFont2, 10)), 215, 265, 0);

                        } else { }

                        if (!string.IsNullOrEmpty(textBox1.Text))
                        {
                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT,new Phrase(Anotacoes.ToString(), new iTextSharp.text.Font(baseFont2, 9)),75, 190, 0);
                        }  else { }

                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_CENTER, new Phrase(DateTime.Now.ToShortDateString(), new iTextSharp.text.Font(baseFont, 8)), 485, 140, 0);
                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(fase, new iTextSharp.text.Font(baseFont, 10)), 485, 725, 0);
                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(obra, new iTextSharp.text.Font(baseFont, 10)), 133, 725, 0);
                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(cliente, new iTextSharp.text.Font(baseFont, 10)), 130, 711, 0);
                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(designacao, new iTextSharp.text.Font(baseFont, 10)), 150, 697, 0);
                            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(classe, new iTextSharp.text.Font(baseFont, 10)), 185, 683, 0);

                            stamper.Close();
                        }

                        MessageBox.Show(this, "PDF atualizado com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, $"Erro ao gerar PDF: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                
            }
            else
            {
                MessageBox.Show(this, "Arquivo 'Plano_Soldadura_Fase.pdf' não encontrado.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void FecharFicheiroPDF(string pdfPath)
        {
            try
            {
                using (FileStream fs = new FileStream(pdfPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    PdfReader reader = new PdfReader(fs);

                    using (PdfStamper stamper = new PdfStamper(reader, new FileStream(pdfPath, FileMode.Create)))
                    {
                        PdfContentByte pbover = stamper.GetOverContent(1);
                                                
                        stamper.Close();
                    }

                    reader.Close();
                }                
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Erro ao fechar o ficheiro PDF: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
    }

