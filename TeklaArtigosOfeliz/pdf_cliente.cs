using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace TeklaArtigosOfeliz
{
    public partial class pdf_cliente : Form
    {
        Frm_ListaOFeliz _FORMPAI = null;

        public pdf_cliente(Frm_ListaOFeliz FORMPAI)
        {

            InitializeComponent();
            _FORMPAI = FORMPAI;

        }
      
        private void pdf_cliente_Load(object sender, EventArgs e)
        {
            int x = 0;
            foreach (DataGridViewColumn Coluna in _FORMPAI.dataGridView1.Columns)
            {
                x++;
                if (x==5)
                {
                    dataGridView1.Columns.Add(x + "=" + Coluna.HeaderText.ToString(), Coluna.HeaderText.ToString());
                }
                if (x == 18)
                {
                    dataGridView1.Columns.Add(x + "=" + Coluna.HeaderText.ToString(), Coluna.HeaderText.ToString());
                }
            }
            foreach (DataGridViewRow row in _FORMPAI.dataGridView1.Rows)
            {
                if (row.Cells[17].Value != "" && row.Cells[17].Value!=null)
                {
                    dataGridView1.Rows.Add(row.Cells[4].Value,row.Cells[17].Value.ToString().Replace("/","-"));
                }
            }
            dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
        }

        private void MERGE_Click(object sender, EventArgs e)
        {
          string[] files = Directory.GetFiles(textBox1.Text, "*.pdf", SearchOption.TopDirectoryOnly);
            List<string> ERROS = new List<string>();

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value != "" && dataGridView1.Rows[i].Cells[0].Value != null)
                {
                    if (dataGridView1.Rows[i].Cells[1].Value != "" && dataGridView1.Rows[i].Cells[1].Value != null)
                    {
                        string ficheiro ="c:\\r\\"+ dataGridView1.Rows[i].Cells[0].Value.ToString().Split('.')[0] + "." + dataGridView1.Rows[i].Cells[0].Value.ToString().Split('.')[1] + "." + dataGridView1.Rows[i].Cells[0].Value.ToString().Split('.')[3] + ".pdf";
                        string ficheiro1 = null;

                        foreach (var item in files)
                        {
                            if (item.Split('\\').Last().ToUpper().Trim() == dataGridView1.Rows[i].Cells[1].Value.ToString().ToUpper().Trim()+".PDF")
                            {
                             ficheiro1 = item;
                            }
                        }
                     

                        if (ficheiro1!=null)
                        {
                            if (File.Exists(ficheiro))
                            {
                                if (File.Exists(ficheiro1))
                                {
                                    List<string> filesmerge = new List<string>();

                                    filesmerge.Add(ficheiro);
                                    filesmerge.Add(ficheiro1);
                                    string outfile = ficheiro.Replace(".pdf", "-.pdf");
                                    MergePDFs(filesmerge, outfile);
                                    File.Delete(ficheiro);
                                    File.Move(outfile, ficheiro);
                                }
                                else
                                {
                                    ERROS.Add("Não existe o ficheiro = " + ficheiro1);
                                }
                            }
                            else
                            {
                                ERROS.Add("Não existe o ficheiro = " + ficheiro);
                            }
                        }
                        else
                        {
                            
                                ERROS.Add("Não existe o ficheiro = " + dataGridView1.Rows[i].Cells[1].Value + " na pasta");
                            
                        }
                    }
                }
            }
            if (ERROS!=null)
            {
                string ERRO = null;
                foreach (var item in ERROS)
                {
                    ERRO += item;
                }

                MessageBox.Show(this, ERRO);

            }
            if (ERROS == null)
            {

                MessageBox.Show(this, "CONCLUIDO");

            }


        }

        public static bool MergePDFs(List<String> InFiles, String OutFile)
        {
            bool merged = true;
            try
            {
                List<PdfReader> readerList = new List<PdfReader>();
                foreach (string filePath in InFiles)
                {
                    PdfReader pdfReader = new PdfReader(filePath);
                    readerList.Add(pdfReader);
                }

                //Define a new output document and its size, type
                Document document = new Document(PageSize.A4, 0, 0, 0, 0);
                //Create blank output pdf file and get the stream to write on it.
                PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(OutFile, FileMode.Create));
                document.Open();

                foreach (PdfReader reader in readerList)
                {
                    PdfReader.unethicalreading = true;
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        PdfImportedPage page = writer.GetImportedPage(reader, i);
                        document.Add(iTextSharp.text.Image.GetInstance(page));
                    }
                }
                document.Close();
                foreach (PdfReader reader in readerList)
                {
                    reader.Close();
                }

            }
            catch (Exception)
            {
                merged = false;
            }


            return merged;
        }


    }
}
