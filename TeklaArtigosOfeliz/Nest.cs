using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TeklaArtigosOfeliz
{
    class Nest
    {
        public void fazernesting(string pastadoscsvexpprimavera)
        {
   
            string[] Files = Directory.GetFiles(pastadoscsvexpprimavera, "*.csv", SearchOption.TopDirectoryOnly);

            foreach (string file in Files)
            {
              
                DataTable pecas = new DataTable();
                DataTable barras = new DataTable();
                pecas.Columns.Add("quantidade");
                pecas.Columns.Add("comprimento");
                pecas.Columns.Add("referencia");
                barras.Columns.Add("quantidade");
                barras.Columns.Add("comprimento");
                barras.Columns.Add("prioridade");
                barras.Columns.Add("referencia");

                bool bopecas = false;
                bool bobarras = false;
                string numerodeobra = null;
                string nomecliente = null;
                string nomeobra = null;
                int a = 0;
                foreach (string line in File.ReadLines(file, Encoding.Default))
                {
                    if (a == 0)
                    {
                        numerodeobra = line;
                    }
                    if (a == 1)
                    {
                        nomecliente = line;
                    }
                    if (a == 2)
                    {
                        nomeobra = line;
                    }
                    a++;


                    if (bopecas)
                    {
                        string[] peca = line.Split(';');
                        if (line.Replace(" ", "") != "")
                        {
                            pecas.Rows.Add(peca[2], peca[1], peca[0]);
                        }
                        else
                        {
                            bopecas = false;
                        }
                    }
                    if (bobarras)
                    {
                        string[] barra = line.Split(';');
                        if (line.Replace(" ", "") != "")
                        {
                            barras.Rows.Add(barra[2], barra[1], 0, barra[0]);
                        }
                        else
                        {
                            bobarras = false;
                        }
                    }


                    if (line == "PIECE;LENGTH;QUANTITY")
                    {
                        bopecas = true;
                    }
                    if (line == "BAR;LENGTH;QUANTITY")
                    {
                        bobarras = true;
                    }


                }

                DataTable teste = otimizar(pecas, barras,0);

                String outputPath = pastadoscsvexpprimavera + file.Split('\\').Last().Replace(".csv", ".xls");


                if (teste != null)
                {
                   
                    Microsoft.Office.Interop.Excel.Application excel = new Excel.Application();
                    Excel.Workbook workbook = excel.Workbooks.Open(Environment.CurrentDirectory + "\\template.xls", 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
                    ((Excel.Range)sheet.Cells[1, 2]).Value = numerodeobra.Replace(";", "|");
                    ((Excel.Range)sheet.Cells[2, 2]).Value = nomecliente.Replace(";", "|");
                    ((Excel.Range)sheet.Cells[3, 2]).Value = nomeobra.Replace(";", "|");
                    ((Excel.Range)sheet.Cells[1, 4]).Value = Percentagemdedesperdicio;
                    int contador = 5;
                    int contador1 = 1;
                    foreach (DataRow item in teste.Rows)
                    {
                        contador1++;
                        if (contador1 % 2 != 0)
                        {

                            ((Excel.Range)sheet.Cells[contador, 1]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            ((Excel.Range)sheet.Cells[contador, 2]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            ((Excel.Range)sheet.Cells[contador, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            ((Excel.Range)sheet.Cells[contador, 5]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            ((Excel.Range)sheet.Cells[contador, 4]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            ((Excel.Range)sheet.Cells[contador, 6]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            ((Excel.Range)sheet.Cells[contador, 7]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                        }
                        else
                        {
                            ((Excel.Range)sheet.Cells[contador, 1]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            ((Excel.Range)sheet.Cells[contador, 2]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            ((Excel.Range)sheet.Cells[contador, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            ((Excel.Range)sheet.Cells[contador, 5]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            ((Excel.Range)sheet.Cells[contador, 4]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            ((Excel.Range)sheet.Cells[contador, 6]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            ((Excel.Range)sheet.Cells[contador, 7]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        }

                        ((Excel.Range)sheet.Cells[contador, 1]).Value = item.ItemArray[2].ToString();
                        ((Excel.Range)sheet.Cells[contador, 2]).Value = item.ItemArray[1].ToString();
                        ((Excel.Range)sheet.Cells[contador, 3]).Value = 1;
                        ((Excel.Range)sheet.Cells[contador, 7]).Value = item.ItemArray[4];
                        int novocontador = 0;
                        List<string> repetidos = new List<string>();
                        foreach (string novamedida in item.ItemArray[3].ToString().Split(','))
                        {


                            if (contador1 % 2 != 0)
                            {

                                ((Excel.Range)sheet.Cells[contador, 1]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                                ((Excel.Range)sheet.Cells[contador, 2]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                                ((Excel.Range)sheet.Cells[contador, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                                ((Excel.Range)sheet.Cells[contador, 5]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                                ((Excel.Range)sheet.Cells[contador, 4]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                                ((Excel.Range)sheet.Cells[contador, 6]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                                ((Excel.Range)sheet.Cells[contador, 7]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                            }
                            else
                            {
                                ((Excel.Range)sheet.Cells[contador, 1]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                                ((Excel.Range)sheet.Cells[contador, 2]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                                ((Excel.Range)sheet.Cells[contador, 3]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                                ((Excel.Range)sheet.Cells[contador, 5]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                                ((Excel.Range)sheet.Cells[contador, 4]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                                ((Excel.Range)sheet.Cells[contador, 6]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                                ((Excel.Range)sheet.Cells[contador, 7]).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                            }

                            ((Excel.Range)sheet.Cells[contador, 5]).Value = novamedida;
                            ((Excel.Range)sheet.Cells[contador, 4]).Value = item.ItemArray[5].ToString().Split('#')[novocontador].Replace("\"", "");
                            int soma = 0;
                            string mysr = item.ItemArray[5].ToString().Split('#')[novocontador].Replace(" ", "");


                            foreach (string con in item.ItemArray[5].ToString().Split('#'))
                            {
                                if (mysr == con.Replace(" ", ""))
                                {
                                    soma = soma + 1;

                                }
                            }
                            ((Excel.Range)sheet.Cells[contador, 6]).Value = soma.ToString();
                            novocontador++;
                            if (!repetidos.Any(str => str.Contains(mysr)))
                            {
                                repetidos.Add(mysr);
                                contador += 1;
                            }
                            else
                            {
                                ((Excel.Range)sheet.Cells[contador, 1]).Value = "";
                                ((Excel.Range)sheet.Cells[contador, 2]).Value = "";
                                ((Excel.Range)sheet.Cells[contador, 3]).Value = "";
                                ((Excel.Range)sheet.Cells[contador, 5]).Value = "";
                                ((Excel.Range)sheet.Cells[contador, 4]).Value = "";
                                ((Excel.Range)sheet.Cells[contador, 6]).Value = "";
                                ((Excel.Range)sheet.Cells[contador, 7]).Value = "";
                            }
                        }
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    workbook.CheckCompatibility = false;
                    workbook.SaveAs(outputPath);
                    workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    excel.Quit();
                    Marshal.ReleaseComObject(sheet);
                    Marshal.ReleaseComObject(workbook);
                    Marshal.ReleaseComObject(excel);
                }
                else
                {
                    MessageBox.Show("ERRO AO CRIAR O PERFIL - " + file.Split('\\').Last().Replace(".csv", ""));
                }
            }
        }

        private static List<propriedadedebarra> PossibleLengths = new List<propriedadedebarra>();
        private static ListaPlank RELATORIO;
        private string Percentagemdeaproveitamento = null;
        public string Percentagemdedesperdicio = null;
        private string Percentagemsemsucata = null;
        private string Percentagemdedesperdiciosemsucata = null;
        private static decimal ESPESSURASERRA;

        public DataTable otimizar(DataTable peças, DataTable barras, decimal _ESPESSURASERRA = 0)
        {
            ESPESSURASERRA = _ESPESSURASERRA;
            PossibleLengths.Clear();
            DataTable resultado = new DataTable();
            resultado.Columns.Add("ID");
            resultado.Columns.Add("barra");
            resultado.Columns.Add("tamanho da barra");
            resultado.Columns.Add("comprimento a cortar");
            resultado.Columns.Add("desperdicio");
            resultado.Columns.Add("marca de peça");
            resultado.Columns.Add("marca de");

            barras.DefaultView.Sort = barras.Columns[2].ColumnName + " ASC";
            //carregar os cortes que queremos 
            List<Part> DesiredLength = new List<Part>();
            try
            {
                for (int i = 0; i < peças.Rows.Count; i++)
                {
                    int a = Convert.ToInt32(peças.Rows[i].ItemArray[0].ToString());
                    while (a > 0)
                    {

                        DesiredLength.Add(new Part(float.Parse(peças.Rows[i].ItemArray[1].ToString()), peças.Rows[i].ItemArray[2].ToString()));
                        a = a - 1;
                    }
                }
            }
            catch (Exception) { }


            List<ListaPlank> aproveitamentos = new List<ListaPlank>();



            //preparar as otimizaçoes
            int prioridades = 0;

            foreach (DataRow item in barras.Rows)
            {
                try
                {
                    int a = 0;
                    int.TryParse(item.ItemArray[2].ToString(), out a);
                    prioridades += a;
                }
                catch (Exception)
                {
                }
            }
            int teste = 2;
            if (prioridades == 0)
            {
                teste = 6;
            }

            for (int i = 0; i < teste; i++)
            {
                //carregar as barras a ser cortadas 
                try
                {
                    for (int a = 0; a < barras.Rows.Count; a++)
                    {
                        int b = Convert.ToInt32(barras.Rows[a].ItemArray[0].ToString());

                        while (b > 0)
                        {
                            PossibleLengths.Add(new propriedadedebarra { Column1 = float.Parse(barras.Rows[a].ItemArray[1].ToString()), Column2 = barras.Rows[a].ItemArray[3].ToString() });
                            b = b - 1;
                        }
                    }
                }
                catch (Exception)
                {

                }

                //enviar os diferentes configuraçoes das otimizaçoes 

                List<Part> DesiredLengths = null;
                if (i == 0)
                {
                    DesiredLengths = new List<Part>(DesiredLength.OrderBy(x => x.OriginalLength));
                }
                else if (i == 1)
                {
                    DesiredLengths = new List<Part>(DesiredLength.OrderBy(x => x.OriginalLength).Reverse());
                }
                else if (i == 2)
                {
                    PossibleLengths = new List<propriedadedebarra>(PossibleLengths.OrderBy(x => x.Column1));
                    DesiredLengths = new List<Part>(DesiredLength.OrderBy(x => x.OriginalLength));
                }
                else if (i == 3)
                {
                    PossibleLengths = new List<propriedadedebarra>(PossibleLengths.OrderBy(x => x.Column1).Reverse());
                    DesiredLengths = new List<Part>(DesiredLength.OrderBy(x => x.OriginalLength));
                }
                else if (i == 4)
                {
                    PossibleLengths = new List<propriedadedebarra>(PossibleLengths.OrderBy(x => x.Column1));
                    DesiredLengths = new List<Part>(DesiredLength.OrderBy(x => x.OriginalLength).Reverse());
                    DesiredLengths = new List<Part>(DesiredLength.OrderBy(x => x.OriginalLength).Reverse());
                    DesiredLengths = new List<Part>(DesiredLength.OrderBy(x => x.OriginalLength).Reverse());
                }
                else if (i == 5)
                {
                    PossibleLengths = new List<propriedadedebarra>(PossibleLengths.OrderBy(x => x.Column1).Reverse());
                    DesiredLengths = new List<Part>(DesiredLength.OrderBy(x => x.OriginalLength).Reverse());
                }

                //curtar as peças retorna as barras com cortes.
                var planks = CalculateCuts(DesiredLengths);
                //armazenar o nest para fazer novo aproveitamento.
                if (planks != null)
                {
                    float somapercentagem = 0;
                    float somapercentagemsucata = 0;
                    float perfi = 0;
                    float desperdicio = 0;
                    float desperdicioSucata = 0;
                    foreach (var plank in planks)
                    {
                        try
                        {
                            perfi += plank.OriginalLength;
                            desperdicio += plank.FreeLength;
                            if (plank.FreeLength <= 2000f)
                            {
                                desperdicioSucata += plank.FreeLength;
                            }

                        }
                        catch (Exception)
                        {
                        }

                    }

                    if (planks.Count != 0)
                    {
                        somapercentagem = (desperdicio / perfi) * 100;
                        somapercentagemsucata = (desperdicioSucata / perfi) * 100;

                        aproveitamentos.Add(new ListaPlank(planks.OrderBy(x => x.FreeLength).ToList(), somapercentagem, planks.Last().FreeLength, somapercentagemsucata));
                    }
                }
            }

            try
            {

                var min = aproveitamentos.Min(t => t.Percentagem);
                List<ListaPlank> conjperfil = new List<ListaPlank>();
                foreach (ListaPlank ite in aproveitamentos.OrderBy(t => t.Percentagem).ToList())
                {
                    if (ite.Percentagem == min)
                    {
                        conjperfil.Add(ite);
                    }
                }
                var perfis = conjperfil.OrderBy(t => t.ultimaponta).Reverse().First();
                int X = 0;





                foreach (var plank in perfis)
                {
                    X++;
                    resultado.Rows.Add(X, plank.OriginalLength, plank.PlankMark, string.Join(",", plank.Cuts), plank.FreeLength, string.Join("#", plank.Cutpartmark), plank.PlankMark);
                    Percentagemdeaproveitamento = (100 - (perfis.Percentagem)).ToString("0.00") + " % ";
                    Percentagemdedesperdicio = perfis.Percentagem.ToString("0.00") + "%";
                    Percentagemsemsucata = (100 - (perfis.PercentagemSucata)).ToString("0.00") + " % ";
                    Percentagemdedesperdiciosemsucata = (perfis.PercentagemSucata).ToString("0.00") + " % ";
                }
                RELATORIO = perfis;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
            return resultado;
        }

        //Calcula a quantidade de desperdicio/comprimento livre deixado na lista de pranchas
        private static float GetFree(List<Plank> planks)
        {

            float free = 0;

            foreach (var plank in planks)
            {
                free += plank.FreeLength;
            }
            return free;
        }

        private static List<Plank> CalculateCuts(List<Part> desired)
        {
            var planks = new List<Plank>(); //Buffer list

            //passar por cortes
            foreach (Part i in desired)
            {
                bool repeat = true;
                while (repeat)
                {
                    //se não forem encontradas pranchas com comprimento disponivel
                    if (!planks.Any(plank => plank.FreeLength >= i.OriginalLength))
                    {
                        //fazer a prancha
                        try
                        {

                            float comp = PossibleLengths.First().Column1;
                            planks.Add(new Plank(comp, PossibleLengths.First().Column2));

                            bool primeirapecaencontrada = true;
                            for (int a = 0; a <= PossibleLengths.Count - 1; a++)
                            {
                                if (PossibleLengths[a].Column1 == comp && primeirapecaencontrada)
                                {
                                    PossibleLengths.RemoveAt(a);
                                    primeirapecaencontrada = false;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            return null;
                        }


                    }

                    //cortar quando possivel

                    foreach (var plank in planks.Where(plank => plank.FreeLength >= (i.OriginalLength + float.Parse(ESPESSURASERRA.ToString()))))
                    {
                        plank.Cut(i.OriginalLength + float.Parse(ESPESSURASERRA.ToString()));
                        plank.Cutpartmarks(i.Mark);
                        repeat = false;
                        break;
                    }
                    if (repeat)
                    {
                        planks.RemoveAt(planks.Count - 1);
                    }
                }
            }

            //reduzir os despedicio minimizando o comprimento da prancha
            foreach (var plank in planks)
            {
                float newLength = plank.OriginalLength;
                foreach (propriedadedebarra possibleLength in PossibleLengths)
                {
                    //possibleLength <= plank.OriginalLength && (plank.OriginalLength - float.Parse(plank.FreeLength.ToString())) <= possibleLength

                    if (possibleLength.Column1 <= plank.OriginalLength && (plank.OriginalLength - float.Parse(plank.FreeLength.ToString())) <= possibleLength.Column1)
                    {
                        newLength = possibleLength.Column1;

                    }
                }
                plank.OriginalLength = newLength;
            }
            PossibleLengths.Clear();
            return planks;
        }
        public class propriedadedebarra
        {
            // obviously you find meaningful names of the 2 properties
            public float columnsfloat;
            public string columnsstring;

            public float Column1 { get; set; }
            public string Column2 { get; set; }
        }

        //class classe para uma 'plank' genérica
        class Plank
        {
            public Plank(float length, string Mark)
            {
                OriginalLength = length;
                PlankMark = Mark;
            }

            public float FreeLength
            {
                get { return OriginalLength - Cuts.Sum(); }
            }

            public float OriginalLength;
            public string PlankMark;


            public List<float> Cuts = new List<float>();
            public List<string> Cutpartmark = new List<string>();
            public void Cut(float cutLength)
            {
                Cuts.Add(cutLength);
            }
            public void Cutpartmarks(string mark)
            {
                Cutpartmark.Add(mark);
            }
        }
        //class classe para uma 'ListaPlank' genérica que engloba a 'plank'
        class ListaPlank
        {
            public List<Plank> perfis = new List<Plank>();
            public float Percentagem;
            public float ultimaponta;
            public float PercentagemSucata;
            public ListaPlank(List<Plank> plank, float percentagem, float comprimento, float Percentagemsucata)
            {
                perfis = plank;
                Percentagem = percentagem;
                ultimaponta = comprimento;
                PercentagemSucata = Percentagemsucata;

            }
            public List<Plank> ListaPlanks
            {
                get { return perfis; }
            }
            public float PercentagemDesperdicio
            {
                get { return Percentagem; }
            }
            public float Percentagemsucata
            {
                get { return PercentagemSucata; }
            }
            public float Ultimaponta
            {
                get { return ultimaponta; }
            }
            private IEnumerable<Plank> Events()
            {
                foreach (var item in perfis)
                {
                    yield return item;
                }
            }
            public IEnumerator<Plank> GetEnumerator()
            {
                return Events().GetEnumerator();
            }
        }

        //class classe para uma 'part' genérica
        class Part
        {
            public Part(float length, string mark)
            {
                OriginalLength = length;
                Mark = mark;
            }

            public float FreeLength
            {
                get { return OriginalLength; }

            }
            public string FreeMark
            {
                get { return Mark; }

            }

            public float OriginalLength;
            public string Mark;
        }

    }

}

