using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using TSM = Tekla.Structures.Model;
using Tekla.Structures.Model;
using Tekla.Structures;
using System.Collections;
using System.Windows.Forms;
using TSD = Tekla.Structures.Drawing;
using System.Reflection;
using Tekla.Structures.Drawing;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using Microsoft.ReportingServices.Diagnostics.Internal;
using System.Web.UI.WebControls.WebParts;
using Tekla.Structures.Catalogs;

namespace TeklaArtigosOfeliz
{
    class ComunicaTekla
    {
        /// <summary>
        /// preenche um atributo de utilizador
        /// </summary>
        /// <param name="parts"></param>
        /// <param name="propriedade"></param>
        /// <param name="valor"></param>
        public static void EnviaproPriedadePeca(TSM.Part parts, string propriedade, string valor)
        {

            parts.SetUserProperty(propriedade, valor);

        }
        /// <summary>
        /// envia para o tekla recebe uma lista de peças o campo do objects.inp a preencher e o valor que quero prencher
        /// </summary>
        /// <param name="parts"></param>
        /// <param name="propriedade"></param>
        /// <param name="valor"></param>
        public static void EnviaproPriedadePeca(ArrayList parts, string propriedade, string valor)
        {
            foreach (TSM.Part part in parts)
            {
                part.SetUserProperty(propriedade, valor);
            }
        }
        /// <summary>
        /// envia para o tekla recebe uma lista de conjuntos o campo do objects.inp a preencher e o valor que quero prencher
        /// </summary>
        /// <param name="Assemblys"></param>
        /// <param name="propriedade"></param>
        /// <param name="valor"></param>
        public static void EnviaproPriedadeConj(ArrayList Assemblys, string propriedade, string valor)
        {
            foreach (TSM.Assembly assembly in Assemblys)
            {
                assembly.SetUserProperty(propriedade, valor);
            }
        }
        /// <summary>
        /// Faz a lista de peças dos conjuntos selecionados do modelo
        /// </summary>
        /// <returns></returns>
        public static ArrayList ListadePecasdoConjSelec()
        {

            ArrayList segParts = new ArrayList();
            TSM.Model model = new TSM.Model();
            TSM.ModelObjectEnumerator modelEnumerator = new TSM.UI.ModelObjectSelector().GetSelectedObjects();


            while (modelEnumerator.MoveNext())
            {
                TSM.Assembly ass = modelEnumerator.Current as TSM.Assembly;
                if (ass != null)
                {
                    foreach (var item in ass.GetSecondaries())
                    {
                        segParts.Add(item);
                    }
                    foreach (TSM.Assembly item in ass.GetSubAssemblies())
                    {
                        foreach (var item1 in item.GetSecondaries())
                        {
                            segParts.Add(item1);
                        }
                    }
                    segParts.Add(ass.GetMainPart());
                }
            }
            return segParts;
        }
        /// <summary>
        /// Faz a lista de peças selecionadas do modelo
        /// </summary>
        /// <returns>Pecas</returns>
        public static ArrayList ListadePecasSelec()
        {
            ArrayList PECAS = new ArrayList();
            TSM.Model model = new TSM.Model();
            TSM.ModelObjectEnumerator modelEnumerator = new TSM.UI.ModelObjectSelector().GetSelectedObjects();

            while (modelEnumerator.MoveNext())
            {
                TSM.Part part = modelEnumerator.Current as TSM.Part;
                if (part != null)
                {
                    PECAS.Add(part);
                }
            }
            return PECAS;
        }
        /// <summary>
        /// Faz a lista de peças selecionadas do modelo
        /// </summary>
        /// <returns>Pecas</returns>
        public static ArrayList ListadeconjdaspecasSelec()
        {
            ArrayList PECAS = new ArrayList();
            ArrayList PECASesp = new ArrayList();
            TSM.Model model = new TSM.Model();
            TSM.ModelObjectEnumerator modelEnumerator = new TSM.UI.ModelObjectSelector().GetSelectedObjects();

            while (modelEnumerator.MoveNext())
            {
                TSM.Part part = modelEnumerator.Current as TSM.Part;
                TSM.Assembly myAssembly = part.GetAssembly();

                if (myAssembly != null && !PECASesp.Contains(myAssembly.Identifier.GUID))
                {
                    PECASesp.Add(myAssembly.Identifier.GUID);
                    PECAS.Add(myAssembly);
                }
            }

            return PECAS;
        }
        /// <summary>
        /// Faz a lista de conjuntos selecionadas do modelo
        /// </summary>
        /// <returns>Conjuntos</returns>
        public static ArrayList ListadeConjuntosSelec()
        {
            ArrayList CONJ = new ArrayList();
            TSM.Model model = new TSM.Model();
            TSM.ModelObjectEnumerator modelEnumerator = new TSM.UI.ModelObjectSelector().GetSelectedObjects();


            while (modelEnumerator.MoveNext())
            {
                TSM.Assembly ass = modelEnumerator.Current as TSM.Assembly;
                if (ass != null)
                {
                    CONJ.Add(ass);
                }
            }
            return CONJ;
        }
        /// <summary>
        /// Preenche o campo artigo "chapa ou perfil"
        /// </summary>
        /// <param name="parts"></param>
        public static void Artigos(ArrayList parts)
        {
            foreach (TSM.Part part in parts)
            {
                string forcadopara = null;
                string material = null;
                part.GetUserProperty("forcar_destino", ref forcadopara);
                if (part.Name.Trim().ToLower() == "br")
                {
                    forcadopara = "CM";
                }

                string perfil = part.Profile.ProfileString.ToLower();
                part.GetReportProperty("MATERIAL", ref material);

                List<string> list = new List<string>();

                ComunicaBDtekla n = new ComunicaBDtekla();
                n.ConectarBD();
                list = n.Procurarbd("SELECT [Perfil] FROM [ArtigoTekla].[dbo].[Perfilagem3] where [Perfil]='" + perfil + "'");
                n.DesonectarBD();


                if (forcadopara == "CL" || forcadopara == "CQ" || forcadopara == "CP")
                {
                    part.SetUserProperty("Artigo", "Chapa");

                }
                else if (forcadopara == "CM")
                {
                    part.SetUserProperty("Artigo", "Perfil");

                }
                else if (perfil.Contains("h60x") )
                {
                    part.SetUserProperty("Artigo", "Chapa");
                }
                else if (perfil.Contains("pl") || perfil.Contains("gradildg") || perfil.Contains("gradilpl") || perfil.Contains("ca") || perfil.Contains("cha") || perfil.Contains("z") || perfil.Contains("c1") || perfil.Contains("c2") || perfil.Contains("c3") || perfil.Contains("p3") || perfil.Contains("p2") || perfil.Contains("p1") || perfil.Contains("p5") || perfil.Contains("p6") || perfil.Contains("p0") || perfil.Contains("h60") || perfil.Contains("chg") || perfil.Contains("saida") || perfil.Contains("omega") || perfil.Contains("VRS") || perfil.Contains("pc-") || perfil.Contains("pf-") || perfil.Contains("pw-") || perfil.Contains("bm") || perfil.Contains("wm") || perfil.Contains("max") || material.ToUpper() == "POLICARBONATO")
                {
                    part.SetUserProperty("Artigo", "Chapa");
                }
                                else if (list.Count > 0)
                {
                    part.SetUserProperty("Artigo", "Chapa");
                }
                else
                {
                    part.SetUserProperty("Artigo", "Perfil");
                }
            }
        }
        /// <summary>
        /// Preenche o destinatário externo da empresa travez da lista de peças
        /// </summary>
        /// <param name="parts"></param>
        /// <param name="incluilaser"></param>
        public static void Destinatarioexterno(ArrayList parts, bool incluilaser)
        {
            foreach (TSM.Part part in parts)
            {
                string pecaforcada = null; part.GetUserProperty("forcar_destino", ref pecaforcada);
                // SE TIVER FUROS ESTE VALOR VAI SER 1//
                int Temfuros = 0; part.GetReportProperty("HAS_HOLES", ref Temfuros);
                /////////////////////////////////////////////////////////////////////
                string Perfil = part.Profile.ProfileString.ToLower();
                string Material = part.Material.MaterialString.ToLower();
                string Artigo = null; part.GetUserProperty("Artigo", ref Artigo);

                double Espessura = 0;
                double Espessura1 = 0;
                double Espessura2 = 0;

                //////////////////////////////////////////////////////////////////////defenir espessura///////////////////////////////////////////////////////////////
                if (Perfil.Contains("x") || Perfil.Contains("*"))
                {
                    try
                    {
                        double.TryParse(part.Profile.ProfileString.ToLower().Split('x')[0].Split('*')[0].Replace("pl", "").Replace("cha", "").Replace("chg", "").Replace(".", ",").ToString(), out Espessura1);
                        double.TryParse(part.Profile.ProfileString.ToLower().Split('x')[1].Split('*')[1].Replace("pl", "").Replace("cha", "").Replace("chg", "").Replace(".", ",").ToString(), out Espessura2);
                    }
                    catch (Exception)
                    {


                    }
                    if (Espessura2 == 0)
                    {
                        Espessura2 = 5000;
                    }
                    if (Espessura1 == 0)
                    {
                        Espessura1 = 5000;
                    }

                    if (Espessura1 <= Espessura2)
                    {
                        Espessura = Espessura1;
                    }
                    else
                    {
                        Espessura = Espessura2;
                    }
                }
                else
                {
                    double.TryParse(part.Profile.ProfileString.ToLower().Replace("pl", "").Replace("cha", "").Replace("chg", "").Replace(".", ",").ToString().ToString(), out Espessura);
                }
                ///////////////////////////////////////////////////////////////////////


                if (pecaforcada != null)
                {
                    part.SetUserProperty("Destinata_ext", pecaforcada);
                }
                else if (Artigo == "Chapa")
                {
                    if (Perfil.Contains("ca"))
                    {
                        part.SetUserProperty("Destinata_ext", "CQ");
                    }
                    else if (Temfuros == 0 && (Material.Contains("gota") || Material.Contains("xadrez") || Material.Contains("zinco")))
                    {
                        part.SetUserProperty("Destinata_ext", "CQ");
                    }
                    else if ((Perfil.Contains("cha") || Perfil.Contains("pl")) && (Material.Contains("jr") || Material.Contains("j1") || Material.Contains("j2") || Material.Contains("j0") || Material.Contains("aisi")))
                    {

                        ///////////////////////////////////////////////////////CHAPAS CORTE TERMICO//////////////////////////////////////////////////////

                        part.SetUserProperty("Destinata_ext", "CL");


                        ////////////////////////////////////////////////////=============//////////////////////////////////////////////////////
                    }
                    else if (Perfil.Contains("fwm"))
                    {
                        List<string> ar = new List<string>();
                        ComunicaBDtekla a = new ComunicaBDtekla();
                        a.ConectarBD();
                        ar = a.Procurarbd("SELECT [dest] FROM [dbo].[Perfilagem3] WHERE [Perfil]='" + Perfil + "'");
                        a.DesonectarBD();
                        part.SetUserProperty("Destinata_ext", ar[0].ToString().Trim());
                    }
                    else if (Perfil.Contains("vrs") || Perfil.Contains("bm") || Perfil.Contains("nm") || Perfil.Contains("wm"))
                    {
                        part.SetUserProperty("Destinata_ext", "08");
                    }
                    else
                    {
                        try
                        {
                            if (Perfil.Contains("gradildg"))
                            {
                                Perfil = "GRADILDG";
                            }
                            else if (Perfil.Contains("gradilpl"))
                            {
                                Perfil = "GRADILPL";
                            }
                            List<string> ar = new List<string>();
                            ComunicaBDtekla a = new ComunicaBDtekla();
                            a.ConectarBD();
                            ar = a.Procurarbd("SELECT [dest] FROM [dbo].[Perfilagem3] WHERE [Perfil]='" + Perfil + "'");
                            a.DesonectarBD();
                            part.SetUserProperty("Destinata_ext", ar[0].ToString().Trim());
                        }
                        catch (Exception)
                        {
                            part.SetUserProperty("Destinata_ext", "CQ");

                        }
                    }
                    if (Perfil.Contains("gradilpl"))
                    {
                        Perfil = "GRADILPL";
                        List<string> ar = new List<string>();
                        ComunicaBDtekla a = new ComunicaBDtekla();
                        a.ConectarBD();
                        ar = a.Procurarbd("SELECT [dest] FROM [dbo].[Perfilagem3] WHERE [Perfil]='" + Perfil + "'");
                        a.DesonectarBD();
                        part.SetUserProperty("Destinata_ext", ar[0].ToString().Trim());
                    }
                }
                else
                {
                    part.SetUserProperty("Destinata_ext", "");
                }
            }
        }

        public static void DestinatarioexternoparaH60(ArrayList parts)
        {
            foreach (TSM.Part part in parts)
            {
                string Perfil = part.Profile.ProfileString.ToLower();
                string Artigo = null; part.GetUserProperty("Artigo", ref Artigo);
                string ArtigoInterno = null; part.GetUserProperty("Artigo_interno", ref ArtigoInterno);


                if (Perfil.Contains("h60"))
                {
                    part.SetUserProperty("Destinata_ext", "CP");
                    part.SetUserProperty("Artigo", "Chapa");

                    string consultaSQL = "SELECT [Artigo] FROM [dbo].[Perfilagem3] WHERE [Perfil] = '" + Perfil + "'";

                    ComunicaBDtekla a = new ComunicaBDtekla();
                    a.ConectarBD();
                    List<string> artigo = a.Procurarbd(consultaSQL);
                    a.DesonectarBD();

                    if (artigo.Count > 0)
                    {
                        part.SetUserProperty("Artigo_interno", artigo[0].ToString().Trim());
                    }
                    else
                    {
                        part.SetUserProperty("Artigo_interno", "");
                    }
                }

            }
        }


        /// <summary>
        /// Preenche as operacoes da empresa atravez da lista de peças e conjuntos
        /// </summary>
        /// <param name="parts"></param>
        /// <param name="Assemblys"></param>
        public static void operacoes(ArrayList parts, ArrayList Assemblys)
        {
            foreach (TSM.Part part in parts)
            {
                TSM.Assembly ass = part.GetAssembly();
                string Destinatario = null; part.GetUserProperty("Destinata_ext", ref Destinatario);
                string pecafurcada = null; part.GetUserProperty("forcar_destino", ref pecafurcada);
                string Departamento = null; ass.GetMainPart().GetUserProperty("Destinata_ext", ref Departamento);
                string Pintura = null; ass.GetReportProperty("USERDEFINED.pintura", ref Pintura);
                string OperacaoPintura = null; ass.GetReportProperty("USERDEFINED.OperacaoFabrica", ref OperacaoPintura);

                if (Pintura == "S\\PINTURA")
                {
                    Pintura = "";
                }


                if (!String.IsNullOrEmpty(Destinatario) || (((part.Profile.ProfileString.ToLower().Contains("PL") || part.Profile.ProfileString.ToLower().Contains("cha")) && (pecafurcada.ToLower() == "cm"))))
                {
                    if (Destinatario == "CP" || Destinatario == "DAP")
                    {


                        if (Pintura == "" && ass.GetSecondaries().Count == 0)
                        {
                            part.SetUserProperty("Operacoes", "Opção 8"); //para peças do cp ou dap
                        }
                        else
                        {
                            part.SetUserProperty("Operacoes", "Preparação de chapas");

                        }
                    }
                    else
                    {
                        part.SetUserProperty("Operacoes", "Preparação de chapas");
                    }

                }

                else
                {
                    int Temfuros = 0; part.GetReportProperty("HAS_HOLES", ref Temfuros);

                    if (Temfuros == 0)
                    {
                        part.SetUserProperty("Operacoes", "Corte");
                    }
                    else
                    {
                        part.SetUserProperty("Operacoes", "Corte e Furação");
                    }
                }
            }
            foreach (TSM.Assembly ass in Assemblys)
            {

                int NumeroObjetos = ass.GetSecondaries().Count;
                string Departamento = null; ass.GetMainPart().GetUserProperty("Destinata_ext", ref Departamento);
                string Pintura = null; ass.GetReportProperty("USERDEFINED.pintura", ref Pintura);
                string OperacaoPintura = null; ass.GetReportProperty("USERDEFINED.OperacaoFabrica", ref OperacaoPintura);

                if (Pintura == "S\\PINTURA")
                {
                    Pintura = "";
                }


                if (Departamento == "CP" || Departamento == "DAP")
                {

                    if (Pintura == "" && ass.GetSecondaries().Count == 0)
                    {
                        ass.SetUserProperty("Operacoes_Conj", "Opção 8");//para peças do cp ou dap
                    }
                    else
                    {
                        if (Pintura == "" && (OperacaoPintura == "" || OperacaoPintura == null))
                        {
                            ass.SetUserProperty("Operacoes_Conj", "Opção 2");
                        }
                        else if (OperacaoPintura == "DECAPADO E PINTADO")
                        {
                            ass.SetUserProperty("Operacoes_Conj", "Opção 3");//para peças soltas 
                        }
                        else if (OperacaoPintura == "DECAPADO")
                        {
                            ass.SetUserProperty("Operacoes_Conj", "Opção 4");//para peças soltas 
                        }
                        else if (string.IsNullOrEmpty(Pintura))
                        {
                            ass.SetUserProperty("Operacoes_Conj", "Opção 1");//para peças soltas 
                        }
                        else
                        {
                            ass.SetUserProperty("Operacoes_Conj", "Opção 15");//para peças soltas 
                        }
                    }



                }
                else if (Departamento == "08")
                {
                    ass.SetUserProperty("Operacoes_Conj", "Opção 9");//para peças de armazem 
                }
                else if (NumeroObjetos == 0 && string.IsNullOrEmpty(Pintura) && !string.IsNullOrEmpty(Departamento))
                {
                    ass.SetUserProperty("Operacoes_Conj", "Opção 8");//para peças que nao tem pintura da fase 500
                }
                else if (NumeroObjetos == 0)
                {
                    if (OperacaoPintura == "DECAPADO E PINTADO")
                    {
                        ass.SetUserProperty("Operacoes_Conj", "Opção 3");//para peças soltas 
                    }
                    else if (OperacaoPintura == "DECAPADO")
                    {
                        ass.SetUserProperty("Operacoes_Conj", "Opção 4");//para peças soltas 
                    }
                    else if (string.IsNullOrEmpty(Pintura))
                    {
                        ass.SetUserProperty("Operacoes_Conj", "Opção 1");//para peças soltas 
                    }
                    else
                    {
                        ass.SetUserProperty("Operacoes_Conj", "Opção 15");//para peças soltas 
                    }
                }
                else if (NumeroObjetos > 0)
                {
                    if (OperacaoPintura == "DECAPADO E PINTADO")
                    {
                        ass.SetUserProperty("Operacoes_Conj", "Opção 5");
                    }
                    else if (OperacaoPintura == "DECAPADO")
                    {
                        ass.SetUserProperty("Operacoes_Conj", "Opção 6");
                    }

                    else if (string.IsNullOrEmpty(Pintura))
                    {
                        ass.SetUserProperty("Operacoes_Conj", "Opção 2");
                    }
                    else
                    {
                        ass.SetUserProperty("Operacoes_Conj", "Opção 16");
                    }
                }
            }
        }
        /// <summary>
        /// Preenche o Artigo interno atravez da lista de peças
        /// </summary>
        /// <param name="parts"></param>
        public static void Artigo_interno(ArrayList parts)
        {
            foreach (TSM.Part part in parts)
            {
                string Destinatario = null; part.GetUserProperty("Destinata_ext", ref Destinatario);
                if (!String.IsNullOrEmpty(Destinatario))
                {
                    //////////////////////////////////variaveis nesseçarias/////////////////////////////////////////
                    //peças quinadas quinagem=1 calandrado = 0 se for verdadeiro//
                    int TemQuinagem = 0; part.GetReportProperty("IS_POLYBEAM", ref TemQuinagem);
                    int Calandrado = 0; part.GetReportProperty("CURVED_SEGMENTS", ref Calandrado);
                    int Temfuros = 0; part.GetReportProperty("HAS_HOLES", ref Temfuros);
                    double Largura = 0; part.GetReportProperty("HEIGHT", ref Largura);
                    double Comprimento = 0; part.GetReportProperty("LENGTH", ref Comprimento);
                    double Espessura = 0;
                    string Perfil = part.Profile.ProfileString.ToLower();
                    string Material = part.Material.MaterialString.ToLower();
                    string Chaparal = null; part.GetUserProperty("CHAPA_LACADA", ref Chaparal);
                    bool BOespessura = false;
                    bool dimlaser = false;
                    if (Perfil.Contains("omega") || Perfil.ToLower().Contains("ca"))
                    {
                        TemQuinagem = 1;
                    }


                    if (Perfil.Contains("x") || Perfil.Contains("*"))
                    {
                        BOespessura = double.TryParse(part.Profile.ProfileString.ToLower().Split('x')[0].Split('*')[0].Replace("pl", "").Replace("cha", "").Replace("chg", "").Replace(".", ",").ToString(), out Espessura);
                    }
                    else
                    {
                        BOespessura = double.TryParse(part.Profile.ProfileString.ToLower().Replace("pl", "").Replace("cha", "").Replace("chg", "").Replace(".", ",").ToString().ToString(), out Espessura);
                    }
                    if ((Largura <= 4000 && Comprimento <= 2000) || (Largura <= 2000 && Comprimento <= 4000))
                    {
                        dimlaser = true;
                    }

                    //////////////////////////////////////////////////////////////////////////////////////
                    if (Destinatario == "CL")
                    {
                        if ((dimlaser == true && Espessura <= 20) && (TemQuinagem == 1 || Calandrado == 0))
                        {
                            part.SetUserProperty("Artigo_interno", "LS0010005");
                        }
                        else if (dimlaser == true && Espessura <= 20)
                        {
                            part.SetUserProperty("Artigo_interno", "LS0010001");
                        }
                        else if ((Espessura <= 12) && (TemQuinagem == 1 || Calandrado == 0))
                        {
                            part.SetUserProperty("Artigo_interno", "LS0030003");
                        }
                        else if ((Espessura > 12) && (TemQuinagem == 1 || Calandrado == 0))
                        {
                            part.SetUserProperty("Artigo_interno", "LS0020003");
                        }
                        else if (Espessura <= 12)
                        {
                            part.SetUserProperty("Artigo_interno", "LS0030001");
                        }
                        else if (Espessura > 12)
                        {
                            part.SetUserProperty("Artigo_interno", "LS0020001");
                        }
                    }
                    else if (Destinatario == "DAP")
                    {
                        string especuradaschapas = "0";
                        part.GetUserProperty("Esp_chapa", ref especuradaschapas);
                        if (Perfil.Contains("gradildg"))
                        {
                            Perfil = "GRADILDG";
                        }
                        else if (Perfil.Contains("gradilpl"))
                        {
                            Perfil = "GRADILPL";
                        }

                        string Artigo = null;
                        string ArtigoInterno = null;

                        part.GetUserProperty("Artigo", ref Artigo);
                        part.GetUserProperty("Artigo_interno", ref ArtigoInterno);
                        part.SetUserProperty("Destinata_ext", "DAP");
                        part.SetUserProperty("Artigo", "Chapa");

                        string consultaSQL = "SELECT [Artigo] FROM [dbo].[Perfilagem3] WHERE [Perfil] = '" + Perfil + "'";

                        ComunicaBDtekla a = new ComunicaBDtekla();
                        a.ConectarBD();
                        List<string> artigo = a.Procurarbd(consultaSQL);
                        a.DesonectarBD();

                        if (artigo.Count > 0)
                        {
                            part.SetUserProperty("Artigo_interno", artigo[0].ToString().Trim());
                        }
                        else
                        {
                            part.SetUserProperty("Artigo_interno", "");
                        }

                    }
                    else if (Destinatario == "CP" || Destinatario == "DAP")
                    {
                        string especuradaschapas = "0";
                        part.GetUserProperty("Esp_chapa", ref especuradaschapas);
                        if (Perfil.Contains("gradildg"))
                        {
                            Perfil = "GRADILDG";
                        }
                        else if (Perfil.Contains("gradilpl"))
                        {
                            Perfil = "GRADILPL";
                        }
                        string material = null;
                        part.GetReportProperty("MATERIAL", ref material);

                        string materialBase = material;
                        string sufixo = null;
                        string consultaSQL = null;

                        if (material.Contains(".2,2"))
                        {
                            sufixo = ".2,2";
                            materialBase = material.Substring(0, material.Length - 4);
                        }
                        else if (material.Contains(".3,1"))
                        {
                            sufixo = ".3,1";
                            materialBase = material.Substring(0, material.Length - 4);

                        }                                                                     
                        
                        materialBase = Regex.Replace(materialBase, @"[^a-zA-Z0-9]", "").Trim();

                       
                        // Para Perfil 'P' 'Madres C e Z' e 'SUPEROMEGA'
                        if (Perfil.Contains("p0") || Perfil.Contains("p1") || Perfil.Contains("p2") || Perfil.Contains("p3") ||
                                 Perfil.Contains("p4") || Perfil.Contains("P5") || Perfil.Contains("p6") || Perfil.Contains("z") ||
                                    Perfil.Contains("c"))
                        {                            
                            consultaSQL = @"SELECT [Artigo] FROM [dbo].[Perfilagem3] WHERE [Material] = '" + materialBase + "' AND [Perfil] = '" + Perfil + "'";
                        }

                        else if (Perfil.ToUpper().StartsWith("SUPEROMEGA")) // Para Os SuperOmegas
                        {                           

                            //MessageBox.Show($"Para o Perfil: {Perfil} e o {materialBase}, os Materiais só podem ser \n\"S280GD Z200\" \n\"S350GD Z200\" \n\"S350GD Z275\" \n\"S350GD ZM310\"");
                           
                            consultaSQL = @"SELECT [Artigo] FROM [dbo].[Perfilagem3] WHERE [Material] = '" + materialBase + "' AND [Perfil] = '" + Perfil + "'";
                        }

                        // Para Perfil 'SAIDAS'
                        else if (Perfil.Contains("saida") && (
                          material.Contains("1.4301") ||
                          material.Contains("1.4307") ||
                          material.Contains("1.4401") ||
                          material.Contains("1.4404") ||
                          material.Contains("1.4845")))
                        {
                                Perfil = Perfil.ToUpper();
                                string materialInox = "inox";
                                consultaSQL = @"SELECT [Artigo] FROM [dbo].[Perfilagem3] WHERE [Material] = '" + materialInox + "' AND [Perfil] = '" + Perfil + "'";
                        }                      

                        else // Para o resto dos perfis do CP 
                        {
                            consultaSQL = "SELECT [Artigo] FROM [dbo].[Perfilagem3] WHERE [Perfil] = '" + Perfil + "'";
                        }

                        ComunicaBDtekla a = new ComunicaBDtekla();
                        a.ConectarBD();
                        List<string> artigo = a.Procurarbd(consultaSQL);
                        a.DesonectarBD();

                        if (artigo.Count > 0)
                        {
                            part.SetUserProperty("Artigo_interno", artigo[0].ToString().Trim());
                        }
                        else
                        {
                            part.SetUserProperty("Artigo_interno", "");
                        }                                            

                    }                  
                    else if (Destinatario == "CQ")
                    {
                        if (Espessura <= 1)
                        {
                            if (TemQuinagem == 1)
                            {
                                ////////////CQ COM QUINAGEM //////////////////
                                if ((Material.Contains("gd") || Material.Contains("dx")) && string.IsNullOrEmpty(Chaparal))//galvanizada
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0017");
                                }
                                else if ((Material.Contains("gd") || Material.Contains("dx")) && !string.IsNullOrEmpty(Chaparal))//prepintada
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0032");
                                }
                                ////////////////////////////////////
                                else if (Material.Contains("+ze"))//zincor
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0044");
                                }
                                else if (Material.Contains("zm"))//magnelis
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0029");
                                }
                                else if (Material.Contains("j0") || Material.Contains("j1") || Material.Contains("j2") || Material.Contains("jr"))//laq
                                {
                                    if (Material.Contains("GOTA"))
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0048");
                                    }
                                    else if (Material.Contains("XADREZ"))
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0050");
                                    }
                                    else
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0014");
                                    }
                                }
                                else if (Material.Contains("wp"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0046");
                                }
                                /////////////////////////////////////
                                else if (Material.Contains("aisi"))//inox
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0023");
                                }
                                else if (Material.ToUpper().Contains("ALUMINIUM"))//aluminio
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0003");
                                }
                                else if (Material.ToUpper().Contains("ZINCO"))//ZINCO
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0038");
                                }
                            }
                            else if (Calandrado == 0)
                            {
                                ////////////CQ calandrado //////////////////
                                if ((Material.Contains("gd") || Material.Contains("dx")) && string.IsNullOrEmpty(Chaparal))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0015");
                                }
                                else if ((Material.Contains("gd") || Material.Contains("dx")) && !string.IsNullOrEmpty(Chaparal))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0030");
                                }
                                ///////////////////////////////////////////
                                else if (Material.Contains("+ze"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0042");
                                }
                                else if (Material.Contains("zm"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0027");
                                }
                                else if (Material.Contains("wp"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0046");
                                }
                                else if (Material.Contains("j0") || Material.Contains("j1") || Material.Contains("j2") || Material.Contains("jr"))//laq
                                {
                                    if (Material.Contains("GOTA"))
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0048");
                                    }
                                    else if (Material.Contains("XADREZ"))
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0050");
                                    }
                                    else
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0012");
                                    }
                                }
                                //////////////////////////////////////
                                else if (Material.Contains("aisi"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0021");
                                }
                                else if (Material.ToUpper().Contains("ALUMINIUM"))//aluminio
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0001");
                                }
                                else if (Material.ToUpper().Contains("ZINCO"))//ZINCO
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0036");
                                }
                            }
                            else
                            {
                                ////////////CQ SEM QUINAGEM //////////////////
                                if ((Material.Contains("gd") || Material.Contains("dx")) && string.IsNullOrEmpty(Chaparal))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0016");
                                }
                                else if ((Material.Contains("gd") || Material.Contains("dx")) && !string.IsNullOrEmpty(Chaparal))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0031");
                                }
                                //////////////////////
                                else if (Material.Contains("+ze"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0043");
                                }
                                else if (Material.Contains("zm"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0028");
                                }
                                else if (Material.Contains("wp"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0045");
                                }
                                else if (Material.Contains("j0") || Material.Contains("j1") || Material.Contains("j2") || Material.Contains("jr"))//laq
                                {
                                    if (Material.Contains("GOTA"))
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0047");
                                    }
                                    else if (Material.Contains("XADREZ"))
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0049");
                                    }
                                    else
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0013");
                                    }
                                }
                                /////////////////////////
                                else if (Material.Contains("aisi"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0022");
                                }
                                else if (Material.ToUpper().Contains("ALUMINIUM"))//aluminio
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0002");
                                }
                                else if (Material.ToUpper().Contains("ZINCO"))//ZINCO
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0037");
                                }
                            }
                        }
                        else
                        {

                            if (TemQuinagem == 1)
                            {
                                ////////////CQ COM QUINAGEM //////////////////
                                if ((Material.Contains("gd") || Material.Contains("dx")) && string.IsNullOrEmpty(Chaparal))//galvanizada
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0020");
                                }
                                else if ((Material.Contains("gd") || Material.Contains("dx")) && !string.IsNullOrEmpty(Chaparal))//prepintada
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0035");
                                }
                                ////////////////////////////////////
                                else if (Material.Contains("+ze"))//zincor
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0044");
                                }
                                else if (Material.Contains("zm"))//magnelis
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0029");
                                }
                                else if (Material.Contains("wp"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0046");
                                }
                                else if (Material.Contains("j0") || Material.Contains("j1") || Material.Contains("j2") || Material.Contains("jr"))//laq
                                {
                                    if (Material.Contains("gota"))
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0048");
                                    }
                                    else if (Material.Contains("xadrez"))
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0050");
                                    }
                                    else
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0014");
                                    }
                                }
                                /////////////////////////////////////
                                else if (Material.Contains("aisi"))//inox
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0026");
                                }
                                else if (Material.ToUpper().Contains("ALUMINIUM"))//aluminio
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0006");
                                }
                                else if (Material.ToUpper().Contains("ZINCO"))//ZINCO
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0041");
                                }
                            }
                            else if (Calandrado == 0)
                            {
                                ////////////CQ calandrado //////////////////
                                if ((Material.Contains("gd") || Material.Contains("dx")) && string.IsNullOrEmpty(Chaparal))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0018");
                                }
                                else if ((Material.Contains("gd") || Material.Contains("dx")) && !string.IsNullOrEmpty(Chaparal))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0033");
                                }
                                ///////////////////////////////////////////
                                else if (Material.Contains("+ze"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0042");
                                }
                                else if (Material.Contains("zm"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0027");
                                }
                                else if (Material.Contains("wp"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0046");
                                }
                                else if (Material.Contains("j0") || Material.Contains("j1") || Material.Contains("j2") || Material.Contains("jr"))//laq
                                {
                                    if (Material.Contains("gota"))
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0048");
                                    }
                                    else if (Material.Contains("xadrez"))
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0050");
                                    }
                                    else
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0012");
                                    }
                                }
                                //////////////////////////////////////
                                else if (Material.Contains("aisi"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0024");
                                }
                                else if (Material.ToUpper().Contains("ALUMINIUM"))//aluminio
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0004");
                                }
                                else if (Material.ToUpper().Contains("ZINCO"))//ZINCO
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0039");
                                }
                            }
                            else
                            {
                                ////////////CQ SEM QUINAGEM //////////////////
                                if ((Material.Contains("gd") || Material.Contains("dx")) && string.IsNullOrEmpty(Chaparal))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0019");
                                }
                                else if ((Material.Contains("gd") || Material.Contains("dx")) && !string.IsNullOrEmpty(Chaparal))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0034");
                                }
                                //////////////////////
                                else if (Material.Contains("+ze"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0043");
                                }
                                else if (Material.Contains("zm"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0028");
                                }
                                else if (Material.Contains("wp"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0045");
                                }
                                else if (Material.Contains("wp") || Material.Contains("j0") || Material.Contains("j1") || Material.Contains("j2") || Material.Contains("jr"))//laq
                                {
                                    if (Material.Contains("gota"))
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0047");
                                    }
                                    else if (Material.Contains("xadrez"))
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0049");
                                    }
                                    else
                                    {
                                        part.SetUserProperty("Artigo_interno", "CMCQ0013");
                                    }
                                }
                                /////////////////////////
                                else if (Material.Contains("aisi"))
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0025");
                                }
                                else if (Material.ToUpper().Contains("ALUMINIUM"))//aluminio
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0005");
                                }
                                else if (Material.ToUpper().Contains("ZINCO"))//ZINCO
                                {
                                    part.SetUserProperty("Artigo_interno", "CMCQ0040");
                                }
                            }
                        }
                    }
                    if ((Destinatario == "CM"))
                    {
                        if (TemQuinagem == 1 || Calandrado == 0)
                        {
                            part.SetUserProperty("Artigo_interno", "CQ0020005");
                        }
                        else
                        {
                            part.SetUserProperty("Artigo_interno", "CM0050001");
                        }
                    }
                }
                else
                {
                    part.SetUserProperty("Artigo_interno", "");
                }

            }
            
        }  

        /// <summary>
        /// altera o prefixo de uma lista de peças 
        /// </summary>
        /// <param name="peças"></param>
        /// <param name="fase"></param>
        /// <param name="fase1000"></param>
        public static void AlteraPrefixo(ArrayList peças, int fase, int fase1000 = 1000)
        {
            foreach (TSM.Part part in peças)
            {
                if (part != null)
                {
                    if (part.Profile.ProfileString.Contains("CHA") || part.Profile.ProfileString.Contains("PL"))
                    {
                        part.PartNumber.StartNumber = 1;
                        part.PartNumber.Prefix = fase + "C";
                        part.AssemblyNumber.StartNumber = 1;
                        part.AssemblyNumber.Prefix = fase + "CJ";
                        part.Modify();
                    }
                    else
                    {

                        if (part.Material.MaterialString.Contains("8,8") || part.Material.MaterialString.Contains("5,8") || part.Material.MaterialString.Contains("10,9") || part.Profile.ProfileString.Contains("NUT_M"))
                        {
                            part.PartNumber.StartNumber = 1;
                            part.PartNumber.Prefix = fase1000 + "H";
                            part.AssemblyNumber.StartNumber = 1;
                            part.AssemblyNumber.Prefix = fase1000 + "H";
                            part.Modify();
                        }
                        else
                        {
                            part.PartNumber.StartNumber = 1;
                            part.PartNumber.Prefix = fase + "P";
                            part.AssemblyNumber.StartNumber = 1;
                            part.AssemblyNumber.Prefix = fase + "CJ";
                            part.Modify();
                        }
                    }
                }
            }
        }
        /// <summary>
        /// altera prefixo de uma peça só
        /// </summary>
        /// <param name="part"></param>
        /// <param name="fase"></param>
        /// <param name="fase1000"></param>
        public static void AlteraPrefixo(TSM.Part part, int fase, int fase1000 = 1000)
        {
            if (part != null)
            {
                if (part.Profile.ProfileString.Contains("CHA") || part.Profile.ProfileString.Contains("PL"))
                {
                    part.PartNumber.StartNumber = 1;
                    part.PartNumber.Prefix = fase + "C";
                    part.AssemblyNumber.StartNumber = 1;
                    part.AssemblyNumber.Prefix = fase + "CJ";
                    part.Modify();
                }
                else
                {

                    if (part.Material.MaterialString.Contains("8,8") || part.Material.MaterialString.Contains("5,8") || part.Material.MaterialString.Contains("10,9") || part.Profile.ProfileString.Contains("NUT_M"))
                    {
                        part.PartNumber.StartNumber = 1;
                        part.PartNumber.Prefix = fase1000 + "H";
                        part.AssemblyNumber.StartNumber = 1;
                        part.AssemblyNumber.Prefix = fase1000 + "H";
                        part.Modify();
                    }
                    else
                    {
                        part.PartNumber.StartNumber = 1;
                        part.PartNumber.Prefix = fase + "P";
                        part.AssemblyNumber.StartNumber = 1;
                        part.AssemblyNumber.Prefix = fase + "CJ";
                        part.Modify();
                    }
                }
            }
        }
        /// <summary>
        /// seleciona no modelo as peças ou conjuntos 
        /// </summary>
        /// <param name="partList"></param>
        public static void selectinmodel(ArrayList primaryList)
        {
            TSM.UI.ModelObjectSelector mos = new TSM.UI.ModelObjectSelector();
            mos.Select(primaryList);
        }
        /// <summary>
        /// verifica se as operaçoes sao de soldadura "Opção 2", "Opção 5", "Opção 6", "Opção 16" 
        /// </summary>
        /// <param name="conjunto"></param>
        /// <returns></returns>
        public static bool TemSolda(TSM.Assembly conjunto)
        {
            bool valida = false;
            string mysr=null;
            conjunto.GetUserProperty("Operacoes_Conj", ref mysr);
            string[] teste = { "Opção 2", "Opção 5", "Opção 6", "Opção 16"};

            if (teste.Contains(mysr))
            {
                valida = true;
            }

            return valida;
        }
        /// <summary>
        /// imprime a partir do tekla 
        /// </summary>
        /// <param name="conjuntos"></param>
        /// <param name="pecas"></param>
        /// <param name="labelrelatorio"></param>
        public static void imprimepdf(ArrayList conjuntos, ArrayList pecas,Label desenho = null)
        {
            string mark=null;
            List<int> lAssemblyDrgIDs = new List<int>();
            List<string> marcaconj = new List<string>();
        
            foreach (TSM.Part item in pecas)
            {
                
                item.GetReportProperty("PART_POS", ref mark);
                if (item != null)
                {
                    marcaconj.Add(mark); //Create a distinct list of parts drawing ID's
                }
            }
            foreach (TSM.Assembly item in conjuntos)
            {
                if (item != null)
                {
                  
                    item.GetReportProperty("ASSEMBLY_POS", ref mark);
                    if (TemSolda(item))
                    {
                        marcaconj.Add(mark + " - 1");
                    }
                   
                    marcaconj.Add(mark);//Create a distinct list of Assembly drawing ID's
                }
            }

            TSD.PrintAttributes pa = new TSD.PrintAttributes(); //Use the default printer values
            pa.PrinterInstance = "PDF";
            TSD.DrawingHandler Handler = new TSD.DrawingHandler();

            desenho.Text = "A obeter desenhos";

            List<TSD.Drawing> a = GETDRAWINGS();

            string erro = "NÃO FOI POSSIVEL IMPRIMIR OS DESENHOS :" + Environment.NewLine;
            List<string> c = marcaconj.Distinct().ToList();

            foreach (string item in c)
            {
                desenho.Text = "A imprimir o desenho " + item;
                TSD.Drawing IMP = a.AsEnumerable().FirstOrDefault(r => r.Mark.Replace("[","").Replace("]", "").Replace(".", "") == item);
                try
                {
                    Handler.PrintDrawing(IMP, pa);
                }
                catch (Exception)
                {
                    erro +=item + Environment.NewLine;             
                }
            
            }
            if (erro != "NÃO FOI POSSIVEL IMPRIMIR OS DESENHOS :" + Environment.NewLine)
            {
                MessageBox.Show(erro, "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// get drawings 
        /// </summary>
        /// <returns></returns>
        public static List<TSD.Drawing> GETDRAWINGS()
        {
            TSD.DrawingHandler Handler = new TSD.DrawingHandler();
            TSD.DrawingEnumerator dEnum = Handler.GetDrawings();
            dEnum.SelectInstances = false;
            List<TSD.Drawing> a = new List<Drawing>();
            while (dEnum.MoveNext())
                a.Add(dEnum.Current as TSD.Drawing);

            return a;
        }
        /// <summary>
        /// abrir desenho
        /// </summary>
        /// <param name="DESENHO"></param>
        public static void OPENDRAWING(Drawing DESENHO)
        {
            TSD.DrawingHandler Handler = new TSD.DrawingHandler();
            Handler.SetActiveDrawing(DESENHO,true);
        }
        /// <summary>
        /// criar desenho no tekla 
        /// </summary>
        /// <param name="pecas"></param>
        /// <returns></returns>
        public bool CriaDesenhos(ArrayList pecas)
        {
            List<Identifier> partIDList = new List<Identifier>();
            Identifier part = null;
            foreach (TSM.Part item in pecas)
            {
                if (item.Profile.ProfileString.Contains("NM") || item.Profile.ProfileString.Contains("WM") || item.Profile.ProfileString.Contains("VRSM") || item.Profile.ProfileString.Contains("NUT_M"))
                {

                }
                else
                {
                    part = item.Identifier;
                    if (part != null)
                    {
                        partIDList.Add(part);
                    }
                }
            }

            var singleRule = new TSD.Automation.AutoDrawingRule(@"rui.dproc");//AssemblyDrawings.dproc
            TSD.Automation.AutoDrawingsStatusEnum singleOperationStatus;

            //gurantee drawing closed
            new TSD.DrawingHandler().CloseActiveDrawing();

            //call the draw code
            bool drawingsGenerated = false;
            drawingsGenerated = TSD.Automation.DrawingCreator.CreateDrawings(singleRule, partIDList, out singleOperationStatus);

            return drawingsGenerated;
        }
        /// <summary>
        /// imprime em pdf com a dpmprinter
        /// </summary>
        /// <param name="conjuntos"></param>
        /// <param name="pecas"></param>
        /// <param name="numeroobra"></param>
        /// <param name="lblestado"></param>
        public static void exportDrawings(ArrayList conjuntos, ArrayList pecas,string numeroobra,Frm_DesenhosFerramentas lblestado)
        {

            if (conjuntos.Count>0||pecas.Count>0)
            {

            string InstallFolder = Tekla.Structures.Dialog.StructuresInstallation.InstallFolder;
            var printerLocation = InstallFolder+@"nt\bin\applications\Tekla\Model\DPMPrinter\DPMPrinterCommand.exe";

            Model mdl = new Model();
                if (mdl.GetConnectionStatus())
                {
                    int drgID;
                    List<int> lAssemblyDrgIDs = new List<int>();
                    List<int> lPartDrgIDs = new List<int>();
                  

                    foreach (TSM.Assembly item in conjuntos)
                    {
                        drgID = 0;
                        item.GetReportProperty("DRAWING.ID", ref drgID);
                        if (drgID > 0 && lAssemblyDrgIDs.Contains(drgID) == false)
                        {
                            lAssemblyDrgIDs.Add(drgID); //Create a distinct list of Assembly drawing ID's
                        }
                    }
                    foreach (TSM.Part item in pecas)
                    {
                        drgID = 0;
                        item.GetReportProperty("DRAWING.ID", ref drgID);
                        if (drgID > 0 && lAssemblyDrgIDs.Contains(drgID) == false)
                        {
                            lPartDrgIDs.Add(drgID); //Create a distinct list of parts drawing ID's
                        }
                    }
                
                    PrintAttributes pa = new PrintAttributes();
                    DrawingHandler Handler = new DrawingHandler();
                    DrawingEnumerator dEnum = Handler.GetDrawings();

                    dEnum.MoveNext();
                    Drawing drg = dEnum.Current;
                    PropertyInfo drgInfo = drg.GetType().GetProperty("Identifier", BindingFlags.Instance | BindingFlags.NonPublic);
                    dEnum.Reset();
                    int conjuntosencontrados=0;
                    int pecasencontradas=0;
                    int i = 0;
                    System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                    startInfo.FileName = printerLocation;

                    

                    while (dEnum.MoveNext())
                    {

                        lblestado.LBLestado.Text = "Total desenhos - "+ dEnum.GetSize().ToString()+" a ver " + (i += 1).ToString() +Environment.NewLine+ " imp-"+ conjuntosencontrados +"- de"+(lAssemblyDrgIDs.Count)+" Conj||"+ pecasencontradas + "- de" + lPartDrgIDs.Count+" Peças";





                        if (lAssemblyDrgIDs.Count<=conjuntosencontrados&&lPartDrgIDs.Count <= pecasencontradas)
                        {
                            break;
                        }



                        drg = dEnum.Current;
                        Identifier id = (drgInfo.GetValue(drg, null)) as Identifier;
                        if (drg is AssemblyDrawing)
                        {
                            if (lAssemblyDrgIDs.Contains(id.ID))
                            {
                                string dwgDir = "c:\\r\\";

                                Directory.CreateDirectory(dwgDir);

                                try
                                {
                               
                                    conjuntosencontrados += 1;
                                    var test = Handler.SetActiveDrawing(drg, false);
                                    string revMark = null;
                                    drg.GetUserProperty("REVISION_MARK", ref revMark);
                                    var drawingName = drg.Mark.Replace("[", "").Replace(".", "").Replace("]", "").Trim().Replace("_1", "");
                                    var drawingNameNoWhiteSpace = "2." + numeroobra + "." + drawingName.Replace(" ", "") + ".pdf";
                                    var fullpath = Path.Combine(dwgDir, drawingNameNoWhiteSpace);
                                    var arg = string.Format("settingsFile:\"X:\\Tekla_configuracoes\\TS\\PdfPrintOptions.xml\" printActive:true printer:pdf paper:Tabloid:false orientation:Auto out:{0}", "\"" + fullpath + "\"");
                                    startInfo.Arguments = arg;
                                    System.Diagnostics.Process.Start(startInfo).WaitForExit();
                               
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.ToString());
                                }
                            }
                        }
                        else if (drg is TSD.SinglePartDrawing)
                        {

                            if (lPartDrgIDs.Contains(id.ID))
                            {

                                string dwgDir = "c:\\r\\";

                                Directory.CreateDirectory(dwgDir);

                                try
                                {
                                   
                                    pecasencontradas += 1;
                                    var test = Handler.SetActiveDrawing(drg, false);
                                    string revMark = null;
                                    drg.GetUserProperty("REVISION_MARK", ref revMark);
                                    var drawingName = drg.Mark.Replace("[", "").Replace(".", "").Replace("]", "").Trim().Replace("_1", "");
                                    var drawingNameNoWhiteSpace = "2." + numeroobra + "." + drawingName.Replace(" ", "") + ".pdf";
                                    var fullpath = Path.Combine(dwgDir, drawingNameNoWhiteSpace);
                                    var arg = string.Format("settingsFile:\"X:\\Tekla_configuracoes\\TS\\PdfPrintOptions.xml\" printActive:true printer:pdf paper:Tabloid:false orientation:Auto out:{0}", "\"" + fullpath + "\"");
                                    startInfo.Arguments = arg;
                                     System.Diagnostics.Process.Start(startInfo).WaitForExit();
                                  
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.ToString());
                                }

                            }
                        }
                        Handler.CloseActiveDrawing();
                    }
                }
            }
        }

        internal string ObterNomeDaObra()
        {
            throw new NotImplementedException();
        }
    }
}

