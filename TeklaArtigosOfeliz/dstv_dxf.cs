using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace TeklaArtigosOfeliz
{
    class dstv_dxf
    {
      

        public static void CRIAR(List<string> paths)
        {
            bool TESTE = true;

            foreach (string path in paths)
            {
                if (path != null)
                {
                   string file = path.ToLower().Replace(".nc1", "").Replace(".nc", "").ToUpper() + ".dxf";
                   string[] lines = File.ReadAllLines(path, Encoding.Default);
           
                    double MAXX = 0;
                    double MAXY = 0;
                    List<Ponto> pontos = new List<Ponto>();
                    bool R = false;
                    int w = lines.Length;
                    if (lines.Contains("AK"))
                    {
                        CRIARCABEÇALHO(file);
                        //MODULO AK CONTORNO EXTERIOR
                        for (int i = 0; i < w; i++)
                        {

                            if (lines.Contains("AK"))
                            {

                                if (R == true)
                                {

                                    try
                                    {
                                        string[] valoresiniciais = lines[i].Split(' ').Where(x => !string.IsNullOrEmpty(x)).ToArray();
                                        string[] valoresfinais = lines[i + 1].Split(' ').Where(x => !string.IsNullOrEmpty(x)).ToArray();
                                        double A;
                                        int B = 0;
                                        int C = 0;
                                        double d = 0;
                                        if (!double.TryParse(valoresiniciais[0].ToString().Replace(".", ","), out A))
                                        {
                                            B = 1;

                                        }
                                        if (!double.TryParse(valoresfinais[0].ToString().Replace(".", ","), out A))
                                        {
                                            C = 1;

                                        }

                                        double XINICIAL = double.Parse(Regex.Replace(valoresiniciais[B].ToString(), "[^0-9.]", "").Replace(".", ","));
                                        double YINICIAL = double.Parse(Regex.Replace(valoresiniciais[B + 1].ToString(), "[^0-9.]", "").Replace(".", ","));
                                        double RAIO = double.Parse(Regex.Replace(valoresiniciais[B + 2].ToString(), @"^-?\d+$", "").Replace(".", ","));
                                        double XFINAL = double.Parse(Regex.Replace(valoresfinais[C].ToString(), "[^0-9.]", "").Replace(".", ","));
                                        double YFINAL = double.Parse(Regex.Replace(valoresfinais[C + 1].ToString(), "[^0-9.]", "").Replace(".", ","));


                                        d = calculoBulge(XINICIAL, YINICIAL, XFINAL, YFINAL, RAIO);

                                        pontos.Add(new Ponto(XINICIAL, YINICIAL, d));


                                        if (MAXX < XINICIAL)
                                        {
                                            MAXX = XINICIAL;
                                        }
                                        if (MAXY < YINICIAL)
                                        {
                                            MAXY = YINICIAL;
                                        }

                                    }
                                    catch (Exception)
                                    {
                                        R = false;
                                    }
                                }
                                if (lines[i].Contains("AK"))
                                {

                                    if (lines[i + 1].Contains("v"))
                                    {
                                        R = true;
                                    }
                                    else if (lines[i].Contains("AK"))
                                    {
                                        R = false;
                                    }

                                }
                            }
                        }
                        if (pontos != null)
                        {
                            CRIARPOLYLINHA(pontos, 0, file);
                        }

                        pontos.Clear();
                        //MODULO AK CONTORNO EXTERIOR
                        for (int i = 0; i < w; i++)
                        {

                            if (lines.Contains("IK"))
                            {

                                if (R == true)
                                {

                                    try
                                    {
                                        string[] valoresiniciais = lines[i].Split(' ').Where(x => !string.IsNullOrEmpty(x)).ToArray();
                                        string[] valoresfinais = lines[i + 1].Split(' ').Where(x => !string.IsNullOrEmpty(x)).ToArray();
                                        double A;
                                        int B = 0;
                                        int C = 0;
                                        double d = 0;
                                        if (!double.TryParse(valoresiniciais[0].ToString().Replace(".", ","), out A))
                                        {
                                            B = 1;

                                        }
                                        if (!double.TryParse(valoresfinais[0].ToString().Replace(".", ","), out A))
                                        {
                                            C = 1;

                                        }

                                        double XINICIAL = double.Parse(Regex.Replace(valoresiniciais[B].ToString(), "[^0-9.]", "").Replace(".", ","));
                                        double YINICIAL = double.Parse(Regex.Replace(valoresiniciais[B + 1].ToString(), "[^0-9.]", "").Replace(".", ","));
                                        double RAIO = double.Parse(Regex.Replace(valoresiniciais[B + 2].ToString(), @"^-?\d+$", "").Replace(".", ","));
                                        double XFINAL = double.Parse(Regex.Replace(valoresfinais[C].ToString(), "[^0-9.]", "").Replace(".", ","));
                                        double YFINAL = double.Parse(Regex.Replace(valoresfinais[C + 1].ToString(), "[^0-9.]", "").Replace(".", ","));


                                        d = calculoBulge(XINICIAL, YINICIAL, XFINAL, YFINAL, RAIO);

                                        pontos.Add(new Ponto(XINICIAL, YINICIAL, d));


                                    }
                                    catch (Exception)
                                    {

                                        R = false;
                                        if (pontos != null)
                                        {
                                            CRIARPOLYLINHA(pontos, 0, file);
                                            pontos.Clear();
                                        }
                                    }
                                }
                                if (lines[i].Contains("IK"))
                                {
                                    if (lines[i + 1].Contains("v"))
                                    {
                                        R = true;
                                    }
                                    else
                                    {
                                        R = false;
                                    }

                                }
                            }
                        }

                        R = false;
                        //MODULO DE FUROS
                        for (int i = 0; i < w; i++)
                        {
                            if (lines.Contains("BO"))
                            {

                                if (R == true)
                                {
                                    try
                                    {
                                        string[] valoresiniciais = lines[i].Split(' ').Where(x => !string.IsNullOrEmpty(x)).ToArray();
                                        bool contersunk = false;
                                        double A;
                                        int B = 0;
                                        double XINICIAL = 0;
                                        double YINICIAL = 0;
                                        double DIAMETRO = 0;
                                        double DISTX = 0;
                                        double DISTY = 0;
                                        Double ANG = 0;
                                        if (!double.TryParse(valoresiniciais[0].ToString().Replace(".", ","), out A))
                                        {
                                            B = 1;
                                        }

                                        if (valoresiniciais.Length == 8)
                                        {
                                            if (valoresiniciais[7].ToString() == "90.00" && valoresiniciais[6].ToString() == "0.00" && valoresiniciais[6].ToString() == "0.00")
                                            {
                                                contersunk = true;
                                            }
                                        }



                                        if (valoresiniciais.Length == 4 || valoresiniciais.Length == 5)
                                        {
                                            XINICIAL = double.Parse(Regex.Replace(valoresiniciais[B].ToString(), "[^0-9.]", "").Replace(".", ","));
                                            YINICIAL = double.Parse(Regex.Replace(valoresiniciais[B + 1].ToString(), "[^0-9.]", "").Replace(".", ","));
                                            DIAMETRO = double.Parse(Regex.Replace(valoresiniciais[B + 2].ToString(), @"^-?[0 - 9.]", "").Replace(".", ","));
                                            CRIARCIRCULO(XINICIAL, YINICIAL, DIAMETRO, 0, file);
                                        }
                                        else
                                        {


                                            XINICIAL = double.Parse(Regex.Replace(valoresiniciais[B].ToString(), "[^0-9.]", "").Replace(".", ","));
                                            YINICIAL = double.Parse(Regex.Replace(valoresiniciais[B + 1].ToString(), "[^0-9.]", "").Replace(".", ","));
                                            DIAMETRO = double.Parse(Regex.Replace(valoresiniciais[B + 2].ToString(), @"^-?[0 - 9.]", "").Replace(".", ","));
                                            DISTX = double.Parse(Regex.Replace(valoresiniciais[B + 4].ToString(), @"^-?\d+$", "").Replace(".", ","));
                                            DISTY = double.Parse(Regex.Replace(valoresiniciais[B + 5].ToString(), @"^-?\d+$", "").Replace(".", ","));
                                            ANG = double.Parse(Regex.Replace(valoresiniciais[B + 6].ToString(), @"/^(?:[\d-]*,?[\d-]*\.?[\d-]*|[\d-]*\.[\d-]*,[\d-]*)$/", "").Replace(".", ","));

                                            double x1 = 0;
                                            double x2 = 0;
                                            double y1 = 0;
                                            double y2 = 0;
                                            double L = 0;
                                            var offset = DIAMETRO / 2;
                                            if (DISTX > 0)
                                            {

                                                //linha inicial central
                                                x1 = XINICIAL;
                                                x2 = XINICIAL + Math.Cos(ANG * Math.PI / 180) * DISTX;
                                                y1 = YINICIAL;
                                                y2 = YINICIAL + Math.Sin(ANG * Math.PI / 180) * DISTX;
                                                L = DISTX;
                                            }
                                            else if (DISTY > 0)
                                            {
                                                x1 = XINICIAL;
                                                x2 = XINICIAL + Math.Cos((ANG - 270) * Math.PI / 180) * DISTY;
                                                y1 = YINICIAL;
                                                y2 = YINICIAL + Math.Sin((ANG - 270) * Math.PI / 180) * DISTY;
                                                L = DISTY;
                                            }

                                            //decobrir a distancia da linha
                                            //var L = Math.Sqrt((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2));

                                            double y2y1 = (y2 - y1) / L;
                                            double x1x2 = (x1 - x2) / L;
                                            if (double.IsNaN(y2y1))
                                            {
                                                y2y1 = 0;
                                            }

                                            if (double.IsNaN(x1x2))
                                            {
                                                x1x2 = 0;
                                            }

                                            // linha paralela 1
                                            var x1p = x1 + offset * y2y1;
                                            var x2p = x2 + offset * y2y1;
                                            var y1p = y1 + offset * x1x2;
                                            var y2p = y2 + offset * x1x2;


                                            // linha paralela 2
                                            var x1p1 = x1 - (offset) * y2y1;
                                            var x2p1 = x2 - (offset) * y2y1;
                                            var y1p1 = y1 - (offset) * x1x2;
                                            var y2p1 = y2 - (offset) * x1x2;
                                            double BULGE = calculoBulge(x2p1, y2p1, x2p, y2p, DIAMETRO / 2);



                                            //enviar dados de polilinha furo avalizado
                                            if (!contersunk)
                                            {
                                                List<Ponto> POLY = new List<Ponto>();
                                                POLY.Add(new Ponto(x1p1, y1p1, 0));
                                                POLY.Add(new Ponto(x2p1, y2p1, -BULGE));
                                                POLY.Add(new Ponto(x2p, y2p, 0));
                                                POLY.Add(new Ponto(x1p, y1p, -BULGE));
                                                CRIARPOLYLINHA(POLY, 0, file);
                                            }
                                            else
                                            {
                                                contersunk = false;
                                            }

                                        }
                                    }
                                    catch (Exception)
                                    {

                                        R = false;

                                    }
                                }
                                if (lines[i].Contains("BO"))
                                {
                                    if (lines[i + 1].Contains("v")||lines[i + 1].Contains("h"))
                                        R = true;
                                }

                            }
                        }

                        R = false;
                        //MODULO DE QUINAGEM
                        for (int i = 0; i < w; i++)
                        {
                            if (lines.Contains("KA"))
                            {

                                if (R == true)
                                {

                                    try
                                    {
                                        int B = 0;
                                        double A;
                                        string[] valoresiniciais = lines[i].Split(' ').Where(x => !string.IsNullOrEmpty(x)).ToArray();
                                        if (!double.TryParse(valoresiniciais[0].ToString().Replace(".", ","), out A))
                                        {
                                            B = 1;
                                        }

                                        double XINICIAL = double.Parse(Regex.Replace(valoresiniciais[B].ToString(), "[^0-9.]", "").Replace(".", ","));
                                        double YINICIAL = double.Parse(Regex.Replace(valoresiniciais[B + 1].ToString(), "[^0-9.]", "").Replace(".", ","));
                                        string teste = "CIMA"+ (180 - double.Parse(valoresiniciais[4].Replace("-", "").Replace(".", ","))).ToString("0.0");
                                        if (valoresiniciais[4].Contains("-"))
                                        {
                                            teste = "BAIXO"+(180-double.Parse(valoresiniciais[4].Replace("-","").Replace(".", ","))).ToString("0.0");
                                        }



                                        if (XINICIAL > 0)
                                        {
                                           
                                            CRIARLINHA(XINICIAL, MAXY-10, XINICIAL, MAXY, 2, file);
                                            CRIARLINHA(XINICIAL, 0, XINICIAL, 10, 2, file);
                                            double Y = MAXY / 2;
                                            if (MAXY < 150)
                                                Y = 5;
                                            CRIARTEXTO(XINICIAL, Y, teste, 2, file,"90");
                                        }
                                        else if (YINICIAL > 0)
                                        {

                                            CRIARLINHA(0, YINICIAL, 10, YINICIAL, 2, file);
                                            CRIARLINHA(MAXX-10, YINICIAL, MAXX, YINICIAL, 2, file);
                                            double X=MAXX/2;
                                            if (MAXX < 150)
                                                X = 5;
                                                CRIARTEXTO(X, YINICIAL, teste, 2, file,"0");
                                        }

                                        



                                    }
                                    catch (Exception)
                                    {
                                        R = false;

                                    }
                                }
                                if (lines[i].Contains("KA"))
                                {
                                    R = true;
                                }
                            }
                        }

                        CRIARTEXTO(MAXX / 2, MAXY / 2, lines[4].Replace(" ", ""), 2, file);
                        CRIARRODAPE(file);
                    }
                    else
                    {
                        if (TESTE)
                        {

                            TESTE = false;
                            MessageBox.Show("ERRO NA CONVERSÃO DE FICHEIROS DXF POR FAVOR CONSULTE O FICHEIRO ERRO.TXT QUE ESTA NA PASTA.");
                            if (File.Exists(file.Replace(file.Split('\\').Last(), "") + "Erro.txt"))
                            {
                                File.Delete(file.Replace(file.Split('\\').Last(), "") + "Erro.txt");
                            }

                        }
                       
                        StreamWriter W = new StreamWriter(file.Replace(file.Split('\\').Last(), "") + "Erro.txt", true, Encoding.Default);
                        W.WriteLine(file);
                        W.Close();

                        CRIARCABEÇALHO(file);
                        //MODULO AK CONTORNO EXTERIOR

                        try
                        {

                            double d = 0;

                            double XINICIAL = double.Parse("0");
                            double YINICIAL = double.Parse("0");
                            double RAIO = double.Parse("0");
                            double XFINAL = double.Parse(lines[10].Replace(".", ","));
                            double YFINAL = double.Parse("0");

                            d = calculoBulge(XINICIAL, YINICIAL, XFINAL, YFINAL, RAIO);

                            pontos.Add(new Ponto(XINICIAL, YINICIAL, d));

                            if (MAXX < XINICIAL)
                            {
                                MAXX = XINICIAL;
                            }
                            if (MAXY < YINICIAL)
                            {
                                MAXY = YINICIAL;
                            }


                            XINICIAL = double.Parse(lines[10].Replace(".", ","));
                            YINICIAL = double.Parse("0");
                            RAIO = double.Parse("0");
                            XFINAL = double.Parse(lines[10].Replace(".", ","));
                            YFINAL = double.Parse(lines[11].Replace(".", ","));

                            d = calculoBulge(XINICIAL, YINICIAL, XFINAL, YFINAL, RAIO);

                            pontos.Add(new Ponto(XINICIAL, YINICIAL, d));

                            if (MAXX < XINICIAL)
                            {
                                MAXX = XINICIAL;
                            }
                            if (MAXY < YINICIAL)
                            {
                                MAXY = YINICIAL;
                            }



                            XINICIAL = double.Parse(lines[10].Replace(".", ","));
                            YINICIAL = double.Parse(lines[11].Replace(".", ","));
                            RAIO = double.Parse("0");
                            XFINAL = double.Parse("0");
                            YFINAL = double.Parse(lines[11].Replace(".", ","));

                            d = calculoBulge(XINICIAL, YINICIAL, XFINAL, YFINAL, RAIO);

                            pontos.Add(new Ponto(XINICIAL, YINICIAL, d));

                            if (MAXX < XINICIAL)
                            {
                                MAXX = XINICIAL;
                            }
                            if (MAXY < YINICIAL)
                            {
                                MAXY = YINICIAL;
                            }

                            XINICIAL = double.Parse("0");
                            YINICIAL = double.Parse(lines[11].Replace(".", ","));
                            RAIO = double.Parse("0");
                            XFINAL = double.Parse("0");
                            YFINAL = double.Parse("0");

                            d = calculoBulge(XINICIAL, YINICIAL, XFINAL, YFINAL, RAIO);

                            pontos.Add(new Ponto(XINICIAL, YINICIAL, d));

                            if (MAXX < XINICIAL)
                            {
                                MAXX = XINICIAL;
                            }
                            if (MAXY < YINICIAL)
                            {
                                MAXY = YINICIAL;
                            }

                            if (pontos != null)
                            {
                                CRIARPOLYLINHA(pontos, 0, file);
                            }

                        }
                        catch (Exception)
                        {


                        }

                        pontos.Clear();
                        //MODULO AK CONTORNO EXTERIOR
                        for (int i = 0; i < w; i++)
                        {

                            if (lines.Contains("IK"))
                            {

                                if (R == true)
                                {

                                    try
                                    {
                                        string[] valoresiniciais = lines[i].Split(' ').Where(x => !string.IsNullOrEmpty(x)).ToArray();
                                        string[] valoresfinais = lines[i + 1].Split(' ').Where(x => !string.IsNullOrEmpty(x)).ToArray();
                                        double A;
                                        int B = 0;
                                        int C = 0;
                                        double d = 0;
                                        if (!double.TryParse(valoresiniciais[0].ToString().Replace(".", ","), out A))
                                        {
                                            B = 1;

                                        }
                                        if (!double.TryParse(valoresfinais[0].ToString().Replace(".", ","), out A))
                                        {
                                            C = 1;

                                        }

                                        double XINICIAL = double.Parse(Regex.Replace(valoresiniciais[B].ToString(), "[^0-9.]", "").Replace(".", ","));
                                        double YINICIAL = double.Parse(Regex.Replace(valoresiniciais[B + 1].ToString(), "[^0-9.]", "").Replace(".", ","));
                                        double RAIO = double.Parse(Regex.Replace(valoresiniciais[B + 2].ToString(), @"^-?[0 - 9.]", "").Replace(".", ","));
                                        double XFINAL = double.Parse(Regex.Replace(valoresfinais[C].ToString(), "[^0-9.]", "").Replace(".", ","));
                                        double YFINAL = double.Parse(Regex.Replace(valoresfinais[C + 1].ToString(), "[^0-9.]", "").Replace(".", ","));


                                        d = calculoBulge(XINICIAL, YINICIAL, XFINAL, YFINAL, RAIO);

                                        pontos.Add(new Ponto(XINICIAL, YINICIAL, d));


                                    }
                                    catch (Exception)
                                    {

                                        R = false;
                                        if (pontos != null)
                                        {
                                            CRIARPOLYLINHA(pontos, 0, file);
                                            pontos.Clear();
                                        }
                                    }
                                }
                                if (lines[i].Contains("IK"))
                                {
                                    if (lines[i + 1].Contains("v")|| lines[i + 1].Contains("h"))
                                    {
                                        R = true;
                                    }
                                    else
                                    {
                                        R = false;
                                    }

                                }
                            }
                        }

                        R = false;
                        //MODULO DE FUROS
                        for (int i = 0; i < w; i++)
                        {
                            if (lines.Contains("BO"))
                            {

                                if (R == true)
                                {
                                    try
                                    {
                                        string[] valoresiniciais = lines[i].Split(' ').Where(x => !string.IsNullOrEmpty(x)).ToArray();
                                        bool contersunk = false;
                                        double A;
                                        int B = 0;
                                        double XINICIAL = 0;
                                        double YINICIAL = 0;
                                        double DIAMETRO = 0;
                                        double DISTX = 0;
                                        double DISTY = 0;
                                        Double ANG = 0;
                                        if (!double.TryParse(valoresiniciais[0].ToString().Replace(".", ","), out A))
                                        {
                                            B = 1;
                                        }

                                        if (valoresiniciais.Length == 8)
                                        {
                                            if (valoresiniciais[7].ToString() == "90.00" && valoresiniciais[6].ToString() == "0.00" && valoresiniciais[6].ToString() == "0.00")
                                            {
                                                contersunk = true;
                                            }
                                        }



                                        if (valoresiniciais.Length == 4 || valoresiniciais.Length == 5)
                                        {
                                            XINICIAL = double.Parse(Regex.Replace(valoresiniciais[B].ToString(), "[^0-9.]", "").Replace(".", ","));
                                            YINICIAL = double.Parse(Regex.Replace(valoresiniciais[B + 1].ToString(), "[^0-9.]", "").Replace(".", ","));
                                            DIAMETRO = double.Parse(Regex.Replace(valoresiniciais[B + 2].ToString(), @"^-?[0 - 9.]", "").Replace(".", ","));
                                            CRIARCIRCULO(XINICIAL, YINICIAL, DIAMETRO, 0, file);
                                        }
                                        else
                                        {


                                            XINICIAL = double.Parse(Regex.Replace(valoresiniciais[B].ToString(), "[^0-9.]", "").Replace(".", ","));
                                            YINICIAL = double.Parse(Regex.Replace(valoresiniciais[B + 1].ToString(), "[^0-9.]", "").Replace(".", ","));
                                            DIAMETRO = double.Parse(Regex.Replace(valoresiniciais[B + 2].ToString(), @"^-?[0 - 9.]", "").Replace(".", ","));
                                            DISTX = double.Parse(Regex.Replace(valoresiniciais[B + 4].ToString(), @"^-?\d+$", "").Replace(".", ","));
                                            DISTY = double.Parse(Regex.Replace(valoresiniciais[B + 5].ToString(), @"^-?\d+$", "").Replace(".", ","));
                                            ANG = double.Parse(Regex.Replace(valoresiniciais[B + 6].ToString(), @"/^(?:[\d-]*,?[\d-]*\.?[\d-]*|[\d-]*\.[\d-]*,[\d-]*)$/", "").Replace(".", ","));

                                            double x1 = 0;
                                            double x2 = 0;
                                            double y1 = 0;
                                            double y2 = 0;
                                            double L = 0;
                                            var offset = DIAMETRO / 2;
                                            if (DISTX > 0)
                                            {

                                                //linha inicial central
                                                x1 = XINICIAL;
                                                x2 = XINICIAL + Math.Cos(ANG * Math.PI / 180) * DISTX;
                                                y1 = YINICIAL;
                                                y2 = YINICIAL + Math.Sin(ANG * Math.PI / 180) * DISTX;
                                                L = DISTX;
                                            }
                                            else if (DISTY > 0)
                                            {
                                                x1 = XINICIAL;
                                                x2 = XINICIAL + Math.Cos((ANG - 270) * Math.PI / 180) * DISTY;
                                                y1 = YINICIAL;
                                                y2 = YINICIAL + Math.Sin((ANG - 270) * Math.PI / 180) * DISTY;
                                                L = DISTY;
                                            }

                                            //decobrir a distancia da linha
                                            //var L = Math.Sqrt((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2));
                                            double y2y1 = (y2 - y1) / L;
                                            double x1x2 = (x1 - x2) / L;
                                            if (double.IsNaN(y2y1))
                                            {
                                                y2y1 = 0;
                                            }

                                            if (double.IsNaN(x1x2))
                                            {
                                                x1x2 = 0;
                                            }

                                            // linha paralela 1
                                            var x1p = x1 + offset * y2y1;
                                            var x2p = x2 + offset * y2y1;
                                            var y1p = y1 + offset * x1x2;
                                            var y2p = y2 + offset * x1x2;


                                            // linha paralela 2
                                            var x1p1 = x1 - (offset) * y2y1;
                                            var x2p1 = x2 - (offset) * y2y1;
                                            var y1p1 = y1 - (offset) * x1x2;
                                            var y2p1 = y2 - (offset) * x1x2;
                                            double BULGE = calculoBulge(x2p1, y2p1, x2p, y2p, DIAMETRO / 2);


                                        



                                            //enviar dados de polilinha furo avalizado
                                            if (!contersunk)
                                            {
                                                List<Ponto> POLY = new List<Ponto>();
                                                POLY.Add(new Ponto(x1p1, y1p1, 0));
                                                POLY.Add(new Ponto(x2p1, y2p1, -BULGE));
                                                POLY.Add(new Ponto(x2p, y2p, 0));
                                                POLY.Add(new Ponto(x1p, y1p, -BULGE));
                                                CRIARPOLYLINHA(POLY, 0, file);
                                            }
                                            else
                                            {
                                                contersunk = false;
                                            }

                                        }
                                    }
                                    catch (Exception)
                                    {

                                        R = false;

                                    }
                                }
                                if (lines[i].Contains("BO"))
                                {

                                    R = true;

                                }
                            }
                        }

                        R = false;
                        //MODULO DE QUINAGEM
                        for (int i = 0; i < w; i++)
                        {
                            if (lines.Contains("KA"))
                            {

                                if (R == true)
                                {

                                    try
                                    {
                                        int B = 0;
                                        double A;
                                        string[] valoresiniciais = lines[i].Split(' ').Where(x => !string.IsNullOrEmpty(x)).ToArray();
                                        if (!double.TryParse(valoresiniciais[0].ToString().Replace(".", ","), out A))
                                        {
                                            B = 1;
                                        }

                                        double XINICIAL = double.Parse(Regex.Replace(valoresiniciais[B].ToString(), "[^0-9.]", "").Replace(".", ","));
                                        double YINICIAL = double.Parse(Regex.Replace(valoresiniciais[B + 1].ToString(), "[^0-9.]", "").Replace(".", ","));
                                        string teste = "CIMA" + (180 - double.Parse(valoresiniciais[4].Replace("-", "").Replace(".", ","))).ToString("0.0");
                                        if (valoresiniciais[4].Contains("-"))
                                        {
                                            teste = "BAIXO" + (180 - double.Parse(valoresiniciais[4].Replace("-", "").Replace(".", ","))).ToString("0.0");
                                        }



                                        if (XINICIAL > 0)
                                        {

                                            CRIARLINHA(XINICIAL, MAXY - 10, XINICIAL, MAXY, 2, file);
                                            CRIARLINHA(XINICIAL, 0, XINICIAL, 10, 2, file);
                                            double Y = MAXY / 2;
                                            if (MAXY < 150)
                                                Y = 5;
                                            CRIARTEXTO(XINICIAL, Y, teste, 2, file, "90");
                                        }
                                        else if (YINICIAL > 0)
                                        {

                                            CRIARLINHA(0, YINICIAL, 10, YINICIAL, 2, file);
                                            CRIARLINHA(MAXX - 10, YINICIAL, MAXX, YINICIAL, 2, file);
                                            double X = MAXX / 2;
                                            if (MAXX < 150)
                                                X = 5;
                                            CRIARTEXTO(X, YINICIAL, teste, 2, file, "0");
                                        }

                                    }
                                    catch (Exception)
                                    {
                                        R = false;

                                    }
                                }
                                if (lines[i].Contains("KA"))
                                {
                                    R = true;
                                }
                            }
                        }
                        CRIARTEXTO(MAXX / 2, MAXY / 2, lines[4].Replace(" ", ""), 2, file);
                        CRIARRODAPE(file);
                    }
                }
            }
        }

        public static void CRIARCABEÇALHO(string FILE)
        {
            StreamWriter WR = new StreamWriter(FILE, false);
            WR.WriteLine("0");
            WR.WriteLine("SECTION");
            WR.WriteLine("2");
            WR.WriteLine("ENTITIES");
            WR.Close();
        }

        public static void CRIARRODAPE(string FILE)
        {
            StreamWriter WR = new StreamWriter(FILE, true);
            WR.WriteLine("0");
            WR.WriteLine("ENDSEC");
            WR.WriteLine("0");
            WR.WriteLine("EOF");
            WR.Close();
        }

        public static void CRIARLINHA(double XINICIAL, double YINICIAL, double XFINAL, double YFINAL, int COR, string FILE)
        {
            StreamWriter WR = new StreamWriter(FILE, true);
            //////////////////////////////////////////////////inicio comando
            WR.WriteLine("0");
            WR.WriteLine("LINE");//comando pode cer circle texte
            WR.WriteLine("8");//codigo de criaçao de layer
            WR.WriteLine("LINHAS");// nome da layer
            WR.WriteLine("10");
            WR.WriteLine(XINICIAL.ToString().Replace(",", "."));//X inicio da linha
            WR.WriteLine("20");
            WR.WriteLine(YINICIAL.ToString().Replace(",", "."));//y inicio da linha
            WR.WriteLine("11");
            WR.WriteLine(XFINAL.ToString().Replace(",", "."));//X fim da linha
            WR.WriteLine("21");
            WR.WriteLine(YFINAL.ToString().Replace(",", "."));//Y fim da linha
            WR.WriteLine("62");//codigo de cor
            WR.WriteLine(COR);// codigo da cor 2 amarelo 0 preto 1 vermelho 3 verde
                              ////////////////////////////////////////////////////fim do comando
            WR.Close();
        }

        public static void CRIARCIRCULO(double XINICIAL, double YINICIAL, double DIAMETRO, int COR, string FILE)
        {
            StreamWriter WR = new StreamWriter(FILE, true);
            //////////////////////////////////////////////////inicio comando
            WR.WriteLine("0");
            WR.WriteLine("CIRCLE");//comando pode cer circle texte
            WR.WriteLine("8");//codigo de criaçao de layer
            WR.WriteLine("FUROS");// nome da layer
            WR.WriteLine("10");
            WR.WriteLine(XINICIAL.ToString().Replace(",", "."));//X inicio da linha
            WR.WriteLine("20");
            WR.WriteLine(YINICIAL.ToString().Replace(",", "."));//y inicio da linha
            WR.WriteLine("30");
            WR.WriteLine("0");//Z inicio da linha
            WR.WriteLine("40");
            WR.WriteLine((DIAMETRO / 2).ToString().Replace(",", "."));//DIAMETRO
            WR.WriteLine("62");//codigo de cor
            WR.WriteLine(COR);// codigo da cor 2 amarelo 0 preto 1 vermelho 3 verde
            WR.Close();
        }

        public static void CRIARARCO(double X, double Y, double DIAMETRO, double ANGULOINICIO, double ANGULOFIM, int COR, string FILE)
        {

            StreamWriter WR = new StreamWriter(FILE, true);
            WR.WriteLine("0");
            WR.WriteLine("ARC");//comando pode ser circle texte
            WR.WriteLine("8");//codigo de criaçao de layer
            WR.WriteLine("FUROS");// nome da layer
            WR.WriteLine("10");
            WR.WriteLine(X);//X inicio da linha
            WR.WriteLine("20");
            WR.WriteLine(Y);//y inicio da linha
            WR.WriteLine("30");
            WR.WriteLine("0");//Z inicio da linha
            WR.WriteLine("40");
            WR.WriteLine((DIAMETRO / 2).ToString().Replace(",", "."));//RAIO
            WR.WriteLine("50");
            WR.WriteLine(ANGULOINICIO.ToString().Replace(",", "."));//INICIO DO ANGULO
            WR.WriteLine("51");
            WR.WriteLine(ANGULOFIM.ToString().Replace(",", "."));//FIM DO ANGULO
            WR.WriteLine("62");//codigo de cor
            WR.WriteLine(COR);// codigo da cor 2 amarelo 0 preto 1 vermelho 3 verde
            WR.Close();
        }

        public static void CRIARTEXTO(double XINICIAL, double YINICIAL, string TEXTO, int COR, string FILE,string rotacao="0")
        {

            StreamWriter WR = new StreamWriter(FILE, true);
            WR.WriteLine("0");
            WR.WriteLine("TEXT");//comando pode ser circle texte
            WR.WriteLine("8");//codigo de criaçao de layer
            WR.WriteLine("TEXTO");// nome da layer
            WR.WriteLine("10");
            WR.WriteLine(XINICIAL.ToString().Replace(",", "."));//X inicio da linha
            WR.WriteLine("20");
            WR.WriteLine(YINICIAL.ToString().Replace(",", "."));//y inicio da linha
            WR.WriteLine("30");
            WR.WriteLine("0");//Z inicio da linha
            WR.WriteLine("40");
            WR.WriteLine("10");//TAMANHO
            WR.WriteLine("1");
            WR.WriteLine(TEXTO);
            WR.WriteLine("50");
            WR.WriteLine(rotacao);
            WR.WriteLine("62");//codigo de cor
            WR.WriteLine(COR);// codigo da cor 2 amarelo 0 preto 1 vermelho 3 verde
            WR.Close();
        }

        public static void CRIARPOLYLINHA(List<Ponto> pontos, int COR, string FILE)
        {

            StreamWriter WR = new StreamWriter(FILE, true);

            WR.WriteLine("0");
            WR.WriteLine("POLYLINE");//comando pode Ser circle texte
            WR.WriteLine("8");//codigo de criaçao de layer
            WR.WriteLine("POLILINHA");// nome da layer
            WR.WriteLine("6");//comando para tipo de linha
            WR.WriteLine("CONTINUOUS");
            WR.WriteLine("62");
            WR.WriteLine(COR);
            WR.WriteLine("66");
            WR.WriteLine("1");
            WR.WriteLine("10");
            WR.WriteLine("0.000");
            WR.WriteLine("20");
            WR.WriteLine("0.000");
            WR.WriteLine("30");
            WR.WriteLine("0.000");
            WR.WriteLine("70");
            WR.WriteLine("1");
            ///////////////////////////////////////////////////////////////////////////////////////
            foreach (Ponto pontoactual in pontos)
            {
                WR.WriteLine("0");
                WR.WriteLine("VERTEX");
                WR.WriteLine("8");
                WR.WriteLine("POLILINHA");
                WR.WriteLine("10");
                WR.WriteLine(pontoactual.PontoX.ToString("0.000000").Replace(',', '.'));
                WR.WriteLine("20");
                WR.WriteLine(pontoactual.PontoY.ToString("0.000000").Replace(',', '.'));
                WR.WriteLine("30");
                WR.WriteLine("0.000");
                WR.WriteLine("42");
                WR.WriteLine(pontoactual.Curvatura.ToString("0.000000").Replace(',', '.'));
            }
            /////////////////////////////////////////////////////////////////////
            WR.WriteLine("0");
            WR.WriteLine("SEQEND");
            WR.Close();

        }

        public static double calculoBulge(double xinicial, double yinicial, double xfinal, double yfinal, double RAIO)
        {
            double Bulge = 0;
            // o quadrado do comprimento da hipotenusa é igual à soma dos quadrados dos comprimentos dos catetos.
            float distance = (float)Math.Sqrt(Math.Pow(xinicial - xfinal, 2) + Math.Pow(yinicial - yfinal, 2));
            // um quarto da tangente de um angulo 
            Bulge = Math.Tan((2 * (Math.Asin(double.Parse((distance / (2 * RAIO)).ToString("0.00"))))) * 0.25);

            if (double.IsNaN(Bulge)) Bulge = 0;
            return Bulge;
        }

    }
    public class Ponto
    {
        public Ponto(double PontoX, double PontoY, double Curvatura)
        {
            _PontoX = PontoX;
            _PontoY = PontoY;
            _Curvatura = Curvatura;
        }
        public double PontoX
        {
            get { return _PontoX; }
        }
        public double Curvatura
        {
            get { return _Curvatura; }
        }
        public double PontoY
        {
            get { return _PontoY; }
        }
        public double _PontoX;
        public double _PontoY;
        public double _Curvatura;
    }
}
