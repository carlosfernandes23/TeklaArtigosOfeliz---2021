using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Text.RegularExpressions;
using iTextSharp.text.pdf.parser;
using System.Drawing;


namespace TeklaArtigosOfeliz
{
    class pdfitext
    {
        public void pdfsoldaduraescreve(string[] oldFiles, string nota)
        {
            foreach (string oldFile in oldFiles)
            {
                Regex regex = new Regex(@"\d.\d\d\d\d\d\d\d.\dCJ\d.pdf");
                Match match = regex.Match(oldFile.Split('\\').Last());
                if (match.Success)
                {
                    //create a document object
                    //var doc = new Document(PageSize.A4);
                    //create PdfReader object to read from the existing document
                    PdfReader reader = new PdfReader(oldFile);
                    //select two pages from the original document
                    reader.SelectPages("1");
                    //create PdfStamper object to write to get the pages from reader
                    PdfStamper stamper = new PdfStamper(reader, new FileStream(oldFile.Replace("\\20004\\", "\\20005\\").Replace(".pdf", " - 1.pdf"), FileMode.Create));
                    // PdfContentByte from stamper to add content to the pages over the original content
                    PdfContentByte pbover = stamper.GetOverContent(1);
                    //add content to the page using ColumnText
                    var baseFont = BaseFont.CreateFont(@"C:\Windows\Fonts\arialbd.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_CENTER, new Phrase(nota, new iTextSharp.text.Font(baseFont, 6)), 540, 140, 0);
                    //Creates an image that is the size i need to hide the text i'm interested in removing
                    creatbmpblank("soldadura       ");
                    iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance("bit.bmp");
                    //Sets the position that the image needs to be placed (ie the location of the text to be removed)
                    //txtX.Text = 33,txtY.Text = 708
                    image.SetAbsolutePosition(775, 158);
                    //Adds the image to the output pdf
                    stamper.GetOverContent(1).AddImage(image, true);
                    //Creates the first copy of the outputted pdf
                    ColumnText.ShowTextAligned(pbover, Element.ALIGN_CENTER, new Phrase("Soldadura", new iTextSharp.text.Font(baseFont, 6)), 795, 160, 0);
                    // PdfContentByte from stamper to add content to the pages under the original content
                    PdfContentByte pbunder = stamper.GetUnderContent(1);
                    //close the stamper
                    stamper.Close();
                }
            }
        }

        void creatbmpblank(string sFileData)
        {
            System.Drawing.Font oFont = new System.Drawing.Font("Arial", 8.5F, FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel);
            var sz = MeasureString(sFileData, oFont);

            var oBitmap = new Bitmap((int)sz.Width, (int)sz.Height);

            using (Graphics oGraphics = Graphics.FromImage(oBitmap))
            {
                oGraphics.Clear(Color.White);
                oGraphics.DrawString(sFileData, oFont, new SolidBrush(System.Drawing.Color.White), 0, 0);
                oGraphics.Flush();

            }

            oBitmap.Save("bit.bmp", System.Drawing.Imaging.ImageFormat.Bmp);
        }

        public static SizeF MeasureString(string s, System.Drawing.Font font)
        {
            SizeF result;
            using (var image = new Bitmap(1, 1))
            {
                using (var g = Graphics.FromImage(image))
                {
                    result = g.MeasureString(s, font);
                }
            }
            return result;
        }

        public static void CriarPlanoSoldadura(string fase,string designacao,string numeroobra,string cliente,string classe)
        {

            //create a document object
            //var doc = new Document(PageSize.A4);
            //create PdfReader object to read from the existing document
            PdfReader reader = new PdfReader("Plano_Soldadura_Fase.pdf");
            //select two pages from the original document
            reader.SelectPages("1");
            //create PdfStamper object to write to get the pages from reader
            PdfStamper stamper = new PdfStamper(reader, new FileStream(@"c:\r\Plano_Soldadura_Fase"+ fase +".pdf", FileMode.Create));
            // PdfContentByte from stamper to add content to the pages over the original content
            PdfContentByte pbover = stamper.GetOverContent(1);
            //add content to the page using ColumnText
            var baseFont = BaseFont.CreateFont(@"C:\Windows\Fonts\arialbd.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            ColumnText.ShowTextAligned(pbover, Element.ALIGN_CENTER, new Phrase(DateTime.Now.ToShortDateString(), new iTextSharp.text.Font(baseFont, 8)), 480, 153, 0);
            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(fase, new iTextSharp.text.Font(baseFont, 10)), 485, 725, 0);
            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(numeroobra, new iTextSharp.text.Font(baseFont, 10)), 133, 725, 0);
            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(cliente, new iTextSharp.text.Font(baseFont, 10)), 130, 711, 0);
            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(designacao, new iTextSharp.text.Font(baseFont, 10)), 150, 697, 0);
            ColumnText.ShowTextAligned(pbover, Element.ALIGN_LEFT, new Phrase(classe, new iTextSharp.text.Font(baseFont, 10)), 185, 683, 0);
            //close the stamper
            stamper.Close();

        }


    }
}
