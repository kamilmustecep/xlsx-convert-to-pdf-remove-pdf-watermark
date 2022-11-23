using Aspose.Cells;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.xtra.iTextSharp.text.pdf.pdfcleanup;
using System.IO;

namespace Excel2PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            //XLSX CONVERT TO PDF
            string fileName = "test.xlsx";
            Workbook workbook = new Workbook(fileName);
            workbook.Save("output.pdf", SaveFormat.Pdf);
            Console.WriteLine(fileName + " PDF oluşturuldu.");


            //WATERMARK REMOVER 
            textsharpie(fileName);

        }


        static void textsharpie(string fileName)
        {

            Console.WriteLine(fileName + " sayfalar temizleniyor...");
            string file = "output.pdf";
            string oldchar = "output.pdf";
            string repChar = "PDFS-OF-DAY/" + fileName.Replace("xlsx", "pdf");
            PdfReader reader = new PdfReader(file);
            PdfStamper stamper = new PdfStamper(reader, new FileStream(file.Replace(oldchar, repChar), FileMode.Create, FileAccess.Write));
            List<PdfCleanUpLocation> cleanUpLocations = new List<PdfCleanUpLocation>();
            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                cleanUpLocations.Add(new PdfCleanUpLocation(i, new iTextSharp.text.Rectangle(0, 830, 1300, 900), iTextSharp.text.BaseColor.WHITE));
            }
            PdfCleanUpProcessor cleaner = new PdfCleanUpProcessor(cleanUpLocations, stamper);
            cleaner.CleanUp();
            stamper.Close();
            reader.Close();
            Console.WriteLine(fileName + " dönüştürme tamamlandı.");


            FileInfo fi = new FileInfo("output.pdf");
            try
            {
                fi.Delete();
                Console.WriteLine("İlk çıkış dosyası silindi.");
            }
            catch (IOException e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
