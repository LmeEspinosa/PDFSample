using System;
using System.Collections;
using Syncfusion.HtmlConverter;
using System.IO;
using System.Web.Mvc;
using iText.Html2pdf;
using Path = System.IO.Path;
using PdfDocument = Syncfusion.Pdf.PdfDocument;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html.simpleparser;

namespace PDFSample.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult PDF()
        {
            return View();
        }

        public ActionResult ConvertSF()
        {
            HtmlToPdfConverter converter = new HtmlToPdfConverter();
            WebKitConverterSettings setting = new WebKitConverterSettings();
            setting.WebKitPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "QtBinariesWindows");
            converter.ConverterSettings = setting;

            PdfDocument document = converter.Convert("http://localhost:57519/Assets/Template/01.htm");
            //PdfDocument document = converter.Convert(Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/Assets/Template"), "01.htm"));


            MemoryStream ms = new MemoryStream();
            document.Save(ms);
            document.Close(true);
            ms.Position = 0;

            FileStreamResult fileStreamResult = new FileStreamResult(ms, "application/PDF");
            //fileStreamResult.FileDownloadName = "SolicitudAdhesion.pdf";

            return fileStreamResult;
        }

        public FileStreamResult ConvertIT()
        {
            try
            {
                //FileStream htmlSource = System.IO.File.Open(Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/Assets/Template"), "01.htm"), FileMode.Open);
                var html =
                    "http://localhost:57519/Assets/Template/01.htm"; //Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/Assets/Template"), "01.htm");
                using (var docStream = new MemoryStream())
                {
                    //docStream.ReadTimeout = int.MaxValue;
                    //docStream.WriteTimeout = int.MaxValue;
                    using (var docWriter = new iText.Kernel.Pdf.PdfWriter(docStream))
                    {
                        docWriter.SetCloseStream(false);
                        using (var doc = new iText.Kernel.Pdf.PdfDocument(docWriter))
                        {
                            doc.SetCloseWriter(false);
                            doc.SetCloseReader(false);
                            HtmlConverter.ConvertToPdf(html, doc, new ConverterProperties());

                            docStream.Position = 0;
                            FileStreamResult fileStreamResult = new FileStreamResult(docStream, "application/PDF");
                            //fileStreamResult.FileDownloadName = "SolicitudAdhesion.pdf";
                            //docStream.Flush(); //Always catches me out
                            //docStream.Position = 0; //Not sure if this is required

                            return fileStreamResult;
                        }
                    }
                }
                //MemoryStream ms = new MemoryStream();
                //PdfWriter pdfWriter = new PdfWriter(ms);
                //iText.Kernel.Pdf.PdfDocument document = new iText.Kernel.Pdf.PdfDocument(pdfWriter);

                //HtmlConverter.ConvertToPdf(htmlSource, document, new ConverterProperties());
                //document.Close();
                //htmlSource.Close();


                //FileStreamResult fileStreamResult = new FileStreamResult(ms, "application/PDF");
                //fileStreamResult.FileDownloadName = "SolicitudAdhesion.pdf";

                //return fileStreamResult;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public FileStreamResult ConvertITS()
        {
            //using (MemoryStream stream = new MemoryStream())
            //{
            //StringReader sr = new StringReader("http://localhost:57519/Assets/Template/01.htm");
            //Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 100f, 0f);
            //iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(pdfDoc, stream);
            //pdfDoc.Open();
            //HTMLWorker.ParseToList()//ParseXHtml(writer, pdfDoc, sr);
            //pdfDoc.Close();
            //return FileStreamResult(stream.ToArray(), "application/pdf", "Grid.pdf");
            string file1 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "File1.pdf");

            using (FileStream fs = new FileStream(file1, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                Document doc = new Document(PageSize.A4);
                iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, fs);
                doc.Open();
                //string html = "<table><tr><th>First Name</th><th>Last Name</th></tr><tr><td>Chris</td><td>Haas</td></tr></table>";
                string html = "http://localhost:57519/Assets/Template/01.htm";
                using (StringReader sr = new StringReader(html))
                {
                    //Create a style sheet
                    StyleSheet styles = new StyleSheet();
                    //...styles omitted for brevity

                    //Convert our HTML to iTextSharp elements
                    ArrayList elements = HTMLWorker.ParseToList(sr, styles);
                    //Loop through each element (in this case there's actually just one PdfPTable)
                    foreach (IElement el in elements)
                    {
                        //If the element is a PdfPTable
                        if (el is PdfPTable)
                        {
                            //Cast it
                            PdfPTable tt = (PdfPTable) el;
                            //Change the widths, these are relative width by the way
                            tt.SetWidths(new float[] {75, 25});
                        }

                        //Add the element to the document
                        doc.Add(el);
                    }
                }

                doc.Close();
                
                FileStreamResult file = new FileStreamResult(fs, "application/PDF");
                return file;
            }
        }
    }
}