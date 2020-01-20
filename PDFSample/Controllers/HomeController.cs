using System;
using System.Collections;
using Syncfusion.HtmlConverter;
using System.IO;
using System.Web.Mvc;
using System.Web.Razor.Text;
using iText.Html2pdf;
using iText.Layout.Element;
using Path = System.IO.Path;
using PdfDocument = Syncfusion.Pdf.PdfDocument;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html.simpleparser;
using IElement = iTextSharp.text.IElement;

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
            FileStreamResult file = new FileStreamResult(RenderHtmlToPdfStream("http://localhost:57519/Assets/Template/01.htm"), "application/PDF");
            return file;
        }


        public static MemoryStream RenderHtmlToPdfStream(string html)
        {
            var memoryStream = new MemoryStream();

            var reader = new StringReader(html);

            var document = new Document(PageSize.A4, 30, 30, 30, 30);
            HTMLWorker worker = new HTMLWorker(document);
            PdfWriter writer = PdfWriter.GetInstance(document, memoryStream);
            writer.CloseStream = false;
            document.Open();
            worker.StartDocument();
            worker.Parse(reader);
            worker.EndDocument();
            worker.Close();
            document.Close();
            memoryStream.Seek(0, 0);
            return memoryStream;
        }
    }
}