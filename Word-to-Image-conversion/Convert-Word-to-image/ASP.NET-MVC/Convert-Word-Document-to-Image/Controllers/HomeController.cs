using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.Drawing;

namespace Convert_Word_Document_to_Image.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        public void ConvertWordtoImage()
        {
            //Open the file as Stream
            using (FileStream docStream = new FileStream(Server.MapPath("~/App_Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                //Loads file stream into Word document
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Docx))
                {
                    //Convert the first page of the Word document into an image.
                    Image image = wordDocument.RenderAsImages(0, ImageType.Bitmap);
                    //Save the image as jpeg.           
                    ExportAsImage(image, "WordToImage.Jpeg", ImageFormat.Jpeg, HttpContext.ApplicationInstance.Response);
                }
            }
        }
        //To download the image file
        protected void ExportAsImage(Image image, string fileName, ImageFormat imageFormat, HttpResponse response)
        {
            if (ControllerContext == null)
                throw new ArgumentNullException("Context");
            string disposition = "content-disposition";
            response.AddHeader(disposition, "attachment; filename=" + fileName);
            if (imageFormat != ImageFormat.Emf)
                image.Save(Response.OutputStream, imageFormat);
            Response.End();
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
    }
}