using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Linq;
using System.Web.UI.WebControls;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Drawing;
using System.Drawing.Imaging;
using System.Web;

namespace Convert_Word_Document_to_Image
{
    public partial class MainPage : System.Web.UI.Page
    {

        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void OnButtonClicked(object sender, EventArgs e)
        {
            //Open existing Word document.
            using (FileStream docStream = new FileStream(Server.MapPath("~/App_Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                //Loads file stream into Word document
                using (WordDocument wordDocument = new WordDocument(docStream, FormatType.Docx))
                {
                    //Convert the first page of the Word document into an image.
                    System.Drawing.Image image = wordDocument.RenderAsImages(0, ImageType.Bitmap);
                    //Save the Image as Jpeg.
                    ExportAsImage(image, "WordToImage.Jpeg", ImageFormat.Jpeg, HttpContext.Current.Response);
                }
            }
        }

        //Download the Image file
        protected void ExportAsImage(System.Drawing.Image image, string fileName, ImageFormat imageFormat, HttpResponse response)
        {
            string disposition = "content-disposition";
            response.AddHeader(disposition, "attachment; filename=" + fileName);
            if (imageFormat != ImageFormat.Emf)
                image.Save(Response.OutputStream, imageFormat);
            Response.End();
        }
    }
}