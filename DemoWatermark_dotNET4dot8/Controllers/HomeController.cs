using DemoWatermark_dotNET4dot8.Models;
using QRCoder;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace DemoWatermark_dotNET4dot8.Controllers
{
    public class HomeController : Controller
    {
        #region Return Views
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
        public ActionResult CreateQRCode(string stringCode)
        {
            return View();
        }
        #endregion
        public JsonResult CreateQRCodeJob(string title, string stringCode, int price, int curStatus, DateTime expDate)
        {
            if (string.IsNullOrEmpty(stringCode)) return null;

            try
            {
                #region Generate QR Image
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(stringCode, QRCodeGenerator.ECCLevel.Q);
                QRCode qrCode = new QRCode(qrCodeData);
                Bitmap qrCodeBitmap = qrCode.GetGraphic(20);
                qrCodeBitmap.SetResolution(75, 75);
                Bitmap qrCodeBitmapResized = new Bitmap(qrCodeBitmap, new Size(375, 375));
                #endregion

                #region Create QR with template
                //
                #region Check name's width and get lines
                //Create temporay graphics wid defined width to check ticket name width
                Bitmap tempBitmap = new Bitmap(656, 450);
                Graphics tempGraphic = Graphics.FromImage(tempBitmap);

                //Measure width and get lines
                Font pFont = new Font("Roboto", 28, FontStyle.Bold);
                SizeF pSize = tempGraphic.MeasureString(title, pFont, tempBitmap.Width - 64);
                SizeF cSize = tempGraphic.MeasureString(stringCode, pFont, tempBitmap.Width - 64);
                int lines = (int)Math.Round(pSize.Height / pFont.Height);
                int cLines = (int)Math.Round(cSize.Height / pFont.Height);
                #endregion
                //
                #region Watermark image Header, Footer and QR
                //Define general frame 
                Bitmap frameBitmap = new Bitmap(656, (int)(958 + cSize.Height + pSize.Height));
                Graphics frameGraphic = Graphics.FromImage(frameBitmap);
                frameGraphic.Clear(Color.White);

                //Draw footer
                Image footerImage = Image.FromFile(System.Web.HttpContext.Current.Server.MapPath("assets/img/new-frame-footer-656-177.png"));
                frameGraphic.DrawImage(footerImage, 0, (int)(958 + cSize.Height + pSize.Height) - 177);

                //Draw QR
                Bitmap qrBitImg = new Bitmap(qrCodeBitmap, new Size(430, 430));
                frameGraphic.DrawImage(qrBitImg, frameBitmap.Width / 2 - qrBitImg.Width / 2, 50);

                //Draw logo
                Image logoImage = Image.FromFile(System.Web.HttpContext.Current.Server.MapPath("assets/img/logo.png"));
                frameGraphic.DrawImage(logoImage, (frameBitmap.Width / 2 - 32), 50 + qrBitImg.Height / 2 - 32, 64, 64);

                //Draw Header
                Image headerImage = Image.FromFile(System.Web.HttpContext.Current.Server.MapPath("assets/img/new-frame-header-656-100.png"));
                frameGraphic.DrawImage(headerImage, 0, 0);
                #endregion
                //
                #region Watermark text
                //-----------------------------------------------------------------------------------------
                #region General setup
                ///Draws to the Graphics object with coordinate (0, 0) at 100% original size => All future drawing will occur on top of the original photograph
                //frameGraphic.SmoothingMode = SmoothingMode.AntiAlias;

                frameGraphic.DrawImage(frameBitmap,
                                        new Rectangle(0, 0, frameBitmap.Width, frameBitmap.Height),
                                        0,
                                        0,
                                        frameBitmap.Width,
                                        frameBitmap.Height,
                                        GraphicsUnit.Pixel);
                //-----------------------------------------------------------------------------------------



                ///Define a StringFormat object and set the StringAlignment to Center. 
                //-----------------------------------------------------------------------------------------
                StringFormat StrFormat = new StringFormat();
                StrFormat.Alignment = StringAlignment.Center;
                //-----------------------------------------------------------------------------------------
                #endregion

                #region Draw Order ID
                var noteColor = System.Drawing.ColorTranslator.FromHtml("#9E9E9E");
                SolidBrush noteBrush = new SolidBrush(noteColor);
                Font noteFont = new Font("Roboto", 22, FontStyle.Regular);

                frameGraphic.DrawString("Trạng thái: " + (curStatus == 0 ? "Avaiable" : (curStatus == 1 ? "Used" : "Others")),
                                    noteFont,
                                    noteBrush,
                                    new PointF(frameBitmap.Width / 2, 100 + 325 + 35),
                                    StrFormat);
                #endregion

                #region Draw Code
                var contentColor = System.Drawing.ColorTranslator.FromHtml("#000000");
                SolidBrush codeBrush = new SolidBrush(contentColor); //SolidBrush codeBrush = new SolidBrush(Color.Black);
                Font contentFont = new Font("Roboto", 28, FontStyle.Bold);


                //Create rectangle frame for wrap text
                RectangleF cRectFrame = new RectangleF(32,
                                                        100 + 325 + 35 + frameGraphic.MeasureString("Y", noteFont).Height + 35,
                                                        frameBitmap.Width - 64,
                                                        cLines * pSize.Height);

                RectangleF cRectFrame2 = new RectangleF(5 / 10 + 32,
                                                        1 / 10 + 100 + 325 + 35 + frameGraphic.MeasureString("Y", noteFont).Height + 35,
                                                        frameBitmap.Width - 64,
                                                        cLines * pSize.Height);

                frameGraphic.DrawString(stringCode,
                                        contentFont,
                                        codeBrush,
                                        cRectFrame,
                                        StrFormat);

                frameGraphic.DrawString(stringCode,
                                        contentFont,
                                        codeBrush,
                                        cRectFrame2,
                                        StrFormat);
                #endregion

                #region Draw Name
                //Create rectangle frame for wrap text
                RectangleF rectFrame = new RectangleF(32,
                                                        100 + 325 + 35 + frameGraphic.MeasureString("Y", noteFont).Height + 35 + cSize.Height + 35,
                                                        frameBitmap.Width - 64,
                                                        lines * pSize.Height);

                RectangleF rectFrame2 = new RectangleF(5 / 10 + 32,
                                                        1 / 10 + 100 + 325 + 35 + frameGraphic.MeasureString("Y", noteFont).Height + 35 + cSize.Height + 35,
                                                        frameBitmap.Width - 64,
                                                        lines * pSize.Height);

                frameGraphic.DrawString(title,
                                        contentFont,
                                        codeBrush,
                                        rectFrame,
                                        StrFormat);

                frameGraphic.DrawString(title,
                                        contentFont,
                                        codeBrush,
                                        rectFrame2,
                                        StrFormat);
                #endregion

                #region Draw EXP date
                frameGraphic.DrawString("HSD: " + expDate.ToString("dd/MM/yyyy"),
                                    noteFont,
                                    noteBrush,
                                    new PointF(frameBitmap.Width / 2,
                                                100 + 325 + 35 + frameGraphic.MeasureString("Y", noteFont).Height + 35 + cSize.Height + 35 + pSize.Height + 35),
                                    StrFormat);
                #endregion

                #region Draw Price
                var priceColor = System.Drawing.ColorTranslator.FromHtml("#F7941D");
                SolidBrush priceBrush = new SolidBrush(priceColor); //SolidBrush codeBrush = new SolidBrush(Color.Black);
                Font priceFont = new Font("Roboto", 45, FontStyle.Bold);

                frameGraphic.DrawString(price.ToString("N0") + "đ",
                                        priceFont,
                                        priceBrush,
                                        new PointF(frameBitmap.Width / 2,
                                                    100 + 325 + 35 + frameGraphic.MeasureString("Y", noteFont).Height + 35 + frameGraphic.MeasureString("Y", noteFont).Height + 35 + frameGraphic.MeasureString("Y", contentFont).Height + 35 + pSize.Height + 35),
                                        StrFormat);

                frameGraphic.DrawString(price.ToString("N0") + "đ",
                                        priceFont,
                                        priceBrush,
                                        new PointF(1 + frameBitmap.Width / 2,
                                                    1 + 100 + 325 + 35 + frameGraphic.MeasureString("Y", noteFont).Height + 35 + frameGraphic.MeasureString("Y", noteFont).Height + 35 + frameGraphic.MeasureString("Y", contentFont).Height + 35 + pSize.Height + 35),
                                        StrFormat);
                #endregion

                #region Round corner
                int CornerRadius = 18;
                CornerRadius *= 2;
                Bitmap RoundedImage = new Bitmap(frameBitmap.Width, frameBitmap.Height);

                using (Graphics g = Graphics.FromImage(RoundedImage))
                {
                    g.Clear(Color.White);
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    Brush brush = new TextureBrush(frameBitmap);
                    GraphicsPath gp2 = new GraphicsPath();
                    gp2.AddArc(0, 0, CornerRadius, CornerRadius, 180, 90);
                    gp2.AddArc(0 + RoundedImage.Width - CornerRadius, 0, CornerRadius, CornerRadius, 270, 90);
                    gp2.AddArc(0 + RoundedImage.Width - CornerRadius, 0 + RoundedImage.Height - CornerRadius, CornerRadius, CornerRadius, 0, 90);
                    gp2.AddArc(0, 0 + RoundedImage.Height - CornerRadius, CornerRadius, CornerRadius, 90, 90);
                    g.FillPath(brush, gp2);
                }

                frameBitmap = RoundedImage;

                //Set transparent background(Corner only)
                for (int x = 0; x < 18; x++)
                {
                    for (int y = frameBitmap.Height - 18; y < frameBitmap.Height; y++)
                    {
                        Color currentColor = frameBitmap.GetPixel(x, y);
                        if (currentColor.R >= 220 && currentColor.G >= 220 && currentColor.B >= 220)
                        {
                            frameBitmap.SetPixel(x, y, Color.Transparent);
                        }
                    }
                }

                for (int x = 656 - 18; x < 656; x++)
                {
                    for (int y = frameBitmap.Height - 18; y < frameBitmap.Height; y++)
                    {
                        Color currentColor = frameBitmap.GetPixel(x, y);
                        if (currentColor.R >= 220 && currentColor.G >= 220 && currentColor.B >= 220)
                        {
                            frameBitmap.SetPixel(x, y, Color.Transparent);
                        }
                    }
                }
                #endregion
                //-----------------------------------------------------------------------------------------
                #endregion
                //
                #region Use memorystream to display image
                using (MemoryStream msDemo = new MemoryStream())
                {
                    frameBitmap.Save(msDemo, System.Drawing.Imaging.ImageFormat.Png);
                    //frameBitmap.Save("C:\\Users\\User\\Downloads\\imgService.Png"); --Save to server
                    ItemInfo fileInfo = new ItemInfo();
                    fileInfo.fileName = "img_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".png";
                    fileInfo.fileData = "data:image/png;base64," + Convert.ToBase64String(msDemo.ToArray());

                    #region TETSING uing byte data
                    ImageConverter converter = new ImageConverter();
                    byte[] byteData = (byte[])converter.ConvertTo(frameBitmap, typeof(byte[]));
                    ViewBag.fileData = "data:image/png;base64," + Convert.ToBase64String(msDemo.ToArray()); ;
                    #endregion

                    return Json(fileInfo, JsonRequestBehavior.AllowGet);
                }
                #endregion
                //
                #endregion
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
        public string UploadFileThenGenQR()
        {
            StringBuilder lstDisplay = new StringBuilder();

            if (Request.Files.Count > 0)
            {
                try
                {
                    #region Upload file
                    //  Get all files from Request object  
                    HttpFileCollectionBase files = Request.Files;
                    var path = System.Web.HttpContext.Current.Server.MapPath(files[0].FileName);

                    HttpPostedFileBase file = files[0];
                    string fname;

                    // Checking for Internet Explorer  
                    if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                    {
                        string[] testfiles = file.FileName.Split(new char[] { '\\' });
                        fname = testfiles[testfiles.Length - 1];
                    }
                    else
                    {
                        fname = file.FileName;
                    }

                    //Save file
                    string fileExtension = Path.GetExtension(file.FileName);
                    file.SaveAs(path);
                    #endregion

                    #region Read Data From imported file
                    DataTable dt = new DataTable();
                    dt = ReadImportExcelFile("Sheet1", System.Web.HttpContext.Current.Server.MapPath(file.FileName)); //Get Excel file with static path(saved above)

                    List<GenerateQRCodeModel> lstCode = (from DataRow row in dt.Rows select new GenerateQRCodeModel { 
                                                                                                                      code = row["Code"].ToString(), 
                                                                                                                      title = row["Title"].ToString(),
                                                                                                                      status = row["Status"].ToString(),
                                                                                                                      expDate = row["ExpDate"].ToString(),
                                                                                                                      price = row["Price"].ToString()
                                                                                                                    }
                                                        ).ToList(); //Select data in Excel ~ Table
                    #endregion

                    #region Generate QR list
                    var rowNum = 0;

                    //Get list QR
                    foreach (var code in lstCode)
                    {
                        rowNum++;
                        #region Generate QR Image
                        QRCodeGenerator qrGenerator = new QRCodeGenerator();
                        QRCodeData qrCodeData = qrGenerator.CreateQrCode(code.code, QRCodeGenerator.ECCLevel.Q);
                        QRCode qrCode = new QRCode(qrCodeData);
                        Bitmap qrCodeBitmap = qrCode.GetGraphic(20);
                        qrCodeBitmap.SetResolution(75, 75);
                        Bitmap qrCodeBitmapResized = new Bitmap(qrCodeBitmap, new Size(375, 375));

                        //Draw logo
                        Image logoImage = Image.FromFile(System.Web.HttpContext.Current.Server.MapPath("assets/img/logo.png"));
                        Graphics tempData = Graphics.FromImage(qrCodeBitmapResized);
                        tempData.DrawImage(logoImage, (qrCodeBitmapResized.Width / 2 - 32), qrCodeBitmapResized.Height / 2 - 32, 64, 64);
                        #endregion

                        using (MemoryStream ms = new MemoryStream())
                        {
                            if (qrCodeBitmapResized != null)
                            {
                                qrCodeBitmapResized.Save(ms, ImageFormat.Png);
                                var base64img = "data:image/png;base64," + Convert.ToBase64String(ms.ToArray());

                                lstDisplay.Append("<div class='col-lg-3 col-md-4 col-sm-6 box-qr'>");
                                lstDisplay.Append("<div class='lst-pos-center'>");
                                lstDisplay.Append("<img src= '" + base64img + "' />");
                                lstDisplay.Append("</div>");
                                lstDisplay.Append("<div class='lst-pos-center'>");
                                lstDisplay.Append("<a id='download-" + rowNum + "' download='qr-" + rowNum + "' href='" + base64img + "' class='btn btn-primary'>Download</a>");
                                lstDisplay.Append("</div>");
                                lstDisplay.Append("</div>");
                            }
                        }
                    }
                    #endregion

                    //Remove Excel file uploaded
                    System.IO.File.Delete(path);

                    return lstDisplay.ToString();
                }
                catch (Exception e)
                {
                    return "";
                }
            }
            else
            {
                return "";
            }
        }
        private DataTable ReadImportExcelFile(string sheetName, string path)
        {
            sheetName = sheetName.Trim();
            path = path.Trim();
            using (OleDbConnection conn = new OleDbConnection())
            {
                DataTable dt = new DataTable();
                string Import_FileName = path;
                string fileExtension = Path.GetExtension(Import_FileName);
                if (fileExtension == ".xls")
                    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 8.0;HDR=YES;'";
                if (fileExtension == ".xlsx")
                    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";
                    comm.Connection = conn;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                        return dt;
                    }
                }
            }
        }
    }
}