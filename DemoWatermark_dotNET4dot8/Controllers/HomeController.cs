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
            #region Initial 
            //Define working folder
            string WorkingDirectory = @"D:\2_demo_project\DemoWatermark_dotNET4dot8\Watermark-.Net4.8\DemoWatermark_dotNET4dot8\wwwroot\assets\watermark";
            //Define watermark text
            string Copyright = "Copyright © 2023 - J.Star";
            #endregion

            #region Create Bitmap object from selected file
            //Creates an image object from file 
            Image imgPhoto = Image.FromFile(WorkingDirectory + "\\watermark_photo.png");
            int phWidth = imgPhoto.Width; //Get image width
            int phHeight = imgPhoto.Height; //Get image height

            //Build a Bitmap object with a 24 bits per pixel format for the color data from above data
            Bitmap bmPhoto = new Bitmap(phWidth, phHeight, PixelFormat.Format24bppRgb);
            bmPhoto.SetResolution(72, 72);

            //Create new Graphics object from the above Bitmap image.
            Graphics grPhoto = Graphics.FromImage(bmPhoto);
            #endregion

            #region Load watermark image
            //Loads the watermark image
            Image imgWatermark = new Bitmap(WorkingDirectory + "\\watermark.png");
            int wmWidth = imgWatermark.Width; //Get image width
            int wmHeight = imgWatermark.Height; //Get image height
            #endregion

            #region Watermark text
            //Draws imgPhoto to the Graphics object positioning it (x= 0, y=0) at 100% of its original size
            //(All future drawing will occur on top of the original photograph)
            grPhoto.SmoothingMode = SmoothingMode.AntiAlias;
            grPhoto.DrawImage(imgPhoto,
                              new Rectangle(0, 0, phWidth, phHeight),
                              0,
                              0,
                              phWidth,
                              phHeight,
                              GraphicsUnit.Pixel);

            //Determining the largest possible size of watermark text
            int[] sizes = new int[] { 28, 26, 24, 22, 20, 18, 16, 14, 12, 10, 8, 6, 4 };
            Font crFont = null;
            SizeF crSize = new SizeF(); //get watermark text height
            for (int i = 0; i < 7; i++)
            {
                crFont = new Font("arial", sizes[i], FontStyle.Bold);
                crSize = grPhoto.MeasureString(Copyright, crFont);

                if ((ushort)crSize.Width < (ushort)phWidth) break;
            }
            
            int yPixlesFromBottom = (int)(phHeight * .05); //Determine a position 5% from the bottom of the image
            float yPosFromBottom = ((phHeight - yPixlesFromBottom) - (crSize.Height / 2)); //Use watermark text height to determine y-coordinate
            float xCenterOfImg = (phWidth / 2); //Determine x-coordinate by calculating the centre of the image

            //Define a StringFormat object and set the StringAlignment to Center. 
            StringFormat StrFormat = new StringFormat();
            StrFormat.Alignment = StringAlignment.Center;


            //Create a SolidBrush with a Color of 60% Black (alpha value of 153)
            SolidBrush semiTransBrush2 = new SolidBrush(Color.FromArgb(153, 0, 0, 0));

            //Draw the Copyright string at the appropriate position offset 1 pixel to the right and 1 pixel down(This offset will create a shadow effect)
            grPhoto.DrawString(Copyright,
                               crFont,
                               semiTransBrush2,
                               new PointF(xCenterOfImg + 1, yPosFromBottom + 1),
                               StrFormat);

            //Repeat this process using a White Brush drawing the same text directly on top of the previously drawn string
            SolidBrush semiTransBrush = new SolidBrush(Color.FromArgb(153, 255, 255, 255));
            grPhoto.DrawString(Copyright,
                               crFont,
                               semiTransBrush,
                               new PointF(xCenterOfImg, yPosFromBottom),
                               StrFormat);
            #endregion

            #region Watermark image
            //Create a Bitmap based on the previously modified photograph
            Bitmap bmWatermark = new Bitmap(bmPhoto);
            bmWatermark.SetResolution(imgPhoto.HorizontalResolution, imgPhoto.VerticalResolution);

            //Load this Bitmap into a new Graphic Object
            Graphics grWatermark = Graphics.FromImage(bmWatermark);

            ImageAttributes imageAttributes = new ImageAttributes();
            ColorMap colorMap = new ColorMap();

            colorMap.OldColor = Color.FromArgb(255, 0, 255, 0);
            colorMap.NewColor = Color.FromArgb(0, 0, 0, 0);
            ColorMap[] remapTable = { colorMap };

            imageAttributes.SetRemapTable(remapTable, ColorAdjustType.Bitmap);


            //The second color manipulation is used to change the opacity of the watermark
            //This is done by applying a 5x5 matrix that contains the coordinates for the RGBA space
            //By setting the 3rd row and 3rd column to 0.3f, we achieve a level of opacity
            //The result is a watermark which slightly shows the underlying image
            float[][] colorMatrixElements = {
                                               new float[] {1.0f,  0.0f,  0.0f,  0.0f, 0.0f},
                                               new float[] {0.0f,  1.0f,  0.0f,  0.0f, 0.0f},
                                               new float[] {0.0f,  0.0f,  1.0f,  0.0f, 0.0f},
                                               new float[] {0.0f,  0.0f,  0.0f,  0.3f, 0.0f},
                                               new float[] {0.0f,  0.0f,  0.0f,  0.0f, 1.0f}
                                            };

            ColorMatrix wmColorMatrix = new ColorMatrix(colorMatrixElements);

            imageAttributes.SetColorMatrix(wmColorMatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);


            //With both color manipulations added to the imageAttributes object, we can now draw the watermark in the upper right hand corner of the photograph

            //Offset the image 10 pixels down and 10 pixels to the left
            int xPosOfWm = ((phWidth - wmWidth) - 10);
            int yPosOfWm = 10;

            grWatermark.DrawImage(imgWatermark,
                                  new Rectangle(xPosOfWm, yPosOfWm, wmWidth, wmHeight),
                                  0,
                                  0,
                                  wmWidth,
                                  wmHeight,
                                  GraphicsUnit.Pixel,
                                  imageAttributes);

            imgPhoto = bmWatermark;
            grPhoto.Dispose();
            grWatermark.Dispose();

            //imgPhoto.Save(WorkingDirectory + "watermark_final.jpg", ImageFormat.Jpeg);
            //Use Memory Stream to display image
            MemoryStream msDemo = new MemoryStream();
            imgPhoto.Save(msDemo, System.Drawing.Imaging.ImageFormat.Png);
            ViewBag.QrCodeUri = "data:image/png;base64," + Convert.ToBase64String(msDemo.ToArray());
            ViewBag.linkDownload = "data:image/png;base64," + Convert.ToBase64String(msDemo.ToArray());
            msDemo.Close();
            msDemo.Flush();
            msDemo.Dispose();
            grPhoto.Dispose();
            imgPhoto.Dispose();
            imgPhoto.Dispose();
            imgWatermark.Dispose();
            #endregion

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
        public string CreateQRCodeJob(string stringCode)
        {
            if (string.IsNullOrEmpty(stringCode)) return "";

            try
            {
                #region Generate QR Image
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(stringCode, QRCodeGenerator.ECCLevel.Q);
                QRCode qrCode = new QRCode(qrCodeData);
                Bitmap qrCodeBitmap = qrCode.GetGraphic(5);
                #endregion

                #region Create QR Image with Background Image (Watermark Image)
                //Define location
                var _environment = new SelfDefinedEnvironment();
                _environment.WebRootPath = @"D:\\2_demo_project\\DemoWatermark_dotNET4dot8\\Watermark-.Net4.8\\DemoWatermark_dotNET4dot8\\wwwroot";
                _environment.ContentRootPath = @"D:\\2_demo_project\\DemoWatermark_dotNET4dot8\\Watermark-.Net4.8\\DemoWatermark_dotNET4dot8";

                Image backgorundImage = Image.FromFile(Path.Combine(_environment.WebRootPath, "assets/img/voucherForm.png"));

                Graphics outputDemo = Graphics.FromImage(backgorundImage);
                outputDemo.DrawImage(qrCodeBitmap, 50, 50);
                #endregion

                #region Use memorystream to display image
                using (MemoryStream msDemo = new MemoryStream())
                {
                    backgorundImage.Save(msDemo, System.Drawing.Imaging.ImageFormat.Png);
                    return ("data:image/png;base64," + Convert.ToBase64String(msDemo.ToArray()));
                }
                #endregion
            }
            catch
            {
                return "";
            }
        }
        public string UploadFileThenGenQR()
        {
            //Define location
            var _environment = new SelfDefinedEnvironment();
            _environment.WebRootPath = @"D:\\2_demo_project\\DemoWatermark_dotNET4dot8\\Watermark-.Net4.8\\DemoWatermark_dotNET4dot8\\wwwroot";
            _environment.ContentRootPath = @"D:\\2_demo_project\\DemoWatermark_dotNET4dot8\\Watermark-.Net4.8\\DemoWatermark_dotNET4dot8";
            StringBuilder lstDisplay = new StringBuilder();

            if (Request.Files.Count > 0)
            {
                try
                {
                    #region Upload file
                    //  Get all files from Request object  
                    HttpFileCollectionBase files = Request.Files;
                    var path = Path.Combine(_environment.WebRootPath, files[0].FileName);

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
                    dt = ReadImportExcelFile("Sheet1", Path.Combine(_environment.WebRootPath, file.FileName)); //Get Excel file with static path(saved above)

                    List<GenerateQRCodeModel> lstCode = (from DataRow row in dt.Rows select new GenerateQRCodeModel { QRCodeText = row["Code"].ToString() }).ToList(); //Select data in Excel ~ Table
                    #endregion

                    #region Generate QR list
                    var rowNum = 0;

                    //Get list QR
                    foreach (var code in lstCode)
                    {
                        rowNum++;
                        // Generate QR Image
                        QRCodeGenerator qrGenerator = new QRCodeGenerator();
                        QRCodeData qrCodeData = qrGenerator.CreateQrCode(code.QRCodeText, QRCodeGenerator.ECCLevel.Q);
                        QRCode qrCode = new QRCode(qrCodeData);
                        Bitmap qrCodeBitmap = qrCode.GetGraphic(5);

                        Image backgorundImage = Image.FromFile(Path.Combine(_environment.WebRootPath, "assets/img/voucherForm.png"));

                        Graphics outputDemo = Graphics.FromImage(backgorundImage);
                        outputDemo.DrawImage(qrCodeBitmap, 50, 50);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            if (backgorundImage != null)
                            {
                                backgorundImage.Save(ms, ImageFormat.Png);
                                var base64img = "data:image/png;base64," + Convert.ToBase64String(ms.ToArray());

                                lstDisplay.Append("<div class='col-lg-3 col-md-4 col-sm-6 box-qr'>");
                                lstDisplay.Append("<div class='pos-center'>");
                                lstDisplay.Append("<img src= '" + base64img + "' />");
                                lstDisplay.Append("</div>");
                                lstDisplay.Append("<div class='pos-center'>");
                                lstDisplay.Append("<a id='download-'" + rowNum + "' download='qr-" + rowNum + "' href='" + base64img + "' class='btn btn-primary'>Download</a>");
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