using DemoWatermark_dotNET4dot8.Models;
using DevExpress.Utils.CommonDialogs.Internal;
using Microsoft.AspNetCore.Http;
using Microsoft.Win32;
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
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace DemoWatermark_dotNET4dot8.Controllers
{
    public class HomeController : Controller
    {
        private static Byte[] BitmapToBytes(Bitmap img)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                img.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                return stream.ToArray();
            }
        }
        private static ImageCodecInfo GetEncoderInfo(String mimeType)
        {
            int j;
            ImageCodecInfo[] encoders;
            encoders = ImageCodecInfo.GetImageEncoders();
            for (j = 0; j < encoders.Length; ++j)
            {
                if (encoders[j].MimeType == mimeType)
                    return encoders[j];
            }
            return null;
        }
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

            #region Generate QR code
            
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
        public string CreateQRCodeJob(string stringCode)
        {
            #region Initial data
            var _environment = new SelfDefinedEnvironment();
            _environment.WebRootPath = @"D:\\2_demo_project\\DemoWatermark_dotNET4dot8\\Watermark-.Net4.8\\DemoWatermark_dotNET4dot8\\wwwroot";
            _environment.ContentRootPath = @"D:\\2_demo_project\\DemoWatermark_dotNET4dot8\\Watermark-.Net4.8\\DemoWatermark_dotNET4dot8";
            #endregion

            if (string.IsNullOrEmpty(stringCode))
            {
                return "";
            }

            try
            {
                #region Generate QR Image
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(stringCode, QRCodeGenerator.ECCLevel.Q);
                QRCode qrCode = new QRCode(qrCodeData);
                Bitmap qrCodeBitmap = qrCode.GetGraphic(5);
                //var byteImg = BitmapToBytes(qrCodeBitmap); //get byte data from image
                #endregion

                #region Save file to memory stream(return QR only)
                //using (MemoryStream ms = new MemoryStream())
                //{
                //    qrCodeBitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                //    ViewBag.QrCodeUri = "data:image/png;base64," + Convert.ToBase64String(ms.ToArray());
                //    ViewBag.linkDownload = "data:image/png;base64," + Convert.ToBase64String(ms.ToArray());
                //}
                //return View();
                #endregion

                #region Save Bitmap as jpg image
                //ImageCodecInfo myImageCodecInfo;
                //Encoder myEncoder;
                //EncoderParameter myEncoderParameter;
                //EncoderParameters myEncoderParameters;

                //// Get an ImageCodecInfo object that represents the JPEG codec.
                //myImageCodecInfo = GetEncoderInfo("image/jpeg");

                //// Create an Encoder object based on the GUID

                //// for the Quality parameter category.
                //myEncoder = Encoder.Quality;

                //// Create an EncoderParameters object.

                //// An EncoderParameters object has an array of EncoderParameter

                //// objects. In this case, there is only one

                //// EncoderParameter object in the array.
                //myEncoderParameters = new EncoderParameters(1);

                //// Save the bitmap as a JPEG file with quality level 75.
                //myEncoderParameter = new EncoderParameter(myEncoder, 75L); // level can be changed to 25L, 50L
                //myEncoderParameters.Param[0] = myEncoderParameter;

                //qrCodeBitmap.Save("qr_" + stringCode + "_L075.jpg", myImageCodecInfo, myEncoderParameters);
                #endregion

                #region Save QR image to Server
                string path = Path.Combine(_environment.WebRootPath, "GeneratedQRCode");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                string filePath = Path.Combine(_environment.WebRootPath, "GeneratedQRCode/qrcode" + stringCode + ".png");

                //render file if it was generated(same text) before
                if (System.IO.File.Exists(filePath))
                {
                    string imgExistedemoUrl = Path.Combine(_environment.WebRootPath, "assets/img/tickets/ticket_" + stringCode + ".png");
                    MemoryStream ms = new MemoryStream();

                    Image image = Image.FromFile(imgExistedemoUrl);
                    image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);

                    return ("data:image/png;base64," + Convert.ToBase64String(ms.ToArray()));
                };

                //save image to file
                qrCodeBitmap.Save(filePath, ImageFormat.Png);
                #endregion

                #region Create QR Image with Background Image (Watermark Image)
                Image backgorundImage = Image.FromFile(Path.Combine(_environment.WebRootPath, "assets/img/voucherForm.png"));
                
                Image imageQR = Image.FromFile(Path.Combine(_environment.WebRootPath, "GeneratedQRCode/qrcode" + stringCode + ".png"));
                Graphics outputDemo = Graphics.FromImage(backgorundImage);
                //outputDemo.DrawImage(imageQR, backgorundImage.Width / 2 + 305, backgorundImage.Height / 2 + 105);
                outputDemo.DrawImage(imageQR, 50, 50);

                #region render image if existed
                if (System.IO.File.Exists("assets/img/tickets/ticket_" + stringCode + ".png"))
                {
                    string imgExistedemoUrl = Path.Combine(_environment.ContentRootPath, "assets/img/tickets/ticket_" + stringCode + ".png");
                    MemoryStream ms = new MemoryStream();

                    Image image = Image.FromFile(imgExistedemoUrl);
                    image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);

                    return ("data:image/png;base64," + Convert.ToBase64String(ms.ToArray()));
                };
                #endregion

                string imgFilePath = Path.Combine(_environment.WebRootPath, "assets/img/tickets/ticket_" + stringCode + ".png");

                backgorundImage.Save(imgFilePath);

                MemoryStream msDemo = new MemoryStream();
                Image img = Image.FromFile(imgFilePath);
                img.Save(msDemo, System.Drawing.Imaging.ImageFormat.Png);

                msDemo.Close();
                msDemo.Flush();
                msDemo.Dispose();
                imageQR.Dispose();
                backgorundImage.Dispose();
                #endregion

                return ("data:image/png;base64," + Convert.ToBase64String(msDemo.ToArray()));
            }
            catch
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
        public async Task<ActionResult> UploadFileThenGenQR(HttpPostedFileBase file)
        {
            #region Initial data
            var _environment = new SelfDefinedEnvironment();
            _environment.WebRootPath = @"D:\\2_demo_project\\DemoWatermark_dotNET4dot8\\Watermark-.Net4.8\\DemoWatermark_dotNET4dot8\\wwwroot";
            _environment.ContentRootPath = @"D:\\2_demo_project\\DemoWatermark_dotNET4dot8\\Watermark-.Net4.8\\DemoWatermark_dotNET4dot8";
            #endregion

            #region Upload file
            string fileExtension = Path.GetExtension(file.FileName);

            var path = Path.Combine(_environment.WebRootPath, file.FileName);

            file.SaveAs(path);
            #endregion

            #region Generate QR code list
            try
            {
                #region Read Data From imported file
                DataTable dt = new DataTable();
                dt = ReadImportExcelFile("Sheet1", Path.Combine(_environment.WebRootPath, file.FileName)); //Get Excel file with static path

                List<DataRow> list = dt.AsEnumerable().ToList();
                List<GenerateQRCodeModel> lstCode = (from DataRow row in dt.Rows
                                                     select new GenerateQRCodeModel
                                                     {
                                                         QRCodeText = row["Code"].ToString()
                                                     }
                                                    ).ToList();
                #endregion

                #region Generate QR list
                var rowNum = 0;

                List<DisplayingCodeInfo> lstDisplay = new List<DisplayingCodeInfo>();
                foreach (var code in lstCode)
                {
                    rowNum++;

                    #region Generate QR Image
                    QRCodeGenerator qrGenerator = new QRCodeGenerator();
                    QRCodeData qrCodeData = qrGenerator.CreateQrCode(code.QRCodeText, QRCodeGenerator.ECCLevel.Q);
                    QRCode qrCode = new QRCode(qrCodeData);
                    Bitmap qrCodeBitmap = qrCode.GetGraphic(5);
                    #endregion

                    string imageQRPath = Path.Combine(_environment.WebRootPath, "GeneratedQRCode");
                    if (!Directory.Exists(imageQRPath))
                    {
                        Directory.CreateDirectory(imageQRPath);
                    }

                    string filePath = Path.Combine(_environment.WebRootPath, "GeneratedQRCode/qrcode_" + code.QRCodeText + ".png");

                    if (!System.IO.File.Exists(filePath))
                    {
                        qrCodeBitmap.Save(filePath, ImageFormat.Png); // Save image
                    };

                    WebClient client = new WebClient();
                    Stream stream = client.OpenRead(filePath);

                    using (MemoryStream ms = new MemoryStream())
                    {
                        using (Bitmap bitMap = new Bitmap(stream))
                        {
                            if (bitMap != null)
                            {
                                bitMap.Save(ms, ImageFormat.Png);
                            }
                            var qrResult = "data:image/png;base64," + Convert.ToBase64String(ms.ToArray());
                            lstDisplay.Add(new DisplayingCodeInfo()
                            {
                                No = rowNum,
                                QRCodeUri = qrResult,
                                LinkDownload = qrResult
                            });
                        }
                    }
                    //Remove file uploaded
                    System.IO.File.Delete(path);
                }

                ViewBag.listDisplay = lstDisplay;
                #endregion

                return View();
            }
            catch(Exception e)
            {
                ViewBag.Error = e.Message;
                return View("~/Views/Shared/Error.cshtml");
            }
            #endregion
        }
    }
}