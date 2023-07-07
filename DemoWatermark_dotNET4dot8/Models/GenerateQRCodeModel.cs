using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace DemoWatermark_dotNET4dot8.Models
{
    public class GenerateQRCodeModel
    {
        [Display(Name = "Enter QR Code Text")]
        public string QRCodeText { get; set; }
    }

    public class DisplayingCodeInfo
    {
        public int No { get; set; }
        public string QRCodeUri { get; set; }
        public string LinkDownload { get; set; }
    }

    public class SelfDefinedEnvironment
    {
        public string WebRootPath { get; set; }
        public string ContentRootPath { get; set; }
    }

    public class ItemInfo
    {
        public string fileName { get; set; }
        public string fileData { get; set; }
    }
}