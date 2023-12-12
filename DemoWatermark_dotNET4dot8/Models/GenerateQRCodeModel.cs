namespace DemoWatermark_dotNET4dot8.Models
{
    public class GenerateQRCodeModel
    {
        public string title { get; set; }
        public string code { get; set; }
        public string status { get; set; }
        public string price { get; set; }
        public string expDate { get; set; }

    }

    public class DisplayingCodeInfo
    {
        public int No { get; set; }
        public string QRCodeUri { get; set; }
        public string LinkDownload { get; set; }
    }

    public class ItemInfo
    {
        public string fileName { get; set; }
        public string fileData { get; set; }
    }
}