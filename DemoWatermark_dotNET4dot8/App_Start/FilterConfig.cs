using System.Web;
using System.Web.Mvc;

namespace DemoWatermark_dotNET4dot8
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
