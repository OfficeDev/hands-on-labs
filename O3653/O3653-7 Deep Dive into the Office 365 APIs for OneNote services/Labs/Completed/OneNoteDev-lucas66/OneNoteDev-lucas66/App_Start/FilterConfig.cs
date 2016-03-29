using System.Web;
using System.Web.Mvc;

namespace OneNoteDev_lucas66
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
