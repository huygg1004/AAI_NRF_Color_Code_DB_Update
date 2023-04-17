using System.Web;
using System.Web.Mvc;

namespace AAI_NRF_Color_Code_DB_Update
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
