using System.Web;
using System.Web.Mvc;

namespace Open_and_save_Word_document
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
