using System.Collections.Generic;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Extensions
{
    public static class StylesMethods
    {
        public static IEnumerable<IVisio.Style> ToEnumerable(this IVisio.Styles styles)
        {
            return VisioAutomation.Styles.StyleHelper.ToEnumerable(styles);
        }
        
        public static string[] GetNamesU(this IVisio.Styles styles)
        {
            return VisioAutomation.Styles.StyleHelper.GetNamesU(styles);
        }
    }
}