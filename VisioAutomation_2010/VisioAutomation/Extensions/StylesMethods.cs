using System.Collections.Generic;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Extensions
{
    public static class StylesMethods
    {
        public static IEnumerable<IVisio.IVStyle> ToEnumerable(this IVisio.IVStyles styles)
        {
            return VisioAutomation.Styles.StyleHelper.ToEnumerable(styles);
        }
        
        public static string[] GetNamesU(this IVisio.IVStyles styles)
        {
            return VisioAutomation.Styles.StyleHelper.GetNamesU(styles);
        }
    }
}