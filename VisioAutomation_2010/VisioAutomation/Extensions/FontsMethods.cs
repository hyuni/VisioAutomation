using System.Collections.Generic;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Extensions
{
    public static class FontsMethods
    {
        public static IEnumerable<IVisio.Font> ToEnumerable(this IVisio.Fonts fonts)
        {
            return VisioAutomation.Fonts.FontHelper.ToEnumerable(fonts);
        }
    }
}