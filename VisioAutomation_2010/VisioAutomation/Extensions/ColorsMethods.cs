using System.Collections.Generic;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Extensions
{
    public static class ColorsMethods
    {
        public static IEnumerable<IVisio.Color> ToEnumerable(this IVisio.Colors colors)
        {
            return VisioAutomation.Colors.ColorHelper.ToEnumerable(colors);
        }
    }
}
