using System.Collections.Generic;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Styles
{
    public static class StyleHelper
    {
        public static IEnumerable<IVisio.IVStyle> ToEnumerable(IVisio.IVStyles styles)
        {
            int count = styles.Count;
            for (int i = 0; i < count; i++)
            {
                yield return (IVisio.Style) styles[i + 1];
            }
        }

        public static string[] GetNamesU(IVisio.IVStyles styles)
        {
            string[] names_sa;
            styles.GetNamesU(out names_sa);
            string[] names = (string[])names_sa;
            return names;
        }
    }
}