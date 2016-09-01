using System.Collections.Generic;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Extensions
{
    public static class SelectionMethods
    {
        public static IEnumerable<IVisio.IVShape> ToEnumerable(this IVisio.IVSelection selection)
        {
            return VisioAutomation.Selections.SelectionHelper.ToEnumerable(selection);
        }
        
        public static Drawing.Rectangle GetBoundingBox(this IVisio.IVSelection selection, IVisio.Enums.VisBoundingBoxArgs args)
        {
            return VisioAutomation.Selections.SelectionHelper.GetBoundingBox(selection, args);
        }

        public static int[] GetIDs(this IVisio.IVSelection selection)
        {
            return VisioAutomation.Selections.SelectionHelper.GetIDs(selection);
        }
    }
}