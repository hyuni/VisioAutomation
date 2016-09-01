using System.Collections.Generic;
using System.Linq;
using IVisio = NetOffice.VisioApi;
using VisioAutomation.Extensions;

namespace VisioAutomation.Models.Charting
{
    internal static class ChartUtil
    {
        public static void GroupShapesIfNeeded(IVisio.IVPage page, List<IVisio.IVShape> shapes)
        {
            if (shapes.Count > 0)
            {
                var app = page.Application;
                var win = app.ActiveWindow;
                win.DeselectAll();
                win.DeselectAll();

                // TODO: NETOFFICE
                foreach (var shape in shapes)
                {
                    win.Select(shape, (short)IVisio.Enums.VisSelectArgs.visSelect);
                }
                var sel = win.Selection;
                sel.Group();
            }
        }

        public static double GetValueRangeDistance(double min, double max)
        {
            double range = -1;

            if (min < 0)
            {
                if (max >= 0)
                {
                    range = max - min;
                }
                else
                {
                    range = System.Math.Abs(min);
                }
            }
            else
            {
                range = max;
            }
            return range;
        }


    }
}