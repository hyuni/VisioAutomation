using System.Collections.Generic;
using IVisio = NetOffice.VisioApi;
using VisioAutomation.Drawing;

namespace VisioAutomation.Extensions
{
    public static class PageMethods
    {
        public static void ResizeToFitContents(this IVisio.IVPage page, Drawing.Size padding)
        {
            Pages.PageHelper.ResizeToFitContents(page, padding);
        }

        public static IVisio.IVShape DrawLine(this IVisio.IVPage page, Drawing.Point p1, Drawing.Point p2)
        {
            return VisioAutomation.Pages.PageHelper.DrawLine(page, p1, p2);
        }

        public static IVisio.IVShape DrawOval(this IVisio.IVPage page, Drawing.Rectangle rect)
        {
            return VisioAutomation.Pages.PageHelper.DrawOval(page, rect);
        }

        public static IVisio.IVShape DrawRectangle(this IVisio.IVPage page, Drawing.Rectangle rect)
        {
            return VisioAutomation.Pages.PageHelper.DrawRectangle(page, rect);
        }

        public static IVisio.IVShape DrawBezier(this IVisio.IVPage page, IList<Drawing.Point> points)
        {
            return VisioAutomation.Pages.PageHelper.DrawBezier(page, points);
        }

        public static IVisio.IVShape DrawBezier(this IVisio.IVPage page, IList<Drawing.Point> points, short degree, short flags)
        {
            return VisioAutomation.Pages.PageHelper.DrawBezier(page, points, degree, flags);
        }

        public static IVisio.IVShape DrawPolyline(this IVisio.IVPage page, IList<Drawing.Point> points)
        {
            return VisioAutomation.Pages.PageHelper.DrawPolyline(page, points);
        }

        public static IVisio.IVShape DrawNURBS(this IVisio.IVPage page, IList<Drawing.Point> controlpoints,
                                             IList<double> knots,
                                             IList<double> weights, int degree)
        {
            return VisioAutomation.Pages.PageHelper.DrawNURBS(page, controlpoints, knots, weights, degree);
        }

        public static IVisio.IVShape Drop(
            this IVisio.IVPage page,
            IVisio.IVMaster master,
            Drawing.Point point)
        {
            return VisioAutomation.Pages.PageHelper.Drop(page, master, point);
        }

        public static short[] DropManyU(
            this IVisio.IVPage page,
            IList<IVisio.IVMaster> masters,
            IEnumerable<Drawing.Point> points)
        {
            // TODO: Put this method in pagehelper class
            var surface = new Drawing.DrawingSurface(page);
            short[] shapeids = surface.DropManyU(masters, points);
            return shapeids;
        }

   	    public static short[] DropManyU(this IVisio.IVPage page, IList<IVisio.IVMaster> masters, IEnumerable<Point> points, IList<string> names)
   	    {
   	        return VisioAutomation.Pages.PageHelper.DropManyU(page, masters, points, names);
        }

        public static IEnumerable<IVisio.IVPage> ToEnumerable(this IVisio.IVPages pages)
        {
            return VisioAutomation.Pages.PageHelper.ToEnumerable(pages);
        }

        public static string[] GetNamesU(this IVisio.IVPages pages)
        {
            return VisioAutomation.Pages.PageHelper.GetNamesU(pages);
        }
    }
}