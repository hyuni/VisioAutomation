using System.Collections.Generic;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Extensions
{

    public static class ShapeMethods
    {
        public static IVisio.IVShape DrawLine(this IVisio.IVShape shape, Drawing.Point p1, Drawing.Point p2)
        {
            return Shapes.ShapeHelper.DrawLine(shape, p1, p2);
        }

        public static IVisio.IVShape DrawQuarterArc(this IVisio.IVShape shape, Drawing.Point p0, Drawing.Point p1, IVisio.Enums.VisArcSweepFlags flags)
        {
            return Shapes.ShapeHelper.DrawQuarterArc(shape, p0, p1, flags);
        }

        public static Drawing.Rectangle GetBoundingBox(this IVisio.IVShape shape, IVisio.Enums.VisBoundingBoxArgs args)
        {
            return Shapes.ShapeHelper.GetBoundingBox(shape, args);
        }

        public static Drawing.Point XYFromPage(this IVisio.IVShape shape, Drawing.Point xy)
        {
            return Shapes.ShapeHelper.XYFromPage(shape, xy);
        }

        public static Drawing.Point XYToPage(this IVisio.IVShape shape, Drawing.Point xy)
        {
            return Shapes.ShapeHelper.XYToPage(shape, xy);
        }

        public static IEnumerable<IVisio.IVShape> ToEnumerable(this IVisio.IVShapes shapes)
        {
            return Shapes.ShapeHelper.ToEnumerable(shapes);
        }

        public static IList<IVisio.IVShape> GetShapesFromIDs(this IVisio.IVShapes shapes, IList<short> shapeids)
        {
            return Shapes.ShapeHelper.GetShapesFromIDs(shapes, shapeids);
        }
    }
}