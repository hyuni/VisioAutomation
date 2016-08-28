﻿using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Drawing.Layout;
using VisioAutomation.Scripting.Layout;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Utilities
{
    internal static class ArrangeHelper
    {

        public static Drawing.Rectangle GetRectangle(Shapes.XFormCells xform)
        {
            var pin = new Drawing.Point(xform.PinX.Result, xform.PinY.Result);
            var locpin = new Drawing.Point(xform.LocPinX.Result, xform.LocPinY.Result);
            var size = new Drawing.Size(xform.Width.Result, xform.Height.Result);
            return new Drawing.Rectangle(pin - locpin, size);
        }

        private static double GetPositionOnShape(Shapes.XFormCells xform, RelativePosition pos)
        {
            switch (pos)
            {
                case RelativePosition.PinY:
                    return xform.PinY.Result;
                case RelativePosition.PinX:
                    return xform.PinX.Result;
            }

            var r = ArrangeHelper.GetRectangle(xform);

            switch (pos)
            {
                case RelativePosition.Left:
                    return r.Left;
                case RelativePosition.Right:
                    return r.Right;
                case RelativePosition.Top:
                    return r.Top;
                case RelativePosition.Bottom:
                    return r.Bottom;
            }

            throw new System.ArgumentOutOfRangeException(nameof(pos));
        }

        internal static IList<int> SortShapesByPosition(TargetShapeIDs targets, RelativePosition pos)
        {
            // First get the transforms of the shapes on the given axis
            var xforms = Shapes.XFormCells.GetCells(targets.Page, targets.ShapeIDs);

            // Then, sort the shapeids pased on the corresponding value in the results

            var sorted_shape_ids = Enumerable.Range(0, targets.ShapeIDs.Count)
                .Select(i => new { index = i, shapeid = targets.ShapeIDs[i], pos = ArrangeHelper.GetPositionOnShape(xforms[i], pos) })
                .OrderBy(i => i.pos)
                .Select(i => i.shapeid)
                .ToList();

            return sorted_shape_ids;
        }

        public static void DistributeWithSpacing(TargetShapeIDs target, Axis axis, double spacing)
        {
            if (spacing < 0.0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(spacing));
            }

            if (target.ShapeIDs.Count < 2)
            {
                return;
            }

            // Calculate the new Xfrms
            var sortpos = axis == Axis.XAxis
                ? RelativePosition.PinX
                : RelativePosition.PinY;

            var delta = axis == Axis.XAxis
                ? new Drawing.Size(spacing, 0)
                : new Drawing.Size(0, spacing);


            var sorted_shape_ids = ArrangeHelper.SortShapesByPosition(target, sortpos);
            var input_xfrms = Shapes.XFormCells.GetCells(target.Page,target.ShapeIDs);
            var output_xfrms = new List<Shapes.XFormCells>(input_xfrms.Count);
            var bb = ArrangeHelper.GetBoundingBox(input_xfrms);
            var cur_pos = new Drawing.Point(bb.Left, bb.Bottom);

            foreach (var input_xfrm in input_xfrms)
            {
                var new_pinpos = axis == Axis.XAxis
                    ? new Drawing.Point(cur_pos.X + input_xfrm.LocPinX.Result, input_xfrm.PinY.Result)
                    : new Drawing.Point(input_xfrm.PinX.Result, cur_pos.Y + input_xfrm.LocPinY.Result);

                var output_xfrm = new Shapes.XFormCells();
                output_xfrm.PinX = new_pinpos.X;
                output_xfrm.PinY = new_pinpos.Y;
                output_xfrms.Add(output_xfrm);

                cur_pos = cur_pos.Add(input_xfrm.Width.Result, input_xfrm.Height.Result).Add(delta);
            }

            // Apply the changes
            ArrangeHelper.update_xfrms(target, output_xfrms);
        }

 
        internal static void update_xfrms(TargetShapeIDs target, IList<Shapes.XFormCells> xfrms)
        {

            var update = new FormulaWriterSIDSRC();
            for (int i = 0; i < target.ShapeIDs.Count; i++)
            {
                var shape_id = target.ShapeIDs[i];
                var xfrm = xfrms[i];
                xfrm.SetFormulas((short)shape_id, update);
            }
            update.Commit(target.Page);
        }

        public static Drawing.Rectangle GetBoundingBox(IEnumerable<Shapes.XFormCells> xfrms)
        {
            var bb = new BoundingBox(xfrms.Select(ArrangeHelper.GetRectangle));
            if (!bb.HasValue)
            {
                throw new System.ArgumentException("Could not calculate bounding box");
            }
            return bb.Rectangle;
        }

        public static void SnapCorner(TargetShapeIDs target, Drawing.Size snapsize, SnapCornerPosition corner)
        {
            // First caculate the new transforms
            var snap_grid = new SnappingGrid(snapsize);
            var input_xfrms = Shapes.XFormCells.GetCells(target.Page,target.ShapeIDs);
            var output_xfrms = new List<Shapes.XFormCells>(input_xfrms.Count);

            foreach (var input_xfrm in input_xfrms)
            {
                var old_rect = ArrangeHelper.GetRectangle(input_xfrm);
                var old_lower_left = old_rect.LowerLeft;
                var new_lower_left = snap_grid.Snap(old_lower_left);
                var output_xfrm = ArrangeHelper._SnapCorner(corner, new_lower_left, input_xfrm);
                output_xfrms.Add(output_xfrm);
            }

            // Now apply them
            ArrangeHelper.update_xfrms(target, output_xfrms);
        }

        private static Shapes.XFormCells _SnapCorner(SnapCornerPosition corner, Drawing.Point new_lower_left, Shapes.XFormCells input_xfrm)
        {
            var new_pin_position = ArrangeHelper.GetPinPositionForCorner(input_xfrm, new_lower_left, corner);

            var output_xfrm = new Shapes.XFormCells();

            // Modify the PinX only if necessary
            if (new_pin_position.X != input_xfrm.PinX.Result)
            {
                output_xfrm.PinX = new_pin_position.X;
            }

            // Modify the PinY only if necessary
            if (new_pin_position.Y != input_xfrm.PinY.Result)
            {
                output_xfrm.PinY = new_pin_position.Y;
            }

            return output_xfrm;
        }

        private static Drawing.Point GetPinPositionForCorner(Shapes.XFormCells input_xfrm, Drawing.Point new_lower_left, SnapCornerPosition corner)
        {
            var size = new Drawing.Size(input_xfrm.Width.Result, input_xfrm.Height.Result);
            var locpin = new Drawing.Point(input_xfrm.LocPinX.Result, input_xfrm.LocPinY.Result);

            switch (corner)
            {
                case SnapCornerPosition.LowerLeft:
                    {
                        return new_lower_left.Add(locpin.X, locpin.Y);
                    }
                case SnapCornerPosition.UpperRight:
                    {
                        return new_lower_left.Subtract(size.Width, size.Height).Add(locpin.X, locpin.Y);
                    }
                case SnapCornerPosition.LowerRight:
                    {
                        return new_lower_left.Subtract(size.Width, 0).Add(locpin.X, locpin.Y);
                    }
                case SnapCornerPosition.UpperLeft:
                    {
                        return new_lower_left.Subtract(0, size.Height).Add(locpin.X, locpin.Y);
                    }
                default:
                    {
                        throw new System.ArgumentOutOfRangeException(nameof(corner), "Unsupported corner");
                    }
            }
        }

        public static void SnapSize(TargetShapeIDs target, Drawing.Size snapsize, Drawing.Size minsize)
        {
            var input_xfrms = Shapes.XFormCells.GetCells(target.Page, target.ShapeIDs);
            var output_xfrms = new List<Shapes.XFormCells>(input_xfrms.Count);

            var grid = new SnappingGrid(snapsize);
            foreach (var input_xfrm in input_xfrms)
            {
                // First snap the size to the grid
                double old_w = input_xfrm.Width.Result;
                double old_h = input_xfrm.Height.Result;
                var input_size = new Drawing.Size(old_w, old_h);
                var snapped_size = grid.Snap(input_size);

                // then account for any minum size requirements
                double new_w = System.Math.Max(snapped_size.Width, minsize.Width);
                double new_h = System.Math.Max(snapped_size.Height, minsize.Height);
                var output_size = new Drawing.Size(new_w, new_h);
                
                // Output the new size for the shape if the size of the shape changed
                bool different_widths = (old_w != new_w);
                bool different_heights = (old_h != new_h);
                if (different_widths || different_heights)
                {
                    var output_xfrm = new Shapes.XFormCells();
                    if (different_widths) 
                    {
                        output_xfrm.Width = output_size.Width;                    
                    }
                    if (different_heights)
                    {
                        output_xfrm.Height = output_size.Height;
                    }
                    output_xfrms.Add(output_xfrm);
                }
            }

            // Now apply the updates to the sizes
            ArrangeHelper.update_xfrms(target, output_xfrms);
        }
    }
}