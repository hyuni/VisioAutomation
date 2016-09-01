using System;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Shapes.Controls
{
    public static class ControlHelper
    {
        public static int Add(IVisio.IVShape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            var ctrl = new ControlCells();

            return ControlHelper.Add(shape, ctrl);
        }

        public static int Add(
            IVisio.IVShape shape,
            ControlCells ctrl)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            short row = shape.AddRow((short)IVisio.Enums.VisSectionIndices.visSectionControls,
                                     (short)IVisio.Enums.VisRowIndices.visRowLast,
                                     (short)IVisio.Enums.VisRowTags.visTagDefault);

            ControlHelper.Set(shape, row, ctrl);

            return row;
        }

        public static int Set(
            IVisio.IVShape shape,
            short row,
            ControlCells ctrl)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }


            if (!ctrl.XDynamics.Formula.HasValue)
            {
                ctrl.XDynamics = string.Format("Controls.Row_{0}", row + 1);
            }

            if (!ctrl.YDynamics.Formula.HasValue)
            {
                ctrl.YDynamics = string.Format("Controls.Row_{0}.Y", row + 1);
            }

            var writer = new FormulaWriterSRC();
            ctrl.SetFormulas(writer, row);
            writer.Commit(shape);

            return row;
        }

        public static void Delete(IVisio.IVShape shape, int index)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            var row = (IVisio.Enums.VisRowIndices)index;
            shape.DeleteRow( (short) IVisio.Enums.VisSectionIndices.visSectionControls, (short)row);
        }

        public static int GetCount(IVisio.IVShape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            return shape.get_RowCount((short)IVisio.Enums.VisSectionIndices.visSectionControls);
        }
    }
}