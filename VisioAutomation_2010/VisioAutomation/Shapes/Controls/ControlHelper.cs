using System;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VACONTROL = VisioAutomation.Shapes.Controls;

namespace VisioAutomation.Shapes.Controls
{
    public static class ControlHelper
    {
        public static int Add(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            var ctrl = new ControlCells();

            return ControlHelper.Add(shape, ctrl);
        }

        public static int Add(
            IVisio.Shape shape,
            ControlCells ctrl)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            short row = shape.AddRow((short)IVisio.VisSectionIndices.visSectionControls,
                                     (short)IVisio.VisRowIndices.visRowLast,
                                     (short)IVisio.VisRowTags.visTagDefault);

            ControlHelper.Set(shape, row, ctrl);

            return row;
        }

        public static int Set(
            IVisio.Shape shape,
            short row,
            ControlCells ctrl)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }


            if (!ctrl.XDynamics.Formula.HasValue)
            {
                ctrl.XDynamics = $"Controls.Row_{row + 1}";
            }

            if (!ctrl.YDynamics.Formula.HasValue)
            {
                ctrl.YDynamics = $"Controls.Row_{row + 1}.Y";
            }

            var update = new ShapeSheet.Update();
            update.SetFormulas(ctrl, row);
            update.Execute(shape);

            return row;
        }

        public static void Delete(IVisio.Shape shape, int index)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            var row = (IVisio.VisRowIndices)index;
            shape.DeleteRow( (short) IVisio.VisSectionIndices.visSectionControls, (short)row);
        }

        public static int GetCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            return shape.RowCount[(short)IVisio.VisSectionIndices.visSectionControls];
        }
    }
}