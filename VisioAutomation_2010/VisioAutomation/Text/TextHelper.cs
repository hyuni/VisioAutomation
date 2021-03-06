using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public static class TextHelper
    {
        private const short tab_section = (short)IVisio.VisSectionIndices.visSectionTab;

        public static List<TabStop> GetTabStops(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            int num_stops = GetTabStopCount(shape);

            if (num_stops < 1)
            {
                return new List<TabStop>(0);
            }

            const short row = 0;
            
            var stream = new VisioAutomation.ShapeSheet.Streams.SrcStreamBuilder(num_stops * 3);
            for (int stop_index = 0; stop_index < num_stops; stop_index++)
            {
                int i = stop_index * 3;

                var src_tabpos = new ShapeSheet.Src(tab_section, row, (short)(i + 1));
                var src_tabalign = new ShapeSheet.Src(tab_section, row, (short)(i + 2));
                var src_tabother = new ShapeSheet.Src(tab_section, row, (short)(i + 3));

                stream.Add(src_tabpos);
                stream.Add(src_tabalign);
                stream.Add(src_tabother);
            }

            var surface = new SurfaceTarget(shape);

            const object[] unitcodes = null;

            var results = surface.GetResults<double>(stream.ToStream(), unitcodes);

            var stops_list = new List<TabStop>(num_stops);
            for (int stop_index = 0; stop_index < num_stops; stop_index++)
            {
                var pos = results[(stop_index * 3) + 1];
                var align = (TabStopAlignment)((int)results[(stop_index * 3) + 2]);
                var ts = new TabStop(pos, align);
                stops_list.Add(ts);
            }

            return stops_list;
        }

        public static void SetTabStops(IVisio.Shape shape, IList<TabStop> stops)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            if (stops == null)
            {
                throw new System.ArgumentNullException(nameof(stops));
            }

            ClearTabStops(shape);
            if (stops.Count < 1)
            {
                return;
            }

            const short row = 0;
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            var vis_tab_stop_count = (short)IVisio.VisCellIndices.visTabStopCount;
            var tabstopcountcell = shape.CellsSRC[tab_section, row, vis_tab_stop_count];
            tabstopcountcell.FormulaU = stops.Count.ToString(culture);

            // set the number of tab stobs allowed for the shape
            var tagtab = GetTabTagForStops(stops.Count);
            shape.RowType[tab_section, (short)IVisio.VisRowIndices.visRowTab] = (short)tagtab;

            // add tab properties for each stop
            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            for (int stop_index = 0; stop_index < stops.Count; stop_index++)
            {
                int i = stop_index * 3;

                var alignment = ((int)stops[stop_index].Alignment).ToString(culture);
                var position = ((int)stops[stop_index].Position).ToString(culture);

                var src_tabpos = new ShapeSheet.Src(tab_section, row, (short)(i + 1));
                var src_tabalign = new ShapeSheet.Src(tab_section, row, (short)(i + 2));
                var src_tabother = new ShapeSheet.Src(tab_section, row, (short)(i + 3));

                writer.SetFormula(src_tabpos, position); // tab position
                writer.SetFormula(src_tabalign, alignment); // tab alignment
                writer.SetFormula(src_tabother, "0"); // tab unknown
            }

            writer.Commit(shape);
        }

        private static IVisio.VisRowTags GetTabTagForStops(int stops)
        {
            if (stops < 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(stops));
            }

            var tagtab = IVisio.VisRowTags.visTagTab0;
            if ((0 <= stops) && (stops <= 2))
            {
                tagtab = IVisio.VisRowTags.visTagTab2;
            }
            else if ((3 <= stops) && (stops <= 10))
            {
                tagtab = IVisio.VisRowTags.visTagTab10;
            }
            else if ((11 <= stops) && (stops <= 60))
            {
                tagtab = IVisio.VisRowTags.visTagTab60;
            }
            else
            {
                throw new System.ArgumentOutOfRangeException(nameof(stops), "unsupported number of tabs");
            }

            return tagtab;
        }

        private static int GetTabStopCount(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            var cell_tabstopcount = shape.CellsSRC[ShapeSheet.SrcConstants.TabStopCount.Section, ShapeSheet.SrcConstants.TabStopCount.Row, ShapeSheet.SrcConstants.TabStopCount.Cell];
            const short rounding = 0;

            return cell_tabstopcount.ResultInt[(short)IVisio.VisUnitCodes.visNumber, rounding];
        }

        /// <summary>
        /// Remove all tab stops on the shape
        /// </summary>
        /// <param name="shape"></param>
        private static void ClearTabStops(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            int num_existing_tabstops = GetTabStopCount(shape);

            if (num_existing_tabstops < 1)
            {
                return;
            }

            var cell_tabstopcount = shape.CellsSRC[ShapeSheet.SrcConstants.TabStopCount.Section, ShapeSheet.SrcConstants.TabStopCount.Row, ShapeSheet.SrcConstants.TabStopCount.Cell];
            cell_tabstopcount.FormulaForce = "0";

            const string formula = "0";

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            for (int i = 1; i < num_existing_tabstops * 3; i++)
            {
                var src = new ShapeSheet.Src(tab_section, (short)IVisio.VisRowIndices.visRowTab,
                    (short)i);
                writer.SetFormula(src, formula);
            }

            writer.Commit(shape);
        }
    }
}