using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioAutomation.Text
{
    public static class TextHelper
    {
        public static void SetTextFormatFields(IVisio.Shape shape, string fmt, params object[] fields)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            if (fields == null)
            {
                throw new ArgumentNullException("fields");
            }

            var t_string = typeof (string);
            var t_field_element = typeof (VA.Text.Markup.Field);

            for (int i = 0; i < fields.Length; i++)
            {
                object field = fields[i];
                if (field == null)
                {
                    string msg = String.Format("Field value {0} is null", i);
                    throw new ArgumentException(msg);
                }

                var ft = field.GetType();

                if ((ft != t_string) && (ft != t_field_element))
                {
                    string msg = String.Format("Field value {0} is must be {1} or {2}. Instead it is {3}", i,
                                               t_string.Name, t_field_element.Name, ft.Name);
                    throw new ArgumentException(msg);
                }
            }

            var fmtparse = new VA.Internal.FormatStringParser(fmt);
            var unique_indices = fmtparse.Segments.Select(f => f.Index).Distinct().ToList();
            if (unique_indices.Count > fields.Length)
            {
                throw new ArgumentOutOfRangeException("fmt", "index out of range for number of insertions");
            }

            // Set the text
            shape.Text = fmt;

            // then Insert the fields from back to front
            for (int i = (fmtparse.Segments.Count - 1); i >= 0; i--)
            {
                var fmt_seg = fmtparse.Segments[i];
                var field_index = fmt_seg.Index;
                object field = fields[field_index];

                var chars = shape.Characters;
                chars.Begin = fmt_seg.Start;
                chars.End = fmt_seg.End;

                var ft = field.GetType();
                if (t_string == ft)
                {
                    // it must be a formula
                    string formula = (string) field;
                    chars.AddCustomFieldU(formula, (short) IVisio.VisFieldFormats.visFmtNumGenNoUnits);
                }
                else if (t_field_element == ft)
                {
                    var field_f = (VA.Text.Markup.Field) field;
                    chars.AddField((short) field_f.Category, (short) field_f.Code, (short) field_f.Format);
                }
                else
                {
                    string msg = String.Format("Unsupported field type {0} for field {1}", ft.Name, i);
                    throw new AutomationException(msg);
                }
            }
        }


        /// <summary>
        /// Tests whether a font is available to the Visio application. The method is not case sensitive
        /// </summary>
        /// <param name="fonts">Visio Fonts Object</param>
        /// <param name="fontname">the name of the font to find.</param>
        /// <returns>null if the font cannot be found, otherwise the font object</returns>
        public static IVisio.Font FindFontWithName(IVisio.Fonts fonts, string fontname)
        {
            if (fontname == null)
            {
                throw new ArgumentNullException("fontname");
            }

            if (String.IsNullOrEmpty(fontname))
            {
                throw new ArgumentException("fontname");
            }

            foreach (var f in fonts.AsEnumerable())
            {
                if (String.Compare(f.Name, fontname, StringComparison.CurrentCultureIgnoreCase) == 0)
                {
                    return f;
                }
            }

            return null;
        }


        public static void FitShapeToText(IVisio.Page page, IEnumerable<IVisio.Shape> shapes)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            if (shapes == null)
            {
                throw new System.ArgumentNullException("shapes");
            }

            var shapeids = shapes.Select(s => s.ID).ToList();

            // Calculate the new sizes for each shape
            var new_sizes = new List<VA.Drawing.Size>(shapeids.Count);
            foreach (var shape in shapes)
            {
                var text_bounding_box = shape.GetBoundingBox(IVisio.VisBoundingBoxArgs.visBBoxUprightText).Size;
                var wh_bounding_box = shape.GetBoundingBox(IVisio.VisBoundingBoxArgs.visBBoxUprightWH).Size;

                double max_w = System.Math.Max(text_bounding_box.Width, wh_bounding_box.Width);
                double max_h = System.Math.Max(text_bounding_box.Height, wh_bounding_box.Height);
                var max_size = new VA.Drawing.Size(max_w, max_h);
                new_sizes.Add(max_size);
            }

            var src_width = VA.ShapeSheet.SRCConstants.Width;
            var src_height = VA.ShapeSheet.SRCConstants.Height;

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            for (int i = 0; i < new_sizes.Count; i++)
            {
                var shapeid = shapeids[i];
                var new_size = new_sizes[i];
                update.SetFormula((short) shapeid, src_width, new_size.Width);
                update.SetFormula((short) shapeid, src_height, new_size.Height);
            }

            update.Execute(page);
        }


        public static IVisio.Font TryGetFont(IVisio.Fonts fonts, string name)
        {
            try
            {
                var font = fonts[name];
                return font;
            }
            catch (System.Runtime.InteropServices.COMException comexc)
            {
                return null;
            }
        }
    }
}