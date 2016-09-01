using VisioAutomation.Exceptions;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Shapes.Geometry
{
    public static class GeometryHelper
    {
        public static short AddSection(IVisio.IVShape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            int num_geometry_sections = shape.GeometryCount;
            short new_sec_index = GeometryHelper.GetGeometrySectionIndex((short)num_geometry_sections);
            short actual_sec_index = shape.AddSection(new_sec_index);

            if (actual_sec_index != new_sec_index)
            {
                throw new InternalAssertionException();
            }
            short row_index = shape.AddRow(new_sec_index, (short)IVisio.Enums.VisRowIndices.visRowComponent, (short)IVisio.Enums.VisRowTags.visTagComponent);

            return new_sec_index;
        }

        private static short GetGeometrySectionIndex(short index)
        {
            short i =
                (short) (((int) IVisio.Enums.VisSectionIndices.visSectionFirstComponent) + (index));
            return i;
        }

        public static void Delete(IVisio.IVShape shape)
        {
            int num = shape.GeometryCount;
            for (int i = num-1; i >=0; i--)
            {
                GeometryHelper.DeleteSection(shape, (short)i);                
            }
        }

        private static void DeleteSection(IVisio.IVShape shape, short index)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            short target_section_index = GeometryHelper.GetGeometrySectionIndex(index);
            shape.DeleteSection(target_section_index);
        }
    }
}