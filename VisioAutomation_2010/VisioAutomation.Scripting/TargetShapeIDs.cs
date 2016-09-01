using System.Collections.Generic;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Scripting
{
    public class TargetShapeIDs
    {
        public readonly IList<int> ShapeIDs;
        public readonly IVisio.IVPage Page;

        public TargetShapeIDs(IVisio.IVPage page, IList<int> shape_ids)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            if (shape_ids == null)
            {
                throw new System.ArgumentNullException(nameof(shape_ids));
            }

            this.Page = page;
            this.ShapeIDs = shape_ids;
        }
    }
}