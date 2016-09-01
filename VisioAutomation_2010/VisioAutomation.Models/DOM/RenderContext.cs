using System.Collections.Generic;
using IVisio=NetOffice.VisioApi;

namespace VisioAutomation.Models.Dom
{
    internal class RenderContext
    {
        private readonly Dictionary<short, IVisio.IVShape> _id_to_shape;
        private readonly IVisio.Shapes _pageshapes;
        public IVisio.IVPage VisioPage { get; private set; }

        public RenderContext(IVisio.IVPage visio_page)
        {
            this._id_to_shape = new Dictionary<short, IVisio.IVShape>();
            this.VisioPage = visio_page;
            this._pageshapes = (IVisio.Shapes)visio_page.Shapes;
        }

        public IVisio.IVShape GetShape(short id)
        {
            IVisio.IVShape vshape;
            if (this._id_to_shape.TryGetValue(id, out vshape))
            {
                return vshape;
            }
            else
            {
                vshape = (IVisio.IVShape) this._pageshapes.get_ItemFromID16(id);
                this._id_to_shape[id] = vshape;
                return vshape;
            }
        }
    }
}