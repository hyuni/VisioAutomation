using System.Collections.Generic;
using IVisio=NetOffice.VisioApi;

namespace VisioAutomation.Models.Dom
{
    internal class RenderContext
    {
        private readonly Dictionary<short, IVisio.Shape> _id_to_shape;
        private readonly IVisio.Shapes _pageshapes;
        public IVisio.Page VisioPage { get; private set; }

        public RenderContext(IVisio.Page visio_page)
        {
            this._id_to_shape = new Dictionary<short, IVisio.Shape>();
            this.VisioPage = visio_page;
            this._pageshapes = (IVisio.Shapes)visio_page.Shapes;
        }

        public IVisio.Shape GetShape(short id)
        {
            IVisio.Shape vshape;
            if (this._id_to_shape.TryGetValue(id, out vshape))
            {
                return vshape;
            }
            else
            {
                vshape = (IVisio.Shape) this._pageshapes.get_ItemFromID16(id);
                this._id_to_shape[id] = vshape;
                return vshape;
            }
        }
    }
}