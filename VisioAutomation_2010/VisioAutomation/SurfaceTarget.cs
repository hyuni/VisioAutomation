using System.Collections.Generic;
using VisioAutomation.Extensions;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation
{
    public struct SurfaceTarget
    {
        public readonly IVisio.IVPage Page;
        public readonly IVisio.IVMaster Master;
        public readonly IVisio.IVShape Shape;

        public SurfaceTarget(IVisio.IVPage page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            this.Page = page;
            this.Master = null;
            this.Shape = null;
        }

        public SurfaceTarget(IVisio.IVMaster master)
        {
            if (master== null)
            {
                throw new System.ArgumentNullException(nameof(master));
            }

            this.Page = null;
            this.Master = master;
            this.Shape = null;
        }

        public SurfaceTarget(IVisio.IVShape shape)
        {
            if (shape== null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            this.Page = null;
            this.Master = null;
            this.Shape = shape;
        }

        public IVisio.IVShapes Shapes
        {
            get
            {

                IVisio.IVShapes shapes;

                if (this.Master != null)
                {

                    shapes = (IVisio.IVShapes) this.Master.Shapes;
                }
                else if (this.Page != null)
                {
                    shapes = (IVisio.IVShapes) this.Page.Shapes;
                }
                else if (this.Shape != null)
                {
                    shapes = (IVisio.IVShapes) this.Shape.Shapes;
                }
                else
                {
                    throw new System.ArgumentException("Unhandled Drawing Surface");
                }
                return shapes;
            }

        }


        public List<IVisio.IVShape> GetAllShapes()
        {
            IVisio.IVShapes shapes;

            if (this.Master != null)
            {

                shapes = (IVisio.IVShapes) this.Master.Shapes;
            }
            else if (this.Page != null)
            {
                shapes = (IVisio.IVShapes) this.Page.Shapes;
            }
            else if (this.Shape != null)
            {
                shapes = (IVisio.IVShapes) this.Shape.Shapes;
            }
            else
            {
                throw new System.ArgumentException("Unhandled Drawing Surface");
            }

            var list = new List<IVisio.IVShape>();
            list.AddRange(shapes.ToEnumerable());

            return list;
        }

        public short ID16
        {
            get
            {
                if (this.Shape != null)
                {
                    return this.Shape.ID16;
                }
                else if (this.Page != null)
                {
                    return this.Page.ID16;
                }
                else if (this.Master != null)
                {
                    return this.Master.ID16;
                }
                else
                {
                    throw new System.ArgumentException("Unhandled Drawing Surface");
                }
            }
        }

    }
}