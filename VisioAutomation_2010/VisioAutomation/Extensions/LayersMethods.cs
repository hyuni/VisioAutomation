using System.Collections.Generic;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Extensions
{
    public static class LayersMethods
    {
        public static IEnumerable<IVisio.IVLayer> ToEnumerable(this IVisio.IVLayers layers)
        {
            return VisioAutomation.Layers.LayerHelper.ToEnumerable(layers);
        }
    }
}
