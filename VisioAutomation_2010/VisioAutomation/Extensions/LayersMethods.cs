using System.Collections.Generic;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Extensions
{
    public static class LayersMethods
    {
        public static IEnumerable<IVisio.Layer> ToEnumerable(this IVisio.Layers layers)
        {
            return VisioAutomation.Layers.LayerHelper.ToEnumerable(layers);
        }
    }
}
