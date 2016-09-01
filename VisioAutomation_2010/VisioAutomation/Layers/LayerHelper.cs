using System.Collections.Generic;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Layers
{
    public static class LayerHelper
    {
        public static IEnumerable<IVisio.IVLayer> ToEnumerable(IVisio.IVLayers layers)
        {
            short count = layers.Count;
            for (int i = 0; i < count; i++)
            {
                yield return (IVisio.IVLayer) layers[i + 1];
            }
        }
    }
}

