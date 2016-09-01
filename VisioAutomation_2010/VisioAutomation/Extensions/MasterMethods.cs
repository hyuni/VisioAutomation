using System.Collections.Generic;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods
    {
        public static Drawing.Rectangle GetBoundingBox(this IVisio.IVMaster master, IVisio.Enums.VisBoundingBoxArgs args)
        {
            return VisioAutomation.Masters.MasterHelper.GetBoundingBox(master, args);
        }

        public static IEnumerable<IVisio.IVMaster> ToEnumerable(this IVisio.IVMasters masters)
        {
            return VisioAutomation.Masters.MasterHelper.ToEnumerable(masters);
        }

        public static string[] GetNamesU(this IVisio.IVMasters masters)
        {
            return VisioAutomation.Masters.MasterHelper.GetNamesU(masters);
        }
    }
}