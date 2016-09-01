using System.Collections.Generic;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Masters
{
    public static class MasterHelper
    {
        public static Drawing.Rectangle GetBoundingBox(IVisio.IVMaster master, IVisio.Enums.VisBoundingBoxArgs args)
        {
            // MSDN: http://msdn.microsoft.com/library/default.asp?url=/library/en-us/vissdk11/html/vimthBoundingBox_HV81900422.asp
            double bbx0, bby0, bbx1, bby1;
            master.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
            var r = new Drawing.Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
        }

        public static IEnumerable<IVisio.IVMaster> ToEnumerable(IVisio.IVMasters masters)
        {
            short count = masters.Count;
            for (int i = 0; i < count; i++)
            {
                yield return (IVisio.IVMaster) masters[i + 1];
            }
        }

        public static string[] GetNamesU(IVisio.IVMasters masters)
        {
            string[] names_sa;
            masters.GetNamesU(out names_sa);
            string[] names = (string[])names_sa;
            return names;
        }
    }
}