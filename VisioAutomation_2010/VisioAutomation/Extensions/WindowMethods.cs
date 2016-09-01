using System.Collections.Generic;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Extensions
{
    public static class WindowMethods
    {
        public static void Select(
            this IVisio.IVWindow window,
            IEnumerable<IVisio.IVShape> shapes,
            IVisio.Enums.VisSelectArgs selectargs)
        {
            VisioAutomation.Windows.WindowHelper.Select(window, shapes, selectargs);
        }



        public static Drawing.Rectangle GetViewRect(this IVisio.IVWindow window)
        {
            return VisioAutomation.Windows.WindowHelper.GetViewRect(window);
        }

        public static System.Drawing.Rectangle GetWindowRect(this IVisio.IVWindow window)
        {
            return VisioAutomation.Windows.WindowHelper.GetWindowRect(window);
        }

        public static void SetWindowRect(
            this IVisio.IVWindow window, 
            System.Drawing.Rectangle rect)
        {
            VisioAutomation.Windows.WindowHelper.SetWindowRect(window, rect);
        }

        public static void SetViewRect(
            this IVisio.IVWindow window, 
            Drawing.Rectangle rect)
        {
            VisioAutomation.Windows.WindowHelper.SetViewRect(window,rect);
        }

        public static IEnumerable<IVisio.Window> ToEnumerable(this IVisio.Windows windows)
        {
            return VisioAutomation.Windows.WindowHelper.ToEnumerable(windows);
        }
    }
}