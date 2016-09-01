using System.Collections.Generic;
using VisioAutomation.Exceptions;
using VisioAutomation.Extensions;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Documents
{
    public static class DocumentHelper
    {
        public static IVisio.IVDocument OpenStencil(IVisio.IVDocuments docs, string filename)
        {
            var stencil = DocumentHelper.TryOpenStencil(docs, filename);
            if (stencil == null)
            {
                string msg = string.Format("Could not open stencil \"{0}\"", filename);
                throw new AutomationException(msg);
            }
            return stencil;
        }

        private static IVisio.IVDocument TryOpenStencil(IVisio.IVDocuments docs, string filename)
        {
            const short flags = (short)IVisio.Enums.VisOpenSaveArgs.visOpenRO | (short)IVisio.Enums.VisOpenSaveArgs.visOpenDocked;
            try
            {
                var doc = docs.OpenEx(filename, flags);
                return (IVisio.IVDocument) doc;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null;
            }
        }

        public static void Activate(IVisio.IVDocument doc)
        {
            var app = doc.Application;
            var cur_active_doc = app.ActiveDocument;

            // if the doc is already active do nothing
            if (doc == cur_active_doc)
            {
                // do nothing
                return;
            }

            // go through each window and check if it is assigned
            // to the target document
            var appwindows = app.Windows;
            // var allwindows = appwindows.ToEnumerable();
            foreach (var curwin in appwindows)
            {
                if (curwin.Document == doc)
                {
                    // we did find one, so activate that window
                    // and then exit the method
                    curwin.Activate();
                    if (app.ActiveDocument != doc)
                    {
                        throw new InternalAssertionException("failed to activate document");
                    }
                    return;
                }
            }

            // If we get here, we couldn't find any matching window
            throw new AutomationException("could not find window for document");
        }

        public static void Close(IVisio.IVDocument doc, bool force_close)
        {
            if (force_close)
            {
                var new_alert_response = Application.AlertResponseCode.No;
                var app = doc.Application;

                using (var alertresponse = new Application.AlertResponseScope((IVisio.Application)app, new_alert_response))
                {
                    doc.Close();
                }
            }
            else
            {
                doc.Close();
            }
        }

        public static IEnumerable<IVisio.IVDocument> ToEnumerable(IVisio.IVDocuments docs)
        {
            short count = docs.Count;
            for (int i = 0; i < count; i++)
            {
                yield return (IVisio.IVDocument) docs[i + 1];
            }
        }
    }
}