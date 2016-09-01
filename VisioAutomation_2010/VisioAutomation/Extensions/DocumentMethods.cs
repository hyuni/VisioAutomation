using System.Collections.Generic;
using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Extensions
{
    public static class DocumentMethods
    {
        public static void Close(this IVisio.IVDocument doc, bool force_close)
        {
            Documents.DocumentHelper.Close(doc, force_close);
        }

        public static IEnumerable<IVisio.IVDocument> ToEnumerable(this IVisio.IVDocuments docs)
        {
            return Documents.DocumentHelper.ToEnumerable(docs);
        }

        public static IVisio.IVDocument OpenStencil(this IVisio.IVDocuments docs, string filename)
        {
            return Documents.DocumentHelper.OpenStencil(docs, filename);
        }

    }
}