using IVisio = NetOffice.VisioApi;

namespace VisioAutomation.Models.Dom
{
    public class Document
    {
        public PageList Pages;
        public IVisio.IVDocument VisioDocument;

        private readonly string _vst_template_file ;
        private readonly IVisio.Enums.VisMeasurementSystem _measurement_system;

        public Document()
        {
            this.Pages = new PageList();
            this._measurement_system = IVisio.Enums.VisMeasurementSystem.visMSDefault;
        }

        public Document(string template, IVisio.Enums.VisMeasurementSystem ms) :
            this()
        {
            this._vst_template_file = template;
            this._measurement_system = ms;
        }

        public IVisio.IVDocument Render(IVisio.IVApplication app)
        {
            var appdocs = app.Documents;
            IVisio.IVDocument doc = null;
            if (this._vst_template_file == null)
            {
                doc = (IVisio.IVDocument)appdocs.Add(string.Empty);
            }
            else
            {
                const int flags = 0; // (int)IVisio.Enums.VisOpenSaveArgs.visAddDocked;
                const int langid = 0;
                doc = (IVisio.IVDocument)appdocs.AddEx(this._vst_template_file, this._measurement_system, flags, langid);
            }
            this.VisioDocument = doc;
            var docpages = doc.Pages;
            var startpage = docpages[1];
            this.Pages.Render((IVisio.IVPage)startpage);
            return doc;
        }
    }
}