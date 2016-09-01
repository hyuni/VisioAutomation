using System.Management.Automation;
using IVisio = NetOffice.VisioApi;

namespace VisioPowerShell.Commands.Copy
{
    [Cmdlet(VerbsCommon.Copy, VisioPowerShell.Nouns.VisioPage)]
    public class Copy_VisioPage : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.IVDocument ToDocument=null;

        protected override void ProcessRecord()
        {
            IVisio.IVPage newpage;
            if (this.ToDocument == null)
            {
                newpage = this.Client.Page.Duplicate();
            }
            else
            {
                newpage = this.Client.Page.Duplicate(this.ToDocument);
            }

            this.WriteObject(newpage);            
        }
    }
}