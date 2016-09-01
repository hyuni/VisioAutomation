using System.Management.Automation;
using IVisio = NetOffice.VisioApi;

namespace VisioPowerShell.Commands.Open
{
    [Cmdlet(VerbsCommon.Open, VisioPowerShell.Nouns.VisioMaster)]
    public class Open_VisioMaster : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        [ValidateNotNull]
        public IVisio.IVMaster Master;

        protected override void ProcessRecord()
        {
            this.Client.Master.OpenForEdit(this.Master);
        }
    }
}