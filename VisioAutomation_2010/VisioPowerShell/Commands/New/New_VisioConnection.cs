using System.Management.Automation;
using IVisio = NetOffice.VisioApi;

namespace VisioPowerShell.Commands.New
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Nouns.VisioConnection)]
    public class New_VisioConnection : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public IVisio.IVShape[] From { get; set; }

        [Parameter(Position = 1, Mandatory = true)]
        public IVisio.IVShape[] To { get; set; }

        [Parameter(Position = 2, Mandatory = false)]
        public IVisio.IVMaster Master { get; set; }

        protected override void ProcessRecord()
        {
            var connectors = this.Client.Connection.Connect(this.From, this.To, this.Master);
            this.WriteObject(connectors, false);
        }
    }
}