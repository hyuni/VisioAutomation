using System.Management.Automation;
using IVisio = NetOffice.VisioApi;

namespace VisioPowerShell.Commands.Remove
{
    [Cmdlet(VerbsCommon.Remove, VisioPowerShell.Nouns.VisioHyperlink)]
    public class Remove_VisioHyperlink : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public int Index { get; set; }

        [Parameter(Mandatory = false)]
        public IVisio.IVShape[] Shapes;

        protected override void ProcessRecord()
        {
            var targets = new VisioAutomation.Scripting.TargetShapes(this.Shapes);
            this.Client.Hyperlink.Delete(targets,this.Index);
        }
    }
}