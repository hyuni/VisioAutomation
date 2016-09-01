using System.Management.Automation;
using IVisio = NetOffice.VisioApi;

namespace VisioPowerShell.Commands.Get
{
    [Cmdlet(VerbsCommon.Get, VisioPowerShell.Nouns.VisioShapeText)]
    public class Get_VisioShapeText : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.IVShape[] Shapes;

        protected override void ProcessRecord()
        {
            var targets = new VisioAutomation.Scripting.TargetShapes(this.Shapes);
            var t = this.Client.Text.Get(targets);
            this.WriteObject(t);
        }
    }
}