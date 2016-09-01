using System.Management.Automation;
using IVisio = NetOffice.VisioApi;

namespace VisioPowerShell.Commands.Remove
{
    [Cmdlet(VerbsCommon.Remove, VisioPowerShell.Nouns.VisioGroup)]
    public class Remove_VisioGroup : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.IVShape[] Shapes;

        protected override void ProcessRecord()
        {
            var targets = new VisioAutomation.Scripting.TargetShapes(this.Shapes);
            this.Client.Grouping.Ungroup(targets);
        }
    }
}