using System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioModelAreaChart)]
    public class NewVisioModelAreaChart : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public double X0 { get; set; }

        [Parameter(Position = 1, Mandatory = true)]
        public double Y0 { get; set; }

        [Parameter(Position = 2, Mandatory = true)]
        public double X1 { get; set; }

        [Parameter(Position = 3, Mandatory = true)]
        public double Y1 { get; set; }

        [Parameter(Position = 4, Mandatory = true)]
        public double[] Values;

        [Parameter(Mandatory = false)]
        public string[] Labels;

        protected override void ProcessRecord()
        {
            var rect = new VA.Geometry.Rectangle(this.X0, this.Y0, this.X1, this.Y1);
            var chart = new VA.Models.Charting.AreaChart(rect);
            chart.DataPoints = new VA.Models.Charting.DataPointList(this.Values, this.Labels);
            this.WriteObject(chart);
        }
    }
}