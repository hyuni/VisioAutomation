﻿using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Test, VisioPowerShell.Commands.Nouns.VisioApplication)]
    public class TestVisioApplication: VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            bool valid_app = this.Client.Application.ValidateApplication();
            this.WriteObject(valid_app);
        }
    }
}