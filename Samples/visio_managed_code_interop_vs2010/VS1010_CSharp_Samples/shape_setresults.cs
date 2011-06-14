﻿using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioInterop;
public static partial class VS2010_CSharp_Samples
{
	public static void Shape_SetResults(IVisio.Document doc)
	{
        var page = VisioInterop.Util.CreateStandardPage(doc, "SSR");
        var shape = VisioInterop.Util.CreateStandardShape(page);
        var request = new[]
        {
              new
                  {
                      Section = (short)IVisio.VisSectionIndices.visSectionObject, 
                      Row=(short)IVisio.VisRowIndices.visRowXFormOut, 
                      Cell=(short)IVisio.VisCellIndices.visXFormWidth,
                      UnitCode=(short) IVisio.VisUnitCodes.visNoCast,
                      Result=8.2
                  },                        
              new
                  {
                      Section = (short)IVisio.VisSectionIndices.visSectionObject, 
                      Row=(short)IVisio.VisRowIndices.visRowXFormOut, 
                      Cell=(short)IVisio.VisCellIndices.visXFormHeight,
                      UnitCode=(short) IVisio.VisUnitCodes.visNoCast,
                      Result=1.3
                  }                        
        };

		// MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        var SRCStream = new short[request.Length * 3];
        var results_objects = new object[request.Length];
        var unitcodes = new object[request.Length];
        for (int i = 0; i < request.Length; i++)
        {
            SRCStream[(i * 3) + 0] = request[i].Section;
            SRCStream[(i * 3) + 1] = request[i].Row;
            SRCStream[(i * 3) + 2] = request[i].Cell;
            results_objects[i] = request[i].Result;
            unitcodes[i] = request[i].UnitCode;
        }

		// EXECUTE THE REQUEST
        short flags = 0;
        int count = shape.SetResults(SRCStream, unitcodes, results_objects, flags);

        // DISPLAY THE INFORMATION
		shape.Text = "SetResults";
	}
}