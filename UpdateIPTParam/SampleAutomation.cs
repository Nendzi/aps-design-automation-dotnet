using Inventor;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace UpdateIPTParam
{
  [ComVisible(true)]
  public class SampleAutomation
  {
    private InventorServer m_server;
    public SampleAutomation(InventorServer app) { m_server = app; }

    public void Run(Document doc)
    {
      try
      {
        // update parameters in the doc
        ChangeParameters(doc);

        // save file                                                                
        doc.Save2();
      }
      catch (Exception e) { LogTrace("Processing failed: {0}", e.ToString()); }
    }

    /// <summary>
    /// Change parameters in Inventor document.
    /// </summary>
    /// <param name="doc">The Inventor document.</param>
    /// <param name="json">JSON with changed parameters.</param>
    public void ChangeParameters(Document doc)
    {
      var theParams = GetParameters(doc);

      Dictionary<string, string> parameters = JsonConvert.DeserializeObject<Dictionary<string, string>>(System.IO.File.ReadAllText("params.json"));
      foreach (KeyValuePair<string, string> entry in parameters)
      {
        var keyName = entry.Key.ToLower();
        try
        {
          Parameter param = theParams[keyName];
          param.Expression = entry.Value;
        }
        catch (Exception e) { LogTrace("Cannot update {0}: {1}", keyName, e.Message); }
      }
      doc.Update();
    }

    /// <summary>
    /// Get parameters for the document.
    /// </summary>
    /// <returns>Parameters. Throws exception if parameters are not found.</returns>
    private static Parameters GetParameters(Document doc)
    {
      var docType = doc.DocumentType;
      switch (docType)
      {
        case DocumentTypeEnum.kAssemblyDocumentObject:
          var asm = doc as AssemblyDocument;
          return asm.ComponentDefinition.Parameters;
        case DocumentTypeEnum.kPartDocumentObject:
          var ipt = doc as PartDocument;
          return ipt.ComponentDefinition.Parameters;
        default:
          throw new ApplicationException(string.Format("Unexpected document type ({0})", docType));
      }
    }

    /// <summary>
    /// This will appear on the Design Automation output
    /// </summary>
    private static void LogTrace(string format, params object[] args) { Trace.TraceInformation(format, args); }
  }
}
