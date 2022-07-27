


using SolidEdgeCommunity.Extensions; // https://github.com/SolidEdgeCommunity/SolidEdge.Community/wiki/Using-Extension-Methods
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportFileProperties
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            SolidEdgeFramework.Application application = null;

            try
            {
                // Register with OLE to handle concurrency issues on the current thread.
                SolidEdgeCommunity.OleMessageFilter.Register();

                // Connect to or start Solid Edge.
                application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);

                // Get a reference to the active assembly document.
                var document = application.GetActiveDocument<SolidEdgeAssembly.AssemblyDocument>(false);

                if (document != null)
                {
                    SExtensions.Helpers.FillOccurrence(document, true, 0);

                    Console.ReadLine();
                }
                else
                {
                    throw new System.Exception("No active document.");
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                SolidEdgeCommunity.OleMessageFilter.Unregister();
            }
        }
    }
}