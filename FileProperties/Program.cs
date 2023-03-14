


using SExtensions;
using SolidEdgeCommunity.Extensions; // https://github.com/SolidEdgeCommunity/SolidEdge.Community/wiki/Using-Extension-Methods
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportFileProperties
{
    class Program
    {
        private static SolidEdgeFramework.Application application => SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);

        [STAThread]
        static void Main(string[] args)
        {
             

            try
            {
                // Register with OLE to handle concurrency issues on the current thread.
                SolidEdgeCommunity.OleMessageFilter.Register();

                // Connect to or start Solid Edge.
                

                // Get a reference to the active assembly document.
                var document = application.GetActiveDocument<SolidEdgeAssembly.AssemblyDocument>(false);

                if (document != null)
                {
                    ExecuteAction();
                    

                    
                   

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

        private static void ExecuteAction()
        {
            //SExtensions.Helpers.FillOccurrence(document, true, 0);
            application.SetOcurrenceProperty( "Certificados", "Requiere Certificado de Material");

            var start = Console.ReadLine();
            if (start == "a")
            {
                ExecuteAction();
            }
        }
        
    }
}