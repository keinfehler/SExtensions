using SolidEdgeCommunity.Extensions;
using SolidEdgeFramework;
using System;

namespace DesignManagerConsoleApp
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            SolidEdgeFramework.Application application = null;
            // Register with OLE to handle concurrency issues on the current thread.
            SolidEdgeCommunity.OleMessageFilter.Register();

            // Connect to or start Solid Edge.
            application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);

            // Get a reference to the active assembly document.
            var document = application.GetActiveDocument<SolidEdgeDocument>(false);

            Helpers.DesignManagerHelpers.ReplaceAndCopy(document, true);

        }
    }


}
