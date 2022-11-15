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

            ChangeName();
            //SolidEdgeFramework.Application application = null;
            //// Register with OLE to handle concurrency issues on the current thread.
            //SolidEdgeCommunity.OleMessageFilter.Register();

            //// Connect to or start Solid Edge.
            //application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);

            //// Get a reference to the active assembly document.
            //var document = application.GetActiveDocument<SolidEdgeDocument>(false);

            //Helpers.DesignManagerHelpers.ReplaceAndCopy(document, true);

        }

        private static void ChangeName()
        {
            string relatedItemfileNameWithoutExtension;

            Console.WriteLine("Enter file name:");
            relatedItemfileNameWithoutExtension = Console.ReadLine();

            Console.WriteLine($"OLDNAME: {relatedItemfileNameWithoutExtension}");

            var newName = Helpers.DesignManagerHelpers.GetRevisionNumber(relatedItemfileNameWithoutExtension);

            Console.WriteLine($"NEWNAME: {newName}");

            Console.ReadLine();
        }
    }


}
