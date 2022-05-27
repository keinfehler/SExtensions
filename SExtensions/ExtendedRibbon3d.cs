using SolidEdgeCommunity.AddIn;

namespace SExtensions
{
    public class ExtendedRibbon3d : Ribbon
    {
        private RibbonCheckBox _ribbonRadioButton;
        const string _embeddedResourceName = "SExtensions.ExtendedRibbon3d.xml";
        public ExtendedRibbon3d()
        {
            // Get a reference to the current assembly. This is where the ribbon XML is embedded.
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();

            // In this example, XML file must have a build action of "Embedded Resource".
            LoadXml(assembly, _embeddedResourceName);

            _ribbonRadioButton = GetCheckBox(1);
        }

        public override void OnControlClick(RibbonControl control)
        {
            switch (control.CommandId)
            {
                case 0:
                    {
                        //ExtendedCommands.AddCopyPlus(_ribbonRadioButton.Checked);
                    }
                    break;
                case 2:
                    //Helpers.FindOccurrencesAndExport();
                    {
                        using (ExcelWaitingForm frm = new ExcelWaitingForm(() => Helpers.FindOccurrencesAndExport()))
                        {
                            frm.ShowDialog();
                        }
                    }
                    break;
                default:
                    break;
            }
            base.OnControlClick(control);
        }


    }
}
