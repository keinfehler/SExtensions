using SolidEdgeCommunity.AddIn;

namespace SExtensions
{
    public class ExtendedRibbon3d : Ribbon
    {
        private RibbonCheckBox _ribbonRadioButton;

        private RibbonCheckBox _ribbonExportarRutasRadioButton;
        private RibbonCheckBox _ribbonUtillajeRadioButton;


        const string _embeddedResourceName = "SExtensions.ExtendedRibbon3d.xml";
        public ExtendedRibbon3d()
        {
            // Get a reference to the current assembly. This is where the ribbon XML is embedded.
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();

            // In this example, XML file must have a build action of "Embedded Resource".
            LoadXml(assembly, _embeddedResourceName);

            _ribbonRadioButton = GetCheckBox(1);
            _ribbonExportarRutasRadioButton = GetCheckBox(4);
            _ribbonUtillajeRadioButton = GetCheckBox(3);
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
                        using (ExcelWaitingForm frm = new ExcelWaitingForm(() => 
                        Helpers.FindOccurrencesAndExport(_ribbonRadioButton.Checked, _ribbonExportarRutasRadioButton.Checked, _ribbonUtillajeRadioButton.Checked)))
                        {
                            frm.ShowDialog();
                        }
                    }
                    break;
                case 5:
                    {
                        IncludeExcludeHelper.IncludeExclude(true);
                        break;
                    }
                case 6:
                    {
                        IncludeExcludeHelper.IncludeExclude(false);
                        break;
                    }
                default:
                    break;
            }
            base.OnControlClick(control);
        }


    }
}
