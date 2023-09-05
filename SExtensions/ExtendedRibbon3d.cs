using SolidEdgeCommunity.AddIn;

namespace SExtensions
{
    public class ExtendedRibbon3d : Ribbon
    {
        public static string directorioSalida = ConfigurationHelper.GetConfigurationValue("DirectorioSalida");
        public static string printFileLocation = ConfigurationHelper.GetConfigurationValue("PrintFileLocation");
        private RibbonCheckBox _ribbonRadioButton;

        private RibbonCheckBox _ribbonExportarRutasRadioButton;
        private RibbonCheckBox _ribbonUtillajeRadioButton;
        public SolidEdgeFramework.Application CurrentApp => SolidEdgeAddIn.Instance.Application;

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
                case 11:
                    {
                        //ExtendedCommands.AddCopyPlus(_ribbonRadioButton.Checked);
                        //08.03.2023 - RE - llamada para cambiar valor de una propiedad en las ocurrencias seleccionadas
                        CurrentApp.SetOcurrenceProperty("Certificados", "- Requiere Certificado de Material");
                        
                    }
                    break;

                case 12:
                    {
                        CurrentApp.SetOcurrenceProperty("Certificados", "- Requiere Certificado : FDA ó EHEDG ó 3A");
                    }
                    break;

                case 13:
                    {
                        CurrentApp.SetOcurrenceProperty("Certificados", "");
                    }
                    break;

                case 16:
                    {
                        CurrentApp.SetOcurrenceProperty("Repuestos", "");
                    }
                    break;

                case 17:
                    {
                        CurrentApp.SetOcurrenceProperty("Repuestos", "R1");
                    }
                    break;

                case 18:
                    {
                        CurrentApp.SetOcurrenceProperty("Repuestos", "R2");
                    }
                    break;

                case 19:
                    {
                        CurrentApp.SetOcurrenceProperty("Repuestos", "R3");
                    }
                    break;

                case 27:
                    {
                        CurrentApp.SetOcurrenceProperty("Repuestos", "R2");
                    }
                    break;

                case 28:
                    {
                        CurrentApp.SetOcurrenceProperty("Repuestos", "R3");
                    }
                    break;

                case 2:
                    {
                        var exportarHijos = false;

                        if (_ribbonRadioButton.Checked != exportarHijos)
                        {
                            exportarHijos = _ribbonRadioButton.Checked;
                        }
                           
                        using (ExcelWaitingDialog frm = new ExcelWaitingDialog(() =>
                            Helpers.FindOccurrencesAndExport(exportarHijos, true, OcurrencesExportMode.None)))
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
                case 7:
                    {
                        using (ExcelWaitingForm frm = new ExcelWaitingForm())
                        {
                            frm.ExportRutas = true;
                            frm.GetPwdFiles = true;
                            frm.Utillaje = true;

                            frm.ShowDialog();
                        }
                        break;
                    }
                case 20:
                    {
                        ExtendedCommands.Readonly();
                        break;
                    }
                case 10:
                    {
                        ExtendedCommands.SaveASM();
                        break;

                    }
                case 8:
                    Helpers.FindOccurrencesAndExport
                        (true, true, OcurrencesExportMode.ParaImprimir, fileTemplate: printFileLocation, outputDirectory: directorioSalida);
                    break;
                default:
                    break;
            }
            base.OnControlClick(control);
        }


    }
}
