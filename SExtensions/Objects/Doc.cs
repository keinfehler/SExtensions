using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SExtensions
{
    public class DocWrapper
    {
        public DocWrapper()
        {
        }
        public DocWrapper(string cat, string docNumber, string revNumber, string keywords, string title, string comments, string material, string ubicacion, string projectname)
        {
            Category = cat;
            DocumentNumber = docNumber;
            RevisionNumber = revNumber;
            Keywords = keywords;
            Title = title;
            Comments = comments;
            Material = material;
            Ubicacion = ubicacion;
            ProjectName = projectname;

        }

        public string Category { get; set; }

        public string DocumentNumber { get; set; }

        public string RevisionNumber { get; set; }

        public string Keywords { get; set; }

        public string Title { get; set; }

        public string Comments { get; set; }

        public string Material { get; set; }
        public string Ubicacion { get; set; }
        public string ProjectName { get; set; }


        public string Empresa { get; set; }
        //public string Maquina { get; set; }
        public string Modelo { get; set; }
        public string NombreArchivo { get; set; }
        public string Ruta { get; set; }
        public bool IsHeader { get; set; }
    }
}
