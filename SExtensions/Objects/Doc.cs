﻿namespace SExtensions
{
    public class DocWrapper
    {
        public DocWrapper()
        {
        }
        public DocWrapper(string cat, string docNumber, string revNumber, string keywords, string title, string comments, string material)
        {
            Category = cat;
            DocumentNumber = docNumber;
            RevisionNumber = revNumber;
            Keywords = keywords;
            Title = title;
            Comments = comments;
            Material = material;
        }
      
        public string Category { get; set; }

        public string DocumentNumber { get; set; }

        public string RevisionNumber { get; set; }

        public string Keywords { get; set; }

        public string Title { get; set; }

        public string Comments { get; set; }

        public string Material { get; set; }
    }
}