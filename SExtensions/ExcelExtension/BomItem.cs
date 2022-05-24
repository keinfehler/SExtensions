using System.Collections.Generic;

namespace SExtensions
{
    /// <summary>
    /// Class to hold BOM data.
    /// </summary>
    public class BomItem
    {
        private List<BomItem> _children = new List<BomItem>();

        public BomItem()
        {
        }

        public BomItem(SolidEdgeAssembly.Occurrence occurrence, int level, int qty)
        {
            Level = level;
            FileName = System.IO.Path.GetFileName(occurrence.OccurrenceFileName);
            IsMissing = occurrence.FileMissing();
            Quantity = qty;
            IsSubassembly = occurrence.Subassembly;



            // If the file exists, extract file properties.
            if (IsMissing.HasValue && !IsMissing.Value)
            {
                var document = (SolidEdgeFramework.SolidEdgeDocument)occurrence.OccurrenceDocument;
                var summaryInfo = (SolidEdgeFramework.SummaryInfo)document.SummaryInfo;

                #region Custom
                SolidEdgeFramework.PropertySets pps = null;

                pps = document.Properties as SolidEdgeFramework.PropertySets;

                if (pps != null)
                {
                    SolidEdgeFramework.Properties mproperties = null;

                    mproperties = pps.Item(7);

                    if (mproperties != null)
                    {
                        SolidEdgeFramework.Property p = null;
                        p = mproperties.Item(1);


                        if (p != null)
                        {
                            var m = p.get_Value();
                            Material = m?.ToString();
                        }
                    }
                }

                PartType = summaryInfo.Category;
                Code = summaryInfo.DocumentNumber;
                ArticleNo = summaryInfo.Keywords;
                CommercialReference = summaryInfo.Comments;

                #endregion

                DocumentNumber = summaryInfo.DocumentNumber;
                Title = summaryInfo.Title;
                Revision = summaryInfo.RevisionNumber;



            }
        }

        public int? Level { get; set; }
        public string DocumentNumber { get; set; }
        public string Revision { get; set; }
        public string Title { get; set; }
        public int? Quantity { get; set; }
        public string FileName { get; set; }

        #region Custom

        public string PartType { get; set; }
        public string Code { get; set; }
        public string ArticleNo { get; set; }
        public string Material { get; set; }

        public string CommercialReference { get; set; }


        #endregion


        //[JsonIgnore]
        public bool? IsSubassembly { get; set; }

        //[JsonIgnore]
        public bool? IsMissing;

        /// <summary>
        /// Returns all direct children.
        /// </summary>
        //[JsonProperty("Child")]
        public List<BomItem> Children { get { return _children; } set { _children = value; } }

        /// <summary>
        /// Returns all direct and descendant children.
        /// </summary>
        //[JsonIgnore]
        public IEnumerable<BomItem> AllChildren
        {
            get
            {
                foreach (var bomItem in Children)
                {
                    yield return bomItem;

                    if (bomItem.IsSubassembly == true)
                    {
                        foreach (var childBomItem in bomItem.AllChildren)
                        {
                            yield return childBomItem;
                        }
                    }
                }
            }
        }

        // Demonstration of how to exclude empty collections during JSON.NET serialization.
        public bool ShouldSerializeChildren()
        {
            return Children.Count > 0;
        }

        public override string ToString()
        {
            return FileName;
        }
    }
}
