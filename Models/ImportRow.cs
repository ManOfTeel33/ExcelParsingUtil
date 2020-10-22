using System;

namespace ExcelParsingUtil.Models
{
    public class ImportRow
    {
        public string IssueNumber { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string Flag { get; set; }
        public DateTimeOffset DatePublished { get; set; }
        public Guid ItemGuid { get; set; }
        public int RowNumber { get; set; }
    }
}
