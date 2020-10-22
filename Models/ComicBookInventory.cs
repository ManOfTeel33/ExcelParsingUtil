using System;

namespace ExcelParsingUtil.Models
{
    public class ComicBookInventory
    {
        public string IssueNumber { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string Flag { get; set; }
        public DateTimeOffset DatePublished { get; set; }
    }
}
