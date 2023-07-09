using System.ComponentModel;

namespace Reports.Core.Integrations.Report
{
    public enum ReportTypes
    {
        [Description("PDF")]
        Pdf,
        [Description("EXCELOPENXML")]
        Excel,
        [Description("CSV")]
        Csv,
        [Description("XML")]
        Xml
    }
    public static class Constants 
    {
        public const string PdfExtension = ".pdf";
        public const string CommaDelimitedExtension = ".csv";
        public const string TabDelimitedExtension = ".txt";
        public const string ExcelExtension = ".xlsx";
        public const string Excel2003Extension = ".xls";
        public const string XmlExtension = ".xml";
        public const string ZipExtension = ".zip";
        public const string HtmlExtension = ".html";
    }

}
