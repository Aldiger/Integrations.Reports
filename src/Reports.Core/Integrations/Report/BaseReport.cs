using Reports.Core.Helpers;
using System.IO.Compression;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Reports.Core.Integrations.Report
{
    public abstract class BaseReport
    {
        private readonly string _tempOutputPath;
        private readonly string[] _sheets;
        private const int DataValidationsMaxRow = 1048576;

        protected internal ReportTypes Type;
        protected internal string SheetPath;
        protected internal string[] SheetPaths;
        protected internal string FileName;
        protected internal string FilePath;

        protected BaseReport(ReportTypes type, string outputPath, string[] sheets)
        {
            Type = type;
            var guid = Guid.NewGuid();
            DirectoryHelper.CheckCreatePath(outputPath);

            switch (type)
            {
                case ReportTypes.Excel:
                    _tempOutputPath = $"{outputPath.TrimEndPath()}\\{guid}";
                    FileName = $"{guid}{Constants.ExcelExtension}";

                    _sheets = sheets.Where(g => g != null).ToArray();
                    if (!_sheets.Any())
                    {
                        _sheets = new[] { "sheet1" };
                    }
                    SheetPaths = _sheets.Select((x, row) => $"{_tempOutputPath}\\xl\\worksheets\\sheet{row + 1}.xml").ToArray();
                    SheetPath = SheetPaths.First();
                    break;
                case ReportTypes.Csv:
                    FileName = $"{guid}{Constants.CommaDelimitedExtension}";
                    break;
                case ReportTypes.Xml:
                    FileName = $"{guid}{Constants.XmlExtension}";
                    break;
            }

            FilePath = $"{outputPath.TrimEndPath()}\\{FileName}";
        }

        #region Excel methods

        protected internal void ExcelStart()
        {
            CreateFolders();
            CreateRels();
            CreateDocProps();
            CreateContentTypes();
            CreateTheme();
            CreateWorkbook();
            foreach (var sheetPath in SheetPaths)
            {
                StartSheet(sheetPath);
            }
        }


        protected internal void ExcelFinish()
        {
            foreach (var sheetPath in SheetPaths)
            {
                FinishSheet(sheetPath);
            }
            ZipFile.CreateFromDirectory(_tempOutputPath, FilePath);
            DirectoryHelper.DeleteDirectory(_tempOutputPath);
        }

        protected internal string GetExcelCellColumn(string value, int row, int col, string style)
        {
            return GetExcelCellData(value, row, col, "str", style);
        }

        protected internal string GetExcelCellData(string value, int row, int col, string dataType, string style,
            bool includeCellEmptyStructure = false)
        {
            if (string.IsNullOrEmpty(value) && !includeCellEmptyStructure)
            {
                return string.Empty;
            }

            var t = string.IsNullOrEmpty(dataType) ? string.Empty : $" t=\"{dataType}\"";
            return $"<c r=\"{GetFullColumnName(col, row)}\"{t} s=\"{style}\"><v>{EscapeInvalidChars(SecurityElement.Escape(value?.Trim() ?? string.Empty))}</v></c>";
        }

        protected internal string GetExcelCellSum(string value, int row, int startRow, int endRow, int col, string dataType, string style)
        {
            var cell = GetColumnName(col);
            return !string.IsNullOrEmpty(value) && string.IsNullOrEmpty(dataType)
                ? $"<c r=\"{cell}{row}\" s=\"{style}\"><f>SUM({cell}{startRow}:{cell}{endRow})</f><v>{value}</v></c>"
                : $"<c r=\"{cell}{row}\" s=\"{style}\" t=\"{dataType}\"><v>{value}</v></c>";
        }

        protected internal string GetExcelColumnDefinition(string width, int col)
        {
            return $"<col width=\"{width}\" min=\"{col}\" max=\"{col}\"/>";
        }

        protected internal string GetExcelDataValidation(string dataValidationString, int column, int row)
        {
            var columnName = GetColumnName(column);
            return "<dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"" + $"{columnName}{row}:{columnName}{DataValidationsMaxRow}" + "\" xr:uid=\"{" + Guid.NewGuid() + "}\"><formula1>\"" + dataValidationString + "\"</formula1></dataValidation>";
        }

        protected internal string GetExcelMergeCellDefinition(int row, int colStart, int colEnd)
        {
            return $"<mergeCell ref=\"{GetFullColumnName(colStart, row)}:{GetFullColumnName(colEnd, row)}\"/>";
        }

        protected internal void CreateExcelStyles(string value)
        {
            using (TextWriter tw = new StreamWriter($"{_tempOutputPath}/xl/styles.xml", false))
            {
                tw.Write(value);
            }
        }

        private void StartSheet(string sheetPath)
        {
            using (TextWriter tw = new StreamWriter(sheetPath))
            {
                tw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac xr xr2 xr3\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\" xr:uid=\"{" + Guid.NewGuid() + "}\"><dimension ref=\"A1:C2\" /><sheetViews><sheetView " + $"{(sheetPath.Contains("sheet1.xml") ? "tabSelected =\"1\"" : "")} workbookViewId=\"0\"><selection activeCell=\"A1\" sqref=\"A1\" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15\" x14ac:dyDescent=\"0.25\" />");
            }
        }

        private void FinishSheet(string sheetPath)
        {
            using (TextWriter tw = new StreamWriter(sheetPath, true))
            {
                tw.Write("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\" /></worksheet>");
            }
        }

        private void CreateFolders()
        {
            Directory.CreateDirectory($"{_tempOutputPath}/_rels");
            Directory.CreateDirectory($"{_tempOutputPath}/docProps");
            Directory.CreateDirectory($"{_tempOutputPath}/xl");
            Directory.CreateDirectory($"{_tempOutputPath}/xl/_rels");
            Directory.CreateDirectory($"{_tempOutputPath}/xl/theme");
            Directory.CreateDirectory($"{_tempOutputPath}/xl/worksheets");
        }

        private void CreateRels()
        {
            using (TextWriter tw = new StreamWriter($"{_tempOutputPath}/_rels/.rels", false))
            {
                tw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
            }
            using (TextWriter tw = new StreamWriter($"{_tempOutputPath}/xl/_rels/workbook.xml.rels", false))
            {
                var relationshipNodes = string.Empty;
                for (var i = 1; i < _sheets.Length + 1; i++)
                {
                    relationshipNodes +=
                        $"<Relationship Id=\"rId{i + 6}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{i}.xml\"/>";
                }

                tw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                         "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                         "<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>" +
                         "<Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
                         "<Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>" +
                            relationshipNodes +
                         "</Relationships>");
            }
        }

        private void CreateDocProps()
        {
            using (TextWriter tw = new StreamWriter($"{_tempOutputPath}/docProps/app.xml", false))
            {
                tw.Write("<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\" ?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:s=\"http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:cdr=\"http://schemas.openxmlformats.org/drawingml/2006/chartDrawing\"><Pages>0</Pages><Words>0</Words><Characters>0</Characters><Lines>0</Lines><Paragraphs>0</Paragraphs><Slides>0</Slides><Notes>0</Notes><TotalTime>0</TotalTime><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>0</ScaleCrop><HeadingPairs><vt:vector baseType=\"variant\" size=\"2\"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector baseType=\"lpstr\" size=\"0\"></vt:vector></TitlesOfParts><LinksUpToDate>0</LinksUpToDate><CharactersWithSpaces>0</CharactersWithSpaces><SharedDoc>0</SharedDoc><HyperlinksChanged>0</HyperlinksChanged><Application>Microsoft Excel</Application><AppVersion>12.0000</AppVersion><DocSecurity>0</DocSecurity></Properties>");
            }
            using (TextWriter tw = new StreamWriter($"{_tempOutputPath}/docProps/core.xml", false))
            {
                tw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"></cp:coreProperties>");
            }
        }

        private void CreateContentTypes()
        {
            using (TextWriter tw = new StreamWriter($"{_tempOutputPath}/[Content_Types].xml", false))
            {
                var sheetContentType = string.Empty;
                for (var i = 1; i < _sheets.Length + 1; i++)
                {
                    sheetContentType += $"<Override PartName=\"/xl/worksheets/sheet{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
                }

                tw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                         sheetContentType +
                         "<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/></Types>");
            }
        }

        private void CreateTheme()
        {
            using (TextWriter tw = new StreamWriter($"{_tempOutputPath}/xl/theme/theme1.xml", false))
            {
                tw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\"><a:themeElements><a:clrScheme name=\"Office\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"44546A\"/></a:dk2><a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2><a:accent1><a:srgbClr val=\"5B9BD5\"/></a:accent1><a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2><a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3><a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4><a:accent5><a:srgbClr val=\"4472C4\"/></a:accent5><a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6><a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink><a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink></a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"Calibri Light\" panose=\"020F0302020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"游ゴシック Light\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"等线 Light\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Times New Roman\"/><a:font script=\"Hebr\" typeface=\"Times New Roman\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"MoolBoran\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Times New Roman\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/></a:majorFont><a:minorFont><a:latin typeface=\"Calibri\" panose=\"020F0502020204030204\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/><a:font script=\"Jpan\" typeface=\"游ゴシック\"/><a:font script=\"Hang\" typeface=\"맑은 고딕\"/><a:font script=\"Hans\" typeface=\"等线\"/><a:font script=\"Hant\" typeface=\"新細明體\"/><a:font script=\"Arab\" typeface=\"Arial\"/><a:font script=\"Hebr\" typeface=\"Arial\"/><a:font script=\"Thai\" typeface=\"Tahoma\"/><a:font script=\"Ethi\" typeface=\"Nyala\"/><a:font script=\"Beng\" typeface=\"Vrinda\"/><a:font script=\"Gujr\" typeface=\"Shruti\"/><a:font script=\"Khmr\" typeface=\"DaunPenh\"/><a:font script=\"Knda\" typeface=\"Tunga\"/><a:font script=\"Guru\" typeface=\"Raavi\"/><a:font script=\"Cans\" typeface=\"Euphemia\"/><a:font script=\"Cher\" typeface=\"Plantagenet Cherokee\"/><a:font script=\"Yiii\" typeface=\"Microsoft Yi Baiti\"/><a:font script=\"Tibt\" typeface=\"Microsoft Himalaya\"/><a:font script=\"Thaa\" typeface=\"MV Boli\"/><a:font script=\"Deva\" typeface=\"Mangal\"/><a:font script=\"Telu\" typeface=\"Gautami\"/><a:font script=\"Taml\" typeface=\"Latha\"/><a:font script=\"Syrc\" typeface=\"Estrangelo Edessa\"/><a:font script=\"Orya\" typeface=\"Kalinga\"/><a:font script=\"Mlym\" typeface=\"Kartika\"/><a:font script=\"Laoo\" typeface=\"DokChampa\"/><a:font script=\"Sinh\" typeface=\"Iskoola Pota\"/><a:font script=\"Mong\" typeface=\"Mongolian Baiti\"/><a:font script=\"Viet\" typeface=\"Arial\"/><a:font script=\"Uigh\" typeface=\"Microsoft Uighur\"/><a:font script=\"Geor\" typeface=\"Sylfaen\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln><a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs><a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs><a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri=\"{05A4C25C-085E-4340-85A3-A5531E510DB2}\"><thm15:themeFamily xmlns:thm15=\"http://schemas.microsoft.com/office/thememl/2012/main\" name=\"Office Theme\" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\" vid=\"{4A3C46E8-61CC-4603-A589-7422A47A8E4A}\"/></a:ext></a:extLst></a:theme>");
            }
        }

        private void CreateWorkbook()
        {
            using (TextWriter tw = new StreamWriter($"{_tempOutputPath}/xl/workbook.xml", false))
            {
                var sheets = string.Empty;
                for (var i = 1; i < _sheets.Length + 1; i++)
                {
                    sheets += $"<sheet name=\"{_sheets[i - 1]}\" sheetId=\"{i}\" r:id=\"rId{i + 6}\"/>";
                }


                tw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x15 xr xr6 xr10 xr2\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr6=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision6\" xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\">" +
                         $"<fileVersion appName=\"xl\" lastEdited=\"{sheets.Length}\" lowestEdited=\"{sheets.Length}\" rupBuild=\"21328\"/>" +
                         "<workbookPr filterPrivacy=\"1\"/><bookViews><workbookView xWindow=\"-120\" yWindow=\"-120\" windowWidth=\"29040\" windowHeight=\"15840\" xr2:uid=\"{00000000-000D-0000-FFFF-FFFF00000000}\"/></bookViews>" +
                 "<sheets>" +
                    sheets +
                 "</sheets>" +
                 "<calcPr calcId=\"162913\"/></workbook>");
            }
        }

        private string GetFullColumnName(int col, int row)
        {
            return $"{GetColumnName(col)}{row}";
        }

        private string GetColumnName(int col)
        {
            string cell;
            if (col <= 26)
            {
                cell = $"{(char)(col + 64)}";
            }
            else
            {
                var firstLetter = (char)((col - 1) / 26 + 64);
                var secondLetter = (char)((col - 1) % 26 + 65);
                cell = $"{firstLetter}{secondLetter}";
            }

            return cell;
        }

        private static string EscapeInvalidChars(string text)
        {
            if (string.IsNullOrEmpty(text) || text.All(XmlConvert.IsXmlChar))
            {
                return text;
            }

            var result = new StringBuilder();
            foreach (var item in text)
            {
                if (XmlConvert.IsXmlChar(item))
                {
                    result.Append(item);
                }
                else
                {
                    result.Append($"_x{(int)item:x4}_");
                }
            }

            return result.ToString();
        }

        #endregion

        #region CSV methods

        protected internal string GetCsvData(string value, bool lastElement)
        {
            var result = value?.Trim() ?? string.Empty;

            if (!string.IsNullOrEmpty(value) && (result.Contains(",") || result.Contains("\"")))
            {
                result = $"\"{result.Replace("\"", "\"\"")}\"";
            }

            if (!lastElement)
            {
                result = $"{result},";
            }

            return result;
        }

        #endregion

        #region XML methods

        protected internal string GetXmlData(string name, string value)
        {
            return $"{name}=\"{RemoveInvalidChars(SecurityElement.Escape(string.IsNullOrEmpty(value) ? string.Empty : value.Trim()))}\" ";
        }

        protected internal string FormatXmlColumnName(string input, int position)
        {
            //leave only digits and letters
            var result = Regex.Replace(input, "[^a-zA-Z0-9]", "");

            //strip digits from the beginning
            var digits = new[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
            result = result.TrimStart(digits);

            //set default name if nothing left (custom column like "2.34567")
            return string.IsNullOrEmpty(result)
                ? $"ColumnName{position}"
                : result;
        }

        private static string RemoveInvalidChars(string text)
        {
            if (string.IsNullOrEmpty(text) || text.All(XmlConvert.IsXmlChar))
            {
                return text;
            }

            return new string(text.Where(XmlConvert.IsXmlChar).ToArray());
        }

        #endregion

        #region Column Model

        protected internal class ColumnModel
        {
            public string Name { get; set; }
            public string Width { get; set; }
            public string DataType { get; set; }
            public string HeaderStyle { get; set; }
            public string DataStyle { get; set; }
            public string TotalStyle { get; set; }
            public bool HasTotal { get; set; }
            public string Total { get; set; }
            public bool Display { get; set; }
            public string DataValidation { get; set; }
            public bool HasDataValidation { get; set; }
        }

        #endregion
    }
}
