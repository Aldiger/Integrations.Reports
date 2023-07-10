using Integrations.Reports.Core.Dto;
using Integrations.Reports.Core.Integrations.Report.Dto;
using Reports.Core.Integrations.Report;

namespace Integrations.Reports.Core.Integrations.Report
{
    public class EmployeeReport : BaseReport
    {
        private readonly string _applicationName;
        private readonly IDictionary<int, ColumnModel> _columns;
        private readonly IList<EmployeeDto> _data;

        public EmployeeReport(ReportingServiceEmployeeReportDto model)
            : base(model.OutputType, model.OutputPath, new[] { "Employees" })
        {
            _data = model.Data;
            _applicationName = model.ApplicationName;

            var index = 0;
            _columns = new Dictionary<int, ColumnModel>
            {
 
                { ++index, new ColumnModel { Name = "Id", Width = "40", DataType = "str", HeaderStyle = "5", DataStyle = "10", TotalStyle = "12", Display = true } },
                { ++index, new ColumnModel { Name = "First Name", Width = "20", DataType = "str", HeaderStyle = "5", DataStyle = "10", TotalStyle = "12", Display = true } },
                { ++index, new ColumnModel { Name = "Last Name", Width = "18", DataType = "str", HeaderStyle = "5", DataStyle = "10", TotalStyle = "12", Display = true } },
                { ++index, new ColumnModel { Name = "Position", Width = "20", DataType = "str", HeaderStyle = "5", DataStyle = "10", TotalStyle = "12", Display = true } },
                { ++index, new ColumnModel { Name = "Date of Birth", Width = "30", DataType = "str", HeaderStyle = "5", DataStyle = "10", TotalStyle = "12", Display = true } },
            };
        }

        public ReportResultDto GenerateReport()
        {
            return GenerateExcel();
        }

        private ReportResultDto GenerateExcel()
        {
            ExcelStart();
            CreateExcelStyles("<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\" ?> <styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:s=\"http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:cdr=\"http://schemas.openxmlformats.org/drawingml/2006/chartDrawing\"> 	<numFmts count=\"2\"> 		<numFmt numFmtId=\"82\" formatCode=\"General\"/> 		<numFmt numFmtId=\"83\" formatCode=\"[$-010409]&quot;$&quot;#,##0.00;(&quot;$&quot;#,##0.00)\"/> 	</numFmts> 	<fonts count=\"7\"> 		<font> 			<sz val=\"11\"/> 			<color rgb=\"FF000000\"/> 			<name val=\"Calibri\"/> 			<family val=\"2\"/> 			<scheme val=\"minor\"/> 		</font> 		<font/> 		<font> 			<b/> 			<i val=\"0\"/> 			<strike val=\"0\"/> 			<u val=\"none\"/> 			<sz val=\"11\"/> 			<color rgb=\"FF000000\"/> 			<name val=\"Arial\"/> 		</font> 		<font> 			<b val=\"0\"/> 			<i val=\"0\"/> 			<strike val=\"0\"/> 			<u val=\"none\"/> 			<sz val=\"9\"/> 			<color rgb=\"FF000000\"/> 			<name val=\"Arial\"/> 		</font> 		<font> 			<b/> 			<i val=\"0\"/> 			<strike val=\"0\"/> 			<u val=\"none\"/> 			<sz val=\"9\"/> 			<color rgb=\"FF000000\"/> 			<name val=\"Arial\"/> 		</font> 		<font> 			<b val=\"0\"/> 			<i val=\"0\"/> 			<strike val=\"0\"/> 			<u val=\"none\"/> 			<sz val=\"8\"/> 			<color rgb=\"FF000000\"/> 			<name val=\"Arial\"/> 		</font> 		<font> 			<b/> 			<i val=\"0\"/> 			<strike val=\"0\"/> 			<u val=\"none\"/> 			<sz val=\"8\"/> 			<color rgb=\"FF000000\"/> 			<name val=\"Arial\"/> 		</font> 	</fonts> 	<fills count=\"2\"> 		<fill> 			<patternFill patternType=\"none\"/> 		</fill> 		<fill> 			<patternFill patternType=\"gray125\"/> 		</fill> 	</fills> 	<borders count=\"3\"> 		<border> 			<left/> 			<right/> 			<top/> 			<bottom/> 			<diagonal/> 		</border> 		<border> 			<left/> 			<right/> 			<top/> 			<bottom style=\"thin\"> 				<color rgb=\"FF000000\"/> 			</bottom> 			<diagonal/> 		</border> 		<border> 			<left/> 			<right/> 			<top style=\"thin\"> 				<color rgb=\"FF000000\"/> 			</top> 			<bottom/> 			<diagonal/> 		</border> 	</borders> 	<cellStyleXfs> 		<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/> 	</cellStyleXfs> 	<cellXfs count=\"14\"> 		<xf applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\"/> 		<xf applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\"> 			<alignment horizontal=\"center\" vertical=\"top\" textRotation=\"0\" wrapText=\"1\" readingOrder=\"1\"/> 		</xf> 		<xf applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" numFmtId=\"0\" fontId=\"3\" fillId=\"0\" borderId=\"0\"> 			<alignment horizontal=\"center\" vertical=\"top\" textRotation=\"0\" wrapText=\"1\" readingOrder=\"1\"/> 		</xf> 		<xf applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" numFmtId=\"0\" fontId=\"4\" fillId=\"0\" borderId=\"1\"> 			<alignment horizontal=\"general\" vertical=\"bottom\" textRotation=\"0\" wrapText=\"1\" readingOrder=\"1\"/> 		</xf> 		<xf applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" numFmtId=\"0\" fontId=\"4\" fillId=\"0\" borderId=\"1\"> 			<alignment horizontal=\"right\" vertical=\"bottom\" textRotation=\"0\" wrapText=\"1\" readingOrder=\"1\"/> 		</xf> 		<xf applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" numFmtId=\"0\" fontId=\"4\" fillId=\"0\" borderId=\"1\"> 			<alignment horizontal=\"center\" vertical=\"bottom\" textRotation=\"0\" wrapText=\"1\" readingOrder=\"1\"/> 		</xf> 		<xf applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" numFmtId=\"82\" fontId=\"5\" fillId=\"0\" borderId=\"0\"> 			<alignment horizontal=\"left\" vertical=\"top\" textRotation=\"0\" wrapText=\"1\" readingOrder=\"1\"/> 		</xf> 		<xf applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" numFmtId=\"0\" fontId=\"5\" fillId=\"0\" borderId=\"0\"> 			<alignment horizontal=\"general\" vertical=\"top\" textRotation=\"0\" wrapText=\"1\" readingOrder=\"1\"/> 		</xf> 		<xf applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" numFmtId=\"83\" fontId=\"5\" fillId=\"0\" borderId=\"0\"> 			<alignment horizontal=\"right\" vertical=\"top\" textRotation=\"0\" wrapText=\"1\" readingOrder=\"1\"/> 		</xf> 		<xf applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" numFmtId=\"82\" fontId=\"5\" fillId=\"0\" borderId=\"0\"> 			<alignment horizontal=\"right\" vertical=\"top\" textRotation=\"0\" wrapText=\"1\" readingOrder=\"1\"/> 		</xf> 		<xf applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" numFmtId=\"0\" fontId=\"5\" fillId=\"0\" borderId=\"0\"> 			<alignment horizontal=\"center\" vertical=\"top\" textRotation=\"0\" wrapText=\"1\" readingOrder=\"1\"/> 		</xf> 		<xf applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" numFmtId=\"0\" fontId=\"5\" fillId=\"0\" borderId=\"0\"> 			<alignment horizontal=\"right\" vertical=\"top\" textRotation=\"0\" wrapText=\"1\" readingOrder=\"1\"/> 		</xf> 		<xf applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" numFmtId=\"0\" fontId=\"6\" fillId=\"0\" borderId=\"2\"> 			<alignment horizontal=\"general\" vertical=\"top\" textRotation=\"0\" wrapText=\"1\" readingOrder=\"1\"/> 		</xf> 		<xf applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\" numFmtId=\"83\" fontId=\"6\" fillId=\"0\" borderId=\"2\"> 			<alignment horizontal=\"general\" vertical=\"top\" textRotation=\"0\" wrapText=\"1\" readingOrder=\"1\"/> 		</xf> 	</cellXfs> 	<cellStyles count=\"1\"> 		<cellStyle xfId=\"0\" name=\"Normal\"/> 	</cellStyles> 	<dxfs count=\"0\"/> 	<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" defaultPivotStyle=\"PivotStyleLight16\"/> 	<colors> 		<indexedColors> 			<rgbColor rgb=\"00000000\"/> 			<rgbColor rgb=\"00FFFFFF\"/> 			<rgbColor rgb=\"00FF0000\"/> 			<rgbColor rgb=\"0000FF00\"/> 			<rgbColor rgb=\"000000FF\"/> 			<rgbColor rgb=\"00FFFF00\"/> 			<rgbColor rgb=\"00FF00FF\"/> 			<rgbColor rgb=\"0000FFFF\"/> 			<rgbColor rgb=\"00000000\"/> 			<rgbColor rgb=\"00D3D3D3\"/> 			<rgbColor rgb=\"00FF0000\"/> 			<rgbColor rgb=\"0000FF00\"/> 			<rgbColor rgb=\"000000FF\"/> 			<rgbColor rgb=\"00FFFF00\"/> 			<rgbColor rgb=\"00FF00FF\"/> 			<rgbColor rgb=\"0000FFFF\"/> 			<rgbColor rgb=\"00800000\"/> 			<rgbColor rgb=\"00008000\"/> 			<rgbColor rgb=\"00000080\"/> 			<rgbColor rgb=\"00808000\"/> 			<rgbColor rgb=\"00800080\"/> 			<rgbColor rgb=\"00008080\"/> 			<rgbColor rgb=\"00C0C0C0\"/> 			<rgbColor rgb=\"00808080\"/> 			<rgbColor rgb=\"009999FF\"/> 			<rgbColor rgb=\"00993366\"/> 			<rgbColor rgb=\"00FFFFCC\"/> 			<rgbColor rgb=\"00CCFFFF\"/> 			<rgbColor rgb=\"00660066\"/> 			<rgbColor rgb=\"00FF8080\"/> 			<rgbColor rgb=\"000066CC\"/> 			<rgbColor rgb=\"00CCCCFF\"/> 			<rgbColor rgb=\"00000080\"/> 			<rgbColor rgb=\"00FF00FF\"/> 			<rgbColor rgb=\"00FFFF00\"/> 			<rgbColor rgb=\"0000FFFF\"/> 			<rgbColor rgb=\"00800080\"/> 			<rgbColor rgb=\"00800000\"/> 			<rgbColor rgb=\"00008080\"/> 			<rgbColor rgb=\"000000FF\"/> 			<rgbColor rgb=\"0000CCFF\"/> 			<rgbColor rgb=\"00CCFFFF\"/> 			<rgbColor rgb=\"00CCFFCC\"/> 			<rgbColor rgb=\"00FFFF99\"/> 			<rgbColor rgb=\"0099CCFF\"/> 			<rgbColor rgb=\"00FF99CC\"/> 			<rgbColor rgb=\"00CC99FF\"/> 			<rgbColor rgb=\"00FFCC99\"/> 			<rgbColor rgb=\"003366FF\"/> 			<rgbColor rgb=\"0033CCCC\"/> 			<rgbColor rgb=\"0099CC00\"/> 			<rgbColor rgb=\"00FFCC00\"/> 			<rgbColor rgb=\"00FF9900\"/> 			<rgbColor rgb=\"00FF6600\"/> 			<rgbColor rgb=\"00666699\"/> 			<rgbColor rgb=\"00969696\"/> 			<rgbColor rgb=\"00003366\"/> 			<rgbColor rgb=\"00339966\"/> 			<rgbColor rgb=\"00003300\"/> 			<rgbColor rgb=\"00333300\"/> 			<rgbColor rgb=\"00993300\"/> 			<rgbColor rgb=\"00993366\"/> 			<rgbColor rgb=\"00333399\"/> 			<rgbColor rgb=\"00333333\"/> 		</indexedColors> 	</colors> </styleSheet>");

            using (TextWriter tw = new StreamWriter(SheetPath, true))
            {
                #region Columns Settings

                tw.Write("<cols>");
                var col = 0;
                foreach (var item in _columns.Where(g => g.Value.Display))
                {
                    tw.Write(GetExcelColumnDefinition(item.Value.Width, ++col));
                }
                tw.Write("</cols>");

                #endregion

                tw.Write("<sheetData>");

                #region Excel Rows

                var row = 0;
                var date = DateTime.UtcNow;

                //Generic headers
                tw.Write("<row>");
                tw.Write(GetExcelCellData($"{_applicationName} Employees", ++row, 1, "str", "1"));
                tw.Write("</row>");


                tw.Write("<row>");
                tw.Write(GetExcelCellData($"{date.ToShortDateString()} {date.ToShortTimeString()} EST", ++row, 1, "str", "2"));
                tw.Write("</row>");
                tw.Write("<row>");
                tw.Write(GetExcelCellData(" ", ++row, 1, "str", "2"));
                tw.Write("</row>");

                #region Column Headers Row

                col = 0;
                row++;
                tw.Write("<row>");
                foreach (var item in _columns.Where(g => g.Value.Display))
                {
                    tw.Write(GetExcelCellColumn(item.Value.Name, row, ++col, item.Value.HeaderStyle));
                }
                tw.Write("</row>");

                #endregion

                #region Detail Rows

                foreach (var item in _data)
                {
                    row++;
                    col = 0;
                    var index = 0;

                    tw.Write("<row>");
           
                    tw.Write(_columns[++index].Display ? GetExcelCellData(item.Id.ToString(), row, ++col, _columns[index].DataType, _columns[index].DataStyle) : string.Empty);
                    tw.Write(_columns[++index].Display ? GetExcelCellData(item.FirstName, row, ++col, _columns[index].DataType, _columns[index].DataStyle) : string.Empty);
                    tw.Write(_columns[++index].Display ? GetExcelCellData(item.LastName, row, ++col, _columns[index].DataType, _columns[index].DataStyle) : string.Empty);
                    tw.Write(_columns[++index].Display ? GetExcelCellData(item.Position, row, ++col, _columns[index].DataType, _columns[index].DataStyle) : string.Empty);
                    tw.Write(_columns[++index].Display ? GetExcelCellData(item.DateOfBirth.ToShortDateString(), row, ++col, _columns[index].DataType, _columns[index].DataStyle) : string.Empty);
                    tw.Write("</row>");
                }

                #endregion


                #endregion

                tw.Write("</sheetData>");

                #region Merge Cels

                tw.Write("<mergeCells>");
                var columnCount = _columns.Count(g => g.Value.Display);
                tw.Write(GetExcelMergeCellDefinition(1, 1, columnCount));
                tw.Write(GetExcelMergeCellDefinition(2, 1, columnCount));
                tw.Write("</mergeCells>");

                #endregion
            }
            ExcelFinish();

            return new ReportResultDto
            {
                FileName = FileName,
                FilePath = FilePath
            };
        }
    }
}
