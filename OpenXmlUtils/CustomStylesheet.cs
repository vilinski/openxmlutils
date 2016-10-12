#region File Information
//
// File: "CustomStylesheet.cs"
// Purpose: "Defines how a spreadsheet will look."
// Author: "Geoplex"
// 
#endregion

#region (c) Copyright 2014 Geoplex
//
// THE SOFTWARE IS PROVIDED "AS-IS" AND WITHOUT WARRANTY OF ANY KIND,
// EXPRESS, IMPLIED OR OTHERWISE, INCLUDING WITHOUT LIMITATION, ANY
// WARRANTY OF MERCHANTABILITY OR FITNESS FOR A PARTICULAR PURPOSE.
//
// IN NO EVENT SHALL GEOPLEX BE LIABLE FOR ANY SPECIAL, INCIDENTAL,
// INDIRECT OR CONSEQUENTIAL DAMAGES OF ANY KIND, OR ANY DAMAGES WHATSOEVER
// RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER OR NOT ADVISED OF THE
// POSSIBILITY OF DAMAGE, AND ON ANY THEORY OF LIABILITY, ARISING OUT OF OR IN
// CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
//
#endregion

using System.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;

namespace OpenXmlUtils
{
    public class CustomStylesheet : Stylesheet
    {
        public enum CustomCellFormats : uint
        {
            // these are referenced by index, must be added in this order
            DefaultText = 0,
            DefaultDate,
            DefaultNumber2DecimalPlace,
            DefaultNumber5DecimalPlace,
            DefaultDateTime,
            HeaderText,
            TotalsNumber,
            TotalsNumber2DecimalPlace,
            TotalsText,
            TitleText,
            SubtitleText,
            Duration,
            TotalsDuration,
            Hyperlink
        }

        public CustomStylesheet()
        {
            NumberingFormat nfDateTime;
            NumberingFormat nf5Decimal;
            NumberingFormat nfDuration;
            NumberingFormat nfTotalDuration;

            Append(CreateNumberingFormats(out nfDateTime, out nf5Decimal, out nfDuration, out nfTotalDuration));
            Append(CreateFonts());
            Append(CreateFills());
            Append(CreateBorders());
            Append(CreateCellStyleFormats());
            Append(CreateCellFormats(nfDateTime, nf5Decimal, nfDuration, nfTotalDuration));
            Append(CreateCellStyles());
            Append(CreateDifferentialFormats());
            Append(CreateTableStyles());
        }

        private static TableStyles CreateTableStyles()
        {
            return new TableStyles
            {
                Count = 0,
                DefaultTableStyle = StringValue.FromString("TableStyleMedium9"),
                DefaultPivotStyle = StringValue.FromString("PivotStyleLight16")
            };
        }

        private static DifferentialFormats CreateDifferentialFormats()
        {
            return new DifferentialFormats {Count = 0};
        }

        private static CellStyles CreateCellStyles()
        {

            // cell style 0
            var cs = new CellStyle
            {
                Name = StringValue.FromString("Normal"),
                FormatId = 0,
                BuiltinId = 0
            };
            var css = new CellStyles();
            css.Count = UInt32Value.FromUInt32((uint) css.ChildElements.Count);
            css.AppendChild(cs);
            return css;
        }

        /// <summary>
        /// Ensure cell formats are added in the order specified by the enumeration
        /// </summary>
        private static CellFormats CreateCellFormats(NumberingFormat nfDateTime, NumberingFormat nf5Decimal,
            NumberingFormat nfDuration, NumberingFormat nfTotalDuration)
        {
            var cfs = new CellFormats();

            // CustomCellFormats.DefaultText
            var cf = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(false)
            };
            cfs.AppendChild(cf);

            // CustomCellFormats.DefaultDate
            cf = new CellFormat
            {
                NumberFormatId = 14,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            // mm-dd-yy
            cfs.AppendChild(cf);

            // CustomCellFormats.DefaultNumber2DecimalPlace
            cf = new CellFormat
            {
                NumberFormatId = 4,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            // #,##0.00
            cfs.AppendChild(cf);

            // CustomCellFormats.DefaultNumber5DecimalPlace
            cf = new CellFormat
            {
                NumberFormatId = nf5Decimal.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cfs.AppendChild(cf);

            // CustomCellFormats.DefaultDateTime
            cf = new CellFormat
            {
                NumberFormatId = nfDateTime.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cfs.AppendChild(cf);

            // CustomCellFormats.HeaderText
            cf = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 1,
                FillId = 2,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(false)
            };
            cfs.AppendChild(cf);

            // CustomCellFormats.TotalsNumber
            cf = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 3,
                BorderId = 2,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cfs.AppendChild(cf);

            // CustomCellFormats.TotalsNumber2DecimalPlace
            cf = new CellFormat
            {
                NumberFormatId = 4,
                FontId = 0,
                FillId = 3,
                BorderId = 2,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            // #,##0.00
            cfs.AppendChild(cf);

            // CustomCellFormats.TotalsText
            cf = new CellFormat
            {
                NumberFormatId = 49,
                FontId = 0,
                FillId = 3,
                BorderId = 2,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            // @
            cfs.AppendChild(cf);

            // CustomCellFormats.TitleText
            cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 2;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(false);
            cf.Alignment = new Alignment
            {
                Vertical = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Bottom)
            };
            cfs.AppendChild(cf);

            // CustomCellFormats.SubtitleText
            cf = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 3,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(false),
                Alignment = new Alignment
                {
                    Vertical = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Top)
                }
            };
            cfs.AppendChild(cf);

            // CustomCellFormats.Duration
            cf = new CellFormat
            {
                NumberFormatId = nfDuration.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
                Alignment = new Alignment
                {
                    Horizontal = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Right)
                }
            };
            // [h]:mm
            cfs.AppendChild(cf);

            // CustomCellFormats.TotalsNumber
            cf = new CellFormat
            {
                NumberFormatId = nfTotalDuration.NumberFormatId,
                FontId = 0,
                FillId = 3,
                BorderId = 2,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true),
                Alignment = new Alignment
                {
                    Horizontal = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Right)
                }
            };
            // d:h:mm
            cfs.AppendChild(cf);

            // CustomCellFormats.Hyperlink
            cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 4;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(false);
            cfs.AppendChild(cf);

            cfs.Count = UInt32Value.FromUInt32((uint) cfs.ChildElements.Count);
            return cfs;
        }

        private static NumberingFormats CreateNumberingFormats(out NumberingFormat nfDateTime,
            out NumberingFormat nf5Decimal, out NumberingFormat nfDuration, out NumberingFormat nfTotalDuration)
        {
            // built-in formats go up to 164
            uint iExcelIndex = 164;

            var nfs = new NumberingFormats();
            nfDateTime = new NumberingFormat();
            nfDateTime.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
            nfDateTime.FormatCode = StringValue.FromString("dd/mm/yyyy hh:mm:ss");
            nfs.AppendChild(nfDateTime);

            nf5Decimal = new NumberingFormat();
            nf5Decimal.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
            nf5Decimal.FormatCode = StringValue.FromString("#,##0.00000");
            nfs.AppendChild(nf5Decimal);

            nfDuration = new NumberingFormat();
            nfDuration.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
            nfDuration.FormatCode = StringValue.FromString("[h]:mm");
            nfs.AppendChild(nfDuration);

            nfTotalDuration = new NumberingFormat();
            nfTotalDuration.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
            nfTotalDuration.FormatCode = StringValue.FromString("d:h:mm");
            nfs.AppendChild(nfTotalDuration);

            nfs.Count = UInt32Value.FromUInt32((uint) nfs.ChildElements.Count);
            return nfs;
        }

        private static CellStyleFormats CreateCellStyleFormats()
        {
            var csfs = new CellStyleFormats();

            // cell style 0
            var cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            csfs.AppendChild(cf);
            csfs.Count = UInt32Value.FromUInt32((uint) csfs.ChildElements.Count);
            return csfs;
        }

        private static Borders CreateBorders()
        {
            var borders = new Borders();

            // boarder index 0
            var border = new Border();
            border.LeftBorder = new LeftBorder();
            border.RightBorder = new RightBorder();
            border.TopBorder = new TopBorder();
            border.BottomBorder = new BottomBorder();
            border.DiagonalBorder = new DiagonalBorder();
            borders.AppendChild(border);

            // boarder Index 1
            border = new Border();
            border.LeftBorder = new LeftBorder();
            border.LeftBorder.Style = BorderStyleValues.Thin;
            border.RightBorder = new RightBorder();
            border.RightBorder.Style = BorderStyleValues.Thin;
            border.TopBorder = new TopBorder();
            border.TopBorder.Style = BorderStyleValues.Thin;
            border.BottomBorder = new BottomBorder();
            border.BottomBorder.Style = BorderStyleValues.Thin;
            border.DiagonalBorder = new DiagonalBorder();
            borders.AppendChild(border);

            // boarder Index 2
            border = new Border();
            border.LeftBorder = new LeftBorder();
            border.RightBorder = new RightBorder();
            border.TopBorder = new TopBorder();
            border.TopBorder.Style = BorderStyleValues.Thin;
            border.BottomBorder = new BottomBorder();
            border.BottomBorder.Style = BorderStyleValues.Thin;
            border.DiagonalBorder = new DiagonalBorder();
            borders.AppendChild(border);

            borders.Count = UInt32Value.FromUInt32((uint) borders.ChildElements.Count);
            return borders;
        }

        private static Fills CreateFills()
        {
            // fill 0
            var fills = new Fills();
            var fill = new Fill();
            var patternFill = new PatternFill {PatternType = PatternValues.None};
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);

            // fill 1 (in-built fill)
            fill = new Fill();
            patternFill = new PatternFill { PatternType = PatternValues.Gray125 };
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);

            // fill 2
            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Solid;
            var fillColor = Color.LightSkyBlue;
            patternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValueFromColor(fillColor) };
            patternFill.BackgroundColor = new BackgroundColor { Rgb = HexBinaryValueFromColor(fillColor) };
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);

            // fill 3
            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Solid;
            fillColor = Color.Orange;
            patternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValueFromColor(fillColor) };
            patternFill.BackgroundColor = new BackgroundColor { Rgb = HexBinaryValueFromColor(fillColor) };
            fill.PatternFill = patternFill;
            fills.AppendChild(fill);

            fills.Count = UInt32Value.FromUInt32((uint) fills.ChildElements.Count);
            return fills;
        }

        private static Fonts CreateFonts()
        {
            var fts = new Fonts();

            // font 0
            var ft = new Font();
            var ftn = new FontName {Val = StringValue.FromString("Arial")};
            var ftsz = new FontSize {Val = DoubleValue.FromDouble(11)};
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            fts.AppendChild(ft);

            // font 1
            ft = new Font();
            ftn = new FontName { Val = StringValue.FromString("Arial") };
            ftsz = new FontSize { Val = DoubleValue.FromDouble(12) };
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            ft.Bold = new Bold();
            fts.AppendChild(ft);

            // font 2
            ft = new Font();
            ftn = new FontName { Val = StringValue.FromString("Arial") };
            ftsz = new FontSize { Val = DoubleValue.FromDouble(18) };
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            ft.Bold = new Bold();
            fts.AppendChild(ft);

            // font 3
            ft = new Font();
            ftn = new FontName { Val = StringValue.FromString("Arial") };
            ftsz = new FontSize { Val = DoubleValue.FromDouble(14) };
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            fts.AppendChild(ft);

            // font 4
            ft = new Font();
            ftn = new FontName { Val = StringValue.FromString("Arial") };
            ftsz = new FontSize { Val = DoubleValue.FromDouble(11) };
            var fontColor = Color.MediumBlue;
            ft.Color = new DocumentFormat.OpenXml.Spreadsheet.Color() { Rgb = HexBinaryValueFromColor(fontColor) };
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            fts.AppendChild(ft);

            fts.Count = UInt32Value.FromUInt32((uint) fts.ChildElements.Count);
            return fts;
        }

        private static HexBinaryValue HexBinaryValueFromColor(Color fillColor)
        {
            return new HexBinaryValue
            {
                Value =
                    ColorTranslator.ToHtml(
                        Color.FromArgb(
                            fillColor.A,
                            fillColor.R,
                            fillColor.G,
                            fillColor.B)).Replace("#", "")
            };
        }
    }
}