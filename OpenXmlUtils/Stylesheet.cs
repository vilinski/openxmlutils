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
            // built-in formats go up to 164
            uint iExcelIndex = 164;

            var nfs = new NumberingFormats();
            var nfDateTime = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                FormatCode = StringValue.FromString("dd/mm/yyyy hh:mm:ss")
            };
            nfs.AppendChild(nfDateTime);

            var nf5Decimal = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                FormatCode = StringValue.FromString("#,##0.00000")
            };
            nfs.AppendChild(nf5Decimal);

            var nfDuration = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                FormatCode = StringValue.FromString("[h]:mm")
            };
            nfs.AppendChild(nfDuration);

            var nfTotalDuration = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                FormatCode = StringValue.FromString("d:h:mm")
            };
            nfs.AppendChild(nfTotalDuration);

            nfs.Count = UInt32Value.FromUInt32((uint) nfs.ChildElements.Count);
            Append(nfs);
            Append(CreateFonts());
            Append(CreateFills());
            Append(CreateBorders());
            Append(CreateCellStyleFormats());
            Append(CreateCellFormats(nfDateTime, nf5Decimal, nfDuration, nfTotalDuration));
            Append(CreateCellStyles());
            Append(new DifferentialFormats {Count = 0});
            Append(new TableStyles
            {
                Count = 0,
                DefaultTableStyle = StringValue.FromString("TableStyleMedium9"),
                DefaultPivotStyle = StringValue.FromString("PivotStyleLight16")
            });
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
            cfs.AppendChild(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(false)
            });

            // CustomCellFormats.DefaultDate
            // mm-dd-yy
            cfs.AppendChild(new CellFormat
            {
                NumberFormatId = 14,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });

            // CustomCellFormats.DefaultNumber2DecimalPlace
            // #,##0.00
            cfs.AppendChild(new CellFormat
            {
                NumberFormatId = 4,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });

            // CustomCellFormats.DefaultNumber5DecimalPlace
            cfs.AppendChild(new CellFormat
            {
                NumberFormatId = nf5Decimal.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });

            // CustomCellFormats.DefaultDateTime
            cfs.AppendChild(new CellFormat
            {
                NumberFormatId = nfDateTime.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });

            // CustomCellFormats.HeaderText
            cfs.AppendChild(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 1,
                FillId = 2,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(false)
            });

            // CustomCellFormats.TotalsNumber
            cfs.AppendChild(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 3,
                BorderId = 2,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });

            // CustomCellFormats.TotalsNumber2DecimalPlace
            // #,##0.00
            cfs.AppendChild(new CellFormat
            {
                NumberFormatId = 4,
                FontId = 0,
                FillId = 3,
                BorderId = 2,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });

            // CustomCellFormats.TotalsText
            // @
            cfs.AppendChild(new CellFormat
            {
                NumberFormatId = 49,
                FontId = 0,
                FillId = 3,
                BorderId = 2,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });

            // CustomCellFormats.TitleText
            cfs.AppendChild(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 2,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(false),
                Alignment = new Alignment
                {
                    Vertical = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Bottom)
                }
            });

            // CustomCellFormats.SubtitleText
            cfs.AppendChild(new CellFormat
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
            });

            // CustomCellFormats.Duration
            // [h]:mm
            cfs.AppendChild(new CellFormat
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
            });

            // CustomCellFormats.TotalsNumber
            // d:h:mm
            cfs.AppendChild(new CellFormat
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
            });

            // CustomCellFormats.Hyperlink
            cfs.AppendChild(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 4,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(false)
            });

            cfs.Count = UInt32Value.FromUInt32((uint) cfs.ChildElements.Count);
            return cfs;
        }

        private static CellStyleFormats CreateCellStyleFormats()
        {
            var csfs = new CellStyleFormats();

            // cell style 0
            csfs.AppendChild(new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            });
            csfs.Count = UInt32Value.FromUInt32((uint) csfs.ChildElements.Count);
            return csfs;
        }

        private static Borders CreateBorders()
        {
            var borders = new Borders();

            // boarder index 0
            borders.AppendChild(new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            });

            // boarder Index 1
            borders.AppendChild(new Border
            {
                LeftBorder = new LeftBorder {Style = BorderStyleValues.Thin},
                RightBorder = new RightBorder {Style = BorderStyleValues.Thin},
                TopBorder = new TopBorder {Style = BorderStyleValues.Thin},
                BottomBorder = new BottomBorder {Style = BorderStyleValues.Thin},
                DiagonalBorder = new DiagonalBorder()
            });

            // boarder Index 2
            borders.AppendChild(new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder {Style = BorderStyleValues.Thin},
                BottomBorder = new BottomBorder {Style = BorderStyleValues.Thin},
                DiagonalBorder = new DiagonalBorder()
            });

            borders.Count = UInt32Value.FromUInt32((uint) borders.ChildElements.Count);
            return borders;
        }

        private static Fills CreateFills()
        {
            // fill 0
            var fills = new Fills();
            fills.AppendChild(new Fill {PatternFill = new PatternFill {PatternType = PatternValues.None}});

            // fill 1 (in-built fill)
            fills.AppendChild(new Fill
            {
                PatternFill = new PatternFill
                {
                    PatternType = PatternValues.Gray125
                }
            });

            // fill 2
            fills.AppendChild(new Fill
            {
                PatternFill = new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor {Rgb = HexBinaryValueFromColor(Color.LightSkyBlue)},
                    BackgroundColor = new BackgroundColor {Rgb = HexBinaryValueFromColor(Color.LightSkyBlue)}
                }
            });

            // fill 3
            fills.AppendChild(new Fill
            {
                PatternFill = new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor {Rgb = HexBinaryValueFromColor(Color.Orange)},
                    BackgroundColor = new BackgroundColor {Rgb = HexBinaryValueFromColor(Color.Orange)}
                }
            });

            fills.Count = UInt32Value.FromUInt32((uint) fills.ChildElements.Count);
            return fills;
        }

        private static Fonts CreateFonts()
        {
            var fts = new Fonts();

            // font 0
            fts.AppendChild(new Font
            {
                FontName = new FontName {Val = StringValue.FromString("Arial")},
                FontSize = new FontSize {Val = DoubleValue.FromDouble(11)}
            });

            // font 1
            fts.AppendChild(new Font
            {
                FontName = new FontName {Val = StringValue.FromString("Arial")},
                FontSize = new FontSize {Val = DoubleValue.FromDouble(12)},
                Bold = new Bold()
            });

            // font 2
            fts.AppendChild(new Font
            {
                FontName = new FontName {Val = StringValue.FromString("Arial")},
                FontSize = new FontSize {Val = DoubleValue.FromDouble(18)},
                Bold = new Bold()
            });

            // font 3
            fts.AppendChild(new Font
            {
                FontName = new FontName {Val = StringValue.FromString("Arial")},
                FontSize = new FontSize {Val = DoubleValue.FromDouble(14)}
            });

            // font 4
            fts.AppendChild(new Font
            {
                Color = new DocumentFormat.OpenXml.Spreadsheet.Color {Rgb = HexBinaryValueFromColor(Color.MediumBlue)},
                FontName = new FontName {Val = StringValue.FromString("Arial")},
                FontSize = new FontSize {Val = DoubleValue.FromDouble(11)}
            });

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