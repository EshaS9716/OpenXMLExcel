using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelAppOpenXML
{
    public static class StyleSheet6
    {
        public static void AddBold(SpreadsheetDocument document, Cell c, int column, bool buChanged)
        {
            Fonts fs = AddFont(document.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts, column);
            Borders bs = AddBorders(document.WorkbookPart.WorkbookStylesPart.Stylesheet.Borders, buChanged);
            AddCellFormat(document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats, document.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts, document.WorkbookPart.WorkbookStylesPart.Stylesheet.Borders);
            c.StyleIndex = (UInt32)(document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Elements<CellFormat>().Count() - 1);
        }

        static Fonts AddFont(Fonts fs, int col)
        {
            Font font2 = new Font();
            Bold bold1 = new Bold();
            Italic italic = new Italic();
            Underline underline = new Underline();
            FontSize fontSize2 = new FontSize();
            Color color = new Color();

            if (col == 1)
            {
                fontSize2.Val = 12.5D;
                font2.Append(bold1);
                font2.Append(italic);
            }
            else if (col == 2)
            {
                color.Rgb = "862d2d";
                fontSize2.Val = 11D;
                font2.Append(bold1);
                font2.Append(italic);
            }
            else if (col == 3)
            {
                color.Rgb = "002b80";
                fontSize2.Val = 11D;
                font2.Append(bold1);
            }
            else if (col == 4)
            {
                color.Rgb = "002b80";
                fontSize2.Val = 11D;
                font2.Append(underline);
                font2.Append(italic);
            }
            else
            {
                color.Rgb = "003366";
                fontSize2.Val = 10D;
                font2.Append(italic);
            }
            font2.Append(fontSize2);
            font2.Append(color);

            fs.Append(font2);
            return fs;
        }

        static Borders AddBorders(Borders borders, bool buChanged)
        {
            Border border = new Border();

            if (buChanged)
            {
                BottomBorder bottomBorder = new BottomBorder() { Style = BorderStyleValues.Thick };
                Color bottomColor = new Color() { Indexed = (UInt32Value)64U };
                bottomBorder.Append(bottomColor);
                border.Append(bottomBorder);
            }

            RightBorder border1 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color1 = new Color() { Indexed = (UInt32Value)64U };
            border1.Append(color1);

            border.Append(border1);

            borders.Append(border);
            return borders;
        }

        static void AddCellFormat(CellFormats cf, Fonts fs, Borders bs)
        {
            CellFormat cellFormat2 = new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) { NumberFormatId = 0, FontId = (UInt32)(fs.Elements<Font>().Count() - 1), BorderId = (UInt32)(bs.Elements<Border>().Count() - 1), FormatId = 0, ApplyBorder = true, ApplyFont = true, ApplyAlignment = true };
            cf.Append(cellFormat2);
        }
    }

    public static class StylesSheet5
    {
        public static void AddBold(SpreadsheetDocument document, Cell c, int column, bool buChanged)
        {
            if (IsBkgColor(column))
            {
                Fonts fs = AddFont(document.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts, column);
                Borders bs = AddBorders(document.WorkbookPart.WorkbookStylesPart.Stylesheet.Borders, buChanged);
                Fills fi = AddFills(document.WorkbookPart.WorkbookStylesPart.Stylesheet.Fills, column);
                AddCellFormat(document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats, document.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts, document.WorkbookPart.WorkbookStylesPart.Stylesheet.Borders, document.WorkbookPart.WorkbookStylesPart.Stylesheet.Fills);
                c.StyleIndex = (UInt32)(document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Elements<CellFormat>().Count() - 1);
            }
            else
            {
                Fonts fs = AddFont(document.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts, column);
                Borders bs = AddBorders(document.WorkbookPart.WorkbookStylesPart.Stylesheet.Borders, buChanged);
                AddCellFormat(document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats, document.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts, document.WorkbookPart.WorkbookStylesPart.Stylesheet.Borders);
                c.StyleIndex = (UInt32)(document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Elements<CellFormat>().Count() - 1);
            }
        }

        static bool IsBkgColor(int col)
        {
            switch (col)
            {
                case 5: case 6: case 8: case 9: case 10: case 11: case 12: case 13: case 15: case 16: case 17: case 18: return true;
                default:
                    return false;
            }
        }

        static Fonts AddFont(Fonts fs, int col)
        {
            Font font2 = new Font();
            Bold bold1 = new Bold();
            Italic italic = new Italic();
            Underline underline = new Underline();
            FontSize fontSize2 = new FontSize();
            Color color = new Color();

            if (col == 1)
            {
                color.Rgb = "002b80";
                fontSize2.Val = 14D;
                font2.Append(bold1);
                font2.Append(italic);
            }
            else if (col == 2)
            {
                fontSize2.Val = 13D;
                font2.Append(italic);
            }
            else if (col == 3)
            {
                color.Rgb = "862d2d";
                fontSize2.Val = 11D;
                font2.Append(bold1);
                font2.Append(italic);
            }
            else if (col == 4)
            {
                color.Rgb = "003366";
                fontSize2.Val = 11D;
                font2.Append(underline);
                font2.Append(italic);
            }
            else if (IsBkgColor(col))
            {
                color.Rgb = "003366";
                fontSize2.Val = 11D;
                font2.Append(bold1);
                font2.Append(italic);
            }
            else
            {
                color.Rgb = "003366";
                fontSize2.Val = 10D;
                font2.Append(italic);
            }
            font2.Append(fontSize2);
            font2.Append(color);

            fs.Append(font2);
            return fs;
        }

        static Fills AddFills(Fills fills, int col)
        {
            Fill fill1 = new Fill();
            PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor3 = new ForegroundColor();
            if (col > 4 && col < 14)
            {
                foregroundColor3 = new ForegroundColor() { Rgb = "e7e7e7" };
            }
            else if (col == 15 || col == 17)
            {
                foregroundColor3 = new ForegroundColor() { Rgb = "e1fccc" };
            }
            else if (col == 16 || col == 18)
            {
                foregroundColor3 = new ForegroundColor() { Rgb = "F2DCDB" };
            }
            patternFill5.Append(foregroundColor3);
            fill1.Append(patternFill5);

            fills.Append(fill1);
            return fills;
        }

        static Borders AddBorders(Borders borders, bool buChanged)
        {
            Border border = new Border();

            if (buChanged)
            {
                BottomBorder bottomBorder = new BottomBorder() { Style = BorderStyleValues.Thick };
                Color bottomColor = new Color() { Indexed = (UInt32Value)64U };
                bottomBorder.Append(bottomColor);
                border.Append(bottomBorder);
            }

            RightBorder border1 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color1 = new Color() { Indexed = (UInt32Value)64U };
            border1.Append(color1);

            border.Append(border1);

            borders.Append(border);
            return borders;
        }

        static void AddCellFormat(CellFormats cf, Fonts fs, Borders bs, Fills fills)
        {
            CellFormat cellFormat2 = new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) { NumberFormatId = 0, FontId = (UInt32)(fs.Elements<Font>().Count() - 1), BorderId = (UInt32)(bs.Elements<Border>().Count() - 1), FillId = (UInt32)(fills.Elements<Fill>().Count() - 1), FormatId = 0, ApplyBorder = true, ApplyFont = true, ApplyAlignment = true, ApplyFill = true};
            cf.Append(cellFormat2);
        }

        static void AddCellFormat(CellFormats cf, Fonts fs, Borders bs)
        {
            CellFormat cellFormat2 = new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) { NumberFormatId = 0, FontId = (UInt32)(fs.Elements<Font>().Count() - 1), BorderId = (UInt32)(bs.Elements<Border>().Count() - 1), FormatId = 0, ApplyBorder = true, ApplyFont = true, ApplyAlignment = true };
            cf.Append(cellFormat2);
        }
    }

    public static class StylesSheet2
    {
        public static void AddBold(SpreadsheetDocument document, Cell c, int row)
        {
            Fonts fs = AddFont(document.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts, row);
            Fills fi = AddFills(document.WorkbookPart.WorkbookStylesPart.Stylesheet.Fills, row);
            Borders bs = AddBorders(document.WorkbookPart.WorkbookStylesPart.Stylesheet.Borders);
            AddCellFormat(document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats, document.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts, document.WorkbookPart.WorkbookStylesPart.Stylesheet.Fills, document.WorkbookPart.WorkbookStylesPart.Stylesheet.Borders);
            c.StyleIndex = (UInt32)(document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Elements<CellFormat>().Count() - 1);
        }

        static Fonts AddFont(Fonts fs, int row)
        {
            Font font2 = new Font();
            Bold bold1 = new Bold();
            Italic italic = new Italic();
            FontSize fontSize2 = new FontSize();
            if (row == 4)
            {
                fontSize2.Val = 20D;
            }
            else if (row == 5)
            {
                fontSize2.Val = 16D;
            }
            else if (row == 6)
            {
                fontSize2.Val = 14D;
            }
            else if (row == 7)
            {
                fontSize2.Val = 12D;
            }
            if (row != 7)
            {
                font2.Append(bold1);
            }
            font2.Append(italic);
            font2.Append(fontSize2);

            fs.Append(font2);
            return fs;
        }

        static Fills AddFills(Fills fills, int row)
        {
            Fill fill1 = new Fill();
            PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor3 = new ForegroundColor() { Rgb = GetColor(row) };

            patternFill5.Append(foregroundColor3);
            fill1.Append(patternFill5);

            fills.Append(fill1);
            return fills;
        }

        static string GetColor(int row)
        {
            switch (row)
            {
                case 4: return "e1fccc";
                case 5: return "d4e3fa";
                case 6: return "faf6d4";
                case 7: return "e7e7e7";
                default: return "FFFFFF";
            }
        }

        static Borders AddBorders(Borders borders)
        {
            Border border = new Border();
            LeftBorder border1 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color1 = new Color() { Indexed = (UInt32Value)64U }; 
            border1.Append(color1);

            border.Append(border1);

            borders.Append(border);
            return borders;
        }

        static void AddCellFormat(CellFormats cf, Fonts fs, Fills fills, Borders bs)
        {
            CellFormat cellFormat2 = new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) { NumberFormatId = 0, FontId = (UInt32)(fs.Elements<Font>().Count() - 1), FillId = (UInt32)(fills.Elements<Fill>().Count() - 1), BorderId = (UInt32)(bs.Elements<Border>().Count() - 1), FormatId = 0, ApplyFill = true, ApplyBorder = true, ApplyFont = true, ApplyAlignment = true};
            cf.Append(cellFormat2);
        }
    }

    public static class StylesSheet1
    {
        public static void AddBold(SpreadsheetDocument document, Cell c, bool isLevel4, bool isLevel3, bool isLevel2, bool isLevel1)
        {
            Fonts fs = AddFont(document.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts, isLevel4, isLevel3, isLevel2, isLevel1);
            AddCellFormat(document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats, document.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts, isLevel1);
            c.StyleIndex = (UInt32)(document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Elements<CellFormat>().Count() - 1);
        }

        static Fonts AddFont(Fonts fs, bool isLevel4, bool isLevel3, bool isLevel2, bool isLevel1)
        {
            Font font2 = new Font();
            Bold bold1 = new Bold();
            Italic italic = new Italic();
            Underline underline = new Underline();

            Color color = new Color();
            FontSize fontSize2 = new FontSize();

            if (isLevel4)
            {
                color.Rgb = "003366";
                fontSize2.Val = 10D;
                font2.Append(bold1);
                font2.Append(underline);
            }
            else if (isLevel3)
            {
                color.Rgb = "862d2d";
                fontSize2.Val = 11D;
                font2.Append(underline);
            }
            else if (isLevel2)
            {
                fontSize2.Val = 14D;
            }
            else if (isLevel1)
            {
                color.Rgb = "002b80";
                fontSize2.Val = 18D;
                font2.Append(bold1);
            }
            else
            {
                color.Rgb = "003366";
                fontSize2.Val = 10D;
            }

            font2.Append(italic);
            font2.Append(color);
            font2.Append(fontSize2);

            fs.Append(font2);
            return fs;
        }

        static void AddCellFormat(CellFormats cf, Fonts fs, bool isLevel1)
        {
            CellFormat cellFormat2;
            if (isLevel1)
            {
                cellFormat2 = new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, TextRotation = (UInt32Value)90U }) { NumberFormatId = 0, FontId = (UInt32)(fs.Elements<Font>().Count() - 1), FillId = 0, BorderId = 0, FormatId = 0, ApplyFont = true, ApplyAlignment = true };
            }
            else
            {
                cellFormat2 = new CellFormat(new Alignment() { Vertical = VerticalAlignmentValues.Top }) { NumberFormatId = 0, FontId = (UInt32)(fs.Elements<Font>().Count() - 1), FillId = 0, BorderId = 0, FormatId = 0, ApplyFont = true, ApplyAlignment = true };
            }
            cf.Append(cellFormat2);
        }
    }
}