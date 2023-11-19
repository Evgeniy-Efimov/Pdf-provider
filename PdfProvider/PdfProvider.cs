using PdfProvider.Helpers;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PdfProvider
{
    public class PdfProvider
    {
        //page settings
        private XSize pageSize = PageSizeConverter.ToSize(PageSize.A4);
        private PageOrientation pageOrientation = PageOrientation.Portrait;
        private Margin pageContentMargin = new Margin()
        {
            MarginTop = 15,
            MarginRight = 35,
            MarginLeft = 35,
            MarginBottom = 50
        };

        //fonts settings
        private const string baseFontName = "Arial";
        private const int mainFontSize = 11;
        private const int headerFontSize = 14;
        private XFont mainFont = new XFont(baseFontName, mainFontSize, XFontStyle.Regular);
        private XBrush mainFontBrush = new XSolidBrush(XColor.FromArgb(68, 68, 68));
        private XFont headerFont = new XFont(baseFontName, headerFontSize, XFontStyle.Regular);
        private XBrush headerFontBrush = new XSolidBrush(XColor.FromArgb(0, 0, 0));
        private XFont commentFont = new XFont(baseFontName, mainFontSize, XFontStyle.Regular);
        private XBrush commentFontBrush = new XSolidBrush(XColor.FromArgb(68, 68, 68));
        private XPen tableBorderPen = new XPen(XColor.FromArgb(0, 0, 0), 0.5f);

        private FontSettings GetFontSettings(CustomFonts font, Aligment aligment)
        {
            var format = new XStringFormat();
            switch (aligment)
            {
                case Aligment.Left:
                    format.Alignment = XStringAlignment.Near;
                    format.LineAlignment = XLineAlignment.Near;
                    break;
                case Aligment.Right:
                    format.Alignment = XStringAlignment.Far;
                    format.LineAlignment = XLineAlignment.Far;
                    break;
                case Aligment.Center:
                    format.Alignment = XStringAlignment.Center;
                    format.LineAlignment = XLineAlignment.Center;
                    break;
            }
            switch (font)
            {
                case CustomFonts.Main: return new FontSettings() { XFont = mainFont, XBrush = mainFontBrush, XStringFormat = format };
                case CustomFonts.Header: return new FontSettings() { XFont = headerFont, XBrush = headerFontBrush, XStringFormat = format };
                case CustomFonts.Comment: return new FontSettings() { XFont = commentFont, XBrush = commentFontBrush, XStringFormat = format };
            }
            throw new Exception("Unknown font, can't get parameters");
        }

        //paragraphs settings
        private int lineVerticalOffset = 5;
        private int GetLineHeight(XFont xFont)
        {
            return (int)xFont.GetHeight();
        }
        private int GetPageLineWidth()
        {
            return (int)CurrentPdfPage.Width - pageContentMargin.MarginLeft - pageContentMargin.MarginRight;
        }
        private int GetBlockXPosition(int blockWidth, Aligment aligment)
        {
            int positionX = 0;
            switch (aligment)
            {
                case Aligment.Left:
                    positionX = pageContentMargin.MarginLeft;
                    break;
                case Aligment.Right:
                    positionX = pageContentMargin.MarginLeft + GetPageLineWidth() - blockWidth;
                    break;
                case Aligment.Center:
                    positionX = (GetPageLineWidth() - blockWidth) / 2;
                    break;
            }
            return positionX;
        }
        private void CheckNewPageNeeded(int blockHeight)
        {
            if (CurrentPageHeight + blockHeight + lineVerticalOffset > CurrentPdfPage.Height - pageContentMargin.MarginBottom)
            {
                AddPageToDocument();
            }
        }

        private PdfDocument pdfDocument;
        private PdfPage CurrentPdfPage { get { return pdfDocument.Pages[CurrentPageNumber]; } }
        private int CurrentPageNumber = 0;
        private int CurrentPageHeight = 0;

        public PdfProvider(string title)
        {
            pdfDocument = new PdfDocument();
            pdfDocument.Info.Title = title;
            AddPageToDocument();
        }

        public void DrawTableRow(string colContent, Margin colTextMargin, CustomFonts font, Aligment textAligment)
        {
            DrawTableRow(1, new List<int>() { 100 }, new List<string>() { colContent },
                colTextMargin, font, textAligment);
        }

        public void DrawTableRow(int columns, List<int> colWidthPercent, List<string> colContent,
            Margin colTextMargin, CustomFonts font, Aligment textAligment)
        {
            var maxWidth = GetPageLineWidth();
            if (colWidthPercent.Sum() != 100) throw new Exception("Invalid table parameters: column width");
            if (colWidthPercent.Count != colContent.Count ||
                colContent.Count != columns) throw new Exception("Invalid table parameters: column count");

            var fontSettings = GetFontSettings(font, textAligment);
            var textBlocks = new List<TextBlock>();
            var colPositionX = new List<int>();
            var colWidth = new List<int>();
            int rowHeight = 0;
            for (int i = 0; i < columns; i++)
            {
                if (i + 1 == columns && columns > 1)
                {
                    colWidth.Add(maxWidth - colWidth.Sum());
                }
                else
                {
                    colWidth.Add((int)(maxWidth * ((double)colWidthPercent[i] / 100f)));
                }
                if (i == 0)
                {
                    colPositionX.Add(pageContentMargin.MarginLeft);
                }
                else
                {
                    colPositionX.Add(colPositionX[i - 1] + colWidth[i - 1]);
                }
                textBlocks.Add(new TextBlock(colWidth[i], colPositionX[i], colContent[i], colTextMargin, fontSettings, CurrentPdfPage));
                if (textBlocks[i].BorderHeight > rowHeight) rowHeight = textBlocks[i].BorderHeight;
            }
            foreach (var textBlock in textBlocks)
            {
                if (textBlock.BorderHeight < rowHeight) textBlock.BorderHeight = rowHeight;
                DrawTextBlock(textBlock, addOffsetAfterBlock: false, absolutePosition: true);
            }
            CurrentPageHeight += rowHeight;
        }

        public void DrawTextBlock(TextBlock textBlock,
            bool addOffsetAfterBlock = true, bool absolutePosition = false)
        {
            CheckNewPageNeeded(textBlock.BorderHeight);

            using (XGraphics graphics = XGraphics.FromPdfPage(CurrentPdfPage))
            {
                var textFormatter = new XTextFormatter(graphics);
                var textBorderRect = new XRect(textBlock.BorderX, CurrentPageHeight, textBlock.BorderWidth, textBlock.BorderHeight);
                var textContentRect = new XRect(textBlock.TextBlockX, textBlock.GetTextBlockY(CurrentPageHeight), textBlock.TextBlockWidth, textBlock.TextBlockHeight);
                graphics.DrawRectangle(tableBorderPen, textBorderRect);
                textFormatter.DrawString(textBlock.Text, textBlock.FontSettings.XFont, textBlock.FontSettings.XBrush, textContentRect,
                    //fontSettings.XStringFormat - Exception: Only TopLeft alignment is currently implemented.
                    XStringFormat.TopLeft);
            }
            if (!absolutePosition)
            {
                CurrentPageHeight += textBlock.BorderHeight + (addOffsetAfterBlock ? lineVerticalOffset : 0);
            }
        }

        public void DrawTextBlock(string text, int blockWidth, CustomFonts font,
            int blockPositionX, Aligment textAligment, Margin textMargin,
            bool addOffsetAfterBlock = true, bool absolutePosition = false)
        {
            var fontSettings = GetFontSettings(font, textAligment);
            var positionX = blockPositionX;
            var textBlock = new TextBlock(blockWidth, positionX, text, textMargin, fontSettings, CurrentPdfPage);
            DrawTextBlock(textBlock, addOffsetAfterBlock, absolutePosition);
        }

        public void DrawTextBlock(string text, int blockWidth, CustomFonts font,
            Aligment blockAligment, Aligment textAligment, Margin textMargin,
            bool addOffsetAfterBlock = true, bool absolutePosition = false)
        {
            var positionX = GetBlockXPosition(blockWidth, blockAligment);
            DrawTextBlock(text, blockWidth, font,
                positionX, textAligment, textMargin,
                addOffsetAfterBlock, absolutePosition);
        }

        public void WriteTextLine(string text, CustomFonts font, Aligment aligment)
        {
            var fontSettings = GetFontSettings(font, aligment);
            var lineHeight = GetLineHeight(fontSettings.XFont);
            var lineWidth = GetPageLineWidth();

            CheckNewPageNeeded(lineHeight);

            using (XGraphics graphics = XGraphics.FromPdfPage(CurrentPdfPage))
            {
                graphics.DrawString(text, fontSettings.XFont, fontSettings.XBrush,
                    new XRect(pageContentMargin.MarginLeft, CurrentPageHeight, lineWidth, lineHeight), fontSettings.XStringFormat);
            }
            CurrentPageHeight += lineHeight + lineVerticalOffset;
        }

        public void DrawImage(byte[] imageData, Aligment aligment = Aligment.Left, bool absolutePosition = false, Tuple<int, int> size = null)
        {
            using (MemoryStream stream = new MemoryStream(imageData))
            {
                var image = XImage.FromStream(stream);
                int positionY = (int)CurrentPageHeight;
                int imageWidth = image.PixelWidth;
                int imageHeight = image.PixelHeight;
                if (size != null)
                {
                    imageWidth = size.Item1;
                    imageHeight = size.Item2;
                }
                int positionX = GetBlockXPosition(imageWidth, aligment);

                CheckNewPageNeeded(imageHeight);

                using (XGraphics graphics = XGraphics.FromPdfPage(CurrentPdfPage))
                {
                    graphics.DrawImage(image, positionX, positionY, imageWidth, imageHeight);
                }
                if (!absolutePosition) CurrentPageHeight += imageHeight + lineVerticalOffset;
            }
        }

        private PdfPage AddPageToDocument()
        {
            var page = pdfDocument.AddPage();
            page.Orientation = pageOrientation;
            page.Width = pageSize.Width;
            page.Height = pageSize.Height;
            CurrentPageNumber = pdfDocument.Pages.Count - 1;
            CurrentPageHeight = pageContentMargin.MarginTop;
            return page;
        }

        public byte[] GetDocumentData()
        {
            byte[] result = null;
            using (MemoryStream stream = new MemoryStream())
            {
                pdfDocument.Save(stream, true);
                result = stream.ToArray();
            }
            return result;
        }
    }

    public enum CustomFonts
    {
        Main = 1,
        Header = 2,
        Comment = 3
    }

    public enum Aligment
    {
        Left = 1,
        Right = 2,
        Center = 3
    }

    public class FontSettings
    {
        public XFont XFont { get; set; }
        public XBrush XBrush { get; set; }
        public XStringFormat XStringFormat { get; set; }
    }

    public class Margin
    {
        public int MarginTop { get; set; }
        public int MarginRight { get; set; }
        public int MarginLeft { get; set; }
        public int MarginBottom { get; set; }
    }

    public class TextBlock
    {
        public Margin TextMargin { get; set; }
        public string Text { get; set; }
        public FontSettings FontSettings { get; set; }
        public int BorderX { get; set; }
        public int BorderWidth { get; set; }
        public int BorderHeight { get; set; }
        public int TextBlockX { get; set; }
        public int GetTextBlockY(int currentPageHeight)
        {
            return currentPageHeight + TextMargin.MarginTop;
        }
        public int TextBlockWidth { get; set; }
        public int TextBlockHeight { get; set; }
        public static int GetEstimatedBlockHeight(int blockWidth, string text, XFont xFont, PdfPage page)
        {
            if (string.IsNullOrEmpty(text)) text = " ";
            double estimatedBlockHeight = 0;
            using (XGraphics graphics = XGraphics.FromPdfPage(page))
            {
                var textFormatterExtended = new XTextFormatterExtended(graphics);
                int lastCharIndex;
                var rect = new XRect(0, 0, blockWidth, double.MaxValue);
                textFormatterExtended.PrepareDrawString(text, xFont, rect, out lastCharIndex, out estimatedBlockHeight);
            }
            return (int)estimatedBlockHeight + 3; //add some height to compensate error
        }
        public TextBlock(int blockWidth, int blockX, string text,
            Margin textMargin, FontSettings fontSettings, PdfPage pdfPage)
        {
            if (string.IsNullOrEmpty(text)) text = " ";
            TextMargin = textMargin;
            Text = text;
            FontSettings = fontSettings;
            BorderWidth = blockWidth;
            BorderX = blockX;
            BorderWidth = blockWidth;
            TextBlockX = blockX + textMargin.MarginLeft;
            TextBlockWidth = blockWidth - textMargin.MarginLeft - textMargin.MarginRight;
            TextBlockHeight = GetEstimatedBlockHeight(TextBlockWidth, text, fontSettings.XFont, pdfPage);
            BorderHeight = TextBlockHeight + textMargin.MarginTop + textMargin.MarginBottom;
        }
    }
}