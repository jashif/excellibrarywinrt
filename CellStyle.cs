// ***********************************************************************
// Assembly         : ExcelManager
// Author           : Dheeraj
// ***********************************************************************
// <copyright file="Cell.cs" company="">
//     Copyright (c) . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
using Windows.UI;
using Windows.UI.Xaml.Media;

namespace ExcelManager
{
    /// <summary>
    /// CellStyle Class for Font,Fill and Allignment
    /// </summary>
    public class CellStyle
    {
        /// <summary>
        /// Internally Gets/Sets Style Id
        /// </summary>
        internal int StyleId { get; set; }

        /// <summary>
        /// Gets or Sets Font Family Property
        /// </summary>
        public FontFamily Font { get; set; }

        /// <summary>
        /// Gets or Sets Font size Property
        /// </summary>
        public double FontSize { get; set; }

        /// <summary>
        /// Gets or Sets IsBold Property
        /// </summary>
        public bool IsBold { get; set; }

        /// <summary>
        /// Gets or Sets IsItalic Property
        /// </summary>
        public bool IsItalic { get; set; }

        /// <summary>
        /// Gets or Sets Fore Colour Property
        /// </summary>
        public Color ForeColor { get; set; }

        /// <summary>
        /// Gets or Sets Fill Colour Property
        /// </summary>
        public Color FillColor { get; set; }

        /// <summary>
        /// Gets or Sets Allignment Property
        /// </summary>
        public Alignment Allignment { get; set; }

        /// <summary>
        /// Gets or Sets Border Colour Property
        /// </summary>
        public Color BorderColor { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public CellStyle()
        {
            Font = new FontFamily("Calibri");
            ForeColor = Colors.Black;
            FillColor = Colors.White;
            BorderColor = Colors.Transparent;
            FontSize = 11;
        }

    }

    /// <summary>
    /// Cell Allignment includes Horizontal, Vertical alignment and Text Wrapping
    /// </summary>
    public class Alignment
    {
        /// <summary>
        /// Gets or Sets HorizontalContentAllignment Property
        /// </summary>
        public ContentAllignment.Horizontal Horizontal { get; set; }

        /// <summary>
        /// Gets or Sets VerticalContentAllignment Property
        /// </summary>
        public ContentAllignment.Vertical Vertical { get; set; }

        /// <summary>
        /// Gets or Sets WrapText Property
        /// </summary>
        public bool WrapText { get; set; }

        public Alignment()
        {
            Horizontal = ContentAllignment.Horizontal.None;
            Vertical = ContentAllignment.Vertical.None;
            WrapText = false;
        }
    }

}
