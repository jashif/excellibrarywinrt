// ***********************************************************************
// Assembly         : ExcelManager
// Author           : Dheeraj
// ***********************************************************************
// <copyright file="Cell.cs" company="">
//     Copyright (c) . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelManager
{
    /// <summary>
    /// Sheet Class Contains Collection of Cells
    /// </summary>
    public class Sheet
    {
        /// <summary>
        /// Gets or Sets Cells Collection Property
        /// </summary>
        internal List<Cell> Cells { get; set; }

        /// <summary>
        /// Gets or Sets Row Style Collecton Property
        /// </summary>
        internal List<RowStyle> RowStyles { get; set; }

        /// <summary>
        /// Gets or Sets Column style Collection Property
        /// </summary>
        internal List<ColumnStyle> ColumnStyles { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        internal Sheet()
        {
            Cells = new List<Cell>();
            RowStyles = new List<RowStyle>();
            ColumnStyles = new List<ColumnStyle>();
        }

        /// <summary>
        /// Add each row style to row  style collection
        /// </summary>
        /// <param name="rowstyle"></param>
        public void AddRowStyle(RowStyle rowstyle)
        {
            RowStyles.Add(rowstyle);
        }

        /// <summary>
        /// Add each column style to column style collection
        /// </summary>
        /// <param name="columnstyle"></param>
        public void AddColumnStyle(ColumnStyle columnstyle)
        {
            ColumnStyles.Add(columnstyle);
        }

        /// <summary>
        /// Add Cells To Cell collection
        /// </summary>
        /// <param name="cell"></param>
        public void AddCell(Cell cell)
        {
            if (null != cell)
            {
                if (null != cell.Style)
                {
                    var style = cell.Style;
                    style = CheckDuplicateStyleEntries(style);
                }
                Cells.Add(cell);
            }
        }

        /// <summary>
        /// Add Cell Value To Cell collection
        /// </summary>
        /// <param name="value"></param>
        /// <param name="address"></param>
        public void AddCellValue(object value, CellAddress address)
        {
            Cells.Add(new Cell() { Value = value, Address = address });
        }

        /// <summary>
        /// Add Cell Value To Cell collection
        /// </summary>
        /// <param name="value"></param>
        /// <param name="address"></param>
        /// <param name="style"></param>
        public void AddCellValue(object value, CellAddress address, CellStyle style)
        {
            if (null != style)
            {
                style = CheckDuplicateStyleEntries(style);
                Cells.Add(new Cell() { Value = value, Address = address, Style = style });

            }
        }

        /// <summary>
        /// Add Cell Value To Cell collection
        /// </summary>
        /// <param name="value"></param>
        /// <param name="address"></param>
        /// <param name="mergcell"></param>
        public void AddCellValue(object value, CellAddress address, CellAddress mergcell)
        {
            if (null != mergcell)
            {
                Cells.Add(new Cell() { Value = value, Address = address, MergeCell = mergcell });
            }
        }

        /// <summary>
        /// Add Cell Value To Cell collection
        /// </summary>
        /// <param name="value"></param>
        /// <param name="address"></param>
        /// <param name="style"></param>
        /// <param name="mergcell"></param>
        public void AddCellValue(object value, CellAddress address, CellStyle style, CellAddress mergcell)
        {
            if (null != style)
            {
                style = CheckDuplicateStyleEntries(style);
                Cells.Add(new Cell() { Value = value, Address = address, Style = style, MergeCell = mergcell });
            }
        }


        #region Private Variables
        private int styleIndex = 0;
        #endregion

        #region Private Method
        /// <summary>
        /// Check the style entries in cells and assign Style Id to each.
        /// </summary>
        /// <param name="style"></param>
        private CellStyle CheckDuplicateStyleEntries(CellStyle style)
        {
            var previousCell = Cells.FirstOrDefault(cell => null != cell.Style &&
                cell.Style.IsBold == style.IsBold &&
    cell.Style.IsItalic == style.IsItalic &&
   cell.Style.FontSize == style.FontSize &&
    cell.Style.ForeColor == style.ForeColor &&
     cell.Style.FillColor == style.FillColor &&
     cell.Style.Font.Source == style.Font.Source &&
     ((null == cell.Style.Allignment && null == style.Allignment) ||
     (null != cell.Style.Allignment && null != style.Allignment &&
       cell.Style.Allignment.Horizontal == style.Allignment.Horizontal &&
       cell.Style.Allignment.Vertical == style.Allignment.Vertical &&
       cell.Style.Allignment.WrapText == style.Allignment.WrapText)));

            if (null != previousCell && null != previousCell.Style)
            {
                style.StyleId = previousCell.Style.StyleId;
            }

            else
            {
                styleIndex++;
                style.StyleId = styleIndex;

            }
            return style;

        }
        #endregion


    }
}
