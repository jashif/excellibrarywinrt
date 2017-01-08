// ***********************************************************************
// Assembly         : ExcelManager
// Author           : Dheeraj
// ***********************************************************************
// <copyright file="Cell.cs" company="">
//     Copyright (c) . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
using ExcelManager.Entity;
using System;
namespace ExcelManager
{
    /// <summary>
    /// The Basic Cell class in a worksheet
    /// </summary>
    public class Cell
    {
        /// <summary>
        /// Gets or Sets Cell value Property
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        /// Gets or Sets Cell Style Property
        /// </summary>
        public CellStyle Style { get; set; }

        /// <summary>
        /// Gets or Sets Cell Address Property
        /// </summary>
        public CellAddress Address { get; set; }

        /// <summary>
        /// Gets or Sets Merge Cells Property
        /// </summary>
        public CellAddress MergeCell { get; set; }

        /// <summary>
        /// Gets or Sets the Cell Data Fomat Property
        /// </summary>
        public CellDataFormat Format { get; set; }

        /// <summary>
        /// Build XML component from cell class
        /// </summary>
        /// <param name="noOfColumns"></param>
        /// <returns></returns>
        internal c ToXml(int cellIndex)
        {
            string type = "s";
            switch (Format)
            {
                case CellDataFormat.SharedString:
                    type = "s";
                    break;
                case CellDataFormat.Boolean:
                    type = "b";
                    break;
                case CellDataFormat.Date:
                    type = "d";
                    break;
                case CellDataFormat.InlineString:
                    type = "inlineStr";
                    break;
                case CellDataFormat.String:
                    type = "str";
                    break;
                case CellDataFormat.Number:
                    type = "n";
                    break;
                default:
                    type = IsNumber(Value) ? "n" : "s";
                    break;
            }

            return new c() { r = ToAddress(), t = type, v = (type == "n" ? Convert.ToString(Value) : cellIndex.ToString()) };
        }

        /// <summary>
        /// Conevrt a Cell to its Address
        /// </summary>
        /// <returns></returns>
        internal string ToAddress()
        {
            return Address.Column.ColumnAddress + Address.Row.ToString();
        }

        private static bool IsNumber(object value)
        {
            return value is sbyte
                    || value is byte
                    || value is short
                    || value is ushort
                    || value is int
                    || value is uint
                    || value is long
                    || value is ulong
                    || value is float
                    || value is double
                    || value is decimal;
        }

    }

    public enum CellDataFormat
    {
        None,
        SharedString,
        Boolean,
        Date,
        InlineString,
        String,
        Number
    }
}
