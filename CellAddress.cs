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
    /// Cell Address Class
    /// </summary>
    public class CellAddress
    {
        /// <summary>
        /// Gets or Sets Column Property
        /// </summary>
        public Column Column { get; set; }

        /// <summary>
        /// Gets or Sets Row Property
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public CellAddress()
        {

        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="column"></param>
        /// <param name="row"></param>
        public CellAddress(Column column, int row)
        {
            Column = column;
            Row = row;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="columnaddress"></param>
        /// <param name="row"></param>
        public CellAddress(string columnaddress, int row)
        {
            Column = new Column(columnaddress);
            Row = row;
        }

    }
    /// <summary>
    /// Column Class
    /// </summary>
    public struct Column
    {
        /// <summary>
        /// Gets or Sets Column Address Property
        /// </summary>
        public string ColumnAddress;

        /// <summary>
        /// Gets or Sets Index Property
        /// </summary>
        public int Index;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="columnAddress"></param>
        public Column(string columnAddress)
        {
            this.Index = 1;
            this.ColumnAddress = columnAddress.ToUpper();
            this.Index = GetIndexFromColumnName(this.ColumnAddress);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="indexcolumnAddress"></param>
        public Column(int indexcolumnAddress)
        {
            this.ColumnAddress = string.Empty;
            this.Index = indexcolumnAddress;
            this.ColumnAddress = GetColumnNameFromIndex(indexcolumnAddress);

        }

        /// <summary>
        /// Returns Column Name/Adderss in uppercase
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return ColumnAddress.ToUpper();
        }

        /// <summary>
        /// Returns Column Index from Column Name/Address
        /// </summary>
        /// <param name="n"></param>
        /// <returns></returns>
        private string GetColumnNameFromIndex(int n)
        {
            string name = "";
            while (n > 0)
            {
                n--;
                name = (char)('A' + n % 26) + name;
                n /= 26;
            }
            return name;
        }

        /// <summary>
        /// Returns Column Name/Address from Column Index
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        private int GetIndexFromColumnName(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }

    }
}
