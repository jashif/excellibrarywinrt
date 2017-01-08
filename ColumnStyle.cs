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
using System.Threading.Tasks;

namespace ExcelManager
{
    /// <summary>
    /// Column Style Class includes Column Width
    /// </summary>
    public class ColumnStyle
    {
        /// <summary>
        /// Gets or Sets Column Index Property
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// Gets or Sets Column Width Property
        /// </summary>
        public double Width { get; set; }
    }
}
