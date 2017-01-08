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
    /// Merge Cell class  
    /// </summary>
    public class MergeCell
    {
        /// <summary>
        /// Gets or Sets Start Address Property
        /// </summary>
        public CellAddress StartAddress { get; set; }

        /// <summary>
        /// Gets or Sets End Address Property
        /// </summary>
        public CellAddress EndAddress { get; set; }
    }
}
