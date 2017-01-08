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
    /// Row style Class Contains Row Height
    /// </summary>
    public class RowStyle
    {
        /// <summary>
        /// Gets or Sets Row Index Property
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// Gets or Sets Row Height Property
        /// </summary>
        public double Height { get; set; }
    }
}
