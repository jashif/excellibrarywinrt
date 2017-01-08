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
    public class ContentAllignment
    {
        /// <summary>
        /// Different Horizontal Content Allignment Types
        /// </summary>
        public enum Horizontal
        {
            Left,
            Center,
            Right,
            None
        }

        /// <summary>
        /// Different Vertical Content Allignment types
        /// </summary>
        public enum Vertical
        {
            Top,
            Center,
            Bottom,
            None
        }
    }
}
