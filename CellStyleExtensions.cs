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
using Windows.Data.Xml.Dom;
using Windows.UI;

namespace ExcelManager
{
    public static class CellStyleExtensions
    {
        /// <summary>
        /// Generate the XML element for Font in Styles
        /// </summary>
        /// <param name="cellstyle"></param>
        /// <param name="xmldocument"></param>
        /// <returns></returns>
        public static XmlElement ToFontElement(this CellStyle cellstyle, XmlDocument xmldocument)
        {
            XmlElement fontelement = xmldocument.CreateElement("font");

            if (cellstyle.IsBold)
            {
                XmlElement bold = xmldocument.CreateElement("b");
                fontelement.AppendChild(bold);
            }
            if (cellstyle.IsItalic)
            {
                XmlElement itallic = xmldocument.CreateElement("i");
                fontelement.AppendChild(itallic);
            }

            if (0 != cellstyle.FontSize)
            {
                XmlElement fontsize = xmldocument.CreateElement("sz");
                fontsize.SetAttribute("val", cellstyle.FontSize.ToString());
                fontelement.AppendChild(fontsize);
            }

            if (null != cellstyle.ForeColor)
            {
                string hex = "FF" + cellstyle.ForeColor.R.ToString("X2") + cellstyle.ForeColor.G.ToString("X2") + cellstyle.ForeColor.B.ToString("X2");
                XmlElement forecolor = xmldocument.CreateElement("color");
                forecolor.SetAttribute("rgb", hex);
                fontelement.AppendChild(forecolor);
            }

            if (null != cellstyle.Font)
            {
                XmlElement fontname = xmldocument.CreateElement("name");
                fontname.SetAttribute("val", cellstyle.Font.Source);
                fontelement.AppendChild(fontname);

                XmlElement fontfamily = xmldocument.CreateElement("family");
                fontfamily.SetAttribute("val", "2");
                fontelement.AppendChild(fontfamily);

                XmlElement fontscheme = xmldocument.CreateElement("scheme");
                fontscheme.SetAttribute("val", "minor");
                fontelement.AppendChild(fontscheme);
            }
            fontelement.RemoveAttribute("xmlns");
            return fontelement;
        }

        /// <summary>
        /// Generate the XML element for Fill in Styles
        /// </summary>
        /// <param name="cellstyle"></param>
        /// <param name="xmldocument"></param>
        /// <returns></returns>
        public static XmlElement ToFillElement(this CellStyle cellstyle, XmlDocument xmldocument)
        {
            XmlElement fillelement = null;
            if (Colors.White != cellstyle.FillColor)
            {
                fillelement = xmldocument.CreateElement("fill");
                string hex = "FF" + cellstyle.FillColor.R.ToString("X2") + cellstyle.FillColor.G.ToString("X2") + cellstyle.FillColor.B.ToString("X2");
                XmlElement patternfill = xmldocument.CreateElement("patternFill");
                patternfill.SetAttribute("patternType", "solid");

                XmlElement fgcolor = xmldocument.CreateElement("fgColor");
                fgcolor.SetAttribute("rgb", hex);
                patternfill.AppendChild(fgcolor);

                XmlElement bgcolor = xmldocument.CreateElement("bgColor");
                bgcolor.SetAttribute("indexed", "64");
                patternfill.AppendChild(bgcolor);

                fillelement.AppendChild(patternfill);
            }

            return fillelement;
        }

        /// <summary>
        /// Generate the XML element for Border in Styles
        /// </summary>
        /// <param name="cellstyle"></param>
        /// <param name="xmldocument"></param>
        /// <returns></returns>
        public static XmlElement ToBorderElement(this CellStyle cellstyle, XmlDocument xmldocument)
        {
            XmlElement borderelement = null;
            if (Colors.Transparent != cellstyle.BorderColor)
            {
                borderelement = xmldocument.CreateElement("border");
                string hex = "FF" + cellstyle.BorderColor.R.ToString("X2") + cellstyle.BorderColor.G.ToString("X2") + cellstyle.BorderColor.B.ToString("X2");
                XmlElement left = xmldocument.CreateElement("left");
                left.SetAttribute("style", "thin");

                XmlElement right = xmldocument.CreateElement("right");
                right.SetAttribute("style", "thin");

                XmlElement top = xmldocument.CreateElement("top");
                top.SetAttribute("style", "thin");

                XmlElement bottom = xmldocument.CreateElement("bottom");
                bottom.SetAttribute("style", "thin");

                XmlElement diagonal = xmldocument.CreateElement("diagonal");

                XmlElement colorl = xmldocument.CreateElement("color");
                colorl.SetAttribute("rgb", hex);
                XmlElement colorr = xmldocument.CreateElement("color");
                colorr.SetAttribute("rgb", hex);
                XmlElement colort = xmldocument.CreateElement("color");
                colort.SetAttribute("rgb", hex);
                XmlElement colorb = xmldocument.CreateElement("color");
                colorb.SetAttribute("rgb", hex);

                left.AppendChild(colorl);
                right.AppendChild(colorr);
                top.AppendChild(colort);
                bottom.AppendChild(colorb);

                borderelement.AppendChild(left);
                borderelement.AppendChild(right);
                borderelement.AppendChild(top);
                borderelement.AppendChild(bottom);
                borderelement.AppendChild(diagonal);
            }

            return borderelement;

        }

        /// <summary>
        /// Generate the XML element for CellStyle in Styles
        /// </summary>
        /// <param name="cellstyle"></param>
        /// <param name="xmldocument"></param>
        /// <param name="fontid"></param>
        /// <param name="fillid"></param>
        /// <param name="borderid"></param>
        /// <returns></returns>
        public static XmlElement ToXFElement(this CellStyle cellstyle, XmlDocument xmldocument, int fontid, int fillid, int borderid)
        {
            XmlElement xfelement = xmldocument.CreateElement("xf");
            if (fillid == 1) fillid = 0;
            xfelement.SetAttribute("numFmtId", "0");
            xfelement.SetAttribute("borderId", "0");
            xfelement.SetAttribute("xfId", "0");
            xfelement.SetAttribute("fontId", fontid.ToString());
            if (Colors.White != cellstyle.FillColor) xfelement.SetAttribute("fillId", fillid.ToString());
            if (Colors.Transparent != cellstyle.BorderColor) xfelement.SetAttribute("borderId", borderid.ToString());
            xfelement.SetAttribute("applyFont", "1");
            if (Colors.White != cellstyle.FillColor) xfelement.SetAttribute("applyFill", "1");
            if (Colors.Transparent != cellstyle.BorderColor) xfelement.SetAttribute("applyBorder", "1");

            if (null != cellstyle.Allignment)
            {
                xfelement.SetAttribute("applyAlignment", "1");
                XmlElement alignment = xmldocument.CreateElement("alignment");

                if (cellstyle.Allignment.Horizontal != ContentAllignment.Horizontal.None)
                {
                    var halgn = cellstyle.Allignment.Horizontal.ToString().ToLower();
                    alignment.SetAttribute("horizontal", halgn);
                }

                if (cellstyle.Allignment.Vertical != ContentAllignment.Vertical.None)
                {
                    var valgn = cellstyle.Allignment.Vertical.ToString().ToLower();
                    alignment.SetAttribute("vertical", valgn);
                }

                if (cellstyle.Allignment.WrapText)
                {
                    alignment.SetAttribute("wrapText", "1");
                }

                xfelement.AppendChild(alignment);
            }


            return xfelement;
        }
    }
}
