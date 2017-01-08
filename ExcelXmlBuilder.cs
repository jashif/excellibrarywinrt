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
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Windows.Data.Xml.Dom;

namespace ExcelManager
{
    internal class ExcelXmlBuilder
    {
        /// <summary>
        /// Build SharedString XML string
        /// </summary>
        /// <param name="cells"></param>
        /// <returns></returns>
        public string BuildSharedXML(List<Cell> cells)
        {
            try
            {
                //sort cells
                cells = cells.OrderBy(x => x.Address.Row).ThenBy(x => x.Address.Column.Index).ToList();

                sst sstRoot = new sst();
                sstRoot.count = cells.Count.ToString();
                sstRoot.uniqueCount = cells.Count.ToString();
                sstRoot.si = new List<si>();
                foreach (var item in cells)
                {
                    sstRoot.si.Add(new si() { t = Convert.ToString(item.Value) });
                }
                var dixt = new Dictionary<string, string>();
                string xmlsrz = Serialize(sstRoot, dixt);
                xmlsrz = xmlsrz.Replace("encoding=\"utf-16\"", "encoding=\"UTF-8\" standalone=\"yes\"");
                return xmlsrz;
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// Build Sheet XML string
        /// </summary>
        /// <param name="excelsheet"></param>
        /// <param name="xmldocument"></param>
        /// <returns></returns>
        public string BuildSheetXML(Sheet excelsheet, ref XmlDocument xmldocument, int sharedstringcount = 0)
        {
            try
            {
                var cells = excelsheet.Cells;
                var rowstyles = excelsheet.RowStyles;
                var columnstyles = excelsheet.ColumnStyles;

                var startRowIndex = cells.Min(x => x.Address.Row);
                var endRowIndex = cells.Max(x => x.Address.Row);
                var startColumnIndex = cells.Min(x => x.Address.Column.Index);
                var endColumnIndex = cells.Max(x => x.Address.Column.Index);
                var noofRows = endRowIndex - startRowIndex + 1;
                var noofCols = endColumnIndex - startRowIndex + 1;

                //sorting by cell index
                cells = cells.OrderBy(x => x.Address.Row).ThenBy(x => x.Address.Column.Index).ToList();

                worksheet worksheet = new worksheet();
                worksheet.sheetviews = new sheetViews();
                worksheet.sheetFormatPr = new sheetFormatPr() { defaultRowHeight = "15", dyDescent = ".25" };
                worksheet.sheetData = new sheetData();
                worksheet.sheetData.row = new List<row>();

                //add column Styles
                if (columnstyles.Count > 0)
                {
                    worksheet.cols = new cols();
                    worksheet.cols.col = new List<col>();
                    foreach (var columnstyle in columnstyles)
                    {
                        var col = new col();
                        col.customWidth = "1";
                        col.max = col.min = columnstyle.Index.ToString();
                        col.width = columnstyle.Width.ToString();
                        worksheet.cols.col.Add(col);
                    }
                }

                List<int> addedStyleIds = new List<int>();
                int fontId = 0; int fillId = 1; int borderId = 0; // initializing fontid,fillid and borderid
                int styles = 0;

                //Gets the Existing XML Nodes/Elements from Style Sheet
                var fonts = xmldocument.GetElementsByTagName("fonts").FirstOrDefault();
                var fills = xmldocument.GetElementsByTagName("fills").FirstOrDefault();
                var borders = xmldocument.GetElementsByTagName("borders").FirstOrDefault();
                var cellXfs = xmldocument.GetElementsByTagName("cellXfs").FirstOrDefault();

                var count = cells.Where(x => null != x.MergeCell).Count();
                if (count > 0)
                {
                    worksheet.mergeCells = new mergeCells() { count = count.ToString() };
                    worksheet.mergeCells.mergeCell = new List<mergeCell>();
                }

                for (var rowindx = startRowIndex; rowindx <= endRowIndex; rowindx++)
                {
                    row row = new row();
                    row.r = rowindx.ToString();
                    row.spans = startColumnIndex.ToString() + ":" + endColumnIndex.ToString();
                    row.dyDescent = "0.25";
                    row.c = new List<c>();

                    //var rstyle = rowstyles.Where(x => x.Index == rowindx).FirstOrDefault();
                    //if (null != rstyle)
                    //{
                    //    row.height = rstyle.Height.ToString();
                    //    row.customHeight = "1";
                    //}

                    for (var colindx = startColumnIndex; colindx <= endColumnIndex; colindx++)
                    {
                        int cellindx = ((rowindx - 1) * columnstyles.Count) + (colindx - 1);
                        var currentcell = cells[cellindx];//.FirstOrDefault(x => x.Address.Row == rowindx && x.Address.Column.Index == colindx);
                        if (null != currentcell)
                        {
                            var index = cellindx + sharedstringcount;// cells.IndexOf(currentcell) + sharedstringcount; //cells.IndexOf(currentcell)
                            var c = currentcell.ToXml(index);

                            if (null != currentcell.Style)
                            {
                                var isAdded = addedStyleIds.Any(x => x == currentcell.Style.StyleId);
                                //Add styles to stylesheet if not added already
                                if (!isAdded)
                                {
                                    addedStyleIds.Add(currentcell.Style.StyleId);
                                    var fntElem = currentcell.Style.ToFontElement(xmldocument);
                                    if (null != fntElem)
                                    {
                                        fonts.AppendChild(fntElem);
                                        fontId++;
                                    }

                                    var fillElem = currentcell.Style.ToFillElement(xmldocument);
                                    if (null != fillElem)
                                    {
                                        fills.AppendChild(fillElem);
                                        fillId++;
                                    }

                                    var borderElem = currentcell.Style.ToBorderElement(xmldocument);
                                    if (null != borderElem)
                                    {
                                        borders.AppendChild(borderElem);
                                        borderId++;
                                    }

                                    var xfElem = currentcell.Style.ToXFElement(xmldocument, fontId, fillId, borderId);
                                    if (null != xfElem)
                                    {
                                        cellXfs.AppendChild(xfElem);
                                        styles++;
                                    }
                                }

                                c.s = currentcell.Style.StyleId.ToString();
                            }


                            if (null != currentcell.MergeCell)
                            {
                                var startaddress = currentcell.ToAddress();
                                var endaddress = new Cell() { Address = currentcell.MergeCell }.ToAddress();
                                worksheet.mergeCells.mergeCell.Add(new mergeCell { reference = startaddress + ":" + endaddress });
                            }

                            row.c.Add(c);

                        }
                    }

                    fonts.Attributes[0].InnerText = (fontId + 1).ToString();
                    cellXfs.Attributes[0].InnerText = (styles + 1).ToString();
                    fills.Attributes[0].InnerText = (fillId + 1).ToString();
                    borders.Attributes[0].InnerText = (borderId + 1).ToString();
                    if (null != row && row.c.Count > 0) worksheet.sheetData.row.Add(row);
                }
                worksheet.pageMargins = new pageMargins();
                worksheet.dimension = new dimension();
                worksheet.dimension.reference = cells[0].ToAddress() + ":" + cells[cells.Count - 1].ToAddress();
                var dixt = new Dictionary<string, string>();
                dixt.Add("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                //dixt.Add("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                //dixt.Add("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
                string ss = Serialize(worksheet, dixt);
                ss = ss.Replace("encoding=\"utf-16\"", "encoding=\"UTF-8\" standalone=\"yes\"");

                return ss;
            }
            catch
            {

                throw;
            }
        }

        /// <summary>
        /// Build Empty Cells
        /// </summary>
        /// <param name="cells"></param>
        /// <param name="rows"></param>
        /// <param name="columns"></param>
        private void BuildEmptyCells(ref List<Cell> cells, int rows, int columns)
        {
            for (var rowindex = cells.Min(x => x.Address.Row); rowindex <= rows; rowindex++)
            {
                for (var columnindex = cells.Min(x => x.Address.Column.Index); columnindex <= columns; columnindex++)
                {
                    var item = cells.FirstOrDefault(x => x.Address.Row == rowindex && x.Address.Column.Index == columnindex);
                    if (null == item)
                    {
                        var cell = new Cell() { Value = string.Empty, Address = new CellAddress(new Column(columnindex), rowindex) };
                        cells.Add(cell);
                    }
                }
            }
        }

        /// <summary>
        /// XML Serializer
        /// </summary>
        /// <param name="soapobj"></param>
        /// <param name="soapnamespaces"></param>
        /// <param name="extraTypes"></param>
        /// <returns></returns>
        private static string Serialize(object soapobj, Dictionary<string, string> soapnamespaces, Type[] extraTypes = null)
        {
            try
            {
                XmlSerializerNamespaces namespaceser = new XmlSerializerNamespaces();
                using (var streamwriter = new StringWriter())
                {
                    foreach (var item in soapnamespaces)
                    {
                        namespaceser.Add(item.Key, item.Value);
                    }
                    var type = soapobj.GetType();
                    var serializer = extraTypes == null ? new XmlSerializer(type) : new XmlSerializer(type, extraTypes);
                    serializer.Serialize(streamwriter, soapobj, namespaceser);
                    return streamwriter.ToString();
                }
            }
            catch
            {
                throw;
            }
        }

    }
}
