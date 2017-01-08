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
using System.IO;
using System.Linq;
using System.Text;
using System.IO.Compression;
using Windows.Storage;
using System.Threading.Tasks;
using Windows.Data.Xml.Dom;
using System.Reflection;
using Windows.ApplicationModel.Resources.Core;
using Windows.ApplicationModel.Resources;
using System.Collections.Generic;
namespace ExcelManager
{
    /// <summary>
    /// Class meant for creating and writing to a new Excel file
    /// </summary>
    public class ExcelWriter
    {
        /// <summary>
        /// Gets or Sets Excel Sheet Property
        /// </summary>
        public Sheet ExcelSheet
        {
            get;
            private set;
        }

        /// <summary>
        /// Get or Sets Excel Sheet2
        /// </summary>
        public Sheet ExcelSheet2
        {
            get;
            private set;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        public ExcelWriter()
        {
            ExcelSheet = new Sheet();
            ExcelSheet2 = new Sheet();
        }

        /// <summary>
        /// Reading from an excel template , entering data in it and write back to a file
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public async Task<StorageFile> WriteToFile(string filename, StorageFile targetfile)
        {
            //StorageFile targetfile = null;

            try
            {
                var exceltemplatestream = Assembly.Load(new AssemblyName("ExcelManager")).GetManifestResourceStream(@"ExcelManager.Resources.ExcelTemplateFile.xlsx");

                System.IO.Compression.ZipArchive zipArchiveExcelTemplate = new System.IO.Compression.ZipArchive(exceltemplatestream);

                //Data Builder Code
                ExcelXmlBuilder excelbuilder = new ExcelXmlBuilder();
                XmlDocument xstyleDoc = null;
                string sheet2string = "", sheet1string;
                List<Cell> sheetcells = new List<Cell>(ExcelSheet.Cells);

                var styles = zipArchiveExcelTemplate.Entries.FirstOrDefault(x => x.FullName == "xl/styles.xml");
                var styledata = GetByteArrayFromStream(styles.Open());
                xstyleDoc = GetStyleXDoc(styledata);
                sheet1string = excelbuilder.BuildSheetXML(ExcelSheet, ref xstyleDoc);

                if (ExcelSheet2 != null && ExcelSheet2.Cells != null && ExcelSheet2.Cells.Count > 0)
                {
                    sheet2string = excelbuilder.BuildSheetXML(ExcelSheet2, ref xstyleDoc, ExcelSheet.Cells.Count);
                    sheetcells.AddRange(ExcelSheet2.Cells);
                }

                using (var zipStream = await targetfile.OpenStreamForWriteAsync())
                {
                    using (System.IO.Compression.ZipArchive newzip = new System.IO.Compression.ZipArchive(zipStream, ZipArchiveMode.Create))
                    {
                        foreach (var file in zipArchiveExcelTemplate.Entries)
                        {
                            System.IO.Compression.ZipArchiveEntry entry = newzip.CreateEntry(file.FullName);
                            using (Stream ZipFile = entry.Open())
                            {
                                byte[] data = null;
                                if (file.FullName == "xl/sharedStrings.xml")
                                {
                                    data = Encoding.UTF8.GetBytes(excelbuilder.BuildSharedXML(sheetcells));
                                    ZipFile.Write(data, 0, data.Length);
                                }
                                else if (file.FullName == "xl/worksheets/sheet1.xml")
                                {
                                    data = Encoding.UTF8.GetBytes(sheet1string);
                                    ZipFile.Write(data, 0, data.Length);
                                }
                                else if (file.FullName == "xl/worksheets/sheet2.xml" && sheet2string != "")
                                {
                                    data = Encoding.UTF8.GetBytes(sheet2string);
                                    ZipFile.Write(data, 0, data.Length);
                                }
                                else if (file.FullName == "xl/styles.xml")
                                {
                                    if (xstyleDoc != null)
                                    {
                                        data = Encoding.UTF8.GetBytes(xstyleDoc.GetXml().Replace("xmlns=\"\"", ""));
                                        ZipFile.Write(data, 0, data.Length);
                                    }
                                }
                                else
                                {
                                    data = GetByteArrayFromStream(file.Open());
                                    ZipFile.Write(data, 0, data.Length);
                                }
                            }
                        }
                    }
                }

            }
            catch
            {
                throw;
            }
            return targetfile;
        }

        /// <summary>
        /// Get Xml Document for Styles from corresponding Byte array
        /// </summary>
        /// <param name="databytes"></param>
        /// <returns></returns>
        private XmlDocument GetStyleXDoc(byte[] databytes)
        {
            var xDoc = new XmlDocument();
            xDoc.LoadXml(Encoding.UTF8.GetString(databytes, 0, databytes.Length));
            return xDoc;

        }

        /// <summary>
        /// Get Byte array from stream
        /// </summary>
        /// <param name="contentstream"></param>
        /// <returns></returns>
        private byte[] GetByteArrayFromStream(Stream contentstream)
        {
            using (var memoryStream = new MemoryStream())
            {
                contentstream.CopyTo(memoryStream);
                return memoryStream.ToArray();
            }
        }
    }
}
