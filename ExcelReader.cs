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
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using System.Xml.Serialization;
using Windows.ApplicationModel;
using Windows.Data.Xml.Dom;
using Windows.Storage;
using ExcelManager.Entity;

namespace ExcelManager
{
    /// <summary>
    /// Class meant for Reading Excel File
    /// </summary>
    public class ExcelReader
    {
        public ExcelReader()
        {
            //throw new NotImplementedException();
        }


        public void ReadExcelData(Stream excelfile)
        {
            //excelfile = Assembly.Load(new AssemblyName("ExcelManager")).GetManifestResourceStream(@"ExcelManager.Resources.ExcelTemplateFile.xlsx");
            System.IO.Compression.ZipArchive zipArchiveExcelTemplate = new System.IO.Compression.ZipArchive(excelfile);

            foreach (var file in zipArchiveExcelTemplate.Entries)
            {

                if (file.FullName == "xl/sharedStrings.xml")
                {
                    var s = file.Open();
                    var xdoc = XDocument.Load(s);
                    var xelement = XElement.Parse(xdoc.ToString());
                    var sst = Deserialize<sst>(xelement.ToString());

                }
                else if (file.FullName == "xl/worksheets/sheet1.xml")
                {
                    var s = file.Open();
                    var xdoc = XDocument.Load(s);
                    var xelement = XElement.Parse(xdoc.ToString());
                    var worksheet = Deserialize<worksheet>(xelement.ToString());

                }
                else if (file.FullName == "xl/styles.xml")
                {
                    //parsing cant be done as of now

                }
                else
                {

                }
            }

        }

        public T Deserialize<T>(string xmlString) where T : class
        {
            try
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(T));
                TextReader xmlText = new StringReader(xmlString);
                T xmlObject = (T)xmlSerializer.Deserialize(xmlText);
                return xmlObject;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
    }
}
