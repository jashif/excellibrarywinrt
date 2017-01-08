// ***********************************************************************
// Assembly         : ExcelManager
// Author           : Shibin
// Created          : 04-30-2014
//
// Last Modified By : Shibin
// Last Modified On : 04-04-2014
// ***********************************************************************
// <copyright file="" company="">
//     Copyright (c) . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Windows.Storage;
using Windows.Storage.Streams;
using System.IO;
using System.Runtime.Serialization;
namespace ExcelManager.Entity
{
    [XmlRoot(ElementName = "sst", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class sst
    {
        [XmlAttribute(AttributeName = "count")]
        public string count { get; set; }

        [XmlAttribute(AttributeName = "uniqueCount")]
        public string uniqueCount { get; set; }

        [XmlElement(ElementName = "si")]
        public List<si> si { get; set; }
    }

    public class si
    {
        [XmlElement(ElementName = "t")]
        public string t { get; set; }
    }

    [XmlRoot(ElementName = "worksheet", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class worksheet
    {
        [XmlElement(ElementName = "dimension")]
        public dimension dimension { get; set; }

        [XmlElement(ElementName = "sheetViews")]
        public sheetViews sheetviews { get; set; }

        [XmlElement(ElementName = "sheetFormatPr")]
        public sheetFormatPr sheetFormatPr { get; set; }

        [XmlElement(ElementName = "cols")]
        public cols cols { get; set; }

        [XmlElement(ElementName = "sheetData")]
        public sheetData sheetData { get; set; }

        [XmlElement(ElementName = "mergeCells")]
        public mergeCells mergeCells { get; set; }

        [XmlElement(ElementName = "pageMargins")]
        public pageMargins pageMargins { get; set; }

    }

    public class sheetViews
    {
        [XmlElement(ElementName = "sheetView")]
        public sheetView sheetview { get; set; }

        public sheetViews()
        {
            sheetview = new sheetView();
        }

    }

    public class dimension
    {
        [XmlAttribute(AttributeName = "ref")]
        public string reference { get; set; }
    }

    public class sheetView
    {
        [XmlAttribute(AttributeName = "tabSelected")]
        public string tabSelected { get; set; }

        [XmlAttribute(AttributeName = "workbookViewId")]
        public string workbookViewId { get; set; }

        public sheetView()
        {
            workbookViewId = "0";
            tabSelected = "1";
        }
    }

    public class sheetFormatPr
    {
        [XmlAttribute(AttributeName = "defaultRowHeight")]
        public string defaultRowHeight { get; set; }

        // [XmlAttribute(AttributeName = "dyDescent", Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")]
        [XmlIgnore]
        public string dyDescent { get; set; }
    }

    public class cols
    {
        [XmlElement(ElementName = "col")]
        public List<col> col { get; set; }
    }

    public class col
    {
        [XmlAttribute(AttributeName = "min")]
        public string min { get; set; }

        [XmlAttribute(AttributeName = "max")]
        public string max { get; set; }

        [XmlAttribute(AttributeName = "width")]
        public string width { get; set; }

        [XmlAttribute(AttributeName = "customWidth")]
        public string customWidth { get; set; }
    }

    public class pageMargins
    {
        [XmlAttribute(AttributeName = "left")]
        public string left { get; set; }

        [XmlAttribute(AttributeName = "right")]
        public string right { get; set; }

        [XmlAttribute(AttributeName = "top")]
        public string top { get; set; }

        [XmlAttribute(AttributeName = "bottom")]
        public string bottom { get; set; }

        [XmlAttribute(AttributeName = "header")]
        public string header { get; set; }

        [XmlAttribute(AttributeName = "footer")]
        public string footer { get; set; }

        public pageMargins()
        {
            left = "0.7";
            right = "0.7";
            top = "0.75";
            bottom = "0.75";
            header = "0.3";
            footer = "0.3";
        }
    }

    public class sheetData
    {
        [XmlElement(ElementName = "row")]
        public List<row> row { get; set; }
    }

    public class row
    {
        [XmlAttribute(AttributeName = "r")]
        public string r { get; set; }

        [XmlAttribute(AttributeName = "spans")]
        public string spans { get; set; }

        //[XmlAttribute(AttributeName = "dyDescent", Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")]
        [XmlIgnore]
        public string dyDescent { get; set; }


        [XmlElement(ElementName = "c")]
        public List<c> c { get; set; }

        [XmlAttribute(AttributeName = "ht")]
        public string height { get; set; }

        [XmlAttribute(AttributeName = "customHeight")]
        public string customHeight { get; set; }
    }

    public class c
    {

        [XmlAttribute(AttributeName = "r")]
        public string r { get; set; }

        [XmlAttribute(AttributeName = "s")]
        public string s { get; set; }

        [XmlAttribute(AttributeName = "t")]
        public string t { get; set; }

        [XmlElement(ElementName = "v")]
        public string v { get; set; }

    }

    public class mergeCells
    {
        [XmlAttribute(AttributeName = "count")]
        public string count { get; set; }

        [XmlElement(ElementName = "mergeCell")]
        public List<mergeCell> mergeCell { get; set; }
    }

    public class mergeCell
    {
        [XmlAttribute(AttributeName = "ref")]
        public string reference { get; set; }
    }
}
