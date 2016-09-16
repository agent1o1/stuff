using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Xml.Linq;

namespace RDLC
{
    public class Utf8StringWriter : StringWriter
    {
        public override Encoding Encoding { get { return Encoding.UTF8; } }
    }

    public class RDLCGenerator
    {
        List<string> subrep = new List<string>();
        MemoryStream ms = new MemoryStream();
        private static Dictionary<string, List<string>> Parameters = new Dictionary<string, List<string>>() 
        {
            { "Name of Report", new List<string>(){"Parameters for report"} } 
        };
        public RDLCGenerator(List<string> subreports)
        {
            subrep = subreports;

            Generate();
        }

        public MemoryStream GetReport()
        {
            return ms;
        }

        private void Generate()
        {
            XNamespace xmlns = "http://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition";
            XDocument doc = new XDocument(new XDeclaration("1.0", "utf-8", null),
                                                        new XElement(xmlns + "Report",
                                                            new XAttribute(XNamespace.Xmlns + "rd", "http://schemas.microsoft.com/SQLServer/reporting/reportdesigner"),
                                                            new XAttribute(XNamespace.Xmlns + "cl", "http://schemas.microsoft.com/sqlserver/reporting/2010/01/componentdefinition"),
                                                            new XElement(xmlns + "AutoRefresh", 0),
                                                            new XElement(xmlns + "ReportSections",
                                                                new XElement(xmlns + "ReportSection",
                                                                    new XElement(xmlns + "Body",
                                                                        new XElement(xmlns + "ReportItems"),
                                                                        new XElement(xmlns + "Height", "6in")
                                                                                ),
                                                                    new XElement(xmlns + "Width", "8in"),
                                                                    new XElement(xmlns + "Page",
                                                                        new XElement(xmlns + "PageHeight", "30cm"),
                                                                        new XElement(xmlns + "PageWidth", "30cm")
                                                                                )
                                                                            )
                                                                        )
                                                                    )
                                          );

            PasteSubreport(subrep, doc);
            var wr = new Utf8StringWriter();
            doc.Save(wr);
            StreamWriter writer = new StreamWriter(ms);
            writer.Write(wr.GetStringBuilder().ToString());
            writer.Flush();
            ms.Position = 0;
        }

        private void PasteSubreport(List<string> names, XDocument doc)
        {
            XNamespace xmlns = "http://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition";
            doc.Element(xmlns + "Report").Element(xmlns + "ReportSections").Element(xmlns + "ReportSection").Element(xmlns + "Body").Element(xmlns + "ReportItems").Add(new XElement(xmlns + "Subreport", new XAttribute("Name", names.FirstOrDefault()),
                                                                new XElement(xmlns + "ReportName", @"../MainReports/Focus.rdlc"),
                                                                new XElement(xmlns + "Height", "2in"),
                                                                new XElement(xmlns + "Width", "2in"),
                                                                new XElement(xmlns + "Style", null)));
            int topcounter = 0;

            //if (Parameters[names.FirstOrDefault()] != null)
            //{
            //    foreach (string item in Parameters[names.FirstOrDefault()])
            //    {
            //        PasteSubrepParams(item, doc, names.FirstOrDefault());
            //    }
            //}

            foreach (string name in names.GetRange(1, names.Count - 1))
            {
                topcounter++;
                doc.Element(xmlns + "Report").Element(xmlns + "ReportSections").Element(xmlns + "ReportSection").Element(xmlns + "Body").Element(xmlns + "ReportItems").Add(new XElement(xmlns + "Subreport", new XAttribute("Name", name + topcounter),
                                                               new XElement(xmlns + "ReportName", @"../MainReports/Focus.rdlc"),
                                                               new XElement(xmlns + "Top", topcounter * 6 + "cm"),
                                                               new XElement(xmlns + "Height", "2in"),
                                                               new XElement(xmlns + "Width", "2in"),
                                                               new XElement(xmlns + "ZIndex", topcounter),
                                                               new XElement(xmlns + "Style", null)));
                //if (Parameters[name] != null)
                //{
                //    foreach (string item in Parameters[name])
                //    {
                //        PasteSubrepParams(item, doc, name + topcounter);
                //    }
                //}
            }

            //doc.Element(xmlns + "Report").Add(new XElement(xmlns + "ReportParameters", 
            //                                                    new XElement(xmlns + "ReportParameter"), new XAttribute("Name", "Path"),
            //                                                        new XElement(xmlns + "DataType", "String"),
            //                                                        new XElement(xmlns + "Prompt", "Path"),
            //                                                        new XElement(xmlns + "Value", AppDomain.CurrentDomain.BaseDirectory + @"MainReports/Focus.rdlc")));
        }

        private void PasteSubrepParams(string param, XDocument doc, string subrepName)
        {
            XNamespace xmlns = "http://schemas.microsoft.com/sqlserver/reporting/2010/01/reportdefinition";
            doc.Element(xmlns + "Report").Element(xmlns + "ReportSections").Element(xmlns + "ReportSection").
                Element(xmlns + "Body").Element(xmlns + "ReportItems").
                Descendants(xmlns + "Subreport").Where(x => (string)x.Attribute("Name") == subrepName).SingleOrDefault().
                Add(new XElement(xmlns + "Parameters",
                        new XElement(xmlns + "Parameter", new XAttribute("Name", param),
                            new XElement(xmlns + "Value", /*"=Parameters!" + param + ".Value"*/"23396"))));
            //todo
 
        }
    }
}