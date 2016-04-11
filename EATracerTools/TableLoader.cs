using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace EATracer
{
    public delegate void ReportElementTraversedHandler(string element);

    public enum TableLoadDirection
    {
        Forwards,
        Backwards,
        Both
    }

    public class TableLoader
    {

        public event ReportElementTraversedHandler ReportElementTraversed;

        Dictionary<string, int> completionDictionary = new Dictionary<string, int>();
        private int rows = 0;
        

        public EA.Repository Repository { get;  }

        public TableLoadDirection Direction { get; set; }

        public TableLoader(EA.Repository repo,TableLoadDirection dir)
        {
            this.Repository = repo;
            this.Direction = dir;
        }

        public void RenderTable(EA.Element element)
        {
            GetConnections(element, Direction.ToString());
            AddHeaders();
        }

        private void GetConnections(EA.Element element, string traversal)
        {
            EA.Collection connectors = element.Connectors;

            for (short i = 0; i < connectors.Count; i++)
            {
                EA.Connector connector = connectors.GetAt(i);
                EA.Element target = this.Repository.GetElementByID(connector.SupplierID);
                EA.Element source = this.Repository.GetElementByID(connector.ClientID);
                String key = target.Name + ":" + source.Name;

                if (completionDictionary.ContainsKey(key))
                {
                    completionDictionary[key] = completionDictionary[key] + 1;
                }
                else
                {
                    completionDictionary.Add(key, 1);
                    AddRow(source, target, traversal, connector);

                    if (Direction == TableLoadDirection.Forwards || Direction == TableLoadDirection.Both)
                    {
                        GetConnections(target, TableLoadDirection.Forwards.ToString());
                    }
                    if (Direction == TableLoadDirection.Backwards || Direction == TableLoadDirection.Both)
                    {
                        GetConnections(source, TableLoadDirection.Backwards.ToString());
                    }
                }
            }
        }

        private void AddRow(EA.Element source, EA.Element destination, string traversal, EA.Connector connector)
        {
            var sheet = Globals.ThisAddIn.Application.ActiveSheet;

            sheet.Range["A1"].EntireRow.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
            sheet.Range["A1"].Value = source.Name;
            sheet.Range["B1"].Value = source.Type;
            sheet.Range["C1"].Value = destination.Name;
            sheet.Range["D1"].Value = destination.Type;
            sheet.Range["E1"].Value = traversal;
            sheet.Range["F1"].Value = connector.Type;
            sheet.Range["G1"].Value = connector.Stereotype;
            sheet.Range["H1"].Value = connector.Direction;
            sheet.Range["I1"].Value = rows;

            ReportElementTraversed?.Invoke($"{source.Name}-->{destination.Name}");
            rows++;

        }

        private void AddHeaders()
        {
            Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
            sheet.Range["A1"].EntireRow.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
            sheet.Range["A1"].Value = "Source";
            sheet.Range["B1"].Value = "Source Type";
            sheet.Range["C1"].Value = "Destination";
            sheet.Range["D1"].Value = "Dest Type";
            sheet.Range["E1"].Value = "Traversal";
            sheet.Range["F1"].Value = "Connector";
            sheet.Range["G1"].Value = "StereoType";
            sheet.Range["H1"].Value = "Direction";
            sheet.Range["I1"].Value = "Order";

            rows++;

            sheet.Columns.AutoFit();
            Range selection = sheet.Range["A1", "I" + rows.ToString()];
            ListObject lo = sheet.ListObjects.AddEx(Microsoft.Office.Interop.Excel.XlListObjectSourceType.xlSrcRange, selection, null, XlYesNoGuess.xlYes);
            lo.Name = "Table1";
            
            selection.Select();
            sheet.ListObjects["Table1"].TableStyle = "TableStyleMedium6";
            //selection.AutoFormat(XlRangeAutoFormat.xlRangeAutoFormat3DEffects1);


        }
    }
}
