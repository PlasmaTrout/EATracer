using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Text;

namespace EAWordTracer
{
    public delegate void ReportElementTraversedHandler(string element,string path, bool direct);

    public class TableLoader
    {

        public event ReportElementTraversedHandler ReportElementTraversed;

        Dictionary<string, int> completionDictionary = new Dictionary<string, int>();
        List<string> directImpacts = new List<string>();
        List<string> indirectImpacts = new List<string>();

        private int rows = 0;
        public Table CurrentTable { get; set; }

        public EA.Repository Repository { get;  }

        public TableLoader(EA.Repository repo,Table wordTable)
        {
            this.Repository = repo;
            this.CurrentTable = wordTable;
        }

        public void RenderTable(EA.Element element)
        {

            GetConnections(element);

        }

        private void GetConnections(EA.Element element)
        {
            EA.Collection connectors = element.Connectors;

            for (short i = 0; i < connectors.Count; i++)
            {
                EA.Connector connector = connectors.GetAt(i);
                EA.Element destination = this.Repository.GetElementByID(connector.SupplierID);
                EA.Element source = this.Repository.GetElementByID(connector.ClientID);
                String key = $"{source.Name}->{destination.Name}";

                if (completionDictionary.ContainsKey(key))
                {
                    completionDictionary[key] = completionDictionary[key] + 1;
                }
                else
                {
                    completionDictionary.Add(key, 1);

                    if(source.Name == element.Name)
                    {
                        ReportElementTraversed?.Invoke(destination.Name, key, false);
                        GetConnections(destination);
                    }

                    if(destination.Name == element.Name)
                    {
                        ReportElementTraversed?.Invoke(source.Name, key, true);
                        GetConnections(source);
                    }
                }
            }
        }

       

        
    }
}
