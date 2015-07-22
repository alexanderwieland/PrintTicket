using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Printing;
using System.IO;
using System.Xml;


namespace ConsoleApplication2
{
    class Program
    {
        static void Main(string[] args)
        {
            PrintTicket oldticket = GetPrintTicketFromPrinter();

            Dictionary<string, string> inputbins = GetInputBins("Phaser 6700DN");

            PrintTicket myPrintTicket = ModifyPrintTicket(oldticket, "psk:JobInputBin", "");
        }

        public static Dictionary<string, string> GetInputBins(string printerName)
        {
            Dictionary<string, string> inputBins = new Dictionary<string, string>();

            // get PrintQueue of Printer from the PrintServer
            LocalPrintServer printServer = new LocalPrintServer();
            PrintQueue printQueue = printServer.GetPrintQueue(printerName);

            // get PrintCapabilities of the printer
            MemoryStream printerCapXmlStream = printQueue.GetPrintCapabilitiesAsXml();

            // read the JobInputBins out of the PrintCapabilities
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(printerCapXmlStream);

            // create NamespaceManager and add PrintSchemaFrameWork-Namespace (should be on DocumentElement of the PrintTicket)
            // Prefix: psf NameSpace: xmlDoc.DocumentElement.NamespaceURI = "http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework"
            XmlNamespaceManager manager = new XmlNamespaceManager(xmlDoc.NameTable);
            manager.AddNamespace(xmlDoc.DocumentElement.Prefix, xmlDoc.DocumentElement.NamespaceURI);

            // and select all nodes of the bins
            XmlNodeList nodeList = xmlDoc.SelectNodes("//psf:Feature[@name='psk:JobInputBin']/psfSurpriseption", manager);

            // fill Dictionary with the bin-names and values
            foreach (XmlNode node in nodeList)
            {
                inputBins.Add(node.LastChild.InnerText, node.Attributes["name"].Value);
            }

            return inputBins;
        }


        public static PrintTicket ModifyPrintTicket(PrintTicket ticket, string featureName, string newValue)
        {
            if (ticket == null)
            {
                throw new ArgumentNullException("ticket");
            }

            // read Xml of the PrintTicket
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(ticket.GetXmlStream());

            // create NamespaceManager and add PrintSchemaFrameWork-Namespace hinzufugen (should be on DocumentElement of the PrintTicket)
            // Prefix: psf NameSpace: xmlDoc.DocumentElement.NamespaceURI = "http://schemas.microsoft.com/windows/2003/08/printing/printschemaframework"
            XmlNamespaceManager manager = new XmlNamespaceManager(xmlDoc.NameTable);
            manager.AddNamespace(xmlDoc.DocumentElement.Prefix, xmlDoc.DocumentElement.NamespaceURI);

            // search node with desired feature we're looking for and set newValue for it
            string xpath = string.Format("//psf:Feature[@name='{0}']/psfSurpriseption", featureName);
            XmlNode node = xmlDoc.SelectSingleNode(xpath, manager);
            if (node != null)
            {
                node.Attributes["name"].Value = newValue;
            }

            // create a new PrintTicket out of the XML
            MemoryStream printTicketStream = new MemoryStream();
            xmlDoc.Save(printTicketStream);
            printTicketStream.Position = 0;
            PrintTicket modifiedPrintTicket = new PrintTicket(printTicketStream);

            // for testing purpose save the printticket to file
            //FileStream stream = new FileStream("modPrintticket.xml", FileMode.CreateNew, FileAccess.ReadWrite);
            //modifiedPrintTicket.GetXmlStream().WriteTo(stream);

            return modifiedPrintTicket;
        }

        static private PrintTicket GetPrintTicketFromPrinter()
        {
            PrintQueue printQueue = null;

            LocalPrintServer localPrintServer = new LocalPrintServer();

            // Retrieving collection of local printer on user machine
            PrintQueueCollection localPrinterCollection =
                localPrintServer.GetPrintQueues();

            System.Collections.IEnumerator localPrinterEnumerator =
                localPrinterCollection.GetEnumerator();

            if (localPrinterEnumerator.MoveNext())
            {
                // Get PrintQueue from first available printer
                printQueue = (PrintQueue)localPrinterEnumerator.Current;
            }
            else
            {
                // No printer exist, return null PrintTicket 
                return null;
            }

            // Get default PrintTicket from printer
            PrintTicket printTicket = printQueue.DefaultPrintTicket;

            PrintCapabilities printCapabilites = printQueue.GetPrintCapabilities();

            // Modify PrintTicket 
            if (printCapabilites.CollationCapability.Contains(Collation.Collated))
            {
                printTicket.Collation = Collation.Collated;
            }

            if (printCapabilites.DuplexingCapability.Contains(
                    Duplexing.TwoSidedLongEdge))
            {
                printTicket.Duplexing = Duplexing.TwoSidedLongEdge;
            }

            if (printCapabilites.StaplingCapability.Contains(Stapling.StapleDualLeft))
            {
                printTicket.Stapling = Stapling.StapleDualLeft;
            }

            return printTicket;
        }// end:GetPrintTicketFromPrinter()
    }

}
