using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace InventoryDataMerge2013
{
    class InvXMLFile
    {
        OpenFileDialog fileXml = new OpenFileDialog();
        internal List<XElement> systems;
        internal InvXMLFile()
        {
            fileXml.Title = "Open IDC Data File - TaxAideInv2013.xml";
            fileXml.Filter = "XML files (*.xml)|*.xml";
            fileXml.FileName = "TaxAideInv2013.xml";
        }

        internal bool GetIDCXmlData()
        {
            DialogResult dlg = fileXml.ShowDialog();
            if (dlg == DialogResult.Cancel)
                return false;
            XElement root = null;
            try
            {
                root = XElement.Load(fileXml.FileName);
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("This File is not a correctly formatted XML file. \rExiting!", Start.mbCaption, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
            var compres = root.Nodes().Where(el => el.NodeType == System.Xml.XmlNodeType.Comment).Any(ele => ele.ToString() == "<!--IDC XML Version 2013.03-->");
            if (!compres)
            {
                XDocument doc = null;
                try
                {
                    doc = XDocument.Load(fileXml.FileName);
                }
                catch (Exception)
                {
                    System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show("This File is not a correctly formatted XML file. \rExiting!", "IDC Data Merge", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return false;
                }
                if (doc.Nodes().FirstOrDefault().ToString() != "<!--IDC XML Version 2013.02-->")
                {
                    MessageBox.Show("This file is not the required 2013 version 2 or 3 IDC Inventory file.\rExiting!", Start.mbCaption, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return false;
                }
                var systemsTemp = root.Elements("system").ToList();
                systemsTemp.RemoveAll(el => el.HasElements == false);
                if (systemsTemp.Count == 0)
                {
                    MessageBox.Show("This file contains no system data\rExiting!", Start.mbCaption, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return false;
                }
                systems = systemsTemp.Select(sysEL => new XElement("system", sysEL.Elements().Select(fld => new XElement(fld.Name.ToString().ToLower(), fld.Value)))).ToList();
            }
            else
            {
                systems = root.Elements("system").ToList();
                systems.RemoveAll(el => el.HasElements == false);
                if (systems.Count == 0)
                {
                    MessageBox.Show("This file contains no system data\rExiting!", Start.mbCaption, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return false;
                }
            }
            //At this point have a List of Xelements each of which is a system
            return true;
        }
    }
}
