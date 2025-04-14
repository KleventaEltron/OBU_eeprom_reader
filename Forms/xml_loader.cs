using EepromReader.Scripts;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EepromReader.Forms
{
    public partial class xml_loader : Form
    {
        XmlEditor xmlEditor = new XmlEditor();

        public xml_loader()
        {
            InitializeComponent();
        }

        private void button_get_xml_Click(object sender, EventArgs e)
        {
            string xmlFilePath = openXmlFile();
            if (xmlFilePath != null)
            {
                Console.WriteLine("File path: " + xmlFilePath);
                xmlEditor.ImportFromExcel(EepromList.mapping, xmlFilePath);
                TestUnfoldedMapping();
            }
        }

        private string openXmlFile()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                // Set the filter to XML files only
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xlsm|All Files|*.*";
                openFileDialog.Title = "Select an XML File";

                // Show the dialog and check if the user selects a file
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedFilePath = openFileDialog.FileName;
                    Console.WriteLine("Selected file: " + selectedFilePath);
                    textBox_xml_location.Text = selectedFilePath;
                    return selectedFilePath;
                }
                else
                {
                    Console.WriteLine("No file selected.");
                    return null;
                }
            }
        }

        private void TestUnfoldedMapping()
        {
            PopulateUnfoldedMapping populateUnfoldedMapping = new PopulateUnfoldedMapping();

            foreach (var item in EepromList.mapping)
            {
                populateUnfoldedMapping.FindUnfoldedMapping(item);
            }

            foreach (var item in EepromList.unfoldedMapping)
            {
                Console.WriteLine($"Address: {item.Address}");
                Console.WriteLine($"Name: {item.Name}");
                Console.WriteLine($"Length: {item.Length}");
                if (item.Value is byte[] byteArray)
                {
                    Console.WriteLine($"Value: {BitConverter.ToString(byteArray)}");
                }
                else
                {
                    Console.WriteLine($"Value: {item.Value}");
                }
                Console.WriteLine($"Datatype: {item.EepromDataType}\n");
            }

            xmlEditor.CreateEepromExcel(EepromList.unfoldedMapping, "C:\\Users\\bobri\\Desktop\\EepromUnfoldedMapping.xlsx");
        }
    }
}
