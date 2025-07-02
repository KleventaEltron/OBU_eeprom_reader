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
        FolderNavigation folderNavigation = new FolderNavigation();

        public xml_loader()
        {
            InitializeComponent();
        }

        private void button_get_xml_Click(object sender, EventArgs e)
        {
            string xmlFilePath = folderNavigation.FileOrFolderNavigation("OpenFile");
            if (xmlFilePath != null)
            {
                Console.WriteLine("File path: " + xmlFilePath);
                textBox_xml_location.Text = xmlFilePath;

                if (checkBox1.Checked)
                {
                    xmlEditor.ImportFromExcel(EepromList.unfoldedMapping, xmlFilePath, true);
                    InsertUnfoldedMapping();
                }
                else
                {
                    xmlEditor.ImportFromExcel(EepromList.mapping, xmlFilePath, false);
                    InsertFoldedMapping();
                }
                
            }
        }

        private void save_eeprom_btn_Click(object sender, EventArgs e)
        {
            string xmlFilePath = folderNavigation.FileOrFolderNavigation("SaveFile");
            save_eeprom_textbox.Text = xmlFilePath;
            xmlEditor.CreateEepromExcel(EepromList.unfoldedMapping, xmlFilePath);
        }

        private void InsertFoldedMapping()
        {
            foreach (var item in EepromList.mapping)
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

            //PopulateFoldedMapping populateFoldedMapping = new PopulateFoldedMapping();
            //List<EepromMapping> MappingList = new List<EepromMapping>();
            //bool listInsert = false;

            //foreach (var item in EepromList.unfoldedMapping)
            //{
            //    var mapping = EepromList.GetMappingByAddress(item.Address);
            //    if (mapping != null && mapping.EepromDataType == typeof(byte[,]))
            //    {
            //        listInsert = true;
            //        MappingList.Add(item);
            //        Console.WriteLine("Found one");
            //        continue;
            //    }

            //    if (listInsert && mapping == null)
            //    {
            //        MappingList.Add(item);
            //        continue;
            //    }

            //    if(listInsert && mapping != null)
            //    {
            //        listInsert = false;
            //        byte[,] temp = populateFoldedMapping.ConvertToByteMatrix(MappingList);
            //        foreach(var item2 in temp)
            //        {
            //            //Console.WriteLine(item2);
            //        }
            //        MappingList.Clear();


            //    }

            //    var data = populateFoldedMapping.ConvertToBytes(item);
            //    Console.WriteLine(data.ToString());
            //populateFoldedMapping.ConvertToBytesAndUpdate(item);
        }

        private void InsertUnfoldedMapping()
        {
            PopulateUnfoldedMapping populateUnfoldedMapping = new PopulateUnfoldedMapping();

            //foreach (var item in EepromList.mapping)
            //{
            //    populateUnfoldedMapping.FindUnfoldedMapping(item);
            //}

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
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
