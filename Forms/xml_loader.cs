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
            if (checkBox1.Checked)
            {
                xmlEditor.CreateEepromExcel(EepromList.unfoldedMapping, xmlFilePath);
            }
            else
            {
                xmlEditor.CreateEepromExcel(EepromList.mapping, xmlFilePath);
            }
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
                    Console.WriteLine($"Value: {string.Join(" ", byteArray)}");
                }
                else if (item.Value is byte[,] byteArray2D)
                {
                    Console.WriteLine("Value:");
                    int dim0 = byteArray2D.GetLength(0);
                    int dim1 = byteArray2D.GetLength(1);
                    for (int i = 0; i < dim0; i++)
                    {
                        for (int j = 0; j < dim1; j++)
                        {
                            Console.Write($"{byteArray2D[i, j]} ");
                        }
                        Console.WriteLine();
                    }
                }
                else if (item.Value is byte[,,] byteArray3D)
                {
                    Console.WriteLine("Value:");
                    int d0 = byteArray3D.GetLength(0);
                    int d1 = byteArray3D.GetLength(1);
                    int d2 = byteArray3D.GetLength(2);
                    for (int i = 0; i < d0; i++)
                    {
                        Console.WriteLine($"Slice {i}:");
                        for (int j = 0; j < d1; j++)
                        {
                            for (int k = 0; k < d2; k++)
                            {
                                Console.Write($"{byteArray3D[i, j, k]} ");
                            }
                            Console.WriteLine();
                        }
                    }
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

        private void button_convert_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                PopulateUnfoldedMapping populateUnfoldedMapping = new PopulateUnfoldedMapping();

                foreach (var item in EepromList.mapping)
                {
                    populateUnfoldedMapping.FindUnfoldedMapping(item);
                }

                MessageBox.Show("Conversie naar unfolded geslaagd", "Conversie popup", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
