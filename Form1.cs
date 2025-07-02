using DocumentFormat.OpenXml.Office.CustomUI;
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

namespace EepromReader
{
    public partial class Form1 : Form
    {
        XmlEditor xmlEditor = new XmlEditor();
        ConvertValues convertValues = new ConvertValues();

        public Form1()
        {
            InitializeComponent();
        }

        // Test button for importing from Excel
        private void button1_Click(object sender, EventArgs e)
        {
            string excelFilePath = "C:\\Users\\bob\\Desktop\\KVM-Manager-test\\EepromUnfoldedMapping.xlsx";
            xmlEditor.ImportFromExcel(EepromList.unfoldedMapping, excelFilePath, true);

            foreach (var item in EepromList.unfoldedMapping)
            {
                Console.WriteLine(item.Address);
                Console.WriteLine(item.Name);
                Console.WriteLine(item.Length);
                Console.WriteLine(item.Value);
                Console.WriteLine(item.EepromDataType);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            TestUnfoldedMapping();

/*            var matrix = new byte[,]
            {
                { 83, 69, 78, 83, 79, 82, 32, 49 },
                { 83, 69, 78, 83, 79, 82, 32, 50 },
                { 83, 69, 78, 83, 79, 82, 32, 51 }
            };

            var cube = new byte[,,]
            {
                {
                    { 83, 69, 78, 83, 79, 82, 32, 49 },
                    { 83, 69, 78, 83, 79, 82, 32, 50 },
                    { 83, 69, 78, 83, 79, 82, 32, 51 }
                },
                {
                    { 83, 69, 78, 83, 79, 82, 32, 52 },
                    { 83, 69, 78, 83, 79, 82, 32, 53 },
                    { 83, 69, 78, 83, 79, 82, 32, 54 }
                }
            };

            //var test = convertValues.ConvertToString(EepromList.GetMappingByAddress(1674));
            //convertValues.ConvertAddress(EepromList.GetMappingByAddress(1)); 
            //var temp = convertValues.ConvertMatrixToStringList(matrix);
            var temp = convertValues.ConvertKubusToStringList(cube);

            foreach (var item in temp)
            {
                Console.WriteLine(item);
            }*/

        }

        private void TestUnfoldedMapping()
        {
            PopulateUnfoldedMapping populateUnfoldedMapping = new PopulateUnfoldedMapping();

            populateUnfoldedMapping.FindUnfoldedMapping(EepromList.GetMappingByAddress(1674));

            foreach(var item in EepromList.mapping) 
            {
                populateUnfoldedMapping.FindUnfoldedMapping(item);
            }

            //EepromList.PrintUnfoldedMapping(1674);

            foreach (var item in EepromList.unfoldedMapping)
            {
                Console.WriteLine(item.Address);
                Console.WriteLine(item.Name);
                Console.WriteLine(item.Length);
                Console.WriteLine(item.Value);
                Console.WriteLine(item.EepromDataType);
            }

        }
    }
}
