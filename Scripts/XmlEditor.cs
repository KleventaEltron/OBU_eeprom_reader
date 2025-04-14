using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace EepromReader.Scripts
{
    public class XmlEditor
    {
        private byte[] TempArray = new byte[4096];

        public void CreateEepromExcel(List<EepromMapping> mappings, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("EEPROM Mapping");
                int currentRow = 1;

                // Add headers
                worksheet.Cell(currentRow, 1).Value = "Address";
                worksheet.Cell(currentRow, 2).Value = "Name";
                worksheet.Cell(currentRow, 3).Value = "Length";
                worksheet.Cell(currentRow, 4).Value = "Data Type";
                worksheet.Cell(currentRow, 5).Value = "Values";
                currentRow++;

                foreach (var mapping in mappings)
                {
                    worksheet.Cell(currentRow, 1).Value = mapping.Address;
                    worksheet.Cell(currentRow, 2).Value = mapping.Name;
                    worksheet.Cell(currentRow, 3).Value = mapping.Length;
                    worksheet.Cell(currentRow, 4).Value = mapping.EepromDataType.Name;

                    var cell = worksheet.Cell(currentRow, 5);
                    var value = mapping.Value;

                    // Set correct typed value
                    if (value is byte b) cell.SetValue(b);
                    else if (value is ushort us) cell.SetValue(us);
                    else if (value is short s) cell.SetValue(s);
                    else if (value is uint ui) cell.SetValue(ui);
                    else if (value is int i) cell.SetValue(i);
                    else if (value is sbyte sb) cell.SetValue(sb);
                    else if (value is string str) cell.SetValue(str);
                    else if (value is byte[] ba) cell.SetValue(string.Join(", ", ba));
                    else if (value is ushort[] usa) cell.SetValue(string.Join(", ", usa));
                    else if (value is short[] sa) cell.SetValue(string.Join(", ", sa));
                    else if (value is uint[] uia) cell.SetValue(string.Join(", ", uia));
                    else if (value is sbyte[] sba) cell.SetValue(string.Join(", ", sba));
                    else if (value is string[] stra) cell.SetValue(string.Join(", ", stra));
                    else if (value is byte[,] byte2DArray) cell.SetValue(Convert2DArrayToString(byte2DArray));
                    else if (value is byte[,,] byte3DArray) cell.SetValue(Convert3DArrayToString(byte3DArray));
                    else cell.SetValue("Unsupported type");

                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                    currentRow++;
                }

                worksheet.Columns().AdjustToContents();
                workbook.SaveAs(filePath);
            }
        }

        private string Convert2DArrayToString(byte[,] array)
        {
            var rows = array.GetLength(0);
            var cols = array.GetLength(1);
            var result = new StringBuilder();
            for (int i = 0; i < rows; i++)
            {
                var row = new List<string>();
                for (int j = 0; j < cols; j++)
                {
                    row.Add(array[i, j].ToString());
                }
                result.AppendLine(string.Join(", ", row));
            }
            return result.ToString();
        }

        private string Convert3DArrayToString(byte[,,] array)
        {
            var dim1 = array.GetLength(0);
            var dim2 = array.GetLength(1);
            var dim3 = array.GetLength(2);
            var result = new StringBuilder();
            for (int i = 0; i < dim1; i++)
            {
                for (int j = 0; j < dim2; j++)
                {
                    var slice = new List<string>();
                    for (int k = 0; k < dim3; k++)
                    {
                        slice.Add(array[i, j, k].ToString());
                    }
                    result.AppendLine(string.Join(", ", slice));
                }
            }
            return result.ToString();
        }

        public void ImportFromExcel(List<EepromMapping> mappings, string filePath)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet("EEPROM Mapping");
                int row = 2; // Start from the second row, the first row is the header

                while (!worksheet.Cell(row, 1).IsEmpty())
                {
                    // Read and parse the 'Address' cell as an integer
                    string addressCell = worksheet.Cell(row, 1).GetValue<string>();
                    if (!int.TryParse(addressCell, out int address))
                    {
                        Console.WriteLine($"Error: Cannot convert cell A{row} value '{addressCell}' to an integer.");
                        row++;
                        continue;
                    }

                    // Read 'Name' cell as a string
                    string name = worksheet.Cell(row, 2).GetValue<string>();

                    // Read 'Data Type' cell as a string
                    string dataTypeString = worksheet.Cell(row, 3).GetValue<string>();

                    // Read 'Values' cell as a string
                    string dataString = worksheet.Cell(row, 4).GetValue<string>();

                    // Find the corresponding mapping
                    var mapping = mappings.FirstOrDefault(m => m.Address == address && m.Name == name);
                    if (mapping != null)
                    {
                        Type dataType;
                        if (dataTypeString == "Byte")
                        {
                            dataType = typeof(byte);
                        }
                        else if (dataTypeString == "Byte[]")
                        {
                            dataType = typeof(byte[]);
                        }
                        else if (dataTypeString == "Byte[,]")
                        {
                            dataType = typeof(byte[,]);
                        }
                        else if (dataTypeString == "Byte[,,]")
                        {
                            dataType = typeof(byte[,,]);
                        }
                        else
                        {
                            dataType = null;
                            Console.WriteLine($"Error: Unrecognized data type '{dataTypeString}' at cell C{row}");
                        }


                        if (dataType != null)
                        {
                            // Parse and assign the value to the mapping
                            mapping.Value = ConvertStringToData(dataString, dataType);

                            ToArray(mapping);
                        }
                        else
                        {
                            Console.WriteLine($"Error: Unrecognized data type '{dataTypeString}' at cell C{row}");
                        }
                    }

                    row++;
                }
            }
        }

        private void ToArray(EepromMapping mapping)
        {
            if (mapping.EepromDataType == typeof(byte))
            {
                TempArray[mapping.Address] = (byte)mapping.Value;
            }
            else if (mapping.EepromDataType == typeof(byte[]))
            {
                byte[] arrayay = mapping.Value as byte[];
                int index = mapping.Address;

                for (int i = 0; i < mapping.Length; i++)
                {
                    TempArray[index++] = arrayay[i];
                }
            }
            else if (mapping.EepromDataType == typeof(byte[,]))
            {
                int index = mapping.Address;
                byte[,] matrix = mapping.Value as byte[,];

                for (int i = 0; i < mapping.RowCount; i++)
                {
                    for (int j = 0; j < mapping.ColumnCount; j++)
                    {
                        TempArray[index++] = matrix[i, j];
                    }
                }
            }
            else if (mapping.EepromDataType == typeof(byte[,,]))
            {
                int index = mapping.Address;
                byte[,,] kubus = mapping.Value as byte[,,];

                for (int i = 0; i < mapping.Depth; i++)
                {
                    for (int j = 0; j < mapping.RowCount; j++)
                    {
                        for (int k = 0; k < mapping.ColumnCount; k++)
                        {
                            TempArray[index++] = kubus[i, j, k];
                        }
                    }
                }
            }
        }

        private object ConvertStringToData(string dataString, Type dataType)
        {
            if (dataType == typeof(byte))
                return byte.Parse(dataString);
            else if (dataType == typeof(byte[]))
                return dataString.Split(',').Select(byte.Parse).ToArray();
            else if (dataType == typeof(byte[,]))
            {
                // Split the data string into rows by newline characters
                string[] rowsData = dataString.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);

                // Determine the number of columns based on the first row
                int cols = rowsData[0].Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length;

                // Create a 2D array with the determined number of rows and columns
                byte[,] data = new byte[rowsData.Length, cols];

                // Populate the 2D array
                for (int i = 0; i < rowsData.Length; i++)
                {
                    string[] rowElements = rowsData[i].Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    for (int j = 0; j < cols; j++)
                    {
                        data[i, j] = byte.Parse(rowElements[j].Trim()); // Trim to handle any spaces
                    }
                }
                return data;
            }
            else if (dataType == typeof(byte[,,]))
            {
                // Split the data string into layers based on a double newline separator
                string[] layers = dataString.Split(new[] { "\n\n" }, StringSplitOptions.RemoveEmptyEntries);
                int depth = 0;

                // Filter out and count valid data layers while skipping lines that are labels (e.g., "Layer X:", "X:")
                List<string[]> validLayers = new List<string[]>();
                foreach (string layer in layers)
                {
                    // Split the layer into rows and filter out any row that starts with "Layer" or a number followed by ":"
                    string[] rowsData = layer.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                             .Where(row => !row.Trim().StartsWith("Layer") && !System.Text.RegularExpressions.Regex.IsMatch(row.Trim(), @"^\d+:"))
                                             .ToArray();

                    if (rowsData.Length > 0)
                    {
                        validLayers.Add(rowsData);
                        depth++;
                    }
                }

                // Determine the number of rows and columns from the first valid layer
                int rows = validLayers[0].Length;
                int cols = validLayers[0][0].Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length;

                // Validate the number of elements
                int totalElements = validLayers.Sum(layer => layer.Sum(row => row.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length));
                string[] elements = dataString.Split(new[] { ' ', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                              .Where(e => !e.StartsWith("Layer") && !System.Text.RegularExpressions.Regex.IsMatch(e.Trim(), @"^\d+:"))
                                              .ToArray();

                if (elements.Length != totalElements)
                {
                    throw new ArgumentException("The number of elements in the data string does not match the calculated depth, rows, and columns.");
                }

                // Create and populate the 3D array
                byte[,,] data = new byte[depth, rows, cols];
                int index = 0;

                for (int d = 0; d < depth; d++)
                {
                    for (int i = 0; i < rows; i++)
                    {
                        string[] rowElements = validLayers[d][i].Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        for (int j = 0; j < cols; j++)
                        {
                            if (index >= elements.Length)
                            {
                                throw new IndexOutOfRangeException("Index exceeds the number of elements in the input string.");
                            }

                            if (byte.TryParse(elements[index].Trim(), out byte parsedValue))
                            {
                                data[d, i, j] = parsedValue;
                            }
                            else
                            {
                                Console.WriteLine($"Invalid byte value at depth {d}, row {i}, col {j}: '{elements[index]}'");
                                data[d, i, j] = 0; // Set a default value or handle as needed
                            }
                            index++;
                        }
                    }
                }
                return data;
            }
            return null;
        }
    }
}
