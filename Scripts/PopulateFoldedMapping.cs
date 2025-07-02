using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EepromReader.Scripts
{
    internal class PopulateFoldedMapping
    {
        public object ConvertToBytes(EepromMapping address)
        {
            var type = address.EepromDataType;
            var value = address.Value;

            if (value == null || type == null)
                throw new ArgumentNullException("EepromMapping must have non-null Value and EepromDataType.");

            if (type == typeof(byte))
                return new byte[] { (byte)value };

            if (type == typeof(sbyte))
                return new byte[] { (byte)(sbyte)value };

            if (type == typeof(char))
                return new byte[] { (byte)(char)value };

            if (type == typeof(ushort))
                return BitConverter.GetBytes((ushort)value);

            if (type == typeof(short))
                return BitConverter.GetBytes((short)value);

            if (type == typeof(uint))
                return BitConverter.GetBytes((uint)value);

            if (type == typeof(int))
                return BitConverter.GetBytes((int)value);

            if (type == typeof(byte[]))
                return (byte[])value;

            if (type == typeof(sbyte[]))
                return ((sbyte[])value).Select(v => (byte)v).ToArray();

            if (type == typeof(string))
                return Encoding.ASCII.GetBytes((string)value);

            throw new NotSupportedException($"Unsupported conversion type: {type}");
        }

        public void ConvertToBytesAndUpdate(EepromMapping address)
        {
            //byte[] temp = ConvertToBytes(address);

        }


        public byte[,] ConvertToByteMatrix(List<EepromMapping> mappingList)
        {
            if (mappingList == null || mappingList.Count == 0)
                throw new ArgumentException("List is null or empty");

            List<byte[]> rows = new List<byte[]>();
            foreach(var mapping in mappingList)
            {
                byte[] rowBytes;

                if (mapping.Value is byte[] b)
                {
                    rowBytes = b;
                }
                else if (mapping.Value is string s)
                {
                    rowBytes = Encoding.ASCII.GetBytes(s);
                }
                else
                {
                    throw new InvalidCastException($"Unsupported value type at address {mapping.Address}: {mapping.Value?.GetType()}");
                }

                rows.Add(rowBytes);
            }

            int rowCount = mappingList.Count;
            int columnCount = mappingList[0].Length;

            byte[,] matrix = new byte[rowCount, columnCount];


            for (int row = 0; row < rowCount; row++)
            {
                //byte[] rowData = (byte[])mappingList[row].Value;
                for (int col = 0; col < columnCount; col++)
                {
                    matrix[row, col] = rows[row][col];
                }
            }

            return matrix;
        }
    }
}
