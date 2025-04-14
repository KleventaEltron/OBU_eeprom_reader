using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EepromReader.Scripts
{
    internal class ConvertValues
    {
        // Convert a byte to a char
        public char ConvertByteToChar(byte Byte)
        {
            return Convert.ToChar(Byte);
        }

        // Convert a byte array to a string
        public string ByteArrayToString(byte[] Byte)
        {
            return System.Text.Encoding.ASCII.GetString(Byte);
        }

        // Convert byte array to Uint
        public uint ByteArrayToUInt16(byte[] byteArray)
        {
            return BitConverter.ToUInt16(byteArray, 0);
        }

        public int ByteArrayToInt16(byte[] byteArray)
        {
            return BitConverter.ToInt16(byteArray, 0);
        }

        public uint ByteArrayToUInt32(byte[] byteArray)
        {
            return BitConverter.ToUInt32(byteArray, 0);
        }

        public int ByteArrayToInt32(byte[] byteArray)
        {
            return BitConverter.ToInt32(byteArray, 0);
        }

        public List<byte[]> ConvertMatrixToByteList(byte[,] matrix)
        {
            var list = new List<byte[]>();
            int rows = matrix.GetLength(0);
            int columns = matrix.GetLength(1);
            for (int i = 0; i < rows; i++)
            {
                byte[] byteArray = new byte[columns];
                for (int j = 0; j < columns; j++)
                {
                    byteArray[j] = matrix[i, j];
                }
                list.Add(byteArray);
            }
            return list;
        }

        public sbyte[] ConvertByteArrayToSbyteArray(byte[] array)
        {
            sbyte[] tempSbyteArray = new sbyte[array.Length];

            for (int i = 0; i < array.Length; i++)
            {
                tempSbyteArray[i] = Convert.ToSByte(array[i]);
            }
            return tempSbyteArray;
        }

        // Convert a matrix to a list of strings
        public List<string> ConvertMatrixToStringList(byte[,] matrix)
        {
            var list = new List<string>();

            int rows = matrix.GetLength(0);
            int columns = matrix.GetLength(1);

            for (int i = 0; i < rows; i++)
            {
                byte[] byteArray = new byte[columns];

                for (int j = 0; j < columns; j++)
                {
                    byteArray[j] = matrix[i, j];
                }
                list.Add(ByteArrayToString(byteArray));
            }

            return list;
        }

        // Convert a kubus to a list of strings
        public List<string> ConvertKubusToStringList(byte[,,] kubus)
        {
            int depth = kubus.GetLength(0);
            int rows = kubus.GetLength(1);
            int columns = kubus.GetLength(2);
            
            var list = new List<string>();

            for(int i = 0; i < depth; i++)
            {
                for (int j = 0; j < rows; j++)
                {
                    byte[] byteArray = new byte[columns];

                    for (int k = 0; k < columns; k++)
                    {
                        byteArray[k] = kubus[i, j, k];
                    }
                    list.Add(ByteArrayToString(byteArray));
                }
            }

            return list;
        }
    }
}
