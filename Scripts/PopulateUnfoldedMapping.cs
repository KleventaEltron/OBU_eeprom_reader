using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using EepromReader.Scripts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EepromReader.Scripts
{
    internal class PopulateUnfoldedMapping
    {
        public void FindUnfoldedMapping(EepromMapping address)
        {
            if(EepromList.doesUnmappingExist(address.Address))
            {
                switch (address.EepromDataType)
                {
                    case Type t when t == typeof(byte):
                        SetUnfoldedMappingValue(address);
                        break;
                    case Type t when t == typeof(byte[]):
                        SetUnfoldedMappingValue(address);   

                        break;
                    case Type t when t == typeof(byte[,]):
                        SetUnfoldedMappingValue(address);

                        break;
                    case Type t when t == typeof(byte[,,]):
                        SetUnfoldedMappingValue(address);
                        break;
                    default:
                        Console.WriteLine("Unknown data type.");
                        break;
                }
            }
        }

        private void SetUnfoldedMappingValue(EepromMapping address)
        {
            ConvertValues convertValues = new ConvertValues();

            var UnFoldedAddress = EepromList.GetUnfoldedMappingByAddress(address.Address);

            switch (UnFoldedAddress.EepromDataType)
            {
                case Type t when t == typeof(char):
                    char tempChar = convertValues.ConvertByteToChar((byte)address.Value);
                    EepromList.UpdateUnfoldedCharValue(UnFoldedAddress.Address, tempChar);
                    break;
                case Type t when t == typeof(int):
                    int tempInt = convertValues.ByteArrayToInt((byte[])address.Value);
                    EepromList.UpdateUnfoldedIntValue(UnFoldedAddress.Address, tempInt);
                    break;
                case Type t when t == typeof(string):
                    List<string> tempStringList;

                    if (address.Value.GetType() == typeof(byte[,]))
                    {
                        tempStringList = convertValues.ConvertMatrixToStringList((byte[,])address.Value);

                        int rows = ((byte[,])address.Value).GetLength(0);
                        int columns = ((byte[,])address.Value).GetLength(1);

                        for (int i = 0; i < rows; i++)
                        {
                            EepromList.UpdateUnfoldedStringValue(UnFoldedAddress.Address + (i * columns), tempStringList[i]);
                        }
                    }
                    else if (address.Value.GetType() == typeof(byte[,,]))
                    {
                        tempStringList = convertValues.ConvertKubusToStringList((byte[,,])address.Value);

                        int depth = ((byte[,,])address.Value).GetLength(0);
                        int rows = ((byte[,,])address.Value).GetLength(1);
                        int columns = ((byte[,,])address.Value).GetLength(2);

                        for (int i = 0; i < rows * depth; i++)
                        {
                            EepromList.UpdateUnfoldedStringValue(UnFoldedAddress.Address + (i * columns), tempStringList[i]);
                        }
                    }

                    break;
                default:
                    Console.WriteLine("Unknown data type.");
                    break;
            }
        }
    }
}

