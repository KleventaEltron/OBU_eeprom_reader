using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
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
            if (EepromList.doesUnmappingExist(address.Address))
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
                        Console.WriteLine($"Unknown data type: {address.EepromDataType}");
                        Console.WriteLine($"at address: {address.Address}");
                        break;
                }
            }
        }

        public void SetUnfoldedMappingValue(EepromMapping address)
        {
            ConvertValues convertValues = new ConvertValues();

            var UnFoldedAddress = EepromList.GetUnfoldedMappingByAddress(address.Address);

            switch (UnFoldedAddress.EepromDataType)
            {
                case Type t when t == typeof(byte):
                    EepromList.UpdateUnfoldedByteValue(UnFoldedAddress.Address, (byte)address.Value);
                    break;
                case Type t when t == typeof(sbyte):
                    sbyte tempSbyte = Convert.ToSByte(address.Value);
                    EepromList.UpdateUnfoldedSByteValue(UnFoldedAddress.Address, tempSbyte);
                    break;
                case Type t when t == typeof(char):
                    char tempChar = convertValues.ConvertByteToChar((byte)address.Value);
                    EepromList.UpdateUnfoldedCharValue(UnFoldedAddress.Address, tempChar);
                    break;
                case Type t when t == typeof(ushort):
                    uint tempUshort = 0;// = convertValues.ByteArrayToUInt16((byte[])address.Value);
                    if (address.Value is ushort ushortVal)
                    {
                        tempUshort = ushortVal;
                    }
                    else if (address.Value is byte[] byteArr)
                    {
                        tempUshort = convertValues.ByteArrayToUInt16(byteArr);
                    }
                    EepromList.UpdateUnfoldedUIntValue(UnFoldedAddress.Address, tempUshort);
                    break;
                case Type t when t == typeof(short):
                    //int tempShort = convertValues.ByteArrayToInt16((byte[])address.Value);
                    int tempShort = 0;
                    if (address.Value is short shortVal)
                    {
                        tempShort = shortVal;
                    }
                    else if (address.Value is byte[] byteArr)
                    {
                        tempShort = convertValues.ByteArrayToInt16(byteArr);
                    }
                    EepromList.UpdateUnfoldedIntValue(UnFoldedAddress.Address, tempShort);
                    break;
                case Type t when t == typeof(uint):
                    //uint tempUint = convertValues.ByteArrayToUInt32((byte[])address.Value);
                    uint tempUint = 0;
                    if (address.Value is uint uintVal)
                    {
                        tempUint = uintVal;
                    }
                    else if (address.Value is byte[] byteArr)
                    {
                        tempUint = convertValues.ByteArrayToUInt32(byteArr);
                    }
                    EepromList.UpdateUnfoldedUIntValue(UnFoldedAddress.Address, tempUint);
                    break;
                case Type t when t == typeof(byte[]):

                    if (address.EepromDataType == typeof(byte[,]))
                    {
                        List<byte[]> tempByteList = convertValues.ConvertMatrixToByteList((byte[,])address.Value);

                        int rows = ((byte[,])address.Value).GetLength(0);
                        int columns = ((byte[,])address.Value).GetLength(1);

                        for (int i = 0; i < rows; i++)
                        {
                            EepromList.UpdateUnfoldedByteArrayValue(UnFoldedAddress.Address + (i * columns), tempByteList[i]);
                        }
                        break;
                    }
                    EepromList.UpdateUnfoldedByteArrayValue(UnFoldedAddress.Address, (byte[])address.Value);
                    break;
                case Type t when t == typeof(sbyte[]):
                    sbyte[] tempSbyteArray = convertValues.ConvertByteArrayToSbyteArray((byte[])address.Value);
                    EepromList.UpdateUnfoldedSbyteArrayValue(address.Address, tempSbyteArray);
                    break;
                case Type t when t == typeof(string):
                    List<string> tempStringList;

                    if (address.Value.GetType() == typeof(byte[]))
                    {
                        string tempString = convertValues.ByteArrayToString((byte[])address.Value);
                        EepromList.UpdateUnfoldedStringValue(UnFoldedAddress.Address, tempString);
                    }
                    else if (address.Value.GetType() == typeof(byte[,]))
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

                    Console.WriteLine($"Unknown data type with {UnFoldedAddress.EepromDataType} at address {UnFoldedAddress.Address}.");
                    break;
            }
        }
    }
}

