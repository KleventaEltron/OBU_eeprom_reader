using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EepromReader.Scripts
{
    internal class FolderNavigation
    {
        public string FileOrFolderNavigation(string mode)
        {
            if (mode == "OpenFile")
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
                        return selectedFilePath;
                    }
                    else
                    {
                        Console.WriteLine("No file selected.");
                        return null;
                    }
                }
            }
            else if (mode == "SaveFile")
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    // Set the filter to XML files only
                    saveFileDialog.Filter = "Excel Files|*.xlsx;*.xlsm|All Files|*.*";
                    saveFileDialog.Title = "Save XML File";

                    // Show the dialog and check if the user selects a location
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string saveFilePath = saveFileDialog.FileName;
                        Console.WriteLine("File will be saved at: " + saveFilePath);
                        return saveFilePath;
                    }
                    else
                    {
                        Console.WriteLine("Save operation canceled.");
                        return null;
                    }
                }
            }
            else if (mode == "SelectFolder")
            {
                using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
                {
                    folderBrowserDialog.Description = "Select a Folder";

                    // Show the dialog and check if the user selects a folder
                    if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                    {
                        string selectedFolderPath = folderBrowserDialog.SelectedPath;
                        Console.WriteLine("Selected folder: " + selectedFolderPath);
                        return selectedFolderPath;
                    }
                    else
                    {
                        Console.WriteLine("No folder selected.");
                        return null;
                    }
                }
            }
            else
            {
                throw new ArgumentException("Invalid mode specified. Use 'OpenFile', 'SaveFile', or 'SelectFolder'.");
            }
        }

    }
}
