using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;

namespace Reusable.IO
{
    class Selection
    {
        /// <summary>
        /// Resuable method to return provide a folder selection window
        /// Returns a path as a string for a single folder
        /// </summary>
        /// <param name="defaultPath"></param>
        /// <returns></returns>
        public static string SelectFolder(string defaultPath)
        {
            string currentPath = defaultPath;
            string newPath = null;
            string returnPath = null;

            currentPath = CheckDirectoryPath(currentPath);

            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.SelectedPath = currentPath;

            DialogResult result = dlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                newPath = dlg.SelectedPath;
            }

            if (Directory.Exists(newPath))
            {
                returnPath = newPath;
            }
            return returnPath;
        }

        /// <summary>
        /// Resuable method to return provide a files selection window
        /// Returns a collection of files or a single file in the list
        /// depending on isMultiSelect option
        /// </summary>
        /// <param name="defaultPath"></param>
        /// <returns></returns>
        public static string[] SelectFile(string defaultPath, bool isSelectMultiple)
        {
            string currentPath = defaultPath;
            string[] returnPath = null;

            currentPath = CheckDirectoryPath(currentPath);

            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Multiselect = isSelectMultiple;
            dlg.InitialDirectory = currentPath;

            DialogResult result = dlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                returnPath = dlg.FileNames;
            }
            return returnPath;
        }

        private static string CheckDirectoryPath(string currentPath)
        {
            if (!Directory.Exists(currentPath))
            {
                //set artificial path
                currentPath = "";
            }
            return currentPath;
        }

    }
}
