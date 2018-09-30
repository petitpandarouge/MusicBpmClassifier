using System;
using System.Configuration;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MusicBpmClassifier
{
    /// <summary>
    /// Class defining the main application program.
    /// </summary>
    public class Program
    {
        #region Methods

        /// <summary>
        /// Program entry point.
        /// </summary>
        /// <param name="pArgs">The command line arguments.</param>
        static void Main(string[] pArgs)
        {
            string lInputPath = ConfigurationManager.AppSettings.Get("InputPath");
            string lOutputPath = ConfigurationManager.AppSettings.Get("OutputPath");

            Console.Title = "MusicBpmClassifier 1.0";
            Console.WriteLine("MusicBpmClassifier 1.0 - Copyright ©  2018\n");
            Console.WriteLine("Classifying the directory :");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(" " + lInputPath);
            Console.ResetColor();
            Console.WriteLine("Into the destination directory :");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(" " + lOutputPath);
            Console.ResetColor();
            Console.WriteLine("\nPlease press the enter key to classify...");
            Console.ReadLine();

            string[] lFilePathes = Directory.GetFiles(lInputPath, "*.*", SearchOption.AllDirectories);
            foreach (string lFilePath in lFilePathes)
            {
                if (Program.IsMusicFile(lFilePath))
                {
                    Console.ForegroundColor = ConsoleColor.DarkGray;
                    string lBpmDirName = Program.GetBpmAsString(lFilePath);
                    if (string.IsNullOrEmpty(lBpmDirName))
                    {
                        Console.ForegroundColor = ConsoleColor.DarkRed;
                        lBpmDirName = "Unknown";
                    }

                    string lBpmDirPath = Path.Combine(lOutputPath, lBpmDirName);
                    if (Directory.Exists(lBpmDirPath) == false)
                    {
                        Directory.CreateDirectory(lBpmDirPath);
                    }

                    Console.WriteLine(lFilePath + " -> " + lBpmDirName);
                    FileInfo lFileInfo = new FileInfo(lFilePath);
                    string lCopiedFilePath = Path.Combine(lBpmDirPath, lFileInfo.Name);
                    File.Copy(lFilePath, lCopiedFilePath, true);
                    Console.ResetColor();
                }
            }

            Console.WriteLine("\nClassification done!");
            Console.WriteLine("Please press the enter key to quit...");
            Console.ReadLine();
        }


        private static string GetBpmAsString(string pFilePath)
        {
            // Instantiate the Application object.
            dynamic lShell = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));

            // Get the folder and the child.
            var lFolder = lShell.NameSpace(Path.GetDirectoryName(pFilePath));
            var lItem = lFolder.ParseName(Path.GetFileName(pFilePath));

            // Get the item's property by it's canonical name. Doc says it's a string.
            return lItem.ExtendedProperty("System.Music.BeatsPerMinute");
        }

        private static bool IsMusicFile(string pFilePath)
        {
            // Instantiate the Application object.
            dynamic lShell = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));

            // Get the folder and the child.
            var lFolder = lShell.NameSpace(Path.GetDirectoryName(pFilePath));
            var lItem = lFolder.ParseName(Path.GetFileName(pFilePath));

            // Get the item's property by it's canonical name. Doc says it's a string.
            string lFormat = lItem.ExtendedProperty("System.Audio.Format");
            return string.IsNullOrEmpty(lFormat) == false;
        }

        #endregion // Methods.
    }
}
