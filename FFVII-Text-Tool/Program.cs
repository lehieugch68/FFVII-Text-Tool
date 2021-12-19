using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace FFVII_Text_Tool
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Title = "Final Fantasy VII Remake Text Tool by LeHieu - viethoagame.com";

            if (args.Length > 0)
            {
                switch (args[0])
                {
                    case "-e":
                        if (args.Length >= 3)
                        {
                            TextExtractor.ExportXLSX(args[1], args[2]);
                        }
                        else
                        {
                            Console.WriteLine("-> Type \"-e [Input Folder] [Output XLSX File]\" to extract the folder.");
                        }
                        break;
                    case "-i":
                        if (args.Length >= 4)
                        {
                            TextExtractor.ImportXLSX(args[1], args[2], args[3]);
                        }
                        else
                        {
                            Console.WriteLine("-> Type \"-i [XLSX File] [Original Folder] [Output Folder]\" to re-import the folder.");
                        }
                        break;
                    default:
                        Help();
                        break;
                }
            }
            else
            {
                Help();
            }
        }
        static void Help()
        {
            Console.WriteLine("Usage:\n-> Type \"-e [Input Folder] [Output XLSX File]\" to extract the folder.\n-> Type \"-i [XLSX File] [Original Folder] [Output Folder]\" to re-import the folder.");
        }
    }
}
