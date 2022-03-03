using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net;

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
                            string xlsx = args[1];
                            if (xlsx.StartsWith("https://") && Uri.IsWellFormedUriString(xlsx, UriKind.RelativeOrAbsolute))
                            {
                                using (WebClient wc = new WebClient())
                                {
                                    xlsx = Path.Combine(AppContext.BaseDirectory, "Translation.xlsx");
                                    wc.DownloadFile(new Uri(args[1]), xlsx);
                                }
                            }
                            TextExtractor.ImportXLSX(xlsx, args[2], args[3]);
                        }
                        else
                        {
                            Console.WriteLine("-> Type \"-i [XLSX Path/URL] [Original Folder] [Output Folder]\" to re-import the folder.");
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
            Console.WriteLine("Usage:\n-> Type \"-e [Input Folder] [Output XLSX File]\" to extract the folder.\n-> Type \"-i [XLSX Path/URL] [Original Folder] [Output Folder]\" to re-import the folder.");
        }
    }
}
