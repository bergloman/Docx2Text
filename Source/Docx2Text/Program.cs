using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Docx2Text {

    class Program {

        static void Main(string[] args) {

            if (args.Length < 2 || args.Length > 3 || args.Length == 2 && args[0] != "-d") {
                Console.WriteLine("Usage:");
                Console.WriteLine("Docx2Text.exe <input file> <output file>");
                Console.WriteLine("Docx2Text.exe -d <input directory> <output directory>");
                return;
            }

            if (args[0] == "-d") {

                var in_dir = args[1];
                var out_dir = args[2];

                if (!Directory.Exists(in_dir)) {
                    Console.WriteLine("Directory doesn't exist: " + in_dir);
                }
                if (!Directory.Exists(out_dir)) {
                    Console.WriteLine("Directory doesn't exist: " + out_dir);
                }

                foreach (var file in Directory.GetFiles(args[1], "*.docx")) {

                    var in_file = Path.GetFileName(file);
                    var in_file_full = Path.Combine(in_dir, in_file);
                    var out_file_full = Path.Combine(out_dir, in_file) + ".txt";
                    try {
                        Console.WriteLine("Extracting file: " + in_file);
                        using (var extractor = new DocxExtractor(in_file_full)) {
                            var s = extractor.ReadWordDocument();
                            File.WriteAllText(out_file_full, s, Encoding.UTF8);
                        }
                    } catch (Exception ex) {
                        Console.WriteLine("ERROR while extracting file: " + ex.ToString());
                    }
                }

            } else {

                if (!File.Exists(args[0])) {
                    Console.WriteLine("File doesn't exist: " + args[0]);
                }

                using (var extractor = new DocxExtractor(args[0])) {
                    Console.WriteLine("Extracting file: " + args[0]);
                    var s = extractor.ReadWordDocument();
                    File.WriteAllText(args[1], s, Encoding.UTF8);
                }
            }
        }
    }
}
