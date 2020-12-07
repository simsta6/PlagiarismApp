using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using Xceed.Words.NET;
using Xceed.Document.NET;

namespace testDocX
{
    class Program
    {
        const string allFilesDir = @"..\..\..\data\all files";
        const string resultsDir = @"..\..\..\data\results\";

        static void Main(string[] args)
        {
            string[] exampleFiles = {
            @"..\..\..\data\example files\0 test file.docx",
            @"..\..\..\data\example files\1 test file.docx",
            @"..\..\..\data\example files\2 test file.docx",
            @"..\..\..\data\example files\3 test file.docx",
            @"..\..\..\data\example files\4 test file.docx",
            @"..\..\..\data\example files\5 test file.docx",
            @"..\..\..\data\example files\6 test file.docx",
            @"..\..\..\data\example files\7 test file.docx",
            @"..\..\..\data\example files\8 test file.docx",
            @"..\..\..\data\example files\9 test file.docx",
            };

            DirectoryInfo di = new DirectoryInfo(allFilesDir);

            for (int i = 0; i < exampleFiles.Length; i++)
            {
                Console.WriteLine("Dirbama su " + GetFileName(exampleFiles[i]));

                List<string> chunks = SplitIntoChuncks(exampleFiles[i]);

                Console.WriteLine("Paprastas metodas užtruko " + NonParallel_CompareFiles(di, chunks, exampleFiles[i]) + " miliseconds.");
                Console.WriteLine("Išlygiagretintas metodas užtruko " + Parallel_CompareFiles(di, chunks, exampleFiles[i]) + " miliseconds.");
                //Console.Write(NonParallel_CompareFiles(di, chunks, exampleFiles[i]) + " ");
                //Console.Write(Parallel_CompareFiles(di, chunks, exampleFiles[i]));

                Console.WriteLine();
            }

            Console.WriteLine("done");

            Console.ReadKey();
        }

        private static long NonParallel_CompareFiles(DirectoryInfo di, List<string> chunks, string orgFilename)
        {
            Stopwatch sw = new Stopwatch();

            sw.Start();

            Console.WriteLine($"Duomenų kiekis: {chunks.Count()}");

            foreach (FileInfo fi in di.GetFiles())
            {
                CompareFiles(chunks, orgFilename, fi.FullName);
            }

            sw.Stop();

            return sw.ElapsedMilliseconds;
        }

        private static long Parallel_CompareFiles(DirectoryInfo di, List<string> chunks, string orgFilename)
        {
            Stopwatch sw = new Stopwatch();

            sw.Start();

            di.GetFiles().AsParallel().WithDegreeOfParallelism(8).ForAll(x => CompareFiles(chunks, orgFilename, x.FullName));

            sw.Stop();

            return sw.ElapsedMilliseconds;
        }

        static void CompareFiles(List<string> chunks, string orgFilename, 
            string filename)
        {
            List<string> chunksToRemove = new List<string>();

            // Plagiato žymėjimas
            var rb = new Formatting();
            rb.Highlight = Highlight.red;
            //

            using (DocX document = DocX.Load(filename))
            {
                for (int i = 0; i < chunks.Count; i++)
                {
                    var a = document.FindAll(chunks[i]);
                    if (a.Count > 0)
                    {
                        chunksToRemove.Add(chunks[i]);
                    }
                    document.ReplaceText(chunks[i], chunks[i], newFormatting: rb);
                }

                int resultPrecent = chunksToRemove.Count() * 100 / chunks.Count();

                if (resultPrecent > 5)
                {
                    Console.WriteLine($"Failas ,,{GetFileName(orgFilename)}\" " +
                        $"sutapo su ,,{GetFileName(filename)}\" {resultPrecent}%");
                    document.SaveAs(resultsDir + "_resultsOf_" +
                        GetFileName(filename) + "_with_" + GetFileName(orgFilename));
                }
            }
        }

        static string GetFileName(string filename)
        {
            return filename.Split('\\').ToList().Last();

        }

        static List<string> SplitIntoChuncks(string filename)
        {
            List<string> groups = new List<string>();

            using (DocX document = DocX.Load(filename))
            {
                for (int i = 0; i < document.Paragraphs.Count; i++)
                {
                    foreach (var item in document.Paragraphs[i].Text.Split(
                        new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        if (item.Length > 35)
                        {
                            var wordslol = item.Split(new string[] { " " }, 
                                StringSplitOptions.RemoveEmptyEntries);
                            int count = 0;
                            string chunk = "";
                            for (int j = 0; j < wordslol.Length; j++)
                            {
                                count++;

                                if (j != wordslol.Length - 1)
                                {
                                    chunk += wordslol[j] + " ";
                                } else
                                {
                                    chunk += wordslol[j];
                                }

                                if (count == 5)
                                {
                                    groups.Add(chunk);
                                    count = 0;
                                    chunk = "";
                                }
                            }
                            if (count != 0)
                                groups.Add(chunk);
                            
                        } else {
                            groups.Add(item);
                        }
                    }
                }
            }

            return groups;
        }
    }
}