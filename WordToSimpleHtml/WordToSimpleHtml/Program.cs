using System;
using System.IO;
using System.Text.RegularExpressions;

namespace WordToSimpleHtml
{
    internal static class Program
	{
        private static void Main(string[] args)
		{
		    if (args.Length != 2)
			{
				Console.WriteLine("Usage: word-file output-file");
				return;
			}

			var docxFile = Path.GetFullPath(args[0]);
			var htmlFile = Path.GetFullPath(args[1]);

            htmlFile = $@"{Path.GetDirectoryName(htmlFile)}\{Regex.Replace(Path.GetFileName(htmlFile), @"\s+", "-").ToLowerInvariant()}";
			var imageFilePrefix = Regex.Replace(Path.GetFileNameWithoutExtension(htmlFile), @"\W", string.Empty) + "-";

            Console.WriteLine($"Converting {docxFile} to {htmlFile} with images to {imageFilePrefix}");

			try
			{
			    string title;
			    new DocxToHtml(Console.WriteLine).Convert(docxFile, htmlFile, imageFilePrefix, out title);
			}
			catch (IOException ioe)
			{
				Console.WriteLine($"{Environment.NewLine}!!!{ioe.Message}");
			}

		}
	}
}
