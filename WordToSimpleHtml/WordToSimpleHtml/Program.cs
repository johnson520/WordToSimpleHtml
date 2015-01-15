using System;
using System.IO;
using System.Text.RegularExpressions;

namespace WordToSimpleHtml
{
	class Program
	{
		static void Main(string[] args)
		{
			string title;

			if (args.Length != 2)
			{
				Console.WriteLine("Usage: word-file output-file");
				return;
			}

			var docxFile = System.IO.Path.GetFullPath(args[0]);
			var htmlFile = System.IO.Path.GetFullPath(args[1]);
// ReSharper disable AssignNullToNotNullAttribute
			var imageFilePrefix = Regex.Replace(Path.GetFileNameWithoutExtension(htmlFile), @"\W", string.Empty) + "-";
// ReSharper restore AssignNullToNotNullAttribute

			Console.WriteLine("Converting {0} to {1} with images to {2}", docxFile, htmlFile, imageFilePrefix);

			try
			{
				new DocxToHtml(Console.WriteLine).Convert(docxFile, htmlFile, imageFilePrefix, out title);
			}
			catch (System.IO.IOException ioe)
			{
				Console.WriteLine(Environment.NewLine + "!!!" + ioe.Message);
			}


		}
	}
}
