using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using McMaster.Extensions.CommandLineUtils;

namespace docxmerge
{
  class Program
  {
    public static int Main(string[] args)
    {
      // https://docs.microsoft.com/ja-jp/dotnet/api/documentformat.openxml.wordprocessing.altchunk?view=openxml-2.8.1

      var app = new CommandLineApplication(throwOnUnexpectedArg: false);
      app.Name = nameof(docxmerge);
      app.Description = "Mearge .docx files. ";
      app.HelpOption("-h|--help");
      app.ExtendedHelpText = @"Merged files should be the end of argument like: 
      docxmerge [options] file1 file2
      ";
      var optionTemplateFile = app.Option(
        template: "-t|--template",
        description: "Template filename. merged files will be appended after the contents of template file. ",
        optionType: CommandOptionType.SingleValue
      );

      var optionOutFile = app.Option(
        template: "-o|--output",
        description: "Output filename",
        optionType: CommandOptionType.SingleValue
      );

      app.OnExecute(() =>
      {
        var templateFile = optionTemplateFile.Value();
        var outFile = optionOutFile.Value();
        // Because position of last paragraph is not updated while looping
        var mergedFiles = app.RemainingArguments;
        mergedFiles.Reverse();
        File.Delete(outFile);
        File.Copy(templateFile, outFile);

        using (WordprocessingDocument myDoc = WordprocessingDocument.Open(outFile, true))
        {
          MainDocumentPart mainPart = myDoc.MainDocumentPart;
          foreach (var f in mergedFiles.Select((Value, Index) => new { Value, Index }))
          {
            string altChunkId = $"AltChunkId{f.Index}";
            AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(
              AlternativeFormatImportPartType.WordprocessingML, altChunkId
            );
            using (FileStream fileStream = File.Open(f.Value, FileMode.Open))
            {
              chunk.FeedData(fileStream);
            }
            AltChunk altChunk = new AltChunk();
            altChunk.Id = altChunkId;
            mainPart.Document.Body.InsertAfter(altChunk, mainPart.Document.Body.Elements<Paragraph>().Last());

          }
          mainPart.Document.Save();
        }
      });
      return app.Execute(args);
    }
  }
}
