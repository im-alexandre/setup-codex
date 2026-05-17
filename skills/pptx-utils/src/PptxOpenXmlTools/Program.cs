using System.Text;
using PptxOpenXmlTools.Commands;

Console.OutputEncoding = Encoding.UTF8;

if (args.Length == 0 || args[0] is "--help" or "-h" or "help")
{
    Console.WriteLine("""
    pptx-utils

    Commands:
      inspect
      text-map
      replace-text
      validate

    Planned:
      duplicate-slide
      delete-slide
      reorder-slides
      replace-image
      render-preview
    """);
    return 0;
}

return args[0] switch
{
    "inspect" => InspectCommand.Run(args),
    "text-map" => TextMapCommand.Run(args),
    "replace-text" => ReplaceTextCommand.Run(args),
    "validate" => ValidateCommand.Run(args),
    _ => Unknown(args[0])
};

static int Unknown(string command)
{
    Console.Error.WriteLine($"Unknown command: {command}");
    return 2;
}
