using CommandLine;

public class ArgsOptions
{
    [Option('s', "save", Required = true, HelpText = "Save directory path")]
    public string? SavePath { get; set; }

    [Option('f', "file", Required = true, HelpText = "Mbox file path")]
    public string? MboxFilePath { get; set; }

    [Option('o', "overwrite", Required = false, HelpText = "If overwrite files with same name")]
    public bool Overwrite { get; set; }

    public override string ToString() => $"Mbox file locations {MboxFilePath}, Save directory {SavePath}, If overwrite files {Overwrite}";
}