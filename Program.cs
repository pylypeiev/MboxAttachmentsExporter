using CommandLine;
using MimeKit;
using System.Diagnostics;
using System.Runtime.InteropServices;

int totalAttachments = 0;

Parser.Default.ParseArguments<ArgsOptions>(args)
              .WithParsed(MainAction);

void MainAction(ArgsOptions options)
{
    if (!CheckIfArgsCorrect(options)) return;

    GetAttachmentsFromMboxFile(options);
    Console.WriteLine($"{totalAttachments} attachments saved successfully");

    if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
    {
        Process.Start("explorer.exe", options.SavePath);
    }
}

bool CheckIfArgsCorrect(ArgsOptions options)
{
    try
    {
        if (!File.Exists(options.MboxFilePath))
        {
            Console.WriteLine($"Please check mbox file path [{options.MboxFilePath}], file is not exist");
            return false;
        }
        else if (!options.MboxFilePath.EndsWith(".mbox", StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine($"Please check mbox file path [{options.MboxFilePath}], program works only with mbox files");
            return false;
        }

        if (!Directory.Exists(options.SavePath))
        {
            Console.WriteLine($"Please check save path directory [{options.SavePath}], directory is not exists");
            return false;
        }
        return true;
    }
    catch (Exception e)
    {

        Console.WriteLine(e.Message);
        return false;
    }
}

void GetAttachmentsFromMboxFile(ArgsOptions args)
{
    try
    {
        using var stream = File.OpenRead(args.MboxFilePath);
        var parser = new MimeParser(stream, MimeFormat.Mbox);

        while (!parser.IsEndOfStream)
        {
            var message = parser.ParseMessage();

            if (!message.Attachments.Any()) continue;

            foreach (var attachment in message.Attachments)
            {
                SaveAttachment(attachment, args);
            }
        }
    }
    catch (Exception e)
    {
        Console.WriteLine(e.Message);
    }
}

void SaveAttachment(MimeEntity attachment, ArgsOptions args)
{
    var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;

    if (string.IsNullOrEmpty(fileName))
    {
        return;
    }

    fileName = fileName.Replace('\\', ' ').Replace('/', ' ');

    if (!args.Overwrite && File.Exists(Path.Combine(args.SavePath, fileName)))
        fileName = Guid.NewGuid().ToString().Substring(1, 6) + fileName;

    using var file = File.Create(Path.Combine(args.SavePath, fileName));

    if (attachment is MessagePart messagePart)
    {
        messagePart.Message.WriteTo(file);
    }
    else if (attachment is MimePart mimePart)
    {
        mimePart.Content.DecodeTo(file);
    }
    Console.WriteLine($"{fileName} saved");
    totalAttachments++;
}
