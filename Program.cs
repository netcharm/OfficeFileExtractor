using System;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Net.WebSockets;
using System.Text;

using Mono.Options;
using OpenMcdf;

public class Program
{
    public enum OfficeFileFormat { Unknown, ZIP, MCDF };

    private static string AppName = Path.GetFileName(AppDomain.CurrentDomain.FriendlyName);

    private static Encoding MBCS = Encoding.Default;

    private static bool overwrite = false;
    private static bool extract_all = true;
    private static bool extract_media = false;
    private static bool extract_embed = false;
    private static string output_folder = string.Empty;
    private static CultureInfo culture = CultureInfo.DefaultThreadCurrentUICulture ?? CultureInfo.DefaultThreadCurrentCulture ?? CultureInfo.CurrentUICulture ?? CultureInfo.CurrentCulture;
    private static int codepage = 1250;

    private static string office_type = string.Empty;
    private static DirectoryInfo? output_folder_info = null;

    private static OptionSet Options { get; set; } = new OptionSet()
    {
        { "h|?|help", "Show Help", v => { ShowHelp(); } },
        { "y|overwrite", "Overwrite exists file", v => { overwrite = v != null; } },
        { "a|all", "Extracting media and embeddings", v => { extract_all = v != null; } },
        { "m|media", "Extracting media ", v => { extract_media = v != null; } },
        { "e|object", "Extracting embeddings", v => { extract_embed = v != null; } },
        { "o|output=", "Output Folder, default is document file name", v => { output_folder = v; } },
        { "c|culture=", "Culture of file name, default is system culture ", v => { culture = CultureInfo.GetCultureInfo(v); } },
        { "p|codepage=", "Codepage of file name, default is system codepage ", v => { int.TryParse(v, out codepage); } },
    };

    public static void Main(string[] args)
    {
        if (args.Length <= 0) { ShowHelp(); return; }

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        codepage = culture.TextInfo.ANSICodePage;

        //var args = Environment.GetCommandLineArgs().Skip(1).ToList();
        var opts = Options.Parse(args);

        MBCS = Encoding.GetEncoding(codepage);

        if (!extract_media && !extract_embed) extract_all = true;

        var doc = opts.FirstOrDefault();
        if (string.IsNullOrEmpty(doc) || !File.Exists(doc)) { ShowHelp(); return; }

        try
        {
            Export(doc);
        }
        catch(Exception ex) { Console.WriteLine(ex.Message + Environment.NewLine + ex.StackTrace); };
    }

    public static void ShowHelp()
    {
        if (Options is OptionSet)
        {
            using (var sw = new StringWriter())
            {
                Console.WriteLine($"Usage: {AppName} [OPTIONS]+ <Office Document file>");
                Console.WriteLine("Options:");
                Options.WriteOptionDescriptions(sw);
                var result = sw.ToString();
                Console.WriteLine(result);
            }
        }
    }

    private static bool ExportOle10Native(byte[] data, string targetFolder)
    {
        var result = false;
        Func<Stream, int, string> GetString = (stream, count) =>
        {
            var result = string.Empty;
            if (stream is Stream && count > 0)
            {
                var pos = stream.Position;
                var buf = new byte[count];
                stream.Read(buf, 0, (int)Math.Max(0, Math.Min(stream.Length - pos, count)));
                var bytes = buf.TakeWhile(b => b != 0).ToArray();
                if (bytes.Length > 0)
                {
                    result = MBCS.GetString(bytes).Trim();
                    stream.Seek(pos + bytes.Length, SeekOrigin.Begin);
                    //Console.WriteLine($"{pos:08N} : {result}");
                }
            }
            return(result);
        };
        Func<Stream, int> GetInt32 = stream =>
        {
            var result = 0;
            if (stream is Stream)
            {
                var pos = stream.Position;
                var buf = new byte[4];
                stream.Read(buf, 0, (int)Math.Max(0, Math.Min(stream.Length - pos, 4)));
                result = BitConverter.ToInt32(buf, 0);
            }
            return(result);
        };

        if (data?.Length > 0 && !string.IsNullOrEmpty(targetFolder) && Directory.Exists(targetFolder))
        {
            using (var olems = new MemoryStream(data))
            {
                //Console.WriteLine(olems.Length);
                var pos = olems.Position;

                olems.Seek(pos + 6, SeekOrigin.Begin);
                var filename = GetString(olems, 260);
                pos = olems.Position;

                olems.Seek(pos + 1, SeekOrigin.Begin);
                var src_file = GetString(olems, 260);
                pos = olems.Position;

                olems.Seek(pos + 4 + 1, SeekOrigin.Begin);
                var temp_len = GetInt32(olems);
                var temp_file = GetString(olems, temp_len);
                pos = olems.Position;

                olems.Seek(pos + 1, SeekOrigin.Begin);
                var cl = GetInt32(olems);
                var cnt = new byte[cl];
                olems.Read(cnt, 0, cl);

                var targetFile = Path.Combine(targetFolder, filename);
                Console.WriteLine("File  : " + targetFile);
                if (overwrite || !File.Exists(targetFile))
                {
                    File.WriteAllBytes(targetFile, cnt);
                    result = true;
                }
            }
        }
        return (result);
    }

    public static OfficeFileFormat DetectOfficeFileFormat(string officeFile)
    {
        var result = OfficeFileFormat.Unknown;
        if (!string.IsNullOrEmpty(officeFile) && File.Exists(officeFile))
        {
            using (var fs = File.OpenRead(officeFile))
            {
                var buf = new byte[4096];
                var ret = fs.Read(buf, 0, buf.Length);
                //fs.ReadExactly(buf, 0, buf.Length);
                if (ret > 0)
                {
                    if (buf[0] == 'P' && buf[1] == 'K' && buf[2] == 0x03 && buf[3] == 0x04)
                        result = OfficeFileFormat.ZIP;
                    else if (buf[0] == 0xD0 && buf[1] == 0xCF && buf[2] == 0x11 && buf[3] == 0xE0 && buf[4] == 0xA1 && buf[5] == 0xB1 && buf[6] == 0x1A && buf[7] == 0xE1)
                        result = OfficeFileFormat.MCDF;
                }
            }
        }
        return (result);
    }

    public static void Export(string officeFile)
    {
        if (string.IsNullOrEmpty(output_folder)) output_folder = Path.ChangeExtension(officeFile, "_files").Replace("._files", "_files");
        output_folder_info = Directory.CreateDirectory(output_folder);

        var officeExt = Path.GetExtension(officeFile).ToLower();
        if (officeExt.StartsWith(".doc")) office_type = "word";
        else if (officeExt.StartsWith(".xls")) office_type = "excel";
        else if (officeExt.StartsWith(".ppt")) office_type = "ppt";

        var file_type = DetectOfficeFileFormat(officeFile);
        if (file_type == OfficeFileFormat.ZIP) ExportX(officeFile);
        else if (file_type == OfficeFileFormat.MCDF) ExportC(officeFile);
    }

    public static void ExportC(string officeFile)
    {
        if (output_folder_info?.Exists ?? false)
        {
            using (var cf = new CompoundFile(officeFile))
            {
                var action = false;
                if (extract_all || extract_media)
                {
                    Console.WriteLine("-------------------------------------------------------------------------------");
                    Console.WriteLine(Path.Combine(output_folder_info.FullName, "media"));

                    ExportMediasC(cf);

                    action |= true;
                }
                if (extract_all || extract_embed)
                {
                    Console.WriteLine("-------------------------------------------------------------------------------");
                    Console.WriteLine(Path.Combine(output_folder_info.FullName, "embeddings"));

                    ExportAttachmentsC(cf);

                    action |= true;
                }
                if (action) Console.WriteLine("-------------------------------------------------------------------------------");
            }
        }
    }

    public static void ExportMediasC(CompoundFile officeFile)
    {
        if (officeFile is CompoundFile)
        {
            Console.WriteLine($"Sorry, unsupported extracting the media files");
            return;

#if DEBUG
            var Data = officeFile.RootStorage.GetStream("Data");
            var buf = Data.GetData();

            for (int sid = 0; sid < officeFile.GetNumDirectories(); sid++)
            {
                var name = officeFile.GetNameDirEntry(sid);
                var data = officeFile.GetDataBySID(sid);
                var stg = officeFile.GetStorageType(sid);
                Console.WriteLine($"{name} : {data.Length}, {stg}");

                if (name.EndsWith("Ole10Native"))
                {
                    var targetFolder = Path.Combine(output_folder, "embeddings");
                    var targetFolder_info = Directory.CreateDirectory(targetFolder);
                    if (targetFolder_info.Exists)
                    {
                        //ExportOle10Native(data, targetFolder);
                    }
                }
            }
#endif
        }
    }

    public static void ExportAttachmentsC(CompoundFile officeFile)
    {
        if (officeFile is CompoundFile)
        {
            for (int sid = 0; sid < officeFile.GetNumDirectories(); sid++)
            {
                var name = officeFile.GetNameDirEntry(sid);
                var data = officeFile.GetDataBySID(sid);
                //Console.WriteLine($"{name} : {data.Length}");
                if (name.EndsWith("Ole10Native"))
                {
                    var targetFolder = Path.Combine(output_folder, "embeddings");
                    var targetFolder_info = Directory.CreateDirectory(targetFolder);
                    if (targetFolder_info.Exists)
                    {
                        Console.WriteLine();
                        var ret = ExportOle10Native(data, targetFolder);
                        Console.WriteLine($"State : Save {(ret ? "successful" : "failed, file already exists")}");
                    }
                }
            }
        }
    }

    public static void ExportX(Package officeFile, string output)
    {
        if (officeFile is Package)
        {
            var action = false;

            if (extract_all || extract_media)
            {
                Console.WriteLine("-------------------------------------------------------------------------------");
                Console.WriteLine(Path.Combine(output, "media"));

                ExportMediasX(officeFile);

                action |= true;
            }
            if (extract_all || extract_embed)
            {
                Console.WriteLine("-------------------------------------------------------------------------------");
                Console.WriteLine(Path.Combine(output, "embeddings"));

                ExportAttachmentsX(officeFile);

                action |= true;
            }
            if (action) Console.WriteLine("-------------------------------------------------------------------------------");
        }

    }

    public static void ExportX(string officeFile)
    {
        if (output_folder_info?.Exists ?? false)
        {
            using (var pkg = Package.Open(officeFile))
            {
                ExportX(pkg, output_folder_info.FullName);
            }
        }
    }

    public static void ExportMediasX(Package officeFile)
    {
        if (officeFile is Package)
        {
            var targetFolder = Path.Combine(output_folder, "media");
            var targetFolder_info = Directory.CreateDirectory(targetFolder);
            if (targetFolder_info.Exists)
            {
                foreach (var mediaPart in officeFile.GetParts().Where(f => f.Uri.ToString().StartsWith($"/{office_type}/media/")))
                {
                    var targetFile = Path.Combine(targetFolder, Path.GetFileName(mediaPart.Uri.ToString()));
                    if (overwrite || !File.Exists(targetFile))
                    {
                        Console.WriteLine();
                        Console.WriteLine("Uri  : " + mediaPart.Uri);
                        using (var ms = mediaPart.GetStream())
                        {
                            var buffer = new byte[ms.Length];
                            ms.Read(buffer, 0, (int)ms.Length);
                            ms.Close();

                            Console.WriteLine("File  : " + targetFile);
                            File.WriteAllBytes(targetFile, buffer);
                            Console.WriteLine($"State : Save successful");
                        }
                    }
                    else Console.WriteLine($"State : Save failed, file already exists");
                }
            }
        }
    }

    public static void ExportMediasX(string officeFile)
    {
        if (string.IsNullOrEmpty(officeFile)) return;

        if (output_folder_info?.Exists ?? false)
        {
            using (var pkg = Package.Open(officeFile))
            {
                ExportMediasX(pkg);
            }
        }
    }

    public static void ExportAttachmentsX(Package officeFile)
    {
        if (officeFile is Package)
        {
            var targetFolder = Path.Combine(output_folder, "embeddings");
            var targetFolder_info = Directory.CreateDirectory(targetFolder);
            if (targetFolder_info.Exists)
            {
                foreach (var mediaPart in officeFile.GetParts().Where(f => f.Uri.ToString().StartsWith($"/{office_type}/embeddings/")))
                {
                    //Console.WriteLine(mediaPart.ContentType);
                    using (var ms = mediaPart.GetStream())
                    {
                        ms.Seek(0, SeekOrigin.Begin);
                        using (var compoundFile = new CompoundFile(ms))
                        {
                            for (var sid = 0; sid < compoundFile.GetNumDirectories(); sid++)
                            {
                                var name = compoundFile.GetNameDirEntry(sid);

                                ////Console.WriteLine(compoundFile.GetNameDirEntry(sid));
                                ////Console.WriteLine(compoundFile.GetStorageType(sid));
                                //var items = compoundFile.GetAllNamedEntries(name);
                                ////Console.WriteLine(items.Count);
                                //foreach (var item in items)
                                //{
                                //    Console.WriteLine(item.Name);
                                //    Console.WriteLine("Root    : " + item.IsRoot);
                                //    Console.WriteLine("Storate : " + item.IsStorage);
                                //    Console.WriteLine("Stream  : " + item.IsStream);
                                //}

                                if (name.EndsWith("Ole10Native"))
                                {
                                    Console.WriteLine();
                                    Console.WriteLine("Uri   : " + mediaPart.Uri);
                                    var data = compoundFile.GetDataBySID(sid);
                                    var ret = ExportOle10Native(data, targetFolder);
                                    Console.WriteLine($"State : Save {(ret ? "successful" : "failed, file already exists")}");
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    public static void ExportAttachmentsX(string officeFile)
    {
        if (string.IsNullOrEmpty(officeFile)) return;

        if (output_folder_info?.Exists ?? false)
        {
            using (var pkg = Package.Open(officeFile))
            {
                ExportAttachmentsX(pkg);
            }
        }
    }
}
