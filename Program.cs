using System;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;

using Mono.Options;
using OpenMcdf;

public class Program
{
    private static string AppName = Path.GetFileName(AppDomain.CurrentDomain.FriendlyName);

    private static Encoding MBCS = Encoding.Default;

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

        Export(doc);
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

    private static void ExportOle10Native(byte[] data, string targetFolder)
    {
        if (data is byte[] && data.Length > 0 && !string.IsNullOrEmpty(targetFolder) && Directory.Exists(targetFolder))
        {
            using (var olems = new MemoryStream(data))
            {
                //Console.WriteLine(olems.Length);

                var pos = olems.Position;

                olems.Seek(pos + 6, SeekOrigin.Begin);
                pos = olems.Position;
                var fn0 = string.Empty;
                var f0 = new byte[260];
                olems.Read(f0, 0, (int)Math.Min(olems.Length, 260));
                var n0 = f0.TakeWhile(b => b != 0).ToArray();
                //Console.WriteLine(n0.Length);
                if (n0.Length > 0)
                {
                    fn0 = MBCS.GetString(n0).Trim();
                    //Console.WriteLine(fn0);
                }
                pos += n0.Length;

                olems.Seek(pos + 1, SeekOrigin.Begin);
                pos = olems.Position;
                var fn1 = string.Empty;
                var f1 = new byte[260];
                olems.Read(f1, 0, (int)Math.Min(olems.Length, 260));
                var n1 = f1.TakeWhile(b => b != 0).ToArray();
                //Console.WriteLine(n1.Length);
                if (n1.Length > 0)
                {
                    fn1 = MBCS.GetString(n1).Trim();
                    //Console.WriteLine(fn1);
                }
                pos += n1.Length + 4;

                olems.Seek(pos + 1, SeekOrigin.Begin);
                pos = olems.Position;
                var s2 = new byte[4];
                olems.Read(s2, 0, 4);
                var l2 = BitConverter.ToInt32(s2, 0);
                //Console.WriteLine(l2);
                var fn2 = string.Empty;
                var f2 = new byte[l2];
                olems.Read(f2, 0, l2);
                var n2 = f2.TakeWhile(b => b != 0).ToArray();
                //Console.WriteLine(n2.Length);
                if (n2.Length > 0)
                {
                    fn2 = MBCS.GetString(n2).Trim();
                    //Console.WriteLine(fn2);
                }
                pos += n2.Length;

                pos = olems.Position;
                //Console.WriteLine(pos);
                var cs = new byte[4];
                olems.Read(cs, 0, 4);
                var cl = BitConverter.ToInt32(cs, 0);
                //Console.WriteLine(cl);
                var cnt = new byte[cl];
                olems.Read(cnt, 0, cl);

                var targetFile = Path.Combine(targetFolder, fn0);
                Console.WriteLine("File : " + targetFile);
                File.WriteAllBytes(targetFile, cnt);
                Console.WriteLine("");
            }
        }
    }

    public static void Export(string officeFile)
    {
        if (string.IsNullOrEmpty(output_folder)) output_folder = Path.ChangeExtension(officeFile, "_files").Replace("._files", "_files");
        output_folder_info = Directory.CreateDirectory(output_folder);

        var officeExt = Path.GetExtension(officeFile).ToLower();
        if (officeExt.StartsWith(".doc")) office_type = "word";
        else if (officeExt.StartsWith(".xls")) office_type = "excel";
        else if (officeExt.StartsWith(".ppt")) office_type = "ppt";

        if (officeExt.EndsWith("x") || officeExt.EndsWith("m")) ExportX(officeFile);
        else ExportC(officeFile);
    }

    public static void ExportC(string officeFile)
    {
        if (output_folder_info?.Exists ?? false)
        {
            using (var cf = new CompoundFile(officeFile))
            {
                var action = false;

                //var data = cf.RootStorage.GetStream("Data");
                //var buf = data.GetData();
                //var objs = cf.RootStorage.GetStorage("ObjectPool");
                //var items = cf.GetAllNamedEntries();

                if (extract_all || extract_media)
                {
                    Console.WriteLine("-------------------------------------------------------------------------------");
                    Console.WriteLine(Path.Combine(output_folder_info.FullName, "media"));
                    Console.WriteLine("");

                    ExportMediasC(cf);

                    action |= true;
                }
                if (extract_all || extract_embed)
                {
                    Console.WriteLine("-------------------------------------------------------------------------------");
                    Console.WriteLine(Path.Combine(output_folder_info.FullName, "embeddings"));
                    Console.WriteLine("");

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
            var Data = officeFile.RootStorage.GetStream("Data");
            var buf = Data.GetData();
            Console.WriteLine($"Sorry, unsupported extracting the media files");
            return;

#if DEBUG
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
                        ExportOle10Native(data, targetFolder);
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
                Console.WriteLine("");

                ExportMediasX(officeFile);

                action |= true;
            }
            if (extract_all || extract_embed)
            {
                Console.WriteLine("-------------------------------------------------------------------------------");
                Console.WriteLine(Path.Combine(output, "embeddings"));
                Console.WriteLine("");

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
                    Console.WriteLine("Uri  : " + mediaPart.Uri);
                    //Console.WriteLine(mediaPart.ContentType);
                    //Console.WriteLine(mediaPart.Package.PackageProperties.Language);
                    using (var ms = mediaPart.GetStream())
                    {
                        var buffer = new byte[ms.Length];
                        ms.Read(buffer, 0, (int)ms.Length);
                        ms.Close();

                        var targetFile = Path.Combine(targetFolder, Path.GetFileName(mediaPart.Uri.ToString()));
                        Console.WriteLine("File : " + targetFile);
                        File.WriteAllBytes(targetFile, buffer);
                        Console.WriteLine("");
                    }
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
                    Console.WriteLine("Uri  : " + mediaPart.Uri);
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
                                    var data = compoundFile.GetDataBySID(sid);
                                    ExportOle10Native(data, targetFolder);
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
