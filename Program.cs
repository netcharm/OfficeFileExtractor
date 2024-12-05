using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;

using OpenMcdf;

public class Program
{
    private static Encoding GBK = Encoding.Default;

    public static void Main(string[] args)
    {
        if (args.Length <= 0) return;
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        GBK = Encoding.GetEncoding("GBK");
        Console.WriteLine(string.Join(" ", args));
        Console.WriteLine("-------------------------------------------------------------------------------");
        ExportMedias(args[0]);
        Console.WriteLine("-------------------------------------------------------------------------------");
        ExportAttachments(args[0]);
        Console.WriteLine("-------------------------------------------------------------------------------");
    }

    public static void ExportMedias(string officeFile)
    {
        if (string.IsNullOrEmpty(officeFile)) return;

        var officeDir = Path.GetDirectoryName(Path.GetFullPath(officeFile));
        var officeExt = Path.GetExtension(officeFile).ToLower();
        var mediaDirectory = Path.Combine(officeDir ?? ".", $"{Path.GetFileNameWithoutExtension(officeFile)}_media");

        var mediaType = string.Empty;
        if (officeExt.StartsWith(".doc")) mediaType = "word";
        else if (officeExt.StartsWith(".xls")) mediaType = "excel";
        else if (officeExt.StartsWith(".ppt")) mediaType = "ppt";

        //Console.WriteLine(mediaDirectory);
        //Console.WriteLine(mediaType);
        var targetDirectoryInfo = Directory.CreateDirectory(mediaDirectory);
        Console.WriteLine(targetDirectoryInfo.FullName);
        Console.WriteLine("");
        if (targetDirectoryInfo.Exists)
        {
            var pkg = Package.Open(officeFile);
            foreach (var mediaPart in pkg.GetParts().Where(f => f.Uri.ToString().StartsWith($"/{mediaType}/media/")))
            {
                //Console.WriteLine(mediaPart.ContentType);
                Console.WriteLine("Uri  : " + mediaPart.Uri);
                using (var ms = mediaPart.GetStream())
                {
                    var buffer = new byte[ms.Length];
                    ms.Read(buffer, 0, (int)ms.Length);
                    ms.Close();
                    var targetFile = Path.Combine(mediaDirectory, Path.GetFileName(mediaPart.Uri.ToString()));
                    Console.WriteLine("File : " + targetFile);
                    File.WriteAllBytes(targetFile, buffer);
                    Console.WriteLine("");
                }
            }
            pkg.Close();
        }
    }

    public static void ExportAttachments(string officeFile)
    {
        if (string.IsNullOrEmpty(officeFile)) return;

        var officeDir = Path.GetDirectoryName(Path.GetFullPath(officeFile));
        var officeExt = Path.GetExtension(officeFile).ToLower();
        var mediaDirectory = Path.Combine(officeDir ?? ".", $"{Path.GetFileNameWithoutExtension(officeFile)}_embeddings");

        var mediaType = string.Empty;
        if (officeExt.StartsWith(".doc")) mediaType = "word";
        else if (officeExt.StartsWith(".xls")) mediaType = "excel";
        else if (officeExt.StartsWith(".ppt")) mediaType = "ppt";

        //Console.WriteLine(mediaDirectory);
        //Console.WriteLine(mediaType);

        var targetDirectoryInfo = Directory.CreateDirectory(mediaDirectory);
        Console.WriteLine(targetDirectoryInfo.FullName);
        Console.WriteLine("");
        if (targetDirectoryInfo.Exists)
        {
            var pkg = Package.Open(officeFile);
            foreach (var mediaPart in pkg.GetParts().Where(f => f.Uri.ToString().StartsWith($"/{mediaType}/embeddings/")))
            {
                //Console.WriteLine(mediaPart.ContentType);
                Console.WriteLine("Uri  : " + mediaPart.Uri);
                using (var ms = mediaPart.GetStream())
                {
                    var buffer = new byte[ms.Length];
                    ms.Read(buffer, 0, (int)ms.Length);
                    //Console.WriteLine(ms.Length);
                    using (var obj = new MemoryStream(buffer))
                    {
                        var compoundFile = new CompoundFile(obj);
                        for (var sid = 0; sid < compoundFile.GetNumDirectories(); sid++)
                        {
                            var name = compoundFile.GetNameDirEntry(sid);

/*
                            //Console.WriteLine(compoundFile.GetNameDirEntry(sid));
                            //Console.WriteLine(compoundFile.GetStorageType(sid));
                            var items = compoundFile.GetAllNamedEntries(name);
                            //Console.WriteLine(items.Count);
                            foreach(var item in items)
                            {
                                Console.WriteLine(item.Name);
                                Console.WriteLine("Root    : " + item.IsRoot);
                                Console.WriteLine("Storate : " + item.IsStorage);
                                Console.WriteLine("Stream  : " + item.IsStream);
                            }
*/

                            if (name.EndsWith("Ole10Native"))
                            {
                                var data = compoundFile.GetDataBySID(sid);
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
                                        fn0 = GBK.GetString(n0).Trim();
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
                                        fn1 = GBK.GetString(n1).Trim();
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
                                        fn2 = GBK.GetString(n2).Trim();
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

                                    var targetFile = Path.Combine(mediaDirectory, fn0);
                                    Console.WriteLine("File : " + targetFile);
                                    File.WriteAllBytes(targetFile, cnt);
                                    Console.WriteLine("");
                                }
                            }
                        }
                        compoundFile.Close();
                    }
                }
            }
            pkg.Close();
        }
    }
}
