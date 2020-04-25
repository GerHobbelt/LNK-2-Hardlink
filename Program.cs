using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Diagnostics.Contracts;
using CommandLine;
using System.IO;

namespace ReplaceLNKsWithHardlinks
{
    class Program
    {
        static void Main(string[] args)
        {
            CommandLine.Parser.Default.ParseArguments<Options>(args)
              .WithParsed(RunOptions)
              .WithNotParsed(HandleParseError);
        }


        class Options
        {
            [Option('d', "directory", Required = false, HelpText = "Directories to be processed.")]
            public IEnumerable<string> InputDirectories { get; set; }

            [Option('f', "file", Required = false, HelpText = "Files to be processed.")]
            public IEnumerable<string> InputFiles { get; set; }

            // Omitting long name, defaults to name of property, ie "--verbose"
            [Option(
              Default = false,
              HelpText = "Prints all messages to standard output.")]
            public bool Verbose { get; set; }

            [Option("stdin",
              Default = false,
              HelpText = "Read from stdin")]
            public bool stdin { get; set; }

            [Value(0, MetaName = "offset", HelpText = "File offset.")]
            public long? Offset { get; set; }
        }


        static void RunOptions(Options opts)
        {
            //handle options

            string path = opts.InputDirectories.First();

            foreach (string file in GetFiles(path))
            {
                Console.WriteLine(file);
                try
                {
                    LinkInfo dest = GetShortcutTarget(file);
                    Console.WriteLine($"{file} --> IsFile:{ dest.IsFile }, dest: { dest.Path }");
                    // double-check before we do this:
                    string lnkExt = file.Substring(file.Length - 4);
                    if (lnkExt.ToLower() == ".lnk")
                    {
                        string source = file.Substring(0, file.Length - 4);
                        Console.WriteLine(source);

                        try
                        {
                            // rare occasion:  LNK points to the file itself:
                            if (source == dest.Path)
                            {
                                Console.WriteLine("### WARNING: LINK points to itself. Deleting link!");
                                File.Delete(file);
                                Console.WriteLine("--> .LNK file deleted.");
                                continue;
                            }
                            // what if the .LNK was pointing to another .LNK?
                            string lnkExt2 = dest.Path.Substring(dest.Path.Length - 4);
                            if (lnkExt2.ToLower() == ".lnk")
                            {
                                Console.WriteLine("### WARNING: LINK points to another LINK!");

                                // And when the file itself *exists*, we take that one:
                                string dest2 = dest.Path.Substring(0, dest.Path.Length - 4);
                                if (File.Exists(dest2))
                                {
                                    dest.Path = dest2;
                                }
                            }

                            CreateHardLinkOrThrow(source, dest.Path, IntPtr.Zero);
                            if (File.Exists(source))
                            {
                                File.Delete(file);
                                Console.WriteLine("--> .LNK file replaced with a HARDLINK.");
                            }
                            else
                            {
                                throw new Exception("Spurious Failure in CreateHardLink code...");
                            }
                        }
                        catch (Exception ex)
                        {
                            // "An attempt was made to create more links on a file than the file system supports"?
                            if (ex.Message.Contains("An attempt was made to create more links on a file than the file system supports"))
                            {
                                // copy file:
                                try
                                {
                                    if (dest.IsFile)
                                    {
                                        File.Copy(dest.Path, source);
                                        if (File.Exists(source))
                                        {
                                            File.Delete(file);
                                            Console.WriteLine("--> file COPIED.");
                                        }
                                        else
                                        {
                                            throw new Exception("Spurious Failure in FileCopy code...");
                                        }
                                    }
                                }
                                catch (Exception ex2)
                                {
                                    Console.WriteLine($"### FILE COPY ERROR: { ex2.Message }");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"### ERROR: { ex.Message }");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"####### LINK ERROR: { ex.Message }");
                }
            }
        }
        static void HandleParseError(IEnumerable<Error> errs)
        {
            //handle errors
        }

        // https://stackoverflow.com/questions/929276/how-to-recursively-list-all-the-files-in-a-directory-in-c

        static IEnumerable<string> GetFiles(string path)
        {
            Queue<string> queue = new Queue<string>();
            queue.Enqueue(path);
            while (queue.Count > 0)
            {
                path = queue.Dequeue();
                try
                {
                    foreach (string subDir in Directory.GetDirectories(path))
                    {
                        queue.Enqueue(subDir);
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"### ERROR for directory { path }: { ex.Message }");
                    continue;
                }

                // NOTE: this next section is coded this way as you CANNOT yield a value within a try/catch block in C#.

                string[] files = null;
                try
                {
                    files = Directory.GetFiles(path, "*.lnk", SearchOption.TopDirectoryOnly);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"### ERROR for directory { path } while looking for .LNK files: { ex.Message }");
                }
                if (files != null)
                {
                    for (int i = 0; i < files.Length; i++)
                    {
                        yield return files[i];
                    }
                }
            }
        }


        // SetLastError: tell the .NET runtime to capture the last Win32
        // error if it fails
        [DllImportAttribute("Kernel32.dll", SetLastError = true)]
        public static extern bool CreateHardLink(string lpFileName, string lpExistingFileName, IntPtr lpSecurityAttributes);

        // Wrapper for CreateHardLink() that throws an exception on failure.
        // Done in C# rather than in PowerShell to ensure PowerShell
        // internals don't clobber the error.
        public static void CreateHardLinkOrThrow(string dest, string source, IntPtr lpSecurityAttributes)
        {
            if (!CreateHardLink(dest, source, lpSecurityAttributes))
            {
                throw new System.ComponentModel.Win32Exception();
            }
        }

        // https://blez.wordpress.com/2013/02/18/get-file-shortcuts-target-with-c/

        struct LinkInfo
        {
            public bool IsFile;
            public string Path;
        }

        static LinkInfo GetShortcutTarget(string file)
        {
            if (System.IO.Path.GetExtension(file).ToLower() != ".lnk")
            {
                throw new Exception("Supplied file must be a .LNK file");
            }

            FileStream fileStream = File.Open(file, FileMode.Open, FileAccess.Read);
            using (System.IO.BinaryReader fileReader = new BinaryReader(fileStream))
            {
                fileStream.Seek(0x14, SeekOrigin.Begin);     // Seek to flags
                uint flags = fileReader.ReadUInt32();        // Read flags
                if ((flags & 1) == 1)
                {
                    // Bit 1 set means we have to skip the shell item ID list
                    fileStream.Seek(0x4c, SeekOrigin.Begin); // Seek to the end of the header
                    uint offset = fileReader.ReadUInt16();   // Read the length of the Shell item ID list
                    fileStream.Seek(offset, SeekOrigin.Current); // Seek past it (to the file locator info)
                }

                long fileInfoStartsAt = fileStream.Position; // Store the offset where the file info structure begins
                uint totalStructLength = fileReader.ReadUInt32(); // read the length of the whole struct
                fileStream.Seek(0xc, SeekOrigin.Current); // seek to offset to base pathname
                uint fileOffset = fileReader.ReadUInt32(); // read offset to base pathname
                                                           // the offset is from the beginning of the file info struct (fileInfoStartsAt)
                fileStream.Seek((fileInfoStartsAt + fileOffset), SeekOrigin.Begin); // Seek to beginning of base pathname (target)
                long pathLength = (totalStructLength + fileInfoStartsAt) - fileStream.Position - 2; // read the base pathname. I don't need the 2 terminating nulls.
                char[] linkTarget = fileReader.ReadChars((int)pathLength); // should be Unicode safe
                string link = new string(linkTarget);

                int begin = link.IndexOf("\0\0");
                if (begin > -1)
                {
                    int end = link.IndexOf("\\\\", begin + 2) + 2;
                    end = link.IndexOf('\0', end) + 1;

                    string firstPart = link.Substring(0, begin);
                    string secondPart = link.Substring(end);

                    link = firstPart + secondPart;
                }

                // check if target exists:
                if (!String.IsNullOrWhiteSpace(link))
                {
                    if (File.Exists(link))
                    {
                        return new LinkInfo()
                        {
                            IsFile = true,
                            Path = link
                        };
                    }
                    if (Directory.Exists(link))
                    {
                        return new LinkInfo()
                        {
                            IsFile = false,
                            Path = link
                        };
                    }
                    if (!link.ToLower().Contains(".lnk") && File.Exists(link + ".LNK"))
                    {
                        return new LinkInfo()
                        {
                            IsFile = true,
                            Path = link + ".LNK"
                        };
                    }
                    throw new Exception($"Invalid link: target does not exist: { link }");
                }
            }
            throw new Exception($"Invalid link: target is empty?!");
        }
    }
}
