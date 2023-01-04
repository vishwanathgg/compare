using Aspose.Email;
using System.Diagnostics;

namespace VishWork.Services
{
    public class CompareService
    {
        string workingFolder = @"C:\Users\Vish\source\repos\VishWork\VishWork\CompareScripts\";
        string compareSessionFolder = @"C:\Users\Vish\source\repos\VishWork\VishWork\CompareScripts\Session\";
        string bCompareExecFile = @"C:\Program Files\Beyond Compare 4\BComp.com";
        string compareOutputFolder = @"C:\Users\Vish\source\repos\VishWork\VishWork\CompareScripts\Session\Output\";
        public List<CompareResult> CompareTwoFiles(string file1, string file2, bool contentCheck = false)
        {
            List<CompareResult> compareResults = new List<CompareResult>();

            if (Path.GetExtension(file1) != Path.GetExtension(file2))
                throw new Exception("Can compare only file with same extension.");

            if((Path.GetExtension(file1) == ".msg" && Path.GetExtension(file2) == ".msg") ||
                (Path.GetExtension(file1) == ".eml" && Path.GetExtension(file2) == ".eml"))
            {
                string parentFile = Path.GetFileName(file1) == Path.GetFileName(file2) ? Path.GetFileName(file1) : Path.GetFileName(file1) + " <> " + Path.GetFileName(file2);
                compareResults.AddRange(CompareFileAttributes(new List<FileDetails>() { new FileDetails() { File1 = file1, File2 = file2, ParentFile = parentFile } }).ToList());
                compareResults.AddRange(CompareEmailMessageAttributes(new List<FileDetails>() { new FileDetails() { File1 = file1, File2 = file2, ParentFile = parentFile } }).ToList());

                var attachmentFolder1 = ExtractAttachmentFolderFromMessage(file1, "_1_");
                var attachmentFolder2 = ExtractAttachmentFolderFromMessage(file2, "_2_");
                
                compareResults.AddRange(CompareMissingFilesinFolders(attachmentFolder1, attachmentFolder2, parentFile));
                var attachmentFiles = GetExistingFilesInFolders(attachmentFolder1, attachmentFolder2, parentFile);
                compareResults.AddRange(CompareFileAttributes(attachmentFiles));

                if (contentCheck)
                {
                    compareResults.AddRange(CompareFileContent(attachmentFiles));
                }
            }
            else
            {
                string parentFile = Path.GetFileName(file1) == Path.GetFileName(file2) ? Path.GetFileName(file1) : Path.GetFileName(file1) + " <> " + Path.GetFileName(file2);
                compareResults.AddRange(CompareFileAttributes(new List<FileDetails>() { new FileDetails() { File1 = file1, File2 = file2, ParentFile = parentFile } }).ToList());

                if (contentCheck)
                {
                    compareResults.AddRange(CompareFileContent(new List<FileDetails>() { new FileDetails() { File1 = file1, File2 = file2, ParentFile = parentFile } }));
                }
            }
            return compareResults.OrderBy(x=>(x.ParentFile, x.File)).ToList();
        }

        public List<CompareResult> CompareTwoFolders(string folder1, string folder2, bool contentCheck = false)
        {
            List<CompareResult> compareResults = new List<CompareResult>();

            DirectoryInfo d1 = new DirectoryInfo(folder1);
            DirectoryInfo d2 = new DirectoryInfo(folder2);

            string[] fileEntriesFolder1 = d1.GetFiles().Select(o => o.Name).ToArray();
            string[] fileEntriesFolder2 = d2.GetFiles().Select(o => o.Name).ToArray();

            var s3 = fileEntriesFolder1.Intersect(fileEntriesFolder2).ToList();

            compareResults.AddRange(CompareMissingFilesinFolders(folder1, folder2));

            var msgFiles = s3.Where(x => Path.GetExtension(x).Equals(".msg") || Path.GetExtension(x).Equals(".eml"))
                .Select(x => new FileDetails() { File1 = folder1 + "\\" + x, File2 = folder2 + "\\" + x, ParentFile = x}).ToList();
            var otherFiles = s3.Where(x => !Path.GetExtension(x).Equals(".msg") && !Path.GetExtension(x).Equals(".eml"))
                .Select(x => new FileDetails() { File1 = folder1 + "\\" + x, File2 = folder2 + "\\" + x, ParentFile = x }).ToList();

            compareResults.AddRange(CompareFileAttributes(s3.Select(x => new FileDetails() { File1 = folder1 + "\\" + x, File2 = folder2 + "\\" + x, ParentFile = x }).ToList()));
            compareResults.AddRange(CompareEmailMessageAttributes(msgFiles.Select(x => new FileDetails() { File1 = x.File1, File2 = x.File2, ParentFile = x.ParentFile }).ToList()));

            var filesForContentCheck = otherFiles.Select(x => new FileDetails() { File1 = x.File1, File2 = x.File2, 
                ParentFile = x.ParentFile,
                }).ToList();

            msgFiles.ForEach(msg =>
            {
                string parentFile = Path.GetFileName(msg.File1);
                var attachmentFolder1 = ExtractAttachmentFolderFromMessage(msg.File1, "_1_");
                var attachmentFolder2 = ExtractAttachmentFolderFromMessage(msg.File2, "_2_");
                compareResults.AddRange(CompareMissingFilesinFolders(attachmentFolder1, attachmentFolder2, parentFile));
                var attachmentFiles = GetExistingFilesInFolders(attachmentFolder1, attachmentFolder2, parentFile);
                compareResults.AddRange(CompareFileAttributes(attachmentFiles));
                filesForContentCheck.AddRange(attachmentFiles);
            });

            if(contentCheck)
            {
                compareResults.AddRange(CompareFileContent(filesForContentCheck));
            }

            return compareResults.OrderBy(x => (x.ParentFile, x.File)).ToList();
        }

        private List<CompareResult> CompareFileAttributes(List<FileDetails> files)
        {
            List<CompareResult> compareResults = new List<CompareResult>();
            files.ForEach(file =>
            {
                FileInfo fileInfo = new FileInfo(file.File1);
                FileInfo fileInfo2 = new FileInfo(file.File2);

                compareResults.Add(GetCompareResult(file.ParentFile, fileInfo.Name, CompareType.FileName, fileInfo.Name, fileInfo2.Name));
                compareResults.Add(GetCompareResult(file.ParentFile, fileInfo.Name, CompareType.FileSize, fileInfo.Length.ToString(), fileInfo.Length.ToString()));
                compareResults.Add(GetCompareResult(file.ParentFile, fileInfo.Name, CompareType.FileExtension, fileInfo.Extension, fileInfo.Extension));
                compareResults.Add(GetCompareResult(file.ParentFile, fileInfo.Name, CompareType.ReadOnly, (File.GetAttributes(file.File1) & FileAttributes.ReadOnly).ToString(),
                    (File.GetAttributes(file.File2) & FileAttributes.ReadOnly).ToString()));
                compareResults.Add(GetCompareResult(file.ParentFile, fileInfo.Name, CompareType.Hidden, (File.GetAttributes(file.File1) & FileAttributes.Hidden).ToString(),
                    (File.GetAttributes(file.File2) & FileAttributes.Hidden).ToString()));
            });
            return compareResults;
        }

        private List<CompareResult> CompareEmailMessageAttributes(List<FileDetails> files)
        {
            List<CompareResult> compareResults = new List<CompareResult>();
            files.ForEach(file =>
            {
                string fileName1 = Path.GetFileName(file.File1);
                string fileName2 = Path.GetFileName(file.File2);
                MailMessage mailMessage1 = MailMessage.Load(file.File1);
                MailMessage mailMessage2 = MailMessage.Load(file.File2);
                string toAddress = String.Join(",", mailMessage1.To.Select(x=>x.Address));
                string fromAddress = mailMessage1.From.Address;
                string ccAddress = String.Join(",", mailMessage1.CC.Select(x => x.Address));
                string bccAddress = String.Join(",", mailMessage1.Bcc.Select(x => x.Address));
                string attachmentNames = String.Join(",", mailMessage1.Attachments.Select(x=>x.Name));

                compareResults.Add(GetCompareResult(file.ParentFile, fileName1, CompareType.Email_To, String.Join(",", mailMessage1.To.Select(x => x.Address)),
                    String.Join(",", mailMessage2.To.Select(x => x.Address))));

                compareResults.Add(GetCompareResult(file.ParentFile, fileName1, CompareType.Email_From, mailMessage1.From.Address, mailMessage2.From.Address));

                compareResults.Add(GetCompareResult(file.ParentFile, fileName1, CompareType.Email_Cc, String.Join(",", mailMessage1.CC.Select(x => x.Address)),
                    String.Join(",", mailMessage2.CC.Select(x => x.Address))));

                compareResults.Add(GetCompareResult(file.ParentFile, fileName1, CompareType.Email_Bcc, String.Join(",", mailMessage1.Bcc.Select(x => x.Address)),
                    String.Join(",", mailMessage2.Bcc.Select(x => x.Address))));

                compareResults.Add(GetCompareResult(file.ParentFile, fileName1, CompareType.Email_Subject, mailMessage1.Subject, mailMessage2.Subject));
                compareResults.Add(GetCompareResult(file.ParentFile, fileName1, CompareType.Email_Body, mailMessage1.Body, mailMessage2.Body));

                compareResults.Add(GetCompareResult(file.ParentFile, fileName1, CompareType.Email_NumberOfAttachment, mailMessage1.Attachments.Count.ToString(), mailMessage2.Attachments.Count.ToString()));

                compareResults.Add(GetCompareResult(file.ParentFile, fileName1, CompareType.Email_AttachmentNames, String.Join(",", mailMessage1.Attachments.Select(x => x.Name)),
                    String.Join(",", mailMessage2.Attachments.Select(x => x.Name))));
            });
            return compareResults;
        }

        private List<CompareResult> CompareMissingFilesinFolders(string folder1, string folder2, string parentFile=null)
        {
            List<CompareResult> compareResults = new List<CompareResult>();
            string[] fileEntriesFolder1 = new DirectoryInfo(folder1).GetFiles().Select(o => o.Name).ToArray();
            string[] fileEntriesFolder2 = new DirectoryInfo(folder2).GetFiles().Select(o => o.Name).ToArray();


            var s1 = fileEntriesFolder1.Except(fileEntriesFolder2).ToList();
            var s2 = fileEntriesFolder2.Except(fileEntriesFolder1).ToList();

            foreach (string file in s1)
            {
                compareResults.Add(GetCompareResult(parentFile ?? file, file, CompareType.Exists, file, ""));
            }

            foreach (string file in s2)
            {
                compareResults.Add(GetCompareResult(parentFile ?? file, file, CompareType.Exists, "", file));
            }
            return compareResults;
        }

        private List<FileDetails> GetExistingFilesInFolders(string folder1, string folder2, string parentFile = "")
        {
            string[] fileEntriesFolder1 = new DirectoryInfo(folder1).GetFiles().Select(o => o.Name).ToArray();
            string[] fileEntriesFolder2 = new DirectoryInfo(folder2).GetFiles().Select(o => o.Name).ToArray();
            var s3 = fileEntriesFolder1.Intersect(fileEntriesFolder2).ToList();
            List<FileDetails> files = new List<FileDetails>();
            foreach (string file in s3)
            {
                files.Add(new FileDetails() { File1 = folder1 + "\\" + file, File2 = folder2 + "\\" + file, ParentFile = parentFile });
            }
            return files;
        }

        private string ExtractAttachmentFolderFromMessage(string msgFile, string suffix="")
        {
            var guID = Guid.NewGuid();
            var path = Directory.CreateDirectory(compareSessionFolder + "\\" + guID + suffix) + "\\";
            MailMessage mailMessage1 = MailMessage.Load(msgFile);
            foreach (var attach in mailMessage1.Attachments)
            {
                attach.Save(path + attach.Name);
            }
            return path;
        }

        private List<CompareResult> CompareFileContent(List<FileDetails> files)
        {
            files = files.ToList();
            List<CompareResult> compareResults = new List<CompareResult>();
            var scriptGuid = Guid.NewGuid();
            string scriptFile = compareSessionFolder + $"GenerateScript_{scriptGuid}.txt";
            string logFile = compareSessionFolder + $"Log_{scriptGuid}.txt";

            List<string> scriptLines = new List<string>();
            scriptLines.Add($"log verbose \"{logFile}\"");
            files.ForEach(file =>
            {
                string outputFile = compareOutputFolder + Path.GetFileName(file.File1) + "_" + Guid.NewGuid() + ".html" ;
                file.OutputFile = outputFile;
                scriptLines.Add(GetScriptLine(scriptFile, file.File1, file.File2, outputFile));
            });
            File.WriteAllLines(scriptFile, scriptLines);

            ProcessStartInfo processInfo = new ProcessStartInfo(bCompareExecFile, "@" + scriptFile + " -silent /silent /closeScript");
            processInfo.CreateNoWindow = true;
            processInfo.UseShellExecute = true;
            var p = Process.Start(processInfo);
            p.WaitForExit();
            
            int res = p.ExitCode;

            string errorLog="";
            if (res != 0)
            {
                var lines = File.ReadAllLines(logFile);
                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i].Contains("Scripting Error:"))
                        errorLog = lines[i] + Environment.NewLine + lines[i-1];
                }
            }
            p.Close();

            //File.Delete(scriptFile);
            //File.Delete(logFile);

            files.ForEach(x =>
            {
                if(File.Exists(x.OutputFile))
                {
                    var mismatched = File.ReadLines(x.OutputFile).Count(x => x.Contains("class=\"TextSegSigDiff\"")) > 0;
                    compareResults.Add(new CompareResult()
                    {
                        CompareType = CompareType.Content,
                        Status = mismatched ? CompareStatus.Fail : CompareStatus.Pass,
                        Link = mismatched ? x.OutputFile : "",
                        Value1 = Path.GetFileName(x.File1),
                        Value2 = Path.GetFileName(x.File2),
                        ParentFile = x.ParentFile,
                        File = Path.GetFileName(x.File1)
                    });
                }
                else
                {
                    compareResults.Add(new CompareResult()
                    {
                        CompareType = CompareType.Content,
                        Status = CompareStatus.Fail,
                        Description = "Failed to compare files",
                        Value1 = Path.GetFileName(x.File1),
                        Value2 = Path.GetFileName(x.File2),
                        Error = errorLog,
                        ParentFile = x.ParentFile,
                        File = Path.GetFileName(x.File1)
                    });
                }

            });

            
            return compareResults;
        }

        private CompareResult GetCompareResult(string parentFile,string file, CompareType compareType, string value1, string value2, string failDescription = "")
        {
            return new CompareResult()
            {
                CompareType = compareType,
                Status = value1 != value2 ? CompareStatus.Fail : CompareStatus.Pass,
                Value1 = value1,
                Value2 = value2,
                ParentFile = parentFile,
                File = file
            };
        }

        private string GetScriptLine(string scriptFile, string file1, string file2, string outputFileName)
        {
            string extn = Path.GetExtension(file1).ToUpper();
            if (extn == ".XLSX" || extn == ".XLS" || extn == ".CSV" )
                return $"text-report layout:side-by-side options:ignore-unimportant,display-context output-to:\"{outputFileName}\" output-options:html-color \"{file1}\" \"{file2}\"";
            else
                return $"text-report layout:side-by-side options:ignore-unimportant,display-context output-to:\"{outputFileName}\" output-options:html-color \"{file1}\" \"{file2}\"";
        }
    }



    public class CompareResult
    {
        public CompareType CompareType { get; set; }
        public CompareStatus Status { get; set; }
        public string Value1 { get; set; }
        public string Value2 { get; set; }
        public string Description { get; set; }
        public string Link { get; set; }
        public string Error { get; set; }
        public string ParentFile { get; set; }
        public string File { get; set; }
    }

    public class FileDetails
    {
        public string File1 { get; set; }
        public string File2 { get; set; }
        public string OutputFile { get; set; }
        public string ParentFile { get; set; }
    }

    public enum CompareType
    {
        Exists,
        FileName,
        FileExtension,
        FileSize,
        ReadOnly,
        Hidden,
        Content,
        Email_To,
        Email_From,
        Email_Cc,
        Email_Bcc,
        Email_Subject,
        Email_Body,
        Email_NumberOfAttachment,
        Email_AttachmentNames
    }

    public enum CompareStatus
    {
        Pass,
        Fail,
        Skipped
    }
}
