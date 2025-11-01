using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Yaml2Docx
{
    /// <summary>
    /// Starts a process in Windows
    /// </summary>
    public class ProcessLauncher
    {
        public static void StartProcess(
            string cmd,
            string args,
            string workDir = "",
            IEnumerable<string>? inputLines = null,
            List<string>? outputLines = null,
            string prefix = "",
            Dictionary<string, string>? argReplacements = null,
            Action<string>? lambdaLog = null)
        {
            // some placeholders
            if (argReplacements != null)
                foreach (var ar in argReplacements)
                {
                    args = args.Replace(ar.Key, ar.Value);
                }

            // log
            lambdaLog?.Invoke($"Starting process: {cmd} {args} ..");

            // start process??
            var proc = new Process();
            proc.StartInfo.FileName = cmd;
            proc.StartInfo.Arguments = args;
            proc.StartInfo.RedirectStandardOutput = true;
            proc.StartInfo.RedirectStandardError = true;
            proc.StartInfo.UseShellExecute = false;
            proc.StartInfo.CreateNoWindow = true;
            proc.EnableRaisingEvents = true;
            proc.StartInfo.WorkingDirectory = workDir;

            // input
            if (inputLines != null)
                proc.StartInfo.RedirectStandardInput = true;

            // finally start
            proc.Start();

            // feed in?
            if (inputLines != null)
            {
                using (var writer = proc.StandardInput)
                {
                    foreach (var line in inputLines)
                        writer.WriteLine(line);
                    proc.StandardInput.Flush();
                    proc.StandardInput.Close();
                }
            }

            proc.OutputDataReceived += (s1, e1) =>
            {
                var msg = e1.Data;
                if (msg != null)
                {
                    if (outputLines != null)
                        outputLines.Add(msg);
                    Console.WriteLine(prefix + msg);
                }
            };

            proc.ErrorDataReceived += (s2, e2) =>
            {
                var msg = e2.Data;
                if (msg != null)
                    Console.WriteLine(prefix + msg);
            };

            proc.Exited += (s3, e3) =>
            {
                Console.WriteLine(prefix + "Process exited.");
            };

            // proc.Start();

            proc.BeginOutputReadLine();
            proc.BeginErrorReadLine();

            proc.WaitForExit();
        }
    }
}
