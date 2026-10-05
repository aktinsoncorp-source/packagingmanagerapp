using System;
using System.Diagnostics;
using System.IO;
using System.Text;

[assembly: System.Reflection.AssemblyVersion("1.11.1.0")]
[assembly: System.Reflection.AssemblyFileVersion("1.11.1.0")]

// Runs the readable engine source with its private, unmodified CPython runtime.
// Keep the existing request/response command line and process exit status.
static class PackingEngineLauncher {
    static string Quote(string value) {
        var result = new StringBuilder("\"");
        int slashes = 0;
        foreach (char c in value) {
            if (c == '\\') { slashes++; continue; }
            if (c == '"') { result.Append('\\', slashes * 2 + 1); result.Append(c); }
            else { result.Append('\\', slashes); result.Append(c); }
            slashes = 0;
        }
        result.Append('\\', slashes * 2);
        return result.Append('"').ToString();
    }

    static int Main(string[] args) {
        try {
            string root = AppDomain.CurrentDomain.BaseDirectory;
            string python = Path.Combine(root, "Runtime", "python.exe");
            string source = Path.Combine(root, "Source", "engine.py");
            if (!File.Exists(python) || !File.Exists(source))
                throw new FileNotFoundException("The processing runtime is incomplete. Keep the Packing List Manager application folder together.");
            var command = new StringBuilder("-B -X utf8 ");
            command.Append(Quote(source));
            foreach (string arg in args) command.Append(' ').Append(Quote(arg));
            var start = new ProcessStartInfo(python, command.ToString()) {
                UseShellExecute = false, CreateNoWindow = true,
                WorkingDirectory = Directory.GetParent(root.TrimEnd(Path.DirectorySeparatorChar)).FullName
            };
            using (var process = Process.Start(start)) {
                process.WaitForExit();
                return process.ExitCode;
            }
        } catch (Exception error) {
            Console.Error.WriteLine(error.Message);
            return 1;
        }
    }
}
