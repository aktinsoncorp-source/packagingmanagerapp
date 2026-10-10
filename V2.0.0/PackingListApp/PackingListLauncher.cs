using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;

[assembly: System.Reflection.AssemblyTitle("Packing List Manager")]
[assembly: System.Reflection.AssemblyProduct("Packing List Manager")]
[assembly: System.Reflection.AssemblyVersion("2.0.0.0")]
[assembly: System.Reflection.AssemblyFileVersion("2.0.0.0")]

// Keep the application in Python, using its own unmodified CPython/Tk runtime.
static class PackingListLauncher {
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

    [STAThread]
    static int Main(string[] args) {
        try {
            string root = AppDomain.CurrentDomain.BaseDirectory;
            string runtime = Path.Combine(root, "App Files", "Runtime");
            string python = Path.Combine(runtime, "pythonw.exe");
            string script = Path.Combine(root, "App Files", "Source", "desktop_start.py");
            if (!File.Exists(python) || !File.Exists(script))
                throw new FileNotFoundException("The application files are incomplete. Run the installer again, or extract the entire portable ZIP before opening the app.");
            var command = new StringBuilder("-B -X utf8 ");
            command.Append(Quote(script));
            foreach (string arg in args) command.Append(' ').Append(Quote(arg));
            var start = new ProcessStartInfo(python, command.ToString()) {
                UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = root
            };
            start.EnvironmentVariables["TCL_LIBRARY"] = Path.Combine(runtime, "tcl", "tcl8.6");
            start.EnvironmentVariables["TK_LIBRARY"] = Path.Combine(runtime, "tcl", "tk8.6");
            start.EnvironmentVariables.Remove("PYTHONHOME");
            start.EnvironmentVariables.Remove("PYTHONPATH");
            using (var process = Process.Start(start)) {
                process.WaitForExit();
                return process.ExitCode;
            }
        } catch (Exception error) {
            MessageBox.Show(error.Message, "Packing List Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
    }
}
