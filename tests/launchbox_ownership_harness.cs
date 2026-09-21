using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using JrpgTranslator.LaunchBox;

public static class LaunchBoxOwnershipHarness
{
    private static int assertions;
    private static void Check(bool condition, string message)
    { assertions++; if (!condition) throw new Exception(message); }

    private static ProcessStartInfo Info(string executable, string script, params string[] args)
    {
        ProcessStartInfo info = new ProcessStartInfo { FileName = executable, WorkingDirectory = Path.GetDirectoryName(script)!, UseShellExecute = false, CreateNoWindow = true };
        info.ArgumentList.Add("/ErrorStdOut"); info.ArgumentList.Add(script);
        foreach (string arg in args) info.ArgumentList.Add(arg);
        return info;
    }
    private static string ReadWhenReady(string path)
    {
        for (int n = 0; n < 100; n++)
        {
            try { if (File.Exists(path)) { string text = File.ReadAllText(path); if (text.EndsWith("READY")) return text; } }
            catch (IOException) { }
            Thread.Sleep(50);
        }
        throw new Exception("Fixture did not become ready: " + path);
    }
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr CommandLineToArgvW(string command, out int count);
    [DllImport("kernel32.dll")] private static extern IntPtr LocalFree(IntPtr memory);

    public static string Run(string executable, string script, string output)
    {
        PluginPaths.DataDirectory = output;
        // Windows command line quoting includes empty values, quotes, trailing
        // backslashes, spaces, and Japanese profile names.
        foreach (string value in new[] { "", "simple", "日本語 profile", "C:\\folder space\\", "embedded\"quote", "\\\"" })
        {
            IntPtr parsed = CommandLineToArgvW("fixture " + OwnedProcess.Quote(value), out int count);
            try { Check(count == 2 && Marshal.PtrToStringUni(Marshal.ReadIntPtr(parsed, IntPtr.Size)) == value, "Argument quoting changed a value"); }
            finally { LocalFree(parsed); }
        }
        string unrelatedFile = Path.Combine(output, "unrelated.txt");
        using Process unrelated = Process.Start(Info(executable, script, "standalone", unrelatedFile))!;
        _ = unrelated.SafeHandle;
        try
        {
            ReadWhenReady(unrelatedFile);
            for (int n = 0; n < 5; n++)
            {
                string rootFile = Path.Combine(output, "owned-" + n + ".txt");
                string childFile = Path.Combine(output, "child-" + n + ".txt");
                using OwnedProcess owned = OwnedProcess.Start(Info(executable, script, n == 4 ? "hung-parent" : "parent", rootFile, childFile));
                string rootText = ReadWhenReady(rootFile);
                int rootId = int.Parse(rootText.Split('\n')[0].Trim());
                int childId = int.Parse(ReadWhenReady(childFile).Split('\n')[0].Trim());
                using Process root = Process.GetProcessById(rootId);
                using Process child = Process.GetProcessById(childId);
                _ = root.SafeHandle; _ = child.SafeHandle;
                string laterFile = Path.Combine(output, "later-" + n + ".txt");
                using Process later = Process.Start(Info(executable, script, "standalone", laterFile))!;
                _ = later.SafeHandle;
                try
                {
                ReadWhenReady(laterFile);
                // Exercise the actual coordinator cleanup method, not just a
                // standalone primitive. Later same-name processes stay outside.
                RuntimeSession session = new RuntimeSession("fixture", "Fixture", new TranslatorGameContext()) { Translator = owned };
                typeof(RuntimeCoordinator).GetMethod("StopTranslator", BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, new object[] { session });
                Check(root.WaitForExit(3000), "Owned root survived cleanup");
                Check(child.WaitForExit(3000), "Owned child survived cleanup");
                Check(!unrelated.HasExited, "Unrelated same-executable process was closed");
                owned.Stop(); // Idempotent, including a dead root.
                Check(!unrelated.HasExited, "Repeated cleanup affected standalone process");
                Check(!later.HasExited, "Cleanup affected later-started same-name process");
                }
                finally { if (!later.HasExited) { later.Kill(); later.WaitForExit(3000); } }
            }
            string deadRoot = Path.Combine(output, "dead-root.txt");
            string orphanFile = Path.Combine(output, "owned-orphan.txt");
            using (OwnedProcess owned = OwnedProcess.Start(Info(executable, script, "exit-parent", deadRoot, orphanFile)))
            {
                ReadWhenReady(deadRoot);
                using Process orphan = Process.GetProcessById(int.Parse(ReadWhenReady(orphanFile).Split('\n')[0].Trim()));
                _ = orphan.SafeHandle;
                Check(owned.WaitForExit(3000), "Early-exit parent fixture did not exit");
                owned.Stop();
                Check(orphan.WaitForExit(3000), "Dead-root cleanup missed its owned child");
                Check(!unrelated.HasExited, "Dead-root cleanup affected standalone process");
            }
            string failPath = Path.Combine(output, "does-not-exist.exe");
            bool failed = false;
            try { using OwnedProcess invalid = OwnedProcess.Start(Info(failPath, script)); }
            catch (System.ComponentModel.Win32Exception) { failed = true; }
            Check(failed && !unrelated.HasExited, "Failed launch damaged standalone process");
            return "PASS: " + assertions + " LaunchBox ownership assertions; full runtime compiled with host stubs, real isolated Windows process jobs tested.";
        }
        finally
        {
            // This exact retained helper was created by this test, not discovered.
            if (!unrelated.HasExited) { unrelated.Kill(); unrelated.WaitForExit(3000); }
        }
    }
}
