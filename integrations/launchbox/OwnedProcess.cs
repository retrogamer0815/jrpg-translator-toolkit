using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace JrpgTranslator.LaunchBox
{
    // Launch suspended, assign to a private job, THEN resume. No descendant can
    // escape the ownership boundary during startup. Never rediscover by PID/name.
    internal sealed class OwnedProcess : IDisposable
    {
        private readonly SafeFileHandle job;
        private readonly SafeFileHandle process;
        private readonly uint processId;
        private bool disposed;

        private OwnedProcess(SafeFileHandle job, SafeFileHandle process, uint processId)
        {
            this.job = job;
            this.process = process;
            this.processId = processId;
        }

        public static OwnedProcess Start(ProcessStartInfo info)
        {
            SafeFileHandle job = CreateJobObject(IntPtr.Zero, null);
            if (job.IsInvalid) { job.Dispose(); throw new Win32Exception(); }
            SafeFileHandle? process = null;
            IntPtr thread = IntPtr.Zero;
            try
            {
                ExtendedLimits limits = new ExtendedLimits();
                limits.Basic.LimitFlags = 0x2000; // KILL_ON_JOB_CLOSE, no breakaway
                if (!SetInformationJobObject(job, 9, ref limits, (uint)Marshal.SizeOf<ExtendedLimits>()))
                    throw new Win32Exception();
                StartupInfo startup = new StartupInfo { Size = Marshal.SizeOf<StartupInfo>() };
                string command = Quote(info.FileName) + " " + string.Join(" ", info.ArgumentList.Select(Quote));
                if (!CreateProcess(info.FileName, new StringBuilder(command), IntPtr.Zero, IntPtr.Zero,
                    false, 0x4 | 0x08000000, IntPtr.Zero, info.WorkingDirectory, ref startup, out ProcessInfo created))
                    throw new Win32Exception();
                process = new SafeFileHandle(created.Process, true);
                thread = created.Thread;
                if (!AssignProcessToJobObject(job, process)) throw new Win32Exception();
                if (ResumeThread(thread) == uint.MaxValue) throw new Win32Exception();
                return new OwnedProcess(job, process, created.ProcessId);
            }
            catch
            {
                // Only the newly created, still-held process; never a PID lookup.
                if (process != null && !process.IsInvalid)
                    TerminateProcess(process, 1);
                process?.Dispose();
                job.Dispose();
                throw;
            }
            finally
            {
                if (thread != IntPtr.Zero) CloseHandle(thread);
            }
        }

        public bool WaitForExit(uint milliseconds) => WaitForSingleObject(process, milliseconds) == 0;

        public void Stop()
        {
            if (disposed) return;
            try
            {
                if (!WaitForExit(0))
                {
                    // The held process handle prevents PID reuse while finding its
                    // hidden AHK main window. WM_CLOSE there runs OnExit handlers.
                    EnumWindows((window, _) =>
                    {
                        GetWindowThreadProcessId(window, out uint owner);
                        if (owner == processId)
                        {
                            StringBuilder name = new StringBuilder(256);
                            GetClassName(window, name, name.Capacity);
                            if (name.ToString() == "AutoHotkey") PostMessage(window, 0x10, IntPtr.Zero, IntPtr.Zero);
                        }
                        return true;
                    }, IntPtr.Zero);
                    WaitForExit(1500);
                }
            }
            finally { Dispose(); }
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            // Closing the private job also catches owned children when the root
            // has exited. Concurrent standalone tools are never job members.
            job.Dispose();
            process.Dispose();
        }

        internal static string Quote(string value)
        {
            StringBuilder result = new StringBuilder("\"");
            int slashes = 0;
            foreach (char ch in value)
            {
                if (ch == '\\') { slashes++; continue; }
                result.Append('\\', ch == '"' ? slashes * 2 + 1 : slashes);
                result.Append(ch);
                slashes = 0;
            }
            return result.Append('\\', slashes * 2).Append('"').ToString();
        }

        [StructLayout(LayoutKind.Sequential)] private struct BasicLimits
        {
            public long ProcessTime, JobTime;
            public uint LimitFlags;
            public UIntPtr MinimumWorkingSet, MaximumWorkingSet;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint PriorityClass, SchedulingClass;
        }
        [StructLayout(LayoutKind.Sequential)] private struct IoCounters
        { public ulong ReadOperations, WriteOperations, OtherOperations, ReadBytes, WriteBytes, OtherBytes; }
        [StructLayout(LayoutKind.Sequential)] private struct ExtendedLimits
        {
            public BasicLimits Basic;
            public IoCounters Io;
            public UIntPtr ProcessMemory, JobMemory, PeakProcessMemory, PeakJobMemory;
        }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] private struct StartupInfo
        {
            public int Size;
            public string? Reserved, Desktop, Title;
            public uint X, Y, XSize, YSize, XCount, YCount, FillAttribute, Flags;
            public ushort ShowWindow, ReservedSize;
            public IntPtr ReservedData, StandardInput, StandardOutput, StandardError;
        }
        [StructLayout(LayoutKind.Sequential)] private struct ProcessInfo
        { public IntPtr Process, Thread; public uint ProcessId, ThreadId; }
        private delegate bool EnumWindowCallback(IntPtr window, IntPtr parameter);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern SafeFileHandle CreateJobObject(IntPtr security, string? name);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool SetInformationJobObject(SafeFileHandle job, int kind, ref ExtendedLimits limits, uint size);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool AssignProcessToJobObject(SafeFileHandle job, SafeFileHandle process);
        [DllImport("kernel32.dll", EntryPoint = "CreateProcessW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool CreateProcess(string application, StringBuilder command, IntPtr processSecurity, IntPtr threadSecurity, bool inheritHandles, uint flags, IntPtr environment, string directory, ref StartupInfo startup, out ProcessInfo result);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern uint ResumeThread(IntPtr thread);
        [DllImport("kernel32.dll")] private static extern bool TerminateProcess(SafeFileHandle process, uint code);
        [DllImport("kernel32.dll")] private static extern uint WaitForSingleObject(SafeFileHandle handle, uint milliseconds);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowCallback callback, IntPtr parameter);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int size);
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr window, uint message, IntPtr wparam, IntPtr lparam);
    }
}
