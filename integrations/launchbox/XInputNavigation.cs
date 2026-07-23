using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;

namespace JrpgTranslator.LaunchBox
{
    internal enum ControllerNavigationCommand
    {
        Up,
        Down,
        Left,
        Right,
        Activate,
        Cancel
    }

    internal readonly struct ControllerNavigationState
    {
        public ControllerNavigationState(ushort buttons)
        {
            Up = (buttons & XInputController.DPadUp) != 0;
            Down = (buttons & XInputController.DPadDown) != 0;
            Left = (buttons & XInputController.DPadLeft) != 0;
            Right = (buttons & XInputController.DPadRight) != 0;
            Activate = (buttons & XInputController.ButtonA) != 0;
            Cancel = (buttons & XInputController.ButtonB) != 0;
        }

        public bool Up { get; }
        public bool Down { get; }
        public bool Left { get; }
        public bool Right { get; }
        public bool Activate { get; }
        public bool Cancel { get; }

        public bool IsPressed(ControllerNavigationCommand command)
        {
            return command switch
            {
                ControllerNavigationCommand.Up => Up,
                ControllerNavigationCommand.Down => Down,
                ControllerNavigationCommand.Left => Left,
                ControllerNavigationCommand.Right => Right,
                ControllerNavigationCommand.Activate => Activate,
                ControllerNavigationCommand.Cancel => Cancel,
                _ => false
            };
        }
    }

    internal static class XInputController
    {
        internal const ushort DPadUp = 0x0001;
        internal const ushort DPadDown = 0x0002;
        internal const ushort DPadLeft = 0x0004;
        internal const ushort DPadRight = 0x0008;
        internal const ushort ButtonA = 0x1000;
        internal const ushort ButtonB = 0x2000;

        private const int ErrorSuccess = 0;

        internal static bool TryReadNavigationState(out ControllerNavigationState state)
        {
            ushort combinedButtons = 0;
            bool connected = false;

            for (uint userIndex = 0; userIndex < 4; userIndex++)
            {
                if (!TryGetState(userIndex, out XInputState xinputState))
                {
                    continue;
                }

                connected = true;
                combinedButtons |= xinputState.Gamepad.Buttons;
            }

            state = new ControllerNavigationState(combinedButtons);
            return connected;
        }

        private static bool TryGetState(uint userIndex, out XInputState state)
        {
            try
            {
                return XInputGetState14(userIndex, out state) == ErrorSuccess;
            }
            catch (DllNotFoundException)
            {
                try
                {
                    return XInputGetState910(userIndex, out state) == ErrorSuccess;
                }
                catch (DllNotFoundException)
                {
                    state = default;
                    return false;
                }
            }
            catch (EntryPointNotFoundException)
            {
                state = default;
                return false;
            }
        }

        [DllImport("xinput1_4.dll", EntryPoint = "XInputGetState")]
        private static extern int XInputGetState14(uint userIndex, out XInputState state);

        [DllImport("xinput9_1_0.dll", EntryPoint = "XInputGetState")]
        private static extern int XInputGetState910(uint userIndex, out XInputState state);

        [StructLayout(LayoutKind.Sequential)]
        private struct XInputState
        {
            public uint PacketNumber;
            public XInputGamepad Gamepad;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct XInputGamepad
        {
            public ushort Buttons;
            public byte LeftTrigger;
            public byte RightTrigger;
            public short ThumbLeftX;
            public short ThumbLeftY;
            public short ThumbRightX;
            public short ThumbRightY;
        }
    }

    internal sealed class XInputDialogCancelScope : IDisposable
    {
        private const uint GetAncestorRootOwner = 3;
        private const uint WindowMessageClose = 0x0010;

        private readonly IntPtr _ownerHandle;
        private readonly DispatcherTimer _timer;
        private ControllerNavigationState _previousState;
        private bool _baselineReady;

        internal XInputDialogCancelScope(Window owner)
        {
            _ownerHandle = new WindowInteropHelper(owner).Handle;
            _timer = new DispatcherTimer(
                TimeSpan.FromMilliseconds(20),
                DispatcherPriority.Input,
                HandleTick,
                owner.Dispatcher);
            _timer.Start();
        }

        public void Dispose()
        {
            _timer.Stop();
        }

        private void HandleTick(object? sender, EventArgs e)
        {
            if (!XInputController.TryReadNavigationState(out ControllerNavigationState state))
            {
                _baselineReady = false;
                _previousState = default;
                return;
            }

            if (!_baselineReady)
            {
                _previousState = state;
                _baselineReady = true;
                return;
            }

            if (state.Cancel && !_previousState.Cancel)
            {
                CloseOwnedDialog();
            }

            _previousState = state;
        }

        private void CloseOwnedDialog()
        {
            IntPtr dialogHandle = GetForegroundWindow();
            if (dialogHandle == IntPtr.Zero || dialogHandle == _ownerHandle)
            {
                return;
            }

            GetWindowThreadProcessId(dialogHandle, out uint processId);
            if (processId != (uint)Environment.ProcessId
                || GetAncestor(dialogHandle, GetAncestorRootOwner) != _ownerHandle)
            {
                return;
            }

            PostMessage(dialogHandle, WindowMessageClose, IntPtr.Zero, IntPtr.Zero);
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr windowHandle, uint flags);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr windowHandle, out uint processId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PostMessage(
            IntPtr windowHandle,
            uint message,
            IntPtr wParam,
            IntPtr lParam);
    }
}
