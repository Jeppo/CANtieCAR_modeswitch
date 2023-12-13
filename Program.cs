using System.Diagnostics;
using System.IO.Ports;
using System.Management;
using System.Text.RegularExpressions;
using IWshRuntimeLibrary;
using Microsoft.Win32;

FileVersionInfo myFileVersionInfo = FileVersionInfo.GetVersionInfo(Process.GetCurrentProcess().MainModule.FileName);

start:
Console.Clear();

Console.Title = AppDomain.CurrentDomain.FriendlyName;
Console.WriteLine(AppDomain.CurrentDomain.FriendlyName + " v" + myFileVersionInfo.ProductVersion);
Console.WriteLine("-----------------------------------------------------------------------");
Console.WriteLine();

string mode;
bool argset = false;
SerialPort port = new();
string CANtieCARport = "";
string serialOutput = "";

void port_DataReceived(object sender, SerialDataReceivedEventArgs e)
{
    serialOutput += port.ReadExisting();
}

string findInterface(string arg = "")
{
    using (ManagementClass i_Entity = new("Win32_PnPEntity"))
    {
        bool interfaceFound = false;
        string setInterface = "";

        foreach (ManagementObject i_Inst in i_Entity.GetInstances())
        {
            Object o_Guid = i_Inst.GetPropertyValue("ClassGuid");
            if (o_Guid == null || o_Guid.ToString().ToUpper() != "{4D36E978-E325-11CE-BFC1-08002BE10318}")
                continue;

            String s_Caption = i_Inst.GetPropertyValue("Caption").ToString();
            String s_Manufact = i_Inst.GetPropertyValue("Manufacturer").ToString();
            String s_DeviceID = i_Inst.GetPropertyValue("PnpDeviceID").ToString();
            String s_RegPath = "HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Enum\\" + s_DeviceID + "\\Device Parameters";
            String s_PortName = Registry.GetValue(s_RegPath, "PortName", "").ToString();

            int s32_Pos = s_Caption.IndexOf(" (COM");
            if (s32_Pos > 0)
                s_Caption = s_Caption.Substring(0, s32_Pos);

            serialOutput = "";
            try
            {
                port = new SerialPort(s_PortName, 9600, Parity.None, 8, StopBits.One);
                port.ReadTimeout = 100;
                port.Open();
                Thread.Sleep(100);
                port.Write("!tieCAR_SC_I\r");
                port.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived);
                Thread.Sleep(100);
            }
            catch (Exception)
            {
            }

            port.Close();

            switch (arg)
            {
                case "ports":
                    Console.Write("Port: " + s_PortName);

                    if (serialOutput.Contains("CANtieCAR"))
                    {
                        Match serialnumberFind = Regex.Match(serialOutput, ".*SN:(.*)\r.*");
                        var serialnumber = serialnumberFind.Groups[1].Value;

                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                        Console.Write(" (CANtieCAR - Serial: " + serialnumber + ")");
                        Console.ResetColor();
                    }

                    Console.Write("\n");
                    Console.WriteLine("Name: " + s_Caption);
                    Console.WriteLine("Manufacturer: " + s_Manufact);
                    Console.WriteLine();
                    break;
                case "modes":
                    if (serialOutput.Contains("CANtieCAR"))
                    {
                        var availableModes = Regex.Replace(serialOutput, ".*\\sModes:\\s(.*)", "$1", RegexOptions.Singleline);
                        foreach (string line in availableModes.Split("\n"))
                        {
                            if (line.Length > 6)
                                Console.WriteLine(line);
                        }
                        return s_PortName;
                    }
                    break;
                default:
                    if (serialOutput.Contains("CANtieCAR"))
                    {
                        if (!interfaceFound)
                        {
                            interfaceFound = true;
                            setInterface = s_PortName;
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Error: More than one CANtieCAR interfaces detected!");
                            Console.WriteLine("Please only connect one interface, when using this tool!");
                            Console.ResetColor();
                            return "error";
                        }
                    }
                    break;
            }
        }

        if (interfaceFound)
        {
                return setInterface;
        }
        else
        {
            return "";
        }
    }
}

try
{
    mode = args[0].ToUpper();
    argset = true;
}
catch (Exception)
{
    Console.WriteLine("Usage: " + Path.GetFileName(Process.GetCurrentProcess().MainModule?.FileName) + " [3-digit ID]");
    Console.WriteLine("E.g. \"" + Path.GetFileName(Process.GetCurrentProcess().MainModule?.FileName) + " 000\" for Loopback.");
    Console.WriteLine("IDs can be found by plugging in your CANtieCAR and pressing \"4\".");
    Console.WriteLine();
    Console.WriteLine("NOTE: Please don't store this program on the CANtieCAR interface!");
    Console.WriteLine("Some modes doesn't expose the thumbdrive,\nso you won't be able to use shortcuts.");
    Console.WriteLine();
    Console.WriteLine("Press \"1\" to create a shortcut.");
    Console.WriteLine("Press \"2\" to manually set mode.");
    Console.WriteLine("Press \"3\" to show current interface mode.");
    Console.WriteLine("Press \"4\" to list available modes from interface.");
    Console.WriteLine("Press \"5\" to list COM-ports.");
    Console.WriteLine();
    Console.Write("Press any other key to exit.. ");

    ConsoleKeyInfo key = Console.ReadKey(true);
    Console.WriteLine();
    Console.WriteLine();
    switch (key.Key)
    {
        case ConsoleKey.D1:
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("Create shortcut");
            Console.ResetColor();

            id:
            Console.Write("3-digit mode ID: ");
            var shortcutMode = Console.ReadLine().ToUpper();

            if (shortcutMode.Length != 3 || !shortcutMode.All(char.IsLetterOrDigit))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: Invalid mode ID!");
                Console.ResetColor();
                Console.WriteLine();
                goto id;
            }

            Console.Write("Shortcut name: ");
            var shortcutName = Console.ReadLine();
            string shortcutLocation = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + shortcutName + ".lnk";
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutLocation);
            shortcut.Arguments = shortcutMode;
            shortcut.TargetPath = Process.GetCurrentProcess().MainModule?.FileName;
            shortcut.Save();

            Console.WriteLine();
            Console.Write("Shortcut ");
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.Write(shortcutName);
            Console.ResetColor();
            Console.Write(" saved to Desktop!\n");
            Console.Write("Click the newly created shortcut to change interface mode to ");
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.Write(shortcutMode);
            Console.ResetColor();
            Console.Write(".\n");
            Console.WriteLine();
            Console.Write("Press any key to return.. ");
            Console.ReadKey(true);
            goto start;
        case ConsoleKey.D2:
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("Manually set mode");
            Console.ResetColor();
            Console.WriteLine();
            Console.WriteLine("Loading available modes from interface..");

            CANtieCARport = findInterface();
            if (string.IsNullOrEmpty(CANtieCARport))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: No CANtieCAR interface found!");
                Console.ResetColor();
                Console.WriteLine();
                Console.Write("Press any key to return.. ");
                Console.ReadKey(true);
                goto start;
            }

            serialOutput = "";
            try
            {
                port = new SerialPort(CANtieCARport, 9600, Parity.None, 8, StopBits.One);
                port.ReadTimeout = 100;
                port.Open();
                Thread.Sleep(100);
                port.Write("!tieCAR_SC_I\r");
                port.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived);
                Thread.Sleep(100);
                port.Close();
            }
            catch (Exception)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: Can't read from interface!");
                Console.ResetColor();
                goto confirm;
            }

            bool isModeAvailable = false;
            var availableModes = Regex.Replace(serialOutput, ".*\\sModes:\\s(.*)", "$1", RegexOptions.Singleline);
            availableModes = Regex.Replace(availableModes, @"^\s+$[\r\n]*", string.Empty, RegexOptions.Multiline);

            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("Available modes for CANtieCAR on " + CANtieCARport + ":");
            Console.ResetColor();

            foreach (string line in availableModes.Split("\n"))
            {
                if (line.Length > 6)
                    Console.WriteLine(line);
            }

            Console.WriteLine();

            manual:
            Console.Write("Enter 3-digit mode ID: ");
            mode = Console.ReadLine().ToUpper();
            Console.WriteLine();

            if (mode.Length != 3 || !mode.All(char.IsLetterOrDigit))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: Invalid mode ID!");
                Console.ResetColor();
                Console.WriteLine();
                goto manual;
            }

            foreach (string modelist in availableModes.Split("\r\n"))
            {
                if (modelist.StartsWith(mode))
                {
                    isModeAvailable = true;
                    break;
                }
            }

            if (!isModeAvailable)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: Mode not available!");
                Console.ResetColor();
                Console.WriteLine();
                goto manual;
            }

            goto main;
        case ConsoleKey.D3:
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("Current mode for CANtieCAR:");
            Console.ResetColor();

            CANtieCARport = findInterface();

            if (CANtieCARport == "error")
            {
                Console.WriteLine();
                Console.Write("Press any key to return.. ");
                Console.ReadKey(true);
                goto start;
            }

            if (string.IsNullOrEmpty(CANtieCARport))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: No CANtieCAR interface found!");
                Console.ResetColor();
                Console.WriteLine();
                Console.Write("Press any key to return.. ");
                Console.ReadKey(true);
                goto start;
            }


            serialOutput = "";
            try
            {
                port = new SerialPort(CANtieCARport, 9600, Parity.None, 8, StopBits.One);
                port.ReadTimeout = 100;
                port.Open();
                Thread.Sleep(100);
                port.Write("!tieCAR_SC_I\r");
                port.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived);
                Thread.Sleep(100);
                port.Close();
            }
            catch (Exception)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: Can't read from interface!");
                Console.ResetColor();
                goto start;
            }

            Match currentModeFind = Regex.Match(serialOutput, ".*Selected Mode: (.*)\\sModes:.*");
            var currentMode = currentModeFind.Groups[1].Value;
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine(currentMode);
            Console.ResetColor();
            Console.WriteLine();
            Console.Write("Press any key to return.. ");
            Console.ReadKey(true);
            goto start;
        case ConsoleKey.D4:
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("Available modes for CANtieCAR:");
            Console.ResetColor();
            if (string.IsNullOrWhiteSpace(findInterface("modes")))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: No CANtieCAR interface found!");
                Console.ResetColor();
            }
            Console.WriteLine();
            Console.Write("Press any key to return.. ");
            Console.ReadKey(true);
            goto start;
        case ConsoleKey.D5:
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("Available COM-ports:");
            Console.ResetColor();

            findInterface("ports");

            Console.Write("Press any key to return.. ");
            Console.ReadKey(true);
            goto start;
        default:
            goto end;
    }
}

main:
if (mode.Length != 3 || !mode.All(char.IsLetterOrDigit))
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("Invalid mode ID!");
    Console.ResetColor();
    goto confirm;
}

Console.WriteLine("Searching for interface");
Console.WriteLine();
Console.ForegroundColor = ConsoleColor.Blue;
Console.WriteLine("Please wait..");
Console.ResetColor();
Console.WriteLine();

//If no interface is found, go to end
CANtieCARport = findInterface();
if (string.IsNullOrEmpty(CANtieCARport))
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("Error: No CANtieCAR interface found!");
    Console.ResetColor();
    goto confirm;
}
else
{
    Console.Write("CANtieCAR interface found on ");
    Console.ForegroundColor = ConsoleColor.DarkYellow;
    Console.Write(CANtieCARport + '\n');
    Console.ResetColor();
}

//Get current interface info
serialOutput = "";
try
{
    port = new SerialPort(CANtieCARport, 9600, Parity.None, 8, StopBits.One);
    port.ReadTimeout = 100;
    port.Open();
    Thread.Sleep(100);
    port.Write("!tieCAR_SC_I\n");
    port.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived);
    Thread.Sleep(100);
    port.Close();
}
catch (Exception)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("Error: Can't read from interface!");
    Console.ResetColor();
    goto confirm;
}

Match currentModePreFind = Regex.Match(serialOutput, ".*Selected Mode: (.*)\\sModes:.*");
var currentModePre = currentModePreFind.Groups[1].Value;

if (currentModePre.StartsWith(mode))
{
    Console.WriteLine();
    Console.ForegroundColor = ConsoleColor.Green;
    Console.Write("Interface already in mode ");
    Console.ForegroundColor = ConsoleColor.DarkYellow;
    Console.Write(mode);
    Console.ForegroundColor = ConsoleColor.Green;
    Console.Write("!\n");
    Console.ResetColor();
    goto confirm;
}

Console.Write("Setting interface mode: ");
Console.ForegroundColor = ConsoleColor.DarkYellow;
Console.Write(mode + '\n');
Console.ResetColor();
Console.WriteLine();

//Check if mode is available, if started with argument
if (argset)
{
    Console.Write("Checking if mode ");
    Console.ForegroundColor = ConsoleColor.DarkYellow;
    Console.Write(mode);
    Console.ResetColor();
    Console.Write(" is avalable.");
    Console.WriteLine();

    bool isModeAvailableMain = false;
    var availableModesMain = Regex.Replace(serialOutput, ".*\\sModes:\\s(.*)", "$1", RegexOptions.Singleline);
    availableModesMain = Regex.Replace(availableModesMain, @"^\s+$[\r\n]*", string.Empty, RegexOptions.Multiline);

    foreach (string modelist in availableModesMain.Split("\r\n"))
    {
        if (modelist.StartsWith(mode))
        {
            isModeAvailableMain = true;
            break;
        }
    }

    if (!isModeAvailableMain)
    {
        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("Error: Mode " + mode + " not available!");
        Console.ResetColor();
        Console.WriteLine("To see available mode IDs, open this program directly and press \"4\".");
        goto confirm;
    }
}

//Set interface mode
serialOutput = "";
try
{
    port = new SerialPort(CANtieCARport, 9600, Parity.None, 8, StopBits.One);
    port.ReadTimeout = 100;
    port.Open();
    Thread.Sleep(100);
    port.Write("!tieCAR_SC_M");
    Thread.Sleep(100);
    port.Write(mode + "\r");
    port.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived);
    Thread.Sleep(100);
    port.Close();
}
catch (Exception)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("Error: Can't write to interface!");
    Console.ResetColor();
    goto confirm;
}

Console.WriteLine("Waiting for interface..");
Console.WriteLine();

bool interfaceFound = false;

for (int i = 0; i < 20; i++)
{
    CANtieCARport = findInterface();

    if (string.IsNullOrEmpty(CANtieCARport))
    {
        Thread.Sleep(500);
    }
    else
    {
        interfaceFound = true;
        break;
    }
}

if (!interfaceFound)
{
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine("Notice: Interface not found after mode changed!");
    Console.ResetColor();
    Console.WriteLine("This usually means the interface mode was changed successfully.");
    Console.WriteLine("A slow or pending driver installation can cause this issue.");
    goto confirm;
}

Console.Write("CANtieCAR interface found on ");
Console.ForegroundColor = ConsoleColor.DarkYellow;
Console.Write(CANtieCARport + '\n');
Console.ResetColor();

//Get new interface mode
serialOutput = "";
try
{
    port = new SerialPort(CANtieCARport, 9600, Parity.None, 8, StopBits.One);
    port.ReadTimeout = 100;
    port.Open();
    Thread.Sleep(100);
    port.Write("!tieCAR_SC_I\n");
    port.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived);
    Thread.Sleep(100);
    port.Close();
}
catch (Exception)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("Error: Can't read from interface!");
    Console.ResetColor();
    goto confirm;
}

Match currentModeMainFind = Regex.Match(serialOutput, ".*Selected Mode: (.*)\\sModes:.*");
var currentModeMain = currentModeMainFind.Groups[1].Value;

Console.WriteLine("Current interface mode:");
Console.ForegroundColor = ConsoleColor.DarkYellow;
Console.WriteLine(currentModeMain);
Console.ResetColor();
Console.WriteLine();

if (currentModeMain.StartsWith(mode))
{
    Console.ForegroundColor = ConsoleColor.Green;
    Console.Write("Success! Interface mode switched to mode ");
    Console.ForegroundColor = ConsoleColor.DarkYellow;
    Console.Write(mode);
    Console.ForegroundColor = ConsoleColor.Green;
    Console.Write("!\n");
    Console.ResetColor();
    Console.WriteLine();
}
else
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("Error: Interface mode NOT changed!");
    Console.ResetColor();
}

confirm:
Console.WriteLine();
Console.Write("Press any key to exit.. ");
Console.ReadKey();

end:
System.Environment.Exit(1);