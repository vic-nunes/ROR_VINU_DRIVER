// -----------------------------------------------------------------------------
// Copyright (c) 2025 Vitor Nunes
// Licensed under the MIT License. See the LICENSE file in the project root.
// IMPORTANT: Use this code and any hardware connected to it at your own risk.
// The author(s) cannot be held responsible or liable for any damages,
// injuries, or losses resulting from the use of this software. Verify and test
// thoroughly before running on physical systems that move or control devices.
// -----------------------------------------------------------------------------
// ASCOM Dome hardware class for ROR_VINU
//
// Description:    Translates ASCOM dome commands to the VINU protocol for Arduino.
//
// Implements:    ASCOM Dome interface version:1.0
// Author:        Vitor Nunes
//
using ASCOM;
using ASCOM.DeviceInterface;
using ASCOM.Utilities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO.Ports;
using System.Threading;
using System.Windows.Forms;

namespace ASCOM.ROR_VINU.Dome
{
    /// <summary>
    /// Represents the state of the roof.
    /// </summary>
    public enum RoofState
    {
        roofOpen = 0,
        roofClosed = 1,
        roofOpening = 2,
        roofClosing = 3,
        roofError = 4
    }

    /// <summary>
    /// Represents the safety state of the telescope.
    /// </summary>
    public enum ScopeState
    {
        ScopeSafe = 0,
        ScopeNotSafe = 1
    }

    /// <summary>
    /// Roll-off roof implementation using ASCOM dome interface and VINU protocol.
    /// </summary>
    [HardwareClass()]
    internal static class DomeHardware
    {
        #region Constants and Fields

        // Profile settings

        internal const string comPortProfileName = "COM Port";
        internal const string comPortDefault = "COM3";
        internal const string traceStateProfileName = "Trace Level";
        internal const string traceStateDefault = "false";
        public const string timeoutProfileName = "Timeout";
        internal const string timeoutDefault = "30"; // seconds
        internal const string scopeSafeProfileName = "Scope Safe";
        internal const string scopeSafeDefault = "true";


        // Driver info
        private static readonly string driverInfo = "ASCOM Arduino Connector for Roll-Off Roof Observatory";
        private static readonly string driverDescription = "ROR using VINU protocol";
        private static readonly string driverName = "Arduino VINU Controller";
        private static readonly string driverVersion = "1.0.0";

        // VINU Commands
        private const string CMD_OPEN = "open#";
        private const string CMD_CLOSE = "close#";
        // private const string CMD_STOP_OPEN = "x#";
        private const string CMD_STOP = "stop#";
        private const string CMD_GET_STATUS = "get#";
        private const string CMD_INIT = "Init#";

        // Statuses
        private const string STATUS_OPEN = "open";
        private const string STATUS_CLOSED = "closed";
        private const string STATUS_OPENING = "opening";
        private const string STATUS_CLOSING = "closing";
        private const string STATUS_ERROR = "error";
        private const string STATUS_UNKNOWN = "unknown";
        private const string STATUS_SCOPE_SAFE = "safe";
        private const string STATUS_SCOPE_UNSAFE = "unsafe";

        private static int command_in_queue = 0;


        public static string DriverProgId = "ASCOM.ROR_VINU.Dome";
        internal static string comPort;
        private static bool isConnected;
        private static bool isMoving;
        private static RoofState currentShutterStatus = RoofState.roofClosed;
        private static ScopeState currentScopeState = ScopeState.ScopeNotSafe;

        internal static Util utilities;
        internal static TraceLogger tl;
        private static List<Guid> uniqueIds = new List<Guid>();
        private static SerialPort serialPort = null;
        private static int operationTimeout = 30; // seconds

        #endregion

        #region Driver Properties
           
        public static string Description => driverDescription;
        public static string DriverInfo => driverInfo;
        public static string Name => driverName;
        public static string DriverVersion => driverVersion;
        public static short InterfaceVersion => 3;

        #endregion

        #region Serial Port Methods

        private static void SendCommand(string command)
        {
            if (serialPort?.IsOpen == true)
            {
                LogMessage("SendCommand", $"Sending: {command}");
                serialPort.Write(command + "\n");
            }
            else
            {
                LogMessage("SendCommand", "Serial port not open.");
                throw new NotConnectedException("Serial port not open.");
            }
        }

        // Pseudocode (detailed plan):
        //1. Check if the serial port is open; if not, return an empty string.
        //2. Set the ReadTimeout to the provided value (milliseconds).
        //3. Create a StringBuilder buffer to accumulate characters.
        //4. Read byte-by-byte using serialPort.ReadByte(), which throws TimeoutException when the timeout elapses.
        //5. For each byte read:
        // a. Convert the byte to a char.
        // b. If the char is '#' -> finish reading (treat '#' as the line terminator).
        // c. Otherwise append the character to the buffer.
        //6. After finding '#', convert the buffer to a string, Trim() it and return.
        //7. Handle TimeoutException by returning an empty string and logging the event.
        //8. Handle other exceptions by logging the error and returning an empty string.
        // Notes:
        // - Use Trim() to remove any remaining CR/LF characters.
        // - Avoid consuming bytes beyond the '#' terminator to not interfere with subsequent messages.

        private static string ReceiveResponse(int timeout)
        {
            if (serialPort?.IsOpen == true)
            {
                try
                {
                    // timeout in milliseconds
                    serialPort.ReadTimeout = timeout;
                    var sb = new System.Text.StringBuilder();

                    while (true)
                    {
                        int read = serialPort.ReadByte(); // throws TimeoutException when the timeout elapses
                        char c = (char)read;

                        if (c == '#')
                        {
                            // terminator found -> end of line
                            // 

                            try
                            {
                                string leftover = serialPort.ReadExisting();
                                if (!string.IsNullOrEmpty(leftover))
                                {
                                    LogMessage("ReceiveResponse", $"Discarded after '#': {leftover}");
                                }
                            }
                            catch (Exception ex)
                            {
                                LogMessage("ReceiveResponse", $"Error on cleaning buffer: {ex.Message}");
                            }
                            break;
                        }
                        if ( c == '\n' || c == '\r')
                        {
                            // ignore line endings
                            continue;
                        }

                        sb.Append(c);
                    }

                    string response = sb.ToString();
                    LogMessage("ReceiveResponse", $"Received: {response}");

                    command_in_queue--;

                    return response.Trim();
                }
                catch (TimeoutException)
                {
                    LogMessage("ReceiveResponse", "Timeout waiting for response.");
                    return string.Empty;
                }
                catch (Exception ex)
                {
                    LogMessage("ReceiveResponse", $"Error reading response: {ex.Message}");
                    return string.Empty;
                }
            }
            return string.Empty;
        }

        #endregion

        private static void UpdateStatus(string statusResponse)
        {
            if (string.IsNullOrWhiteSpace(statusResponse))
            {
                LogMessage("UpdateStatus", "Empty status response, setting error state.");
                currentShutterStatus = RoofState.roofError;
                isMoving = false;
                return;
            }

            statusResponse = statusResponse.ToLowerInvariant().Trim();

            // Split by ',' and trim each token
            string[] parts = statusResponse.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < parts.Length; i++)
            {
                parts[i] = parts[i].Trim();
            }

            if (parts.Length != 3 )
            {
                LogMessage("UpdateStatus", $"Incomplete status response: {statusResponse}");
                return;
            }

            // Default to unknown if tokens missing
            string shutterToken = parts.Length > 0 ? parts[0] : string.Empty;
            string scopeToken = parts.Length > 1 ? parts[1] : string.Empty;
            string motionToken = parts.Length > 2 ? parts[2] : string.Empty;

            //1) Parse shutter token (first position)
            if (!string.IsNullOrEmpty(shutterToken))
            {
                if (shutterToken.Contains(STATUS_OPENING))
                {
                    currentShutterStatus = RoofState.roofOpening;
                    // Opening implies moving
                    isMoving = true;
                }
                else if (shutterToken.Contains(STATUS_CLOSING))
                {
                    currentShutterStatus = RoofState.roofClosing;
                    isMoving = true;
                }
                else if (shutterToken.Contains("opened") || shutterToken.Contains(STATUS_OPEN))
                {
                    currentShutterStatus = RoofState.roofOpen;
                }
                else if (shutterToken.Contains("closed") || shutterToken.Contains(STATUS_CLOSED))
                {
                    currentShutterStatus = RoofState.roofClosed;
                }
                else
                {
                    currentShutterStatus = RoofState.roofError;
                }
            }
            else
            {
                currentShutterStatus = RoofState.roofError;
            }

            //2) Parse scope safety (second position)
            if (!string.IsNullOrEmpty(scopeToken))
            {
                if (scopeToken.Equals(STATUS_SCOPE_SAFE))
                {
                    currentScopeState = ScopeState.ScopeSafe;
                }
                else
                {
                    currentScopeState = ScopeState.ScopeNotSafe;
                }
            }

            //3) Parse motion token (third position) - may override isMoving and/or shutter state
            if (!string.IsNullOrEmpty(motionToken))
            {
                if (motionToken.StartsWith("not_moving"))
                {
                    isMoving = false;
                    // not_moving_o => not moving and open
                    if (motionToken.EndsWith("_o"))
                    {
                        currentShutterStatus = RoofState.roofOpen;
                    }
                    else if (motionToken.EndsWith("_c"))
                    {
                        currentShutterStatus = RoofState.roofClosed;
                    }
                }
                else if (motionToken.StartsWith("moving"))
                {
                    isMoving = true;
                }
                else
                {
                    isMoving = false;
                }
            }

            LogMessage("UpdateStatus", $"Parsed status -> Shutter: {currentShutterStatus}, Scope: {currentScopeState}, IsMoving: {isMoving}");
        }

        private static void PollStatus()
        {
            if (!isConnected || serialPort?.IsOpen != true)
            {
                throw new NotConnectedException("Not connected.");
            }

            try
            {
                if (command_in_queue > 0)
                {
                    string response = ReceiveResponse(1000); //4-second timeout for status
                    if (!string.IsNullOrEmpty(response))
                    {
                        UpdateStatus(response);
                    }
                    command_in_queue--;
                }
                else
                {
                    SendCommand(CMD_GET_STATUS);
                    command_in_queue++;

                    string response = ReceiveResponse(1000); //4-second timeout for status
                    if (!string.IsNullOrEmpty(response))
                    {
                        UpdateStatus(response);
                    }
                }
                    
            }
            catch (Exception ex)
            {
                LogMessage("PollStatus", $"Error polling status: {ex.Message}");
                currentShutterStatus = RoofState.roofError;
            }
        }


        #region Roll-Off Roof Methods

        internal static void OpenShutter()
        {
            if (!isConnected) throw new NotConnectedException("Not connected.");
            if (isMoving) throw new InvalidOperationException("Roof is already moving.");
            if (currentShutterStatus == RoofState.roofOpen)
            {
                LogMessage("OpenShutter", "Roof is already open.");
                return;
            }
            if (currentScopeState == ScopeState.ScopeNotSafe)
            {
                LogMessage("OpenShutter", "Roof is not safe.");
                return;
            }

            LogMessage("OpenShutter", "Attempting to open shutter.");
            SendCommand(CMD_OPEN);

            if (!WaitForShutterCompletion( RoofState.roofOpen ))
            {
                AbortSlew();
                throw new DriverException("Timeout opening shutter.");
            }
        }

        internal static void CloseShutter()
        {
            if (!isConnected) throw new NotConnectedException("Not connected.");
            if (isMoving) throw new InvalidOperationException("Roof is already moving.");
            if (currentShutterStatus == RoofState.roofClosed)
            {
                LogMessage("CloseShutter", "Roof is already closed.");
                return;
            }
            if (currentScopeState == ScopeState.ScopeNotSafe)
            {
                LogMessage("CloseShutter", "Roof is not safe.");
                return;
            }

            LogMessage("CloseShutter", "Attempting to close shutter.");
            SendCommand(CMD_CLOSE);

            if (!WaitForShutterCompletion(RoofState.roofClosed))
            {
                AbortSlew();
                throw new DriverException("Timeout closing shutter.");
            }
        }

        internal static void AbortSlew()
        {
            if (!isConnected) throw new NotConnectedException("Not connected.");

            LogMessage("AbortSlew", "Aborting slew.");

            if ( isMoving || currentShutterStatus == RoofState.roofOpening || currentShutterStatus == RoofState.roofClosing)
            {
                SendCommand(CMD_STOP);
                isMoving = false;
            }

            PollStatus();
        }

        internal static ShutterState ShutterStatus
        {
            get
            {
                if (!isConnected) throw new NotConnectedException("Not connected.");
                PollStatus(); // Actively poll for the current status.
                return (ShutterState)currentShutterStatus;
            }
        }

        private static bool WaitForShutterCompletion(RoofState targetState)
        {
            int timeoutMs = operationTimeout * 1000;

            LogMessage("WaitForShutterCompletion", "Defined timeout: " + timeoutMs);

            
            int pollIntervalMs = 500;
            int elapsedMs = 0;

            while (elapsedMs < timeoutMs)
            {
                Thread.Sleep(pollIntervalMs);
                elapsedMs += pollIntervalMs;

                PollStatus();

                if (currentScopeState == ScopeState.ScopeNotSafe)
                {
                    isMoving = false;
                    LogMessage("WaitForShutterCompletion", "Operation aborted: Scope not safe.");
                    return false;
                }

                if (currentShutterStatus == targetState || !isMoving  )
                {
                    isMoving = false;
                    LogMessage("WaitForShutterCompletion", $"Operation successful. Reached target state: {targetState}");
                    return true;
                }

                if (currentShutterStatus == RoofState.roofError)
                {
                    isMoving = false;
                    LogMessage("WaitForShutterCompletion", "Operation failed with error state.");
                    return false;
                }
            }

            LogMessage("WaitForShutterCompletion", "Operation timed out.");
            return false; // Timeout
        }





        #endregion

        #region Setup and Connection

        internal static void InitialiseHardware()
        {
            tl = new TraceLogger("", "ROR_VINU.DomeHardware");
            utilities = new Util();
            ReadProfileSettings();
        }

        internal static void SetConnected(Guid uniqueId, bool connected)
        {
            if (connected)
            {
                if (!uniqueIds.Contains(uniqueId))
                {
                    if (uniqueIds.Count == 0)
                    {
                        try
                        {
                            LogMessage("SetConnected", "First client connecting.");
                            serialPort = new SerialPort(comPort, 9600, Parity.None, 8, StopBits.One)
                            {
                                NewLine = "#\n",
                                ReadTimeout = operationTimeout * 1000,
                                WriteTimeout = operationTimeout * 1000,
                                ReadBufferSize = 1024,
                                WriteBufferSize = 1024
                            };
                            serialPort.Open();

                            SendCommand(CMD_INIT);
                            string response = ReceiveResponse(operationTimeout * 1000);
                            if (response.IndexOf("VINU", StringComparison.OrdinalIgnoreCase) < 0)
                            {
                                throw new DriverException("Device is not an VINU controller.");
                            }

                            isConnected = true;
                            LogMessage("SetConnected", "Connection successful.");
                            command_in_queue = 0;
                        }
                        catch (Exception ex)
                        {
                            LogMessage("SetConnected", $"Connection failed: {ex.Message}");
                            if (serialPort?.IsOpen == true) serialPort.Close();
                            serialPort?.Dispose();
                            serialPort = null;
                            isConnected = false;
                            throw;
                        }
                    }
                    uniqueIds.Add(uniqueId);
                    LogMessage("SetConnected", $"Client {uniqueId} added. Total clients: {uniqueIds.Count}");
                }
            }
            else
            {
                if (uniqueIds.Contains(uniqueId))
                {
                    uniqueIds.Remove(uniqueId);
                    LogMessage("SetConnected", $"Client {uniqueId} removed. Total clients: {uniqueIds.Count}");
                    if (uniqueIds.Count == 0)
                    {
                        LogMessage("SetConnected", "Last client disconnected.");
                        if (isMoving) AbortSlew();
                        if (serialPort?.IsOpen == true) serialPort.Close();
                        serialPort?.Dispose();
                        serialPort = null;
                        isConnected = false;
                        currentShutterStatus = RoofState.roofError;
                    }
                }
            }
        }

        internal static void SetupDialog()
        {
            if (isConnected)
            {
                MessageBox.Show("Cannot show setup dialog when connected.");
                return;
            }

            using (var f = new SetupDialogForm(tl))
            {
                if (f.ShowDialog() == DialogResult.OK)
                {
                    WriteProfileSettings();
                }
            }
        }

        internal static void ReadProfileSettings()
        {
            using (var driverProfile = new Profile { DeviceType = "Dome" })
            {
                tl.Enabled = Convert.ToBoolean(driverProfile.GetValue(DriverProgId, traceStateProfileName, string.Empty, traceStateDefault));
                comPort = driverProfile.GetValue(DriverProgId, comPortProfileName, string.Empty, comPortDefault);
                operationTimeout = Convert.ToInt32(driverProfile.GetValue(DriverProgId, timeoutProfileName, string.Empty, timeoutDefault));
            }
        }

        internal static void WriteProfileSettings()
        {
            using (var driverProfile = new Profile { DeviceType = "Dome" })
            {
                driverProfile.WriteValue(DriverProgId, traceStateProfileName, tl.Enabled.ToString());
                driverProfile.WriteValue(DriverProgId, comPortProfileName, comPort);
                driverProfile.WriteValue(DriverProgId, timeoutProfileName, operationTimeout.ToString());
            }
        }

        internal static void SetTimeout(int seconds)
        {
            if (seconds >= 1 && seconds <= 300)
            {
                operationTimeout = seconds;
            }
        }





        #endregion

        #region ASCOM Dome Capabilities

        internal static bool CanSetShutter => true;
        internal static bool Slewing => isMoving;

        // Unsupported capabilities
        internal static bool CanFindHome => false;
        internal static bool CanPark => false;
        internal static bool CanSetAltitude => false;
        internal static bool CanSetAzimuth => false;
        internal static bool CanSetPark => false;
        internal static bool CanSlave => false;
        internal static bool CanSyncAzimuth => false;

        internal static double Altitude => throw new PropertyNotImplementedException("Altitude", false);
        internal static bool AtHome => throw new PropertyNotImplementedException("AtHome", false);
        internal static bool AtPark => throw new PropertyNotImplementedException("AtPark", false);
        internal static double Azimuth => throw new PropertyNotImplementedException("Azimuth", false);
        internal static ArrayList SupportedActions => new ArrayList();
        internal static string Action(string actionName, string actionParameters) => throw new ActionNotImplementedException(actionName);
        internal static void CommandBlind(string command, bool raw) => throw new MethodNotImplementedException("CommandBlind");
        internal static bool CommandBool(string command, bool raw) => throw new MethodNotImplementedException("CommandBool");
        internal static string CommandString(string command, bool raw) => throw new MethodNotImplementedException("CommandString");
        internal static void FindHome() => throw new MethodNotImplementedException("FindHome");
        internal static void Park() => throw new MethodNotImplementedException("Park");
        internal static void SetPark() => throw new MethodNotImplementedException("SetPark");
        internal static void SlewToAltitude(double altitude) => throw new MethodNotImplementedException("SlewToAltitude");
        internal static void SlewToAzimuth(double azimuth) => throw new MethodNotImplementedException("SlewToAzimuth");
        internal static void SyncToAzimuth(double azimuth) => throw new MethodNotImplementedException("SyncToAzimuth");
        internal static bool Slaved
        {
            get => false;
            set => throw new PropertyNotImplementedException("Slaved", true);
        }



        #endregion

        internal static void LogMessage(string identifier, string message)
        {
            tl?.LogMessageCrLf(identifier, message);
        }
    }
}
