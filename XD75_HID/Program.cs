using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ColorMine.ColorSpaces;
using HidSharp;
using HidSharp.Utility;

namespace XD75_HID
{
    public static class Enums
    {
        public const string DefaultCounterFilename = "counter.csv";
        public enum RAW_COMMAND_ID : byte
        {
            RAW_COMMAND_GET_PROTOCOL_VERSION = 0x01,
            RAW_COMMAND_ENABLE_KEY_EVENT_REPORT = 0x02,
            RAW_COMMAND_DISABLE_KEY_EVENT_REPORT = 0x03,
            RAW_COMMAND_HEARTBEAT_PING = 0x04,
            RAW_COMMAND_CHANGE_COLOR = 0x05,
            RAW_COMMAND_REPORT_KEY_EVENT = 0xA0,
            RAW_COMMAND_UNDEFINED = 0xff,
        };
    }
    public static class Extensions
    {
        public static byte ToByte(this Enums.RAW_COMMAND_ID id)
        {
            return (byte)id;
        }
        public static byte[] ConstructRawCommand(this Enums.RAW_COMMAND_ID CommandID, byte[] payload = null)
        {
            byte[] rt = null;
            switch (CommandID)
            {
                case Enums.RAW_COMMAND_ID.RAW_COMMAND_GET_PROTOCOL_VERSION:
                    rt = new byte[3] { 0x00, Enums.RAW_COMMAND_ID.RAW_COMMAND_GET_PROTOCOL_VERSION.ToByte(), 0x00 };
                    break;
                case Enums.RAW_COMMAND_ID.RAW_COMMAND_ENABLE_KEY_EVENT_REPORT:
                    rt = new byte[3] { 0x00, Enums.RAW_COMMAND_ID.RAW_COMMAND_ENABLE_KEY_EVENT_REPORT.ToByte(), 0x00 };
                    break;
                case Enums.RAW_COMMAND_ID.RAW_COMMAND_DISABLE_KEY_EVENT_REPORT:
                    rt = new byte[3] { 0x00, Enums.RAW_COMMAND_ID.RAW_COMMAND_DISABLE_KEY_EVENT_REPORT.ToByte(), 0x00 };
                    break;
                case Enums.RAW_COMMAND_ID.RAW_COMMAND_CHANGE_COLOR:
                    if (payload.Length < 3)
                    {
                        rt = null;
                    }
                    else
                    {
                        rt = new byte[6] { 0x00, Enums.RAW_COMMAND_ID.RAW_COMMAND_CHANGE_COLOR.ToByte(), 0x03, payload[0], payload[1], payload[2] };
                    }
                    break;
                case Enums.RAW_COMMAND_ID.RAW_COMMAND_HEARTBEAT_PING:
                    rt = new byte[4] { 0x00, Enums.RAW_COMMAND_ID.RAW_COMMAND_HEARTBEAT_PING.ToByte(), 0x01, Program.HeartBeatCount };
                    break;
                default:
                    break;
            }
            return rt;
        }
        public static void DumpCounter(this int[] counter, string Filename = Enums.DefaultCounterFilename)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < 75; i++)
            {
                sb.Append(counter[i].ToString()+",");
                if (i % 15 ==14)
                {
                    sb.AppendLine();
                }
            }
            File.WriteAllText(Filename, sb.ToString());
        }

        public static int[] ReadCounter()
        {

            int[] counter = new int[75];
            try
            {
                if (File.Exists(Enums.DefaultCounterFilename))
                {

                    var values = File.ReadAllText(Enums.DefaultCounterFilename).Split(',').Select(str => str.Trim()).ToArray() ;
                    int count = 0;
                    foreach (var v in values)
                    {
                        if (v == null || v == "")
                            continue;
                        counter[count] = Convert.ToInt32(v);
                        count++;
                    }
                }
            }
            catch
            {
                Console.WriteLine("Something wrong with the counter file...");
                File.Delete(Enums.DefaultCounterFilename);
            }
            return counter;
        }

    }
    class Program
    {
        public const byte PROTOCOL_VERSION = 0x02;
        public const byte SUCCESS = 0x01;
        public const byte FAILED = 0xff;
        public static int[] keyCounter = new int[75];
        public static object StreamLock = new object();
        public static byte HeartBeatCount = 1;

        static void Main(string[] args)
        {

            var list = DeviceList.Local;
            HidDevice XD75 = null;
            HidStream XD75Strm = null;
            bool reportEnabled = false;
            keyCounter = Extensions.ReadCounter();

            XD75 = TryConnectXD75();

            DeviceList.Local.Changed += (object sender, DeviceListChangedEventArgs a) =>
            {
                if (XD75 == null)
                {
                    XD75 = TryConnectXD75();
                }
            };

            Thread heartBeatThread = new Thread(() =>
            {
                while (true)
                {
                    Thread.Sleep(5000);
                    if (XD75 != null && XD75Strm != null)
                    {
                        var heartBeatMessage = Enums.RAW_COMMAND_ID.RAW_COMMAND_HEARTBEAT_PING.ConstructRawCommand();
                        lock(StreamLock)
                        {
                            XD75Strm.Write(heartBeatMessage);
                        }
                    }
                }
            });
            Thread mainThread = new Thread(() =>
            {
                while (true)
                {
                    if (XD75 == null)
                    {
                        Console.WriteLine("No XD75 found, searching in background...");
                        while (XD75 == null) Thread.Sleep(100);
                    }

                    if (XD75.TryOpen(out XD75Strm))
                    {
                        reportEnabled = EnableReport(XD75Strm);

                        if (reportEnabled)
                        {
                            XD75Strm.ReadTimeout = Timeout.Infinite;

                            Console.CancelKeyPress += delegate
                            {
                                var DisableKeyEventReportCmd = Enums.RAW_COMMAND_ID.RAW_COMMAND_DISABLE_KEY_EVENT_REPORT.ConstructRawCommand();
                                lock (StreamLock)
                                {
                                    XD75Strm?.Write(DisableKeyEventReportCmd);
                                    XD75Strm.ReadTimeout = 50;
                                    XD75Strm?.Close();
                                }
                                keyCounter.DumpCounter();
                                XD75 = null;
                                XD75Strm = null;
                                heartBeatThread.Abort();
                                System.Environment.Exit(0);
                            };

                            while (true)
                            {
                                if (!heartBeatThread.IsAlive)
                                    heartBeatThread.Start();
                                try
                                {
                                    var report = XD75Strm.Read();

                                    if (report[1] == Enums.RAW_COMMAND_ID.RAW_COMMAND_HEARTBEAT_PING.ToByte())
                                    {
                                        Console.WriteLine("Received heartbeat response.");
                                        var heartBeatMsg = Enums.RAW_COMMAND_ID.RAW_COMMAND_HEARTBEAT_PING.ConstructRawCommand();
                                        var response = report.Take(heartBeatMsg.Count()).ToArray();
                                        if (!response.SequenceEqual(heartBeatMsg))
                                        {
                                            Console.WriteLine($"Heart beat message of no.{HeartBeatCount} mismatches...");
                                        }
                                        else
                                        {
                                            Console.WriteLine($"Heart beat message of no.{HeartBeatCount} matches.");
                                            HeartBeatCount++;
                                            continue;
                                        }
                                    }

                                    if (report[1] == Enums.RAW_COMMAND_ID.RAW_COMMAND_CHANGE_COLOR.ToByte())
                                    {
                                        Console.WriteLine("Received change underglow color response.");
                                        if(report[2] == 1)
                                        {
                                            if (report[3] == FAILED)
                                                Console.WriteLine("Command failed.");
                                            else
                                                Console.WriteLine("Command succeded unexpectely.");
                                        }
                                        else if (report[2] == 4)
                                        {
                                            if(report[6] == SUCCESS)
                                            {
                                                Hsv hsv = new Hsv();
                                                hsv.H = report[3] * 360d / 255d;
                                                hsv.S = report[4] / 255d;
                                                hsv.V = report[5] / 255d;
                                                Rgb rgb = hsv.To<Rgb>();
                                                Console.WriteLine($"Command succeded with value:{rgb.R} {rgb.G} {rgb.B}");
                                            }
                                            else
                                            {
                                                Console.WriteLine($"Command failed unexpectely.");
                                            }
                                        }
                                        continue;
                                    }

                                    if (report[1] != Enums.RAW_COMMAND_ID.RAW_COMMAND_REPORT_KEY_EVENT.ToByte())
                                    {
                                        Console.WriteLine("Received unknown command:");
                                        Console.WriteLine(BitConverter.ToString(report));
                                        continue;
                                    }

                                    int len = report[2];
                                    if (len != 0x2)
                                    {
                                        Console.WriteLine("Received unknown report:");
                                        Console.WriteLine(BitConverter.ToString(report));
                                        continue;
                                    }
                                    int col = (int)report[3];//0-14
                                    int row = (int)report[4];//0-4
                                    keyCounter[row * 15 + col]++;
                                    Console.WriteLine($"Received key event: {col},{row}");
                                    Console.WriteLine($"Total KeyCount:     {keyCounter.Sum()}");
                                    if (keyCounter.Sum() % 100 == 0)
                                    {
                                        Console.WriteLine("Saving keyCounter...");
                                        keyCounter.DumpCounter();
                                    }
                                }
                                catch (Exception e)
                                {
                                    XD75Strm?.Close();
                                    reportEnabled = false;
                                    XD75 = null;
                                    XD75Strm = null;
                                    break;
                                }

                            }
                            continue;
                        }
                        else
                        {
                            Console.WriteLine("Report enable failed.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Error communicating with XD75.");
                        continue;
                    }

                    XD75Strm?.Close();
                }
            });
            mainThread.Start();
            while (true)
            {
                var inputs = Console.ReadLine()?.Split();
                if (inputs == null) break;
                if (inputs[0].ToLower() == "c")
                {
                    if (inputs.Length == 4)
                    {
                        try
                        {
                            var payload = inputs.Skip(1).Select(s =>
                            {
                                if (s[0] == 'd')
                                {
                                    return Convert.ToByte(s.Trim('d'), 10);
                                }
                                return Convert.ToByte(s, 16);
                            }).ToArray();

                            Rgb rgb = new Rgb();
                            rgb.R = payload[0];
                            rgb.G = payload[1];
                            rgb.B = payload[2];

                            Console.WriteLine("RGB: "+payload[0].ToString("X2")+ payload[1].ToString("X2")+ payload[2].ToString("X2"));

                            var hsv = rgb.To<Hsv>();
                            payload[0] = (byte)(hsv.H / 360d * 255d);
                            payload[1] = (byte)(hsv.S * 255d);
                            payload[2] = (byte)(hsv.V * 255d);



                            Console.WriteLine(payload[0].ToString() + " " + payload[1].ToString() + " " + payload[2].ToString());

                            if (XD75 != null && XD75Strm != null)
                            {
                                var colorMessage = Enums.RAW_COMMAND_ID.RAW_COMMAND_CHANGE_COLOR.ConstructRawCommand(payload);
                                lock (StreamLock)
                                {
                                    XD75Strm.Write(colorMessage);
                                }
                            }
                        }
                        catch { };
                    }
                }
            }
        }

        private static bool EnableReport(HidStream XD75Strm)
        {
            bool reportEnabled = false;

            var EnableKeyEventReportCmd = Enums.RAW_COMMAND_ID.RAW_COMMAND_ENABLE_KEY_EVENT_REPORT.ConstructRawCommand();
            XD75Strm.Write(EnableKeyEventReportCmd);

            var response = XD75Strm.Read();

            if (response[1] != Enums.RAW_COMMAND_ID.RAW_COMMAND_ENABLE_KEY_EVENT_REPORT.ToByte())
            {
                Console.WriteLine("Unsupported feature: Enable key event report");
            }
            else
            {
                if (response[2] != 0x01)
                {
                    Console.WriteLine("Unsupported feature: Enable key event report");
                }
                else
                {
                    if (response[3] != SUCCESS)
                    {
                        Console.WriteLine("Unsupported feature: Enable key event report");
                    }
                    else
                    {
                        Console.WriteLine("Key event report enabled succusfully.");
                        reportEnabled = true;
                    }
                }
            }

            return reportEnabled;
        }

        private static HidDevice TryConnectXD75()
        {
            HidDevice XD75 = null;
            foreach (var device in DeviceList.Local.GetHidDevices(0xcdcd, 0x7575))
            {
                HidStream hstrm = null;
                try
                {
                    device.TryOpen(out hstrm);
                    if (hstrm == null) continue;
                    var GetProtocolRequest = Enums.RAW_COMMAND_ID.RAW_COMMAND_GET_PROTOCOL_VERSION.ConstructRawCommand();

                    hstrm.Write(GetProtocolRequest);

                    byte[] response = hstrm.Read();

                    if (response[1] == Enums.RAW_COMMAND_ID.RAW_COMMAND_GET_PROTOCOL_VERSION.ToByte())
                    {
                        if (response[2] != 0x01)
                        {

                            throw new Exception($"Expected Length is wrong: {response[1].ToString()}");
                        }
                        if (response[3] != PROTOCOL_VERSION)
                        {
                            throw new Exception($"Unsupported protocol version:{response[2].ToString()}");
                        }

                        Console.WriteLine("Negotiation successful.");
                        XD75 = device;
                        break;
                    }
                    else
                    {
                        throw new Exception($"Unsupported firmware:{response[0].ToString()}");
                    }
                }
                catch (Exception e)
                {
                    //Console.WriteLine(e.ToString());
                }
                finally
                {
                    hstrm?.Close();
                }
            }
            return XD75;
        }
    }
}