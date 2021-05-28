using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Text.RegularExpressions;


public static class SerialPorts
{
    public static List<Info> GetSerialPortsInfo()
    {
        var retList = new List<Info>();

        // System.IO.Ports.SerialPort.GetPortNames() returns port names from Windows Registry
        var registryPortNames = SerialPort.GetPortNames().ToList();

        var managementObjectSearcher = new ManagementObjectSearcher(@"SELECT * FROM Win32_PnPEntity");
        var managementObjectCollection = managementObjectSearcher.Get();

        foreach (var n in registryPortNames)
        {
            //Console.WriteLine($"Searching for {n}");

            foreach (var p in managementObjectCollection)
            {
                if (p["Caption"] != null)
                {
                    string caption = p["Caption"].ToString();
                    string pnpdevid = p["PnPDeviceId"].ToString();

                    if (caption.Contains("(" + n + ")"))
                    {
                        //Console.WriteLine("PnPEntity port found: " + caption);
                        var props = p.Properties.Cast<PropertyData>().ToArray();
                        retList.Add(new Info(n, caption, pnpdevid));
                        break;
                    }
                }
            }
        }

        retList.Sort();

        return retList;
    }

    public class Info : IComparable
    {
        public Info() { }

        public Info(string name, string caption, string pnpdeviceid)
        {
            Name = name;
            Caption = caption;
            PNPDeviceID = pnpdeviceid;

            // build shorter version of PNPDeviceID for better reading
            // from this: BTHENUM\{00001101-0000-1000-8000-00805F9B34FB}_LOCALMFG&0000\7&11A88E8E&0&000000000000_0000000C
            // to this:  BTHENUM\{00001101-0000-1000-8000-00805F9B34FB}_LOCALMFG&0000
            try
            {
                // split by "\" and take 2 elements
                PNPDeviceIDShort = string.Join($"\\", pnpdeviceid.Split('\\').Take(2));
            }
            // todo: check if it can be split instead of using Exception
            catch (Exception)
            {
                // or just take 32 characters if split by "\" is impossible
                PNPDeviceIDShort = pnpdeviceid.Substring(0, 32) + "...";
            }
        }

        /// <summary>
        /// COM port name, "COM3" for example
        /// </summary>
        public string Name { get; }
        /// <summary>
        /// COM port caption from device manager
        /// "Intel(R) Active Management Technology - SOL (COM3)" for example
        /// </summary>
        public string Caption { get; }
        /// <summary>
        /// PNPDeviceID from device manager
        /// "PCI\VEN_8086&DEV_A13D&SUBSYS_224D17AA&REV_31\3&11583659&0&B3" for example
        /// </summary>
        public string PNPDeviceID { get; }

        /// <summary>
        /// Shorter version of PNPDeviceID
        /// "PCI\VEN_8086&DEV_A13D&SUBSYS_224D17AA&REV_31" for example
        /// </summary>
        public string PNPDeviceIDShort { get; }

        /// <summary>
        /// Comparer required to sort by COM port properly (number as number, not string (COM3 before COM21))
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            try
            {
                int a, b;
                string sa, sb;

                sa = Regex.Replace(Name, "[^0-9.]", "");
                sb = Regex.Replace(((Info)obj).Name, "[^0-9.]", "");

                if (!int.TryParse(sa, out a))
                    throw new ArgumentException(nameof(Info) + ": Cannot convert {0} to int32", sa);
                if (!int.TryParse(sb, out b))
                    throw new ArgumentException(nameof(Info) + ": Cannot convert {0} to int32", sb);

                return a.CompareTo(b);
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public override string ToString()
        {
            return $"\nName: { Name}\nCaption: { Caption }\nPNPDeviceId: { PNPDeviceID }\n";
        }
    }
}
