using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
using System.IO;

namespace SerialPortReader
{
    class Program
    {
        public const int DEFAULT = 0;
        public const int SEARCH_SUCCESS = 1;
        public const int SEARCH_FAILED = 2;
        public const int SEARCH_MULTIPLE = 3;

        /// <summary>
        /// Args[0]: SearchPattern
        /// Args[1]: FilePathToWrite
        /// </summary>
        static int Main(string[] args)
        {
            List<SerialPorts.Info> infos = SerialPorts.GetSerialPortsInfo();

            string searchString = "";

            if(args.Length > 0)
            {
                searchString = args[0];
            }

            if (!string.IsNullOrEmpty(searchString))
            {
                Console.WriteLine("Searching COM ports for: " + searchString);

                List<SerialPorts.Info> matchingPorts = new List<SerialPorts.Info>();

                for (int i = 0; i < infos.Count; i++)
                {
                    var info = infos[i];

                    if(info.Caption.Contains(searchString))
                    {
                        matchingPorts.Add(info);
                    }
                }

                if(matchingPorts.Count == 0)
                {
                    Console.WriteLine("-No matching ports found-");
                    return SEARCH_FAILED;
                }
                else if(matchingPorts.Count == 1)
                {
                    Console.WriteLine("-Search successfull-");
                    Console.WriteLine(matchingPorts[0].ToString());

                    string writePath = "Port.txt";

                    if(args.Length > 1)
                    {
                        writePath = args[1];
                    }

                    File.WriteAllText(writePath, matchingPorts[0].Name);
                    Console.WriteLine("Writing to file: " + writePath);
                    return SEARCH_SUCCESS;
                }
                else
                {
                    Console.WriteLine("-Multiple ports found-");

                    for (int i = 0; i < matchingPorts.Count; i++)
                    {
                        Console.WriteLine(matchingPorts[i].ToString());
                    }

                    return SEARCH_MULTIPLE;
                }
            }
            else
            {
                Console.WriteLine("Printing all available COM ports");

                for (int i = 0; i < infos.Count; i++)
                {
                    var info = infos[i];
                    Console.WriteLine(info.ToString());
                }
            }

            return DEFAULT;
        }
    }
}
