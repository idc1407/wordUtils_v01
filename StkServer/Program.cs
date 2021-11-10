using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Sockets;
using System.IO;
using System.Text;

namespace StkServer
{
    class Program
    {
        static void Main(string[] args)
        {
            // Listen on port 1234    
            TcpListener tcpListener = new TcpListener(IPAddress.Any, 1234);
            tcpListener.Start();

            Console.WriteLine("Server started");

            //Infinite loop to connect to new clients    
            while (true)
            {
                // Accept a TcpClient    
                TcpClient tcpClient = tcpListener.AcceptTcpClient();

                Console.WriteLine("Connected to client");

                StreamReader reader = new StreamReader(tcpClient.GetStream());

                // The first message from the client is the file size    
                string cmdFileSize = reader.ReadLine();

                // The second message from the client is the filename.
                string cmdFileName = reader.ReadLine();


                int length = Convert.ToInt32(cmdFileSize);
                byte[] buffer = new byte[length];
                int received = 0;
                int read = 0;
                int size = 1024;
                int remaining = 0;

                // Read bytes from the client using the length sent from the client    
                while (received < length)
                {
                    remaining = length - received;
                    if (remaining < size)
                    {
                        size = remaining;
                    }
                    read = tcpClient.GetStream().Read(buffer, received, size);
                    received += read;
                }

                File.WriteAllBytes(Path.Combine(@"d:\itemp\skFiles\", cmdFileName), buffer);
                Console.WriteLine("File received and saved in " + Path.Combine(@"d:\itemp\skFiles\", cmdFileName));
            }
        }
    }
}