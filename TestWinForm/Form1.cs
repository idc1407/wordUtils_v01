using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net.Sockets;

namespace TestWinForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btn_to_base64_Click(object sender, EventArgs e)
        {

            byte[] bytes = File.ReadAllBytes(@"d:\itemp\t1.docx");
            string base64String = Convert.ToBase64String(bytes, 0, bytes.Length);

            richTextBox1.Text = base64String;

            byte[] newBytes = Convert.FromBase64String(base64String);

            File.Delete(@"d:\itemp\t2.docx");


            File.WriteAllBytes(@"d:\itemp\t2.docx", newBytes);

            
        }

        private void btn_send_file_Click(object sender, EventArgs e)
        {
            try
            {
                TcpClient tcpClient = new TcpClient("127.0.0.1", 1234);
                string fileName = @"t1.docx";
                string fullFileName = Path.Combine(@"d:\itemp\", fileName);

                StreamWriter sWriter = new StreamWriter(tcpClient.GetStream());
                byte[] bytes = File.ReadAllBytes(fullFileName);

                sWriter.WriteLine(bytes.Length.ToString());
                sWriter.Flush();

                sWriter.WriteLine(fileName);
                sWriter.Flush();

                tcpClient.Client.SendFile(fullFileName);

            }
            catch (Exception ex)
            {
                richTextBox1.Text = ex.ToString();
            }
        }
    }
}
