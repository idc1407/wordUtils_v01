using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAspx
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void process_Click(object sender, EventArgs e)
        {
            if ((FileUpload1.PostedFile != null) && (FileUpload1.PostedFile.ContentLength > 0))
            {
                string fn = System.IO.Path.GetFileName(FileUpload1.PostedFile.FileName);
                string SaveLocation = Server.MapPath("App_Data") + "\\" + fn;
                try
                {
                    FileUpload1.PostedFile.SaveAs(SaveLocation);
                    byte[] bytes;
                    using (Stream fs = FileUpload1.PostedFile.InputStream)
                    {
                        using (BinaryReader br = new BinaryReader(fs))
                        {
                            bytes = br.ReadBytes((Int32)fs.Length);
                        }
                    }


                    using (TcpClient tcpClient = new TcpClient("127.0.0.1", 1234))
                    {
                        string fileName = FileUpload1.PostedFile.FileName;
                        string fullFileName = Path.Combine(@"d:\itemp\", fileName);

                        StreamWriter sWriter = new StreamWriter(tcpClient.GetStream());

                        sWriter.WriteLine(bytes.Length.ToString());
                        sWriter.Flush();

                        sWriter.WriteLine(fileName);
                        sWriter.Flush();

                        tcpClient.Client.SendFile(fullFileName);

                    }



                }
                catch (Exception ex)
                {
                    smessage.Text = "Error: " + ex.Message;
                }
                smessage.Text = "Job Completed Successfully!!";
            }
            else
            {
                smessage.Text = "Please select a file to upload.";
            }
        }
    }
}