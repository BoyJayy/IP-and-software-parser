using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApp5
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        static IPHostEntry iphost = Dns.GetHostEntry(Dns.GetHostName());
        static string fthreeBytes = iphost.AddressList[1].MapToIPv4().ToString().Split('.')[0] + "." + iphost.AddressList[1].MapToIPv4().ToString().Split('.')[1] + "." + iphost.AddressList[1].MapToIPv4().ToString().Split('.')[2] + ".";
        static string log = Application.StartupPath + "\\CheckResult_ofIPS" + ".txt";
        static string progs_inf = Application.StartupPath + "\\list_progs" + ".txt";
        static string in_inf = Application.StartupPath + "\\Progs_to" + ".txt";
        static string out_inf = Application.StartupPath + "\\CheckResult_ofprogs" + ".txt";
        private static int found = Application.StartupPath.IndexOf("\\bin\\Debug");
        private static string strCoreData = Application.StartupPath.Substring(0, found);
        static string link_ports = Application.StartupPath + "\\ports.txt";
        Dictionary<string, string> ports_list = ReadAndSplitFile(link_ports);

        static Dictionary<string, string> ReadAndSplitFile(string filePath)
        {
            Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();

            try
            {
                using (StreamReader reader = new StreamReader(filePath))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {

                        keyValuePairs[line.Split('/')[0]] = line.Split('/')[1];
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error: ports info doesn't exist :(");
            }

            return keyValuePairs;
        }

        public void GetInstalled()
        {
            StreamWriter sw = new StreamWriter(progs_inf);
            ListBox lstDisplayHardware = new ListBox();
            string uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(uninstallKey))
            {
                foreach (string skName in rk.GetSubKeyNames())
                {
                    using (RegistryKey sk = rk.OpenSubKey(skName))
                    {
                        try
                        {

                            var displayName = sk.GetValue("DisplayName");
                            var size = sk.GetValue("EstimatedSize");

                            ListViewItem item;
                            if (displayName != null)
                            {
                                if (size != null)
                                    item = new ListViewItem(new string[] {displayName.ToString(),
                                                       size.ToString()});
                                else
                                    item = new ListViewItem(new string[] { displayName.ToString() });
                                lstDisplayHardware.Items.Add(item);
                            }
                        }
                        catch (Exception ex)
                        { }
                    }
                }
                for (int i = 0; i < lstDisplayHardware.Items.Count; i++)
                {
                    sw.WriteLine(lstDisplayHardware.Items[i]);
                }
                sw.Close();
            }
        }
        static class PortScanHelper
        {
            public static async Task<bool> IsPortOpen(string ipAddress, int port)
            {
                using (var ping = new Ping())
                {
                    var result = await ping.SendPingAsync(ipAddress, 100);
                    if (result.Status != IPStatus.Success)
                    {
                        return false;
                    }
                }

                try
                {
                    using (var client = new System.Net.Sockets.TcpClient())
                    {
                        await client.ConnectAsync(ipAddress, port);
                        return true;
                    }
                }
                catch
                {
                    return false;
                }
            }

            public static int[] GetAllPorts()
            {
                return new int[1024].Select((_, index) => index + 1).ToArray();
            }
        }

        static async Task<bool> IsPortOpen(string ipAddress, int port)
        {
            try
            {
                using (var client = new TcpClient())
                {
                    await client.ConnectAsync(ipAddress, port);
                    return true;
                }
            }
            catch (SocketException)
            {
                return false;
            }
        }
        StreamWriter sw1 = new StreamWriter(log);
        private async void button1_Click(object sender, EventArgs e)
        {
            resText.Clear();
            int st = trackBar1.Value;
            int f = trackBar2.Value;
            progressBar1.Maximum = f;
            progressBar1.Minimum = st;
            progressBar1.Value = st;
            for (int i = st; i <= f; i++)
            {
                progressBar1.Value = i;
                if (PingHost(fthreeBytes + i))
                {
                    string ipAdress = fthreeBytes + i;
                    await Task.Run(() =>
                    {
                        Parallel.ForEach(PortScanHelper.GetAllPorts(), async port =>
                        {
                            bool isOpen = await PortScanHelper.IsPortOpen(ipAdress, port);
                            string result = ipAdress + "/" + port + "/" + $"{(isOpen ? "Открыт" : "Закрыт")}" + "/";
                            string key = port.ToString();
                            Invoke(new Action(() =>
                            {
                                if (ports_list.ContainsKey(key))
                                    resText.AppendText(result + ports_list[port.ToString()] + Environment.NewLine);
                                else resText.AppendText(result + "Неизвестно" + Environment.NewLine);
                            }));
                        });
                    });
                }
                else
                {
                }
            }
        }

        public static bool PingHost(string nameOrAddress)
        {
            bool pingable = false;
            Ping pinger = null;
            try
            {
                pinger = new Ping();
                PingReply reply = pinger.Send(nameOrAddress);
                pingable = reply.Status == IPStatus.Success;
            }
            catch (PingException)
            {
                // Discard PingExceptions and return false;
            }
            finally
            {
                if (pinger != null)
                {
                    pinger.Dispose();
                }
            }
            return pingable;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            sw1.WriteLine(resText.Text);
            sw1.Close();
        }

        void wdReplace(Microsoft.Office.Interop.Word.Application app, string ft, string rt)
        {
            object findText = ft;
            object replaceWith = rt;
            object replace = 2;
            object missing = Type.Missing;
            app.Selection.Find.Execute(ref findText, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith,
            ref replace, ref missing, ref missing, ref missing, ref missing);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            File.Delete(Application.StartupPath + "\\IP_report.docx");
            File.Copy(Application.StartupPath + "\\IP_report_blank.docx", Application.StartupPath + "\\IP_report.docx");
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = null;
            object fileName = Application.StartupPath + "\\IP_report.docx";
            object falseValue = false;
            object trueValue = true;
            object missing = Type.Missing;
            doc = app.Documents.Open(ref fileName, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing);
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();
            wdReplace(app, "<now>", DateTime.Now.ToString());
            Microsoft.Office.Interop.Word.Table tbl = app.ActiveDocument.Tables[1];
            for (int i = 0; i < resText.Lines.Count() - 1; i++)
            {
                tbl.Rows.Add();
                tbl.Rows[i + 2].Cells[1].Range.Text = (i + 1).ToString();
                tbl.Rows[i + 2].Cells[2].Range.Text = resText.Lines[i].ToString().Split('/')[0];
                tbl.Rows[i + 2].Cells[3].Range.Text = resText.Lines[i].ToString().Split('/')[1];
                tbl.Rows[i + 2].Cells[4].Range.Text = resText.Lines[i].ToString().Split('/')[2];
                tbl.Rows[i + 2].Cells[5].Range.Text = resText.Lines[i].ToString().Split('/')[3];

            }
            app.ActiveDocument.Save();
            app.Visible = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {

            GetInstalled();
            StreamWriter sw1 = new StreamWriter(in_inf);
            sw1.WriteLine(textBox1.Text);
            sw1.Close();
            StreamReader sr = new StreamReader(in_inf);
            StreamReader sr1 = new StreamReader(progs_inf);
            StreamWriter sw2 = new StreamWriter(out_inf);
            string line;
            string line1;
            while ((line = sr.ReadLine()) != null)
            {
                while ((line1 = sr1.ReadLine()) != null)
                {
                    string t = line1.Split('{')[1];
                    string tocheck = t.Split('}')[0];
                    if (tocheck.Contains(line.Split('/')[0]))
                    {
                        sw2.WriteLine(tocheck + "/" + line.Split('/')[1]);
                    }
                }
                sr1.BaseStream.Position = 0;
            }
            sr.Close();
            sr1.Close();
            sw2.Close();
            if (File.Exists(out_inf))
            {
                StreamReader sr2 = new StreamReader(out_inf);
                string lines;
                while((lines = sr2.ReadLine()) != null)
                {
                    textBox2.AppendText(lines + Environment.NewLine);
                }
                sr2.Close();
            }
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            label4.Text = fthreeBytes + trackBar1.Value.ToString();
            trackBar2.Minimum = trackBar1.Value;
        }

        private void trackBar2_Scroll(object sender, EventArgs e)
        {
            label5.Text = fthreeBytes + trackBar2.Value.ToString();
            trackBar1.Maximum = trackBar2.Value;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (trackBar1.Minimum < trackBar1.Value) trackBar1.Value--;
            label4.Text = fthreeBytes + trackBar1.Value.ToString();
            trackBar2.Minimum = trackBar1.Value;
            trackBar1.Maximum = trackBar2.Value;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (trackBar1.Maximum > trackBar1.Value) trackBar1.Value++;
            label4.Text = fthreeBytes + trackBar1.Value.ToString();
            trackBar2.Minimum = trackBar1.Value;
            trackBar1.Maximum = trackBar2.Value;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (trackBar2.Minimum < trackBar2.Value) trackBar2.Value--;
            label5.Text = fthreeBytes + trackBar2.Value.ToString();
            trackBar2.Minimum = trackBar1.Value;
            trackBar1.Maximum = trackBar2.Value;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (trackBar2.Maximum > trackBar2.Value) trackBar2.Value++;
            label5.Text = fthreeBytes + "." + trackBar2.Value.ToString();
            trackBar2.Minimum = trackBar1.Value;
            trackBar1.Maximum = trackBar2.Value;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label4.Text = fthreeBytes + "0";
            label5.Text = fthreeBytes + "255";
        }

        private void button9_Click(object sender, EventArgs e)
        {
            File.Delete(Application.StartupPath + "\\Programs_report.docx");
            File.Copy(Application.StartupPath + "\\Programs_report_blank.docx", Application.StartupPath + "\\Programs_report.docx");
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = null;
            object fileName = Application.StartupPath + "\\Programs_report.docx";
            object falseValue = false;
            object trueValue = true;
            object missing = Type.Missing;
            doc = app.Documents.Open(ref fileName, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing);
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();
            wdReplace(app, "<now>", DateTime.Now.ToString());
            Microsoft.Office.Interop.Word.Table tbl = app.ActiveDocument.Tables[1];
            for (int i = 0; i < textBox2.Lines.Count() - 1; i++)
            {
                tbl.Rows.Add();
                tbl.Rows[i + 2].Cells[1].Range.Text = (i + 1).ToString();
                tbl.Rows[i + 2].Cells[2].Range.Text = textBox2.Lines[i].ToString().Split('/')[0];
                tbl.Rows[i + 2].Cells[3].Range.Text = textBox2.Lines[i].ToString().Split('/')[1];

            }
            app.ActiveDocument.Save();
            app.Visible = true;
        }
    }
}
