using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.Administration;
using Binding = Microsoft.Web.Administration.Binding;

namespace IISUtil_DefaultWebsiteSelector
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == WindowState)
                Hide();
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //Show();
            //WindowState = FormWindowState.Normal;
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadContextMenu();
        }

        private void LoadContextMenu()
        {
            contextMenuStrip1.Items.Clear();

            var header = new ToolStripLabel("Manage Default Website Binding");
            contextMenuStrip1.Items.Add(header);
            contextMenuStrip1.Items.Add(new ToolStripSeparator());

            var dropDownButton = new ToolStripDropDownButton("No default website active on :80");
            using (var serverManager = new ServerManager())
            {
                if (IsValidSecurityContext(serverManager)) return;

                // Get default website IP
                var activeSite = serverManager.Sites.FirstOrDefault(
                    s =>
                    s.Bindings.Any(b => b.Protocol == "http" && string.IsNullOrWhiteSpace(b.Host)) &&
                    s.State == ObjectState.Started);

                if (activeSite != null)
                {
                    dropDownButton.Text = activeSite.Name;

                    var activeBinding =
                        activeSite.Bindings.FirstOrDefault(
                            b => b.Protocol == "http" && string.IsNullOrWhiteSpace(b.Host));

                    if (activeBinding != null)
                    {
                        LoadContextMenuWithCurrentIPBindings(activeBinding);
                    }
                }

                foreach (var site in serverManager.Sites)
                {
                    var siteStrip = new ToolStripMenuItem(string.Format("{0} ({1} on {2})"
                                                                        , site.Name
                                                                        , site.State
                                                                        , GetBindingDescription(site)));

                    if (site == activeSite)
                    {
                        siteStrip.Checked = true;
                    }

                    if (site.State != ObjectState.Started)
                    {
                        siteStrip.BackColor = Color.DimGray;
                        siteStrip.Checked = false;
                    }

                    siteStrip.Click += delegate(object o, EventArgs args) { SetRawIpBinding(site); };


                    dropDownButton.DropDownItems.Add(siteStrip);
                }
            }
            contextMenuStrip1.Items.Add(dropDownButton);
            contextMenuStrip1.Items.Add(new ToolStripSeparator());
            contextMenuStrip1.Items.Add(new ToolStripMenuItem("About", null,
                                                              (o, args) =>
                                                              Process.Start(
                                                                  "http://github.com/1508/IISUtil-DefaultWebsiteSelector")));
            contextMenuStrip1.Items.Add(new ToolStripMenuItem("Quit", null, (o, args) => Close()));
        }

        private string GetBindingDescription(Site site)
        {
            if (!site.Bindings.Any())
                return string.Empty;

            var binding = site.Bindings.Select(b => b.BindingInformation).ToArray();
            return binding.Any() ? string.Join(";", binding) : string.Empty;
        }

        private void LoadContextMenuWithCurrentIPBindings(Binding activeBinding)
        {
            if (activeBinding.EndPoint.Address.ToString() != "0.0.0.0")
            {
                var textBox = new ToolStripTextBox();
                textBox.TextBox.AcceptsReturn = false;
                textBox.TextBox.AcceptsTab = false;
                textBox.TextBox.Text = activeBinding.EndPoint.Address.ToString();
                contextMenuStrip1.Items.Add(textBox);
            }
            else
            {
                IPHostEntry host = Dns.GetHostEntry(Dns.GetHostName());
                foreach (IPAddress ip in host.AddressList)
                {
                    if (ip.AddressFamily == AddressFamily.InterNetwork)
                    {
                        var textBox = new ToolStripTextBox();
                        textBox.TextBox.AcceptsReturn = false;
                        textBox.TextBox.AcceptsTab = false;
                        textBox.TextBox.Text = ip.ToString();
                        contextMenuStrip1.Items.Add(textBox);
                    }
                }
            }
        }

        private bool IsValidSecurityContext(ServerManager serverManager)
        {
            try
            {
                var validatedElevatePriviledges = serverManager.Sites.FirstOrDefault();
            }
            catch (UnauthorizedAccessException)
            {
                MessageBox.Show(
                    "Utility must be started with elevated administrator priviledges\nIt is required for allowing reading and changing the IIS configuration.",
                    "Security Exception occured starting IISUtil");
                Close();
                return true;
            }
            catch (Exception exception)
            {
                MessageBox.Show("Do you have IIS installed ?\n\n" + exception.ToString(),
                                "Exception occured starting IISUtil");
                Close();
                return true;
            }
            return false;
        }

        public void SetRawIpBinding(Site selectedSite)
        {
            using (var serverManager = new ServerManager())
            {
                // Binding needs to be removed prior to setting it up on the selected site. 
                foreach (var site in serverManager.Sites)
                {
                    var rawBinding = site.Bindings.FirstOrDefault(binding => binding.Protocol == "http" && string.IsNullOrWhiteSpace(binding.Host));
                    if (rawBinding == null) continue;

                    if (site.Bindings.Count != 1)
                    {
                        site.Bindings.Remove(rawBinding);
                        serverManager.CommitChanges();
                    }
                    else
                    {
                        // We want to always leave at least one binding on a site.
                        site.Stop();
                    }
                }

                // the selectedSite must be matched with a live site from serverManager
                foreach (var site in serverManager.Sites.Where(site => site.Id == selectedSite.Id))
                {
                    // Since we want to leave at least one binding it might be the raw ip binding of the current active site.på 
                    if (!site.Bindings.Any(binding => binding.Protocol == "http" && string.IsNullOrWhiteSpace(binding.Host)))
                    {
                        site.Bindings.Add("*:80:", "http");
                        serverManager.CommitChanges();
                    }
                    
                    if (site.State != ObjectState.Started)
                    {
                        site.Start();
                    }

                    break;
                }
            }
            // Reset the menu
            LoadContextMenu();
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            Hide();
        }
    }
}
