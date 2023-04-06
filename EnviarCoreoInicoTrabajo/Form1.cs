using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using EnviarCoreoInicoTrabajo.Properties;

namespace EnviarCoreoInicoTrabajo
{
    public partial class Form1 : Form
    {
        DateTime fechaActiva;
        public Form1()
        {
            InitializeComponent();
            TSTBHora.Text = Settings.Default.hora.ToString();
            fechaActiva = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 18, 00, 0);
        }

        public void SendEmailWithOutlook(string mailDirection, string mailCopia, string mailSubject, string mainContent)
        {
            try
            {
                var oApp = new Microsoft.Office.Interop.Outlook.Application();

                Outlook.NameSpace ns = oApp.GetNamespace("MAPI");
                var f = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

                System.Threading.Thread.Sleep(1000);

                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                mailItem.Subject = mailSubject;
                mailItem.HTMLBody = mainContent;
                mailItem.To = mailDirection;
                mailItem.CC = mailCopia;
                mailItem.Send();
                MessageBox.Show("Mensaje Envíado...", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                timer1.Enabled = false;
                activarTemporizadorToolStripMenuItem.Enabled = true;
            }
            catch
            {
                MessageBox.Show("Mensaje Fellido...", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Settings.Default.enviado = false;
                Settings.Default.Save();
                timer1.Enabled = true;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            textBox1.Text += Settings.Default.enviado.ToString();
            textBox2.Text = Settings.Default.fecha;
            textBox3.Text = DateTime.Today.ToString("d");
            if (Settings.Default.fecha != DateTime.Today.ToString("d")) 
            {
                Settings.Default.fecha = DateTime.Today.ToString("d");
                Settings.Default.Save();
                enviarCorreoToolStripMenuItem.Enabled = true;
                Settings.Default.enviado = false;
                Settings.Default.Save();
            }
            if ((DateTime.Now.ToString("HH:mm") == Settings.Default.hora.ToString()) && !Settings.Default.enviado )
            {
                //string mailDirection = @"Jhonny Paul Bedoya Guerrero <jhonny.bedoya@stracon.com>; Carmen Gavilan Vargas <carmen.gavilan@stracon.com>; Manuel Rodriguez Gutierrez <manuel.rodriguez@stracon.com>";
                //string mailCopia = @"Gustavo Gomez Llanos <gustavo.gomez@stracon.com>; Ricardo Torres Inca <ricardo.torres@stracon.com>; Victor Antonio Calle Angeles <victor.calle@stracon.com>";

                string mailDirection = @"Gregorio Jose Oduber Quevedo <gregorio.oduber@stracon.com>";
                string mailCopia = @"Gregorio Jose Oduber Quevedo <gregorio.oduber@stracon.com>";

                string mailSubject = "Inicio de Actividades del " + DateTime.Now.ToString("d");

                string mainContent;
                mainContent = "Estimados,";
                mainContent += "<br>" + "<br>" + "Buenos Dias, Reportando Inicio de actividades Home Office el " + DateTime.Now.ToString("d");
                mainContent += "<br>" + "<br>" + "Saludos,";
                mainContent += "<br>" + "Gregorio Jose Oduber Quevedo";

                SendEmailWithOutlook(mailDirection, mailCopia, mailSubject, mainContent);
                Settings.Default.enviado = true;
                Settings.Default.Save();
                enviarCorreoToolStripMenuItem.Enabled = false;
            }
            //if (enviado == true)
            //{
            if (DateTime.Now.ToString("HH:mm") != Settings.Default.hora.ToString() && Settings.Default.enviado)
                enviarCorreoToolStripMenuItem.Enabled = false;
            //}
            if (!Settings.Default.enviado && enviarCorreoToolStripMenuItem.Enabled == false)
                enviarCorreoToolStripMenuItem.Enabled = true;
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            notifyIcon1.Icon = this.Icon;
            notifyIcon1.ContextMenuStrip = contextMenuStrip1;
            notifyIcon1.Text = Application.ProductName;
            notifyIcon1.Visible = true;
            this.Visible = false;
        }

        private void salirToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
                this.Visible = false;
        }

        private void enviarCorreoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Confirmar Envío de Correo Electronico", "Correo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                string mailDirection = @"Jhonny Paul Bedoya Guerrero <jhonny.bedoya@stracon.com>; Carmen Gavilan Vargas <carmen.gavilan@stracon.com>; Manuel Rodriguez Gutierrez <manuel.rodriguez@stracon.com>";
                string mailCopia = @"Gustavo Gomez Llanos <gustavo.gomez@stracon.com>; Ricardo Torres Inca <ricardo.torres@stracon.com>; Victor Antonio Calle Angeles <victor.calle@stracon.com>";
                string mailSubject = "Inicio de Actividades del " + DateTime.Now.ToString("d");

                string mainContent;
                mainContent = "Estimados,";
                mainContent += "<br>" + "<br>" + "Buenos Dias, Reportando Inicio de actividades Home Office el " + DateTime.Now.ToString("d");
                mainContent += "<br>" + "<br>" + "Saludos,";
                mainContent += "<br>" + "Gregorio Jose Oduber Quevedo";

                SendEmailWithOutlook(mailDirection, mailCopia, mailSubject, mainContent);
                Settings.Default.enviado = true;
                enviarCorreoToolStripMenuItem.Enabled = false;
            }
            else
                Settings.Default.enviado = false;
            Settings.Default.Save();
        }

        private void TSTBHora_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (TSTBHora.ToString().Length > 0)
                {
                    Settings.Default.hora = TSTBHora.Text.ToString();
                    Settings.Default.Save();
                }
                else
                {
                    MessageBox.Show("Especificar Hora para enviar correo:", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Settings.Default.hora = "07:00";
                    Settings.Default.Save();
                }
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (DateTime.Now > fechaActiva && !timer2.Enabled)
            {
                timer1.Enabled = true;
                activarTemporizadorToolStripMenuItem.Enabled = false;
            }
        }

        private void activarTemporizadorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            activarTemporizadorToolStripMenuItem.Enabled = false;
        }
    }
}
