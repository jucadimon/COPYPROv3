using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using iTextSharp.text;
using iTextSharp.text.pdf;
using COPYPROv2.Properties;

namespace COPYPROv2
{
    public partial class FormCopypro : Form
    {
        public FormCopypro()
        {
            InitializeComponent();
            rbNormal.Tag = new[] {Color.SteelBlue, Color.Snow};
            rbIronMan.Tag = new[] {Color.DarkRed, Color.Gold};
            rbJarvis.Tag = new[] {Color.Black, Color.SteelBlue};
            rbCereza.Tag = new[] {Color.MintCream, Color.MediumOrchid};
            rbRadar.Tag = new[] {Color.Black, Color.DarkCyan};
            rbGris.Tag = new[] {Color.Gray, Color.DarkSlateGray};
            rbVainilla.Tag = new[] {Color.BlanchedAlmond, Color.Blue};
        }
     
        // Para mover la ventana con el mouse inicio
        // descargado desde la pagina:
        // http://csharpmaniax.blogspot.com.co/2012/05/como-mover-form-sin-bordes.html 
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private static extern void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private static extern void SendMessage(IntPtr hWnd, int wMsg, int wParam, int lParam);

        private void MostrarInfo_Click(object sender, EventArgs e)
        {
            MostrarInfo();
        }

        private void MostrarInfo()
        {
            var infoForm = new Form {
                Name = "infoForm",
                AutoScaleBaseSize = new Size(5, 13),
                ClientSize = new Size(640, 480),
                Opacity = 0.85,
                BackColor = txtCaja.BackColor,
                ForeColor = txtCaja.ForeColor,
                FormBorderStyle = FormBorderStyle.None,
                StartPosition = FormStartPosition.CenterScreen
            };

            var salirButton = new Button {
                Name = "salirButton",
                Text = Resources.CloseButtonText,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(45, 25),
                Font = new System.Drawing.Font("Segoe UI", 8F, FontStyle.Bold),
                Location = new Point(585, 10),
                FlatStyle = FlatStyle.Flat
            };
            infoForm.Controls.Add(salirButton);
            infoForm.CancelButton = salirButton;

            var tituloLabel = new Label {
                Name = "tituloLabel",
                Text = Resources.InfoTitulo,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(400, 25),
                Font = new System.Drawing.Font("Segoe UI", 16F, FontStyle.Bold)
            };

            tituloLabel.Location = new Point(320 - (tituloLabel.Width / 2), 10);
            infoForm.Controls.Add(tituloLabel);

            var infoLabel = new Label {
                Name = "infoLabel",
                Text = Resources.InfoText,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(580, 380),
                Location = new Point(40, 35),
                Font = new System.Drawing.Font("Segoe UI", 8F, FontStyle.Bold)
            };

            infoForm.Controls.Add(infoLabel);
            infoForm.ShowDialog();
        }

        private void trackBar1_ValueChanged(object sender, EventArgs e)
        {
            Opacity = trackBar1.Value / 100d;
        }

        private void btnMinimizar_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void btnMaximizar_Click(object sender, EventArgs e)
        {
            switch (WindowState) {
                case FormWindowState.Maximized:
                    WindowState = FormWindowState.Normal;
                    break;
                case FormWindowState.Normal:
                    WindowState = FormWindowState.Maximized;
                    break;
            }
        }

        private void Salir_Click(object sender, EventArgs e)
        {
            MostrarMensaje(Resources.MensajeVuelvaPronto);
            Close();
        }

        private void FormCopypro_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture(); 
            SendMessage(Handle, 0x112, 0xf012, 0); 
        }

        private void abrirTxtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AbrirArchivoDeTexto();
        }

        private void AbrirArchivoDeTexto()
        {
            openFileDialog.FileName = Resources.SeleccioneArchivoDeTexto;
            openFileDialog.Filter = Resources.ArchivosDeTextoFilter;
            openFileDialog.Title = Resources.TituloAbrirArchivoDeTexto;
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                using (var sr = File.OpenText(openFileDialog.FileName))
                {
                    txtCaja.Text = sr.ReadToEnd();
                }             
            }
        }

        private void aTxtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExportarATxt();
        }

        private void ExportarATxt()
        {
            saveFileDialog.Title = Resources.TituloGuardarArchivoDeTexto;
            saveFileDialog.Filter = Resources.ArchivosDeTextoFilter;
            saveFileDialog.DefaultExt = "txt";
            saveFileDialog.AddExtension = true;
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                using (var sw = File.AppendText(saveFileDialog.FileName))
                {
                    sw.Write(txtCaja.Text);
                }

                MostrarMensaje(Resources.ExportacionExitosa);
            }
        }

        private void aPdfToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExportarAPdf();
        }

        private void ExportarAPdf()
        {
            saveFileDialog.Title = Resources.ExportarPdfTitle;
            saveFileDialog.Filter = Resources.ArchivoPdfFilter;
            saveFileDialog.DefaultExt = "pdf";
            saveFileDialog.AddExtension = true;
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                var doc = new Document(PageSize.LETTER, 60, 30, 50, 30);

                var writer = PdfWriter.GetInstance(doc,
                    new FileStream(saveFileDialog.FileName,
                        FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite));

                // Nota: Esto no será visible en el documento
                doc.AddTitle("INFORME PDF");
                doc.AddCreator("COPYPRO");

                doc.Open(); 
               
                var standardFont =
                    new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8,
                        iTextSharp.text.Font.NORMAL, BaseColor.BLACK);

                doc.Add(new Paragraph(string.Empty, standardFont));  
                doc.Add(Chunk.NEWLINE);  
 
                var parrafo = new Paragraph {
                    Alignment = Element.ALIGN_JUSTIFIED,
                    Font = standardFont
                };  
                parrafo.Add(txtCaja.Text);  
                doc.Add(parrafo); 

                doc.Close(); 
                writer.Close();

                MostrarMensaje(Resources.ExportacionExitosa);
            }

        }

        private void LimpiarTexto_Click(object sender, EventArgs e)
        {
            txtCaja.Clear();
        }

        private void MostrarMensaje(string mensaje)
        {
            const int ancho = 320, alto = 160;

            var mensajeForm = new Form {
                AutoScaleBaseSize = new Size(5, 13),
                ClientSize = new Size(ancho, alto),
                Opacity = .90,
                BackColor = txtCaja.BackColor,
                ForeColor = txtCaja.ForeColor,
                FormBorderStyle = FormBorderStyle.None,
                Name = "messageForm",
                StartPosition = FormStartPosition.CenterScreen
            };

            var aceptarButton = new Button {
                Font = new System.Drawing.Font("Segoe UI", 8F, FontStyle.Bold),
                Location = new Point(ancho/2 - 75/2, alto - 35),
                Name = "aceptarButton",
                Size = new Size(75, 25),
                Text = Resources.AcceptButtonText,
                TextAlign = ContentAlignment.MiddleCenter,
                FlatStyle = FlatStyle.Flat
            };

            mensajeForm.Controls.Add(aceptarButton);
            mensajeForm.CancelButton = aceptarButton;

            var tituloLabel = new Label {
                Name = "tituloLabel",
                Text = Resources.MensajeTitulo,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(400, 25),
                Location = new Point(ancho/2 - 400/2, 10),
                Font = new System.Drawing.Font("Segoe UI", 16F, FontStyle.Bold)
            };

            mensajeForm.Controls.Add(tituloLabel);

            var infoLabel = new Label {
                Name = "infoLabel",
                Text = mensaje,
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(ancho - 100, alto - 100),
                Font = new System.Drawing.Font("Segoe UI", 8F, FontStyle.Bold)
            };

            infoLabel.Location= new Point(mensajeForm.Width / 2 - infoLabel.Width / 2, mensajeForm.Height / 2 - infoLabel.Height / 2);

            mensajeForm.Controls.Add(infoLabel);      
            mensajeForm.ShowDialog();
        }

        private void radioTheme_CheckedChanged(object sender, EventArgs e)
        {
            var radioButton = sender as RadioButton;
            if (radioButton != null)
            {
                var colors = (Color[])radioButton.Tag;
                SetColors(colors[0], colors[1]);
            }    
        }

        private void SetColors(Color backColor, Color foreColor)
        {
            BackColor = backColor;
            ForeColor = foreColor;
            txtCaja.BackColor = backColor;
            txtCaja.ForeColor = foreColor;
            menuStrip1.BackColor = foreColor;
            menuStrip1.ForeColor = backColor;
            menuToolStripMenuItem.ForeColor = backColor;
            abrirTxtToolStripMenuItem.BackColor = foreColor;
            abrirTxtToolStripMenuItem.ForeColor = backColor;
            limpiarToolStripMenuItem.BackColor = foreColor;
            limpiarToolStripMenuItem.ForeColor = backColor;
            infoToolStripMenuItem.BackColor = foreColor;
            infoToolStripMenuItem.ForeColor = backColor;
            salirToolStripMenuItem.BackColor = foreColor;
            salirToolStripMenuItem.ForeColor = backColor;
            exportarToolStripMenuItem.ForeColor = backColor;
            aTxtToolStripMenuItem.BackColor = foreColor;
            aTxtToolStripMenuItem.ForeColor = backColor;
            aPdfToolStripMenuItem.BackColor = foreColor;
            aPdfToolStripMenuItem.ForeColor = backColor;
        }

    }
}
