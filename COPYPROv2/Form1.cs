using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using iTextSharp; // Exoprtar a pdf con itextSharp - inicio
using iTextSharp.awt;
using iTextSharp.testutils;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.io;
using iTextSharp.xmp; // Exoprtar a pdf con itextSharp - fin

using System.Runtime.InteropServices;//para mover la ventana con el mouse

namespace COPYPROv2
{
    public partial class FormCopypro : Form
    {
        public string ostipo = "", tipoArchivo = "";
        public bool SistemaOperativo, cambio;
        public StringBuilder archivoGenerado;
        //Creamos el escritor de archivos y damos la ruta, nombre del archivo y modo de creado
        FileStream archivoGeneradoA;
        string directorio1;

        public FormCopypro()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //variables de inicio del formulario
            this.BackColor = Color.SteelBlue;
            this.ForeColor = Color.Snow;
            txtCaja.BackColor = Color.SteelBlue;
            txtCaja.ForeColor = Color.Snow;
            menuStrip1.BackColor = Color.Snow;
            menuStrip1.ForeColor = Color.SteelBlue;
            //primer  menu
            menuToolStripMenuItem.ForeColor = Color.SteelBlue;

            abrirTxtToolStripMenuItem.BackColor = Color.Snow;
            abrirTxtToolStripMenuItem.ForeColor = Color.SteelBlue;
            limpiarToolStripMenuItem.BackColor = Color.Snow;
            limpiarToolStripMenuItem.ForeColor = Color.SteelBlue;
            infoToolStripMenuItem.BackColor = Color.Snow;
            infoToolStripMenuItem.ForeColor = Color.SteelBlue;
            salirToolStripMenuItem.BackColor = Color.Snow;
            salirToolStripMenuItem.ForeColor = Color.SteelBlue;
            //menu exportar
            exportarToolStripMenuItem.ForeColor = Color.SteelBlue;

            aTxtToolStripMenuItem.BackColor = Color.Snow;
            aTxtToolStripMenuItem.ForeColor = Color.SteelBlue;
            aPdfToolStripMenuItem.BackColor = Color.Snow;
            aPdfToolStripMenuItem.ForeColor = Color.SteelBlue;
        }

        /// 
        /// //para mover la ventana con el mouse inicio
        /// descargado desde la pagina:
        /// http://csharpmaniax.blogspot.com.co/2012/05/como-mover-form-sin-bordes.html
        /// 
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);
        /// 
        /// //para mover la ventana con el mouse fin
        /// 
        private void btnInfo_Click(object sender, EventArgs e)
        {
            MostrarInfo();
        }
        private void MostrarInfo()
        {
            Button bsalirSub = new Button();
            Label Titulo_SubFormulario = new Label();
            Label Informacion = new Label();
            Form Subformulario = new Form();
            //  datos del boton salir
            bsalirSub.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            bsalirSub.Location = new Point(585, 10);
            bsalirSub.Name = "bsalirSub";
            bsalirSub.Size = new System.Drawing.Size(45, 25);
            bsalirSub.Text = "X";
            bsalirSub.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;//MiddleCenter/TopCenter
            bsalirSub.FlatStyle = FlatStyle.Flat;
            //  datos del label titulo
            Titulo_SubFormulario.Name = "Titulo_SubFormulario";
            Titulo_SubFormulario.Text = "I N F O R M A C I O N";
            Titulo_SubFormulario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;//MiddleCenter/TopCenter
            Titulo_SubFormulario.Size = new Size(400, 25);//Tamaño del label
            Titulo_SubFormulario.Location = new Point(320 - (Titulo_SubFormulario.Width / 2), 10);//localizacion del label en el formulario
            Titulo_SubFormulario.Font = new System.Drawing.Font("Segoe UI", 16F, System.Drawing.FontStyle.Bold);
            //  datos del label Informacion
            Informacion.Name = "Informacion";
            Informacion.Text = "COPYPRO v3"+
                "\n\nINSTRUCCIONES DE USO: " +
                "\nLocalice la informacion que desea escribir desde un archivo, pagina web, "
                + "u otro programa, a continuación abra el programa COPYPRO y escriba dentro de la caja de "
                + "texto y al final de clic en el boton exportar a un archivo de texto a su escritorio."
                + "\nAl presionar el boton ExportarTXT se crea un archivo de texto multiple con el historial "
                + "del contenido de la caja de texto."
                + "\n"
                + "Puede dar clic en el boton limpiar para borrar todo el contenido de la caja de texto."
                + "\nCOPYPRO: es un pequeño proyecto sobre un programa para hacer resumenes a mano, es un editor de texto minimalista "
                + "que te permite ir guardando informacion por lotes de por ej: una cancion, un resumen de una novela literaria, codigo fuente, etc; "
                + "con eL programa podrás leer y escribir al mismo tiempo... o esa es la idea."
                + "\nPuede utilizar el programa de cualquier forma que se le ocurra, por favor, si encuentra utilidad, "
                + "escribanos a nuestro correo, con asunto copypro y agregue la pagina del programa en la red social facebook: COPYPRO"
                + "\nPara donativos :v escribanos al correo."

                + "\n\nSOFTWARE DE ESCRITURA RAPIDA"
            + "\n\nSOFTWARE CREADO POR:\nJUANCARLOS DIAZ MONTIEL\nESTUDIANTE INGENIERIA CIVIL" +
            "\nUNIVERSIDAD DE SUCRE\n2016\nCEL: 3016922644\nCORREO: jucadimon@hotmail.com\nWEB: " +
            "http://ingenieria-civil-y-pi.blogspot.com.co/ "
            + "\n\nUnete al mejor grupo de programación en facebook: '.NET PROGRAMADORES'";
            Informacion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;//MiddleCenter/TopCenter
            Informacion.Size = new Size(580, 380);//Tamaño del label
            Informacion.Location = new Point(40, 35);//localizacion del label en el formulario
            Informacion.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            //  datos de la ventana
            Subformulario.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            Subformulario.ClientSize = new System.Drawing.Size(640, 480);
            Subformulario.Opacity = 0.85;
            //tomamos el color de un control existente en el form principal y lo pasamos a este form
            Subformulario.BackColor = txtCaja.BackColor;
            Subformulario.ForeColor = txtCaja.ForeColor;
            //Subformulario.BackColor = Color.MintCream;//cambia color de fondo, Black, MintCream, MediumOrchid, 
            //Subformulario.ForeColor = Color.MediumOrchid;//cambia color de la fuente SlateBlue
            Subformulario.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;//FixedDialog, None, Fixed3D
            Subformulario.Name = "Subformulario";
            Subformulario.Text = "SUB VENTANA X 1.0";
            Subformulario.CancelButton = bsalirSub;
            Subformulario.StartPosition = FormStartPosition.CenterScreen;
            //
            Subformulario.Controls.Add(bsalirSub);
            Subformulario.Controls.Add(Titulo_SubFormulario);
            Subformulario.Controls.Add(Informacion);
            //
            Subformulario.ShowDialog();
        }

        private void trackBar1_ValueChanged(object sender, EventArgs e)
        {
            //CAMBIAMOS LA OPACIDAD DEL FORMULARIO
            this.Opacity = trackBar1.Value / 100d;
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            // ExportarATxtDirectoAlEscritorio(); 
        }
        private void ExportarATxtDirectoAlEscritorio()
        {
            //          EXPORTAR TXT 
            ostipo = Environment.OSVersion.ToString();//saber el sistema operativo que se usa
            SistemaOperativo = ostipo.Contains("Unix");

            archivoGenerado = new StringBuilder();//Definimos un comodin publico
            // string fecha = "";
            // fecha = DateTime.Now.ToString();    //pedimos la hora al sistema y almacenamos en fecha

            archivoGenerado.AppendLine(txtCaja.Text);//añadimos la fecha al comodin

            String reg = "";    //creamos un objeto tipo String del espacio de nombres Text
            reg = archivoGenerado.ToString();    //convertimos el comodin a texto

            if (SistemaOperativo)
            {
                directorio1 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/texto.txt";
                // string directorio2 = Environment.CurrentDirectory + "/texto"+fecha+".txt";
            }
            else
            {
                directorio1 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\texto.txt";
                // string directorio2 = Environment.CurrentDirectory + "\\texto"+fecha+".txt";
            }

            archivoGeneradoA = new FileStream(directorio1, FileMode.Append);
            archivoGeneradoA.Write(Encoding.GetEncoding("windows-1252").GetBytes(reg), 0, reg.Length);//escribimos dentro del archivo
            archivoGeneradoA.Close();//Cerra el archivo creado
            // Encoding.GetEncoding("windows-1252")     cambiar esta linea en los programas de exportar a txt 
            // para solucionar le problema de la stildes y de las "ñ"
            // archivoGeneradoA.Write(ASCIIEncoding.ASCII.GetBytes(reg), 0, reg.Length); Esta es la linea vieja

            MessageBox.Show("Exportación correcta! Revise el archivo en su escritorio.");
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            txtCaja.Text = "";
        }

        private void btnMinimizar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;//Minimized, Maximized, Normal
        }

        private void btnMaximizar_Click(object sender, EventArgs e)
        {
            
            if(cambio)
            {
                this.WindowState = FormWindowState.Normal;//Minimized, Maximized, Normal
                cambio = false;
            }
            else
            {
                this.WindowState = FormWindowState.Maximized;//Minimized, Maximized, Normal
                cambio = true;
            }
            
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            MostrarMensaje("Va a salir del programa, vuelva pronto!");
            Application.Exit();//Codigo para boton salir
        }

        private void FormCopypro_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();/// para mover la ventana con el mouse
            SendMessage(this.Handle, 0x112, 0xf012, 0);/// para mover la ventana con el mouse
        }

        private void rbNormal_CheckedChanged(object sender, EventArgs e)
        {
            this.BackColor = Color.SteelBlue;
            this.ForeColor = Color.Snow;
            txtCaja.BackColor = Color.SteelBlue;
            txtCaja.ForeColor = Color.Snow;
            menuStrip1.BackColor = Color.Snow;
            menuStrip1.ForeColor = Color.SteelBlue;
            //primer  menu
            menuToolStripMenuItem.ForeColor = Color.SteelBlue;

            abrirTxtToolStripMenuItem.BackColor = Color.Snow;
            abrirTxtToolStripMenuItem.ForeColor = Color.SteelBlue;
            limpiarToolStripMenuItem.BackColor = Color.Snow;
            limpiarToolStripMenuItem.ForeColor = Color.SteelBlue;
            infoToolStripMenuItem.BackColor = Color.Snow;
            infoToolStripMenuItem.ForeColor = Color.SteelBlue;
            salirToolStripMenuItem.BackColor = Color.Snow;
            salirToolStripMenuItem.ForeColor = Color.SteelBlue;
            //menu exportar
            exportarToolStripMenuItem.ForeColor = Color.SteelBlue;

            aTxtToolStripMenuItem.BackColor = Color.Snow;
            aTxtToolStripMenuItem.ForeColor = Color.SteelBlue;
            aPdfToolStripMenuItem.BackColor = Color.Snow;
            aPdfToolStripMenuItem.ForeColor = Color.SteelBlue;
        }

        private void rbGris_CheckedChanged(object sender, EventArgs e)
        {
            this.BackColor = Color.Gray;
            this.ForeColor = Color.DarkSlateGray;
            txtCaja.BackColor = Color.Gray;
            txtCaja.ForeColor = Color.DarkSlateGray;
            menuStrip1.BackColor = Color.DarkSlateGray;
            menuStrip1.ForeColor = Color.Gray;
            //primer  menu
            menuToolStripMenuItem.ForeColor = Color.Gray;

            abrirTxtToolStripMenuItem.BackColor = Color.DarkSlateGray;
            abrirTxtToolStripMenuItem.ForeColor = Color.Gray;
            limpiarToolStripMenuItem.BackColor = Color.DarkSlateGray;
            limpiarToolStripMenuItem.ForeColor = Color.Gray;
            infoToolStripMenuItem.BackColor = Color.DarkSlateGray;
            infoToolStripMenuItem.ForeColor = Color.Gray;
            salirToolStripMenuItem.BackColor = Color.DarkSlateGray;
            salirToolStripMenuItem.ForeColor = Color.Gray;
            //menu exportar
            exportarToolStripMenuItem.ForeColor = Color.Gray;

            aTxtToolStripMenuItem.BackColor = Color.DarkSlateGray;
            aTxtToolStripMenuItem.ForeColor = Color.Gray;
            aPdfToolStripMenuItem.BackColor = Color.DarkSlateGray;
            aPdfToolStripMenuItem.ForeColor = Color.Gray;
        }

        private void rbRadar_CheckedChanged(object sender, EventArgs e)
        {
            this.BackColor = Color.Black;
            this.ForeColor = Color.DarkCyan;
            txtCaja.BackColor = Color.Black;
            txtCaja.ForeColor = Color.DarkCyan;
            menuStrip1.BackColor = Color.DarkCyan;
            menuStrip1.ForeColor = Color.Black;
            //primer  menu
            menuToolStripMenuItem.ForeColor = Color.Black;

            abrirTxtToolStripMenuItem.BackColor = Color.DarkCyan;
            abrirTxtToolStripMenuItem.ForeColor = Color.Black;
            limpiarToolStripMenuItem.BackColor = Color.DarkCyan;
            limpiarToolStripMenuItem.ForeColor = Color.Black;
            infoToolStripMenuItem.BackColor = Color.DarkCyan;
            infoToolStripMenuItem.ForeColor = Color.Black;
            salirToolStripMenuItem.BackColor = Color.DarkCyan;
            salirToolStripMenuItem.ForeColor = Color.Black;
            //menu exportar
            exportarToolStripMenuItem.ForeColor = Color.Black;

            aTxtToolStripMenuItem.BackColor = Color.DarkCyan;
            aTxtToolStripMenuItem.ForeColor = Color.Black;
            aPdfToolStripMenuItem.BackColor = Color.DarkCyan;
            aPdfToolStripMenuItem.ForeColor = Color.Black;
        }

        private void rbCereza_CheckedChanged(object sender, EventArgs e)
        {
            this.BackColor = Color.MintCream;
            this.ForeColor = Color.MediumOrchid;
            txtCaja.BackColor = Color.MintCream;
            txtCaja.ForeColor = Color.MediumOrchid;
            menuStrip1.BackColor = Color.MediumOrchid;
            menuStrip1.ForeColor = Color.MintCream;
            //primer  menu
            menuToolStripMenuItem.ForeColor = Color.MintCream;

            abrirTxtToolStripMenuItem.BackColor = Color.MediumOrchid;
            abrirTxtToolStripMenuItem.ForeColor = Color.MintCream;
            limpiarToolStripMenuItem.BackColor = Color.MediumOrchid;
            limpiarToolStripMenuItem.ForeColor = Color.MintCream;
            infoToolStripMenuItem.BackColor = Color.MediumOrchid;
            infoToolStripMenuItem.ForeColor = Color.MintCream;
            salirToolStripMenuItem.BackColor = Color.MediumOrchid;
            salirToolStripMenuItem.ForeColor = Color.MintCream;
            //menu exportar
            exportarToolStripMenuItem.ForeColor = Color.MintCream;

            aTxtToolStripMenuItem.BackColor = Color.MediumOrchid;
            aTxtToolStripMenuItem.ForeColor = Color.MintCream;
            aPdfToolStripMenuItem.BackColor = Color.MediumOrchid;
            aPdfToolStripMenuItem.ForeColor = Color.MintCream;
        }

        private void rbJarvis_CheckedChanged(object sender, EventArgs e)
        {
            this.BackColor = Color.Black;
            this.ForeColor = Color.SteelBlue;
            txtCaja.BackColor = Color.Black;
            txtCaja.ForeColor = Color.SteelBlue;
            menuStrip1.BackColor = Color.SteelBlue;
            menuStrip1.ForeColor = Color.Black;
            //primer  menu
            menuToolStripMenuItem.ForeColor = Color.Black;

            abrirTxtToolStripMenuItem.BackColor = Color.SteelBlue;
            abrirTxtToolStripMenuItem.ForeColor = Color.Black;
            limpiarToolStripMenuItem.BackColor = Color.SteelBlue;
            limpiarToolStripMenuItem.ForeColor = Color.Black;
            infoToolStripMenuItem.BackColor = Color.SteelBlue;
            infoToolStripMenuItem.ForeColor = Color.Black;
            salirToolStripMenuItem.BackColor = Color.SteelBlue;
            salirToolStripMenuItem.ForeColor = Color.Black;
            //menu exportar
            exportarToolStripMenuItem.ForeColor = Color.Black;

            aTxtToolStripMenuItem.BackColor = Color.SteelBlue;
            aTxtToolStripMenuItem.ForeColor = Color.Black;
            aPdfToolStripMenuItem.BackColor = Color.SteelBlue;
            aPdfToolStripMenuItem.ForeColor = Color.Black;          
        }

        private void rbIronMan_CheckedChanged(object sender, EventArgs e)
        {
            this.BackColor = Color.DarkRed;
            this.ForeColor = Color.Gold;
            txtCaja.BackColor = Color.DarkRed;
            txtCaja.ForeColor = Color.Gold;
            menuStrip1.BackColor = Color.Gold;
            menuStrip1.ForeColor = Color.DarkRed;
            //primer  menu
            menuToolStripMenuItem.ForeColor = Color.DarkRed;

            abrirTxtToolStripMenuItem.BackColor = Color.Gold;
            abrirTxtToolStripMenuItem.ForeColor = Color.DarkRed;
            limpiarToolStripMenuItem.BackColor = Color.Gold;
            limpiarToolStripMenuItem.ForeColor = Color.DarkRed;
            infoToolStripMenuItem.BackColor = Color.Gold;
            infoToolStripMenuItem.ForeColor = Color.DarkRed;
            salirToolStripMenuItem.BackColor = Color.Gold;
            salirToolStripMenuItem.ForeColor = Color.DarkRed;
            //menu exportar
            exportarToolStripMenuItem.ForeColor = Color.DarkRed;

            aTxtToolStripMenuItem.BackColor = Color.Gold;
            aTxtToolStripMenuItem.ForeColor = Color.DarkRed;
            aPdfToolStripMenuItem.BackColor = Color.Gold;
            aPdfToolStripMenuItem.ForeColor = Color.DarkRed;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void abrirTxtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            abrirTxt();
        }
        private void abrirTxt()
        {
            //Abrir archivo de texto desde ventana
            openFileDialog1.FileName = "Seleccione un archivo de texto";
            openFileDialog1.Filter = "Archivos de texto (*.txt)|*.txt";
            openFileDialog1.Title = "Apertura de archivo de texto.";
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string nombreArchivo = "";//variable que almacena la ruta completa del archivo
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                nombreArchivo = openFileDialog1.FileName;//almacenamos la ruta completa del archivo que se va abrir
                // debemos crear us stream para leer el archivo, como parametros de entrada en el 
                // constructor: 1-ruta completa del archivo que se va abrir con un @ por delante
                // y 2-codificacion de caracteres para que pueda reconocer tildes y eñes y otros 
                // simbolos especiales
                StreamReader lector = new StreamReader(@nombreArchivo, Encoding.GetEncoding("windows-1252"));
                string leido = lector.ReadToEnd(); //pasamos todo lo que se lea a un string
                txtCaja.Text = leido;// pasamos la info del string leido al textbox :)
                lector.Close();// cerramos el stream 
            }
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MostrarMensaje("Va a salir del programa, vuelva pronto!");
            Application.Exit();//Codigo para boton salir
        }

        private void aTxtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExportarATxt();
        }
        private void ExportarATxt()
        {
            //          EXPORTAR TXT - METODO 2: saveFileDialog

            archivoGenerado = new StringBuilder();//Definimos un comodin publico
            archivoGenerado.AppendLine(txtCaja.Text);//añadimos la fecha al comodin

            String reg = "";    //creamos un objeto tipo String del espacio de nombres Text
            reg = archivoGenerado.ToString();    //convertimos el comodin a texto

            saveFileDialog1.Title = "Guardar Archivo de Texto";
            saveFileDialog1.Filter = "Archivo de Texto (.txt) |*.txt";
            saveFileDialog1.DefaultExt = "txt";
            saveFileDialog1.AddExtension = true;
            saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string directorio1 = saveFileDialog1.FileName;

                archivoGeneradoA = new FileStream(directorio1, FileMode.Append);
                archivoGeneradoA.Write(Encoding.GetEncoding("windows-1252").GetBytes(reg), 0, reg.Length);//escribimos dentro del archivo
                archivoGeneradoA.Close();//Cerra el archivo creado

                MostrarMensaje("Exportación exitosa!");
            }

            // FIN EXPORTAR
        }

        private void aPdfToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExportarAPdf();
        }
        private void ExportarAPdf()
        {
            //          EXPORTAR PDF - METODO 1: saveFileDialog

            saveFileDialog1.Title = "Guardar Archivo de Texto en PDF";
            saveFileDialog1.Filter = "Archivo de Texto (.pdf) |*.pdf";
            saveFileDialog1.DefaultExt = "pdf";
            saveFileDialog1.AddExtension = true;
            saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string directorio1 = saveFileDialog1.FileName;

                // Creamos el documento con el tamaño de página tradicional
                // ademas definimos los margenes
                iTextSharp.text.Document doc = new iTextSharp.text.Document(PageSize.LETTER, 60, 30, 50, 30);
                // Indicamos donde vamos a guardar el documento
                PdfWriter writer = PdfWriter.GetInstance(doc,
                    new FileStream(directorio1,
                        FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite));
                // Le colocamos el título y el autor
                // **Nota: Esto no será visible en el documento
                doc.AddTitle("INFORME PDF");
                doc.AddCreator("COPYPRO");

                doc.Open(); // Abrimos el archivo
                // Creamos el tipo de Font que vamos utilizar
                iTextSharp.text.Font _standardFont =
                    new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8,
                        iTextSharp.text.Font.NORMAL, BaseColor.BLACK);


                doc.Add(new Paragraph("", _standardFont)); // Escribimos el encabezamiento en el documento
                doc.Add(Chunk.NEWLINE); // salto de linea

                // Forma 1 de agregar texto:
                Paragraph parrafo = new Paragraph(); //creamos un elemento parrafo
                parrafo.Alignment = Element.ALIGN_JUSTIFIED; // lo justificamos
                parrafo.Font = _standardFont; // definimos la font del parrafo
                parrafo.Add(txtCaja.Text); //agreagmos el texto
                doc.Add(parrafo); // añadimos el elemento tipo parrafo al documento

                // Forma 2 de agregar texto:
                // doc.Add(new Paragraph(txtCaja.Text,_standardFont)); //Añadir texto a exportar

                doc.Close(); // y cerramos el documento
                writer.Close();

                MostrarMensaje("Exportación exitosa!");
            }

            // FIN EXPORTAR
        }

        private void tema1ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void infoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MostrarInfo();
        }

        private void limpiarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            txtCaja.Text = "";
        }

        private void txtCaja_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if (e.KeyChar == (char)Keys.Tab)
            //{
            //    txtCaja.SelectionStart = txtCaja.TextLength;
            //    txtCaja.Text = txtCaja.Text.Insert(txtCaja.SelectionStart, "     ");
            //    // MessageBox.Show("xd1");
            //}
        }
        private void MostrarMensaje(string MENSAJE)
        {
            int ancho = 320, alto = 160;
            Button bsalirSub = new Button();
            Label Titulo_SubFormulario = new Label();
            Label Informacion = new Label();
            Form SubformularioMensaje = new Form();
            //  datos del boton salir
            bsalirSub.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            bsalirSub.Location = new Point(ancho/2-75/2, alto-35);
            bsalirSub.Name = "bsalirSub";
            bsalirSub.Size = new System.Drawing.Size(75, 25);
            bsalirSub.Text = "ACEPTAR";
            bsalirSub.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;//MiddleCenter/TopCenter
            bsalirSub.FlatStyle = FlatStyle.Flat;
            //  datos del label titulo
            Titulo_SubFormulario.Name = "Titulo_SubFormulario";
            Titulo_SubFormulario.Text = "MENSAJE";
            Titulo_SubFormulario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;//MiddleCenter/TopCenter
            Titulo_SubFormulario.Size = new Size(400, 25);//Tamaño del label
            Titulo_SubFormulario.Location = new Point(ancho/2-400/2,10);//localizacion del label en el formulario
            Titulo_SubFormulario.Font = new System.Drawing.Font("Segoe UI", 16F, System.Drawing.FontStyle.Bold);
            //  datos del label Informacion
            Informacion.Name = "Informacion";
            Informacion.Text = MENSAJE;

            Informacion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;//MiddleCenter/TopCenter
            Informacion.Size = new Size(ancho-100,alto-100);//Tamaño del label
            Informacion.Location = new Point(10, 35);//localizacion del label en el formulario
            Informacion.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            //  datos de la ventana
            SubformularioMensaje.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            SubformularioMensaje.ClientSize = new System.Drawing.Size(ancho, alto);
            SubformularioMensaje.Opacity = .90;
            //tomamos el color de un control existente en el form principal y lo pasamos a este form
            SubformularioMensaje.BackColor = txtCaja.BackColor;
            SubformularioMensaje.ForeColor = txtCaja.ForeColor;
            
            SubformularioMensaje.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;//FixedDialog, None, Fixed3D
            SubformularioMensaje.Name = "Subformulario";
            SubformularioMensaje.Text = "SUB VENTANA X 1.0";
            SubformularioMensaje.CancelButton = bsalirSub;
            SubformularioMensaje.StartPosition = FormStartPosition.CenterScreen;
            //
            SubformularioMensaje.Controls.Add(bsalirSub);
            SubformularioMensaje.Controls.Add(Titulo_SubFormulario);
            SubformularioMensaje.Controls.Add(Informacion);
            //
            SubformularioMensaje.ShowDialog();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            this.BackColor = Color.BlanchedAlmond;
            this.ForeColor = Color.Blue;
            txtCaja.BackColor = Color.BlanchedAlmond;
            txtCaja.ForeColor = Color.Blue;
            menuStrip1.BackColor = Color.Blue;
            menuStrip1.ForeColor = Color.BlanchedAlmond;
            //primer  menu
            menuToolStripMenuItem.ForeColor = Color.BlanchedAlmond;

            abrirTxtToolStripMenuItem.BackColor = Color.Blue;
            abrirTxtToolStripMenuItem.ForeColor = Color.BlanchedAlmond;
            limpiarToolStripMenuItem.BackColor = Color.Blue;
            limpiarToolStripMenuItem.ForeColor = Color.BlanchedAlmond;
            infoToolStripMenuItem.BackColor = Color.Blue;
            infoToolStripMenuItem.ForeColor = Color.BlanchedAlmond;
            salirToolStripMenuItem.BackColor = Color.Blue;
            salirToolStripMenuItem.ForeColor = Color.BlanchedAlmond;
            //menu exportar
            exportarToolStripMenuItem.ForeColor = Color.BlanchedAlmond;

            aTxtToolStripMenuItem.BackColor = Color.Blue;
            aTxtToolStripMenuItem.ForeColor = Color.BlanchedAlmond;
            aPdfToolStripMenuItem.BackColor = Color.Blue;
            aPdfToolStripMenuItem.ForeColor = Color.BlanchedAlmond;
        }
    }
}
