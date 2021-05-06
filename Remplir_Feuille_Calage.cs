using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using Microsoft.VisualBasic;
using System.Configuration;
using Oracle.DataAccess.Client;
using System.Data.SqlClient;

namespace Calage_Inserts
{
    public partial class Remplir_Feuille_Calage : Form
    {
        string connString = @"Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=fra2exa01-sxdir1-vip.europe.essilor.group)(PORT=1561)))
                             (CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=PUE1)));User Id=combo;Password=combo;";


        public Remplir_Feuille_Calage()
        {
            InitializeComponent();

        }

        private void DefinirColonnes()
        {

        }


        private void button_start_Click(object sender, EventArgs e)
        {

        }

        private void button__Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Suivie_scrap_Click(object sender, EventArgs e)
        {

        }

        private void Remplir_une_feuille_de_calage_Click(object sender, EventArgs e)
        {

        }

        private void Rechercher_une_feuille_Click(object sender, EventArgs e)
        {

        }

        private void Stock_Click(object sender, EventArgs e)
        {

        }

        private void Remplacer_un_insert_Click(object sender, EventArgs e)
        {

        }

        private void Ajouter_un_insert_Click(object sender, EventArgs e)
        {

        }

        private void extraction_jobs_Click(object sender, EventArgs e)
        {

        }

        private void retirer_insert_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            Connexion c = new Connexion();
            c.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        public class donnees2
        {
            public donnees2() { }
            public string numero { get; set; }
            public double epaisseur_bord { get; set; }
            public double hauteur_bord { get; set; }
        }
        public class donnees3
        {
            public donnees3() { }
            public string numero { get; set; }
            public double epaisseur_bord { get; set; }
            public double hauteur_bord { get; set; }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string myConnectionString = @"user id=sa; password=K@rdexlsadm21!; data source=FRESD32615\SQLEXPRESS";
            SqlConnection myConnection = new SqlConnection(myConnectionString);
            Excel._Application xlsp;
            Excel.Workbook xlworkbook;
            Excel.Worksheet xlworkSheet;
            object missValue = System.Reflection.Missing.Value;
            xlsp = new Excel.Application();

            // affichage dossier excel
            if (true)
            {
                xlworkbook = xlsp.Workbooks.Add(missValue);
                xlworkSheet = (Excel.Worksheet)xlworkbook.Worksheets.get_Item(1);
                xlworkSheet.get_Range("A1", "I30").Borders.Weight = Excel.XlBorderWeight.xlThin;
                xlworkSheet.get_Range("A1", "I30").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                xlworkSheet.get_Range("A1", "I30").ColumnWidth = 8.5;
                xlworkSheet.Cells[1, 1] = "Date";
                xlworkSheet.Cells[1, 2] = DateTime.Now.ToString("yyyy/MM/dd");
                xlworkSheet.Cells[1, 5] = "Semaine";
                xlworkSheet.Cells[1, 6] = CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                xlworkSheet.Cells[1, 7] = "Validation préparation";
                xlworkSheet.get_Range("A1:A3", "I1:I3").Font.Size = 10;
                xlworkSheet.get_Range("B1", "D1").Merge(false);
                xlworkSheet.get_Range("G1", "H1").Merge(false);
                xlworkSheet.get_Range("G2", "H2").Merge(false);
                xlworkSheet.get_Range("G2", "H2").Font.Size = 6;
                xlworkSheet.Cells[2, 7] = "VALIDATION RECEPTION REGLEUR";
                xlworkSheet.get_Range("A2:A3", "F2:F3").Merge(false);
                xlworkSheet.get_Range("G3", "I3").Merge(false);
                xlworkSheet.Cells[4, 1] = "Job";
                //xlworkSheet.Cells[4, 5] = cd_mold;
                xlworkSheet.Cells[4, 2] = textBox1.Text;
                xlworkSheet.get_Range("B4", "C4").Merge(false);
                //xlworkSheet.Cells[4, 4] = "Moule";
                xlworkSheet.get_Range("E4", "F4").Merge(false);
                xlworkSheet.Cells[4, 7] = "Presse";
                //xlworkSheet.Cells[4, 8] = cd_press;
                xlworkSheet.get_Range("H4", "I4").Merge(false);
                xlworkSheet.Cells[5, 1] = "Info";
                xlworkSheet.get_Range("B5", "C5").Merge(false);
                xlworkSheet.get_Range("D5", "F5").Merge(false);
                xlworkSheet.get_Range("G5", "H5").Merge(false);
                xlworkSheet.Cells[5, 7] = "Run";
                xlworkSheet.Cells[6, 1] = "Produit";
                //xlworkSheet.Cells[6, 2] = routing_name;
                xlworkSheet.get_Range("B6", "F6").Merge(false);
                xlworkSheet.Cells[6, 7] = "Num Shot";
                //xlworkSheet.Cells[6, 8] = shots_nb;
                xlworkSheet.get_Range("H6", "I6").Merge(false);
                xlworkSheet.get_Range("A7", "I7").Merge(false);
                xlworkSheet.Cells[8, 1] = "Inserts convexes";
                xlworkSheet.get_Range("A8", "I8").Merge(false);
                xlworkSheet.get_Range("A9:A10", "I9:I10").Merge(false);
                xlworkSheet.Cells[11, 1] = "Cavités";
                xlworkSheet.get_Range("B11", "I11").NumberFormat = "@";
                xlworkSheet.Cells[11, 2] = "01";
                xlworkSheet.Cells[11, 3] = "02";
                xlworkSheet.Cells[11, 4] = "03";
                xlworkSheet.Cells[11, 5] = "04";
                xlworkSheet.Cells[11, 6] = "05";
                xlworkSheet.Cells[11, 7] = "06";
                xlworkSheet.Cells[11, 8] = "07";
                xlworkSheet.Cells[11, 9] = "08";
                xlworkSheet.Cells[12, 1] = "Base/Sphère";
                xlworkSheet.get_Range("A12", "I12").Font.Size = 8.5;
                xlworkSheet.get_Range("A12", "I12").NumberFormat = "0.00";
                xlworkSheet.get_Range("A14", "A15").Font.Size = 9;
                xlworkSheet.get_Range("A20", "A21").Font.Size = 9;
                xlworkSheet.Cells[13, 1] = "Ins CC";
                xlworkSheet.get_Range("B13", "I13").Font.Bold = 1;
                xlworkSheet.Cells[14, 1] = "Epais.Bord";
                xlworkSheet.get_Range("B14", "I14").NumberFormat = "0.00";
                xlworkSheet.Cells[15, 1] = "Epais.Centre";
                xlworkSheet.get_Range("B15", "I15").NumberFormat = "0.00";
                xlworkSheet.Cells[16, 1] = "Cales CC";
                xlworkSheet.get_Range("B16", "I16").NumberFormat = "0.00";
                xlworkSheet.get_Range("A17", "I17").Merge(false);
                xlworkSheet.Cells[18, 1] = "CYL";
                xlworkSheet.get_Range("B18", "I18").NumberFormat = "0.00";
                xlworkSheet.Cells[19, 1] = "ins CX";
                xlworkSheet.get_Range("B19", "I19").Font.Bold = 1;
                xlworkSheet.Cells[20, 1] = "Epais.Ctr.CX";
                xlworkSheet.get_Range("B20", "I20").NumberFormat = "0.00";
                xlworkSheet.Cells[21, 1] = "Epais verre";
                xlworkSheet.get_Range("B21", "I21").NumberFormat = "0.00";
                xlworkSheet.Cells[22, 1] = "Cales CX";
                xlworkSheet.Cells[23, 1] = "INS CC";
                xlworkSheet.Cells[24, 1] = "INS CX";
                xlworkSheet.Cells[25, 1] = "Changement d'inserts";
                xlworkSheet.get_Range("A25", "I25").Merge(false);
                xlworkSheet.Cells[26, 1] = "Cavités";
                xlworkSheet.get_Range("B26", "I26").NumberFormat = "@";
                xlworkSheet.Cells[26, 2] = "01";
                xlworkSheet.Cells[26, 3] = "02";
                xlworkSheet.Cells[26, 4] = "03";
                xlworkSheet.Cells[26, 5] = "04";
                xlworkSheet.Cells[26, 6] = "05";
                xlworkSheet.Cells[26, 7] = "06";
                xlworkSheet.Cells[26, 8] = "07";
                xlworkSheet.Cells[26, 9] = "08";
            }

            List<donnees2> list2 = new List<donnees2>();
            List<donnees3> list3 = new List<donnees3>();

            // ajout a liste de CC
            if (true)
            {
                list2.Add(new donnees2
                {
                    numero = textBox2.Text
                });

                list2.Add(new donnees2
                {
                    numero = textBox3.Text
                });
                
                list2.Add(new donnees2
                {
                    numero = textBox4.Text
                });

                list2.Add(new donnees2
                {
                    numero = textBox5.Text
                });
                
                list2.Add(new donnees2
                {
                    numero = textBox9.Text
                });

                list2.Add(new donnees2
                {
                    numero = textBox8.Text
                });

                list2.Add(new donnees2
                {
                    numero = textBox7.Text
                });
                
                list2.Add(new donnees2
                {
                    numero = textBox6.Text
                });

            }

            // ajout a liste de CX
            if (true)
            {
                list3.Add(new donnees3
                {
                    numero = textBox17.Text
                });

                list3.Add(new donnees3
                {
                    numero = textBox16.Text
                });

                list3.Add(new donnees3
                {
                    numero = textBox15.Text
                });

                list3.Add(new donnees3
                {
                    numero = textBox14.Text
                });

                list3.Add(new donnees3
                {
                    numero = textBox13.Text
                });

                list3.Add(new donnees3
                {
                    numero = textBox12.Text
                });

                list3.Add(new donnees3
                {
                    numero = textBox11.Text
                });

                list3.Add(new donnees3
                {
                    numero = textBox10.Text
                });

            }





            xlworkbook.SaveAs(@"R:\Commun\ACI\Data\Jobs_Combo\job_"+textBox1.Text, Excel.XlSaveAction.xlSaveChanges, missValue, missValue, missValue, missValue, Excel.XlSaveAsAccessMode.xlExclusive, missValue, missValue, missValue, missValue, missValue);
            xlworkbook.Close(true, missValue, missValue);
            xlsp.Quit();


            releaseObject(xlworkSheet);
            releaseObject(xlworkbook);
            releaseObject(xlsp);
        }



        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception e)
            {
                obj = null;
                MessageBox.Show("Exception occured while releasing object " + e.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Accueil r = new Accueil("");
            r.Show();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }
    }
}