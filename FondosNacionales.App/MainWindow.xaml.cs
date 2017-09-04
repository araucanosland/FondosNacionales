using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace FondosNacionales.App
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        System.ComponentModel.BackgroundWorker wrkrHistoricos;
        System.ComponentModel.BackgroundWorker wrkrFondoCesantia;
        System.ComponentModel.BackgroundWorker wrkrFondoAsfam;
        System.ComponentModel.BackgroundWorker wrkrFondoSIL;
        System.ComponentModel.BackgroundWorker wrkrFondoMaternal;
        System.ComponentModel.BackgroundWorker wrkAnexos;

        IA.FondosNacionales.Excel.ControladorPrincipal x;

        string periodo_ejec;



        public MainWindow()
        {
            InitializeComponent();
            x = new IA.FondosNacionales.Excel.ControladorPrincipal();
            //worker Historicos
            wrkrHistoricos = new System.ComponentModel.BackgroundWorker();
            wrkrHistoricos.WorkerReportsProgress = true;
            wrkrHistoricos.DoWork += wrkr_DoWorkHistoricos;
            wrkrHistoricos.ProgressChanged += wrkrHistoricos_ProgressChanged;
            wrkrHistoricos.RunWorkerCompleted += wrkrHistoricos_RunWorkerCompleted;


            //worker Cesantia
            wrkrFondoCesantia = new System.ComponentModel.BackgroundWorker();
            wrkrFondoCesantia.WorkerReportsProgress = true;
            wrkrFondoCesantia.DoWork += wrkr_DoWorkFondoCesantia;
            wrkrFondoCesantia.ProgressChanged += wrkr_ProgressChanged;
            wrkrFondoCesantia.RunWorkerCompleted += wrkrCesantia_RunWorkerCompleted;

            //worker Asfam
            wrkrFondoAsfam = new System.ComponentModel.BackgroundWorker();
            wrkrFondoAsfam.WorkerReportsProgress = true;
            wrkrFondoAsfam.DoWork += wrkr_DoWorkFondoAsfam;
            wrkrFondoAsfam.ProgressChanged += wrkr_ProgressChanged;
            wrkrFondoAsfam.RunWorkerCompleted += wrkrAsfam_RunWorkerCompleted;


            //worker SIL
            wrkrFondoSIL = new System.ComponentModel.BackgroundWorker();
            wrkrFondoSIL.WorkerReportsProgress = true;
            wrkrFondoSIL.DoWork += wrkr_DoWorkSIL;
            wrkrFondoSIL.ProgressChanged += wrkr_ProgressChanged;
            wrkrFondoSIL.RunWorkerCompleted += wrkrSIL_RunWorkerCompleted;


            //worker Maternal
            wrkrFondoMaternal = new System.ComponentModel.BackgroundWorker();
            wrkrFondoMaternal.WorkerReportsProgress = true;
            wrkrFondoMaternal.DoWork += wrkr_DoWorkMaternal;
            wrkrFondoMaternal.ProgressChanged += wrkr_ProgressChanged;
            wrkrFondoMaternal.RunWorkerCompleted += wrkrMaternal_RunWorkerCompleted;

            //worker Maternal
            wrkAnexos = new System.ComponentModel.BackgroundWorker();
            wrkAnexos.WorkerReportsProgress = true;
            wrkAnexos.DoWork += wrkr_DoWorkAnexos;
            wrkAnexos.ProgressChanged += wrkr_ProgressChanged;
            wrkAnexos.RunWorkerCompleted += wrkrMaternal_RunWorkerCompleted;

            /*List<IA.FondosNacionales.Entity.Periodo> periodos = new List<IA.FondosNacionales.Entity.Periodo>();
            periodos.Add(new IA.FondosNacionales.Entity.Periodo(201701, "Enero de 2017"));
            periodos.Add(new IA.FondosNacionales.Entity.Periodo(201702, "Febrero de 2017"));
            periodos.Add(new IA.FondosNacionales.Entity.Periodo(201703, "Marzo de 2017"));
            periodos.Add(new IA.FondosNacionales.Entity.Periodo(201704, "Abril de 2017"));*/

            DateTime di = new DateTime(2017, 01, 01);
            DateTime df = new DateTime(2019, 12, 01);

            while(di <= df)
            {
                ComboBoxItem g = new ComboBoxItem();
                g.Content = di.Year.ToString() + "" + di.Month.ToString().PadLeft(2,'0');
                cb_periodo.Items.Add(g);

                di = di.AddMonths(1);
            }

            /*for (int i = 201701; i<=201912; i++)
            {
                ComboBoxItem g = new ComboBoxItem();
                g.Content = i;
                cb_periodo.Items.Add(g);
            }*/

            cb_periodo.SelectedIndex = 0;

        }

        private void btn_Historicos_Click(object sender, RoutedEventArgs e)
        {
            lb_status.Content = "Procesando Historicos..";
            this.periodo_ejec = ((ComboBoxItem)cb_periodo.SelectedItem).Content.ToString();
            wrkrHistoricos.RunWorkerAsync();
        }

        private void wrkr_DoWorkHistoricos(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            x.ProcesarHistoricos((System.ComponentModel.BackgroundWorker)sender, this.periodo_ejec);
        }

        private void wrkrHistoricos_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            pg_avance.Value = e.ProgressPercentage;
        }

        private void wrkrHistoricos_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            lb_status.Content = "Proceso Historicos Finalizado.";
            ((System.ComponentModel.BackgroundWorker)sender).Dispose();
        }
        
        private void wrkr_DoWorkFondoCesantia(object sender, System.ComponentModel.DoWorkEventArgs e)
        {

            x.ProcesarFondoCesantia((System.ComponentModel.BackgroundWorker)sender, this.periodo_ejec);
        }
                
        private void wrkrCesantia_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            lb_status.Content = "Proceso Fondo Cesantia Finalizado.";
            ((System.ComponentModel.BackgroundWorker)sender).Dispose();
        }

        private void btn_Cesantia_Click(object sender, RoutedEventArgs e)
        {
            lb_status.Content = "Procesando Fondo Cesantia..";
            this.periodo_ejec = ((ComboBoxItem)cb_periodo.SelectedItem).Content.ToString();
            wrkrFondoCesantia.RunWorkerAsync();
        }
        
        private void wrkr_DoWorkFondoAsfam(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            x.ProcesarFondoAsfam((System.ComponentModel.BackgroundWorker)sender, this.periodo_ejec);
        }

        private void wrkr_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            pg_avance.Value = e.ProgressPercentage;
        }

        private void wrkrAsfam_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            lb_status.Content = "Proceso Fondo Asfam Finalizado.";
            ((System.ComponentModel.BackgroundWorker)sender).Dispose();
        }

        private void btn_ASFAM_Click(object sender, RoutedEventArgs e)
        {
            lb_status.Content = "Procesando Fondo Asfam..";
            this.periodo_ejec = ((ComboBoxItem)cb_periodo.SelectedItem).Content.ToString();
            wrkrFondoAsfam.RunWorkerAsync();
        }

        private void btn_SIL_Click(object sender, RoutedEventArgs e)
        {
            lb_status.Content = "Procesando Fondo SIL..";
            this.periodo_ejec = ((ComboBoxItem)cb_periodo.SelectedItem).Content.ToString();
            wrkrFondoSIL.RunWorkerAsync();
        }

        private void btm_Maternal_Click(object sender, RoutedEventArgs e)
        {
            lb_status.Content = "Procesando Fondo Maternal..";
            this.periodo_ejec = ((ComboBoxItem)cb_periodo.SelectedItem).Content.ToString();
            wrkrFondoMaternal.RunWorkerAsync();
        }


        private void wrkrSIL_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            lb_status.Content = "Proceso Fondo SIL Finalizado.";
            ((System.ComponentModel.BackgroundWorker)sender).Dispose();
        }

        private void wrkrMaternal_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            lb_status.Content = "Proceso Fondo Maternal Finalizado.";
            ((System.ComponentModel.BackgroundWorker)sender).Dispose();
        }

        private void wrkr_DoWorkSIL(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            x.ProcesarFondoSIL((System.ComponentModel.BackgroundWorker)sender, this.periodo_ejec);
        }

        private void wrkr_DoWorkMaternal(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            x.ProcesarFondoMaternal((System.ComponentModel.BackgroundWorker)sender, this.periodo_ejec);
        }


        private void wrkr_DoWorkAnexos(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            x.ProcesarAnexos((System.ComponentModel.BackgroundWorker)sender, this.periodo_ejec);
        }

        private void btn_anexos_Click(object sender, RoutedEventArgs e)
        {
            lb_status.Content = "Procesando Anexos..";
            this.periodo_ejec = ((ComboBoxItem)cb_periodo.SelectedItem).Content.ToString();
            wrkAnexos.RunWorkerAsync();
        }
    }
}
