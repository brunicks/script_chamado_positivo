using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace AbrirChamadoPositivo
{
    public partial class Form1 : Form
    {
        private BackgroundWorker bgWorker;

        public Form1()
        {
            InitializeComponent();
            InitializeBackgroundWorker();
        }

        private void InitializeBackgroundWorker()
        {
            bgWorker = new BackgroundWorker();
            bgWorker.DoWork += BgWorker_DoWork;
            bgWorker.ProgressChanged += BgWorker_ProgressChanged;
            bgWorker.RunWorkerCompleted += BgWorker_RunWorkerCompleted;
            bgWorker.WorkerReportsProgress = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Obtém os dados dos campos do formulário
            string ns = textBoxNs.Text;
            string defeito = textBoxDefeito.Text;
            string endereco = textBoxEndereco.Text;
            string complemento = textBoxComplemento.Text;
            string nomecontato = textBoxPessoaContato.Text;
            string tel = textBoxTelContato.Text;
            string exped = textBoxExped.Text;

            // Caminho dos arquivos
            string inputFilePath = @"C:\Users\est.07545996984\Documents\teste_script_chamado\email_input.txt";
            string logFilePath = @"C:\Users\est.07545996984\Documents\teste_script_chamado\python_log.txt";

            // Cria o conteúdo do arquivo
            string input = $"{ns},{defeito},{endereco},{complemento},{nomecontato},{tel},{exped}";
            File.WriteAllText(inputFilePath, input); // Salva os dados no arquivo

            // Inicia o processo do Python em uma thread separada
            labelDebug.Text = "Iniciando o processo do Python...";
            bgWorker.RunWorkerAsync(new string[] { inputFilePath, logFilePath });
        }

        private void BgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var args = (string[])e.Argument;
            string inputFilePath = args[0];
            string logFilePath = args[1];

            // Configura e inicia o processo do script Python
            var process = new Process();
            process.StartInfo.FileName = @"C:\Users\est.07545996984\AppData\Local\Programs\Python\Python312\python.exe";
            process.StartInfo.Arguments = $"\"C:\\Users\\est.07545996984\\Documents\\teste_script_chamado\\script_chamado_positivo.py\" \"{inputFilePath}\" \"{logFilePath}\"";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            process.StartInfo.CreateNoWindow = true;

            process.OutputDataReceived += (o, dataArgs) =>
            {
                if (dataArgs.Data != null)
                {
                    bgWorker.ReportProgress(0, $"Saída: {dataArgs.Data}");
                }
            };

            process.ErrorDataReceived += (o, errorArgs) =>
            {
                if (errorArgs.Data != null)
                {
                    bgWorker.ReportProgress(0, $"Erro: {errorArgs.Data}");
                }
            };

            try
            {
                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                process.WaitForExit();
                bgWorker.ReportProgress(0, "Processo do Python concluído com sucesso.");
            }
            catch (Exception ex)
            {
                bgWorker.ReportProgress(0, $"Erro ao iniciar o processo: {ex.Message}");
            }
        }

        private void BgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            labelDebug.Text += $"{e.UserState}{Environment.NewLine}";
        }

        private void BgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // Ações após o término do trabalho em segundo plano, se necessário
        }
    }
}
