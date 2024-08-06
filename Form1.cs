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
        private string appDirectory;

        public Form1()
        {
            InitializeComponent();
            appDirectory = AppDomain.CurrentDomain.BaseDirectory;

            // Verificar se o diretório da aplicação foi corretamente definido
            if (string.IsNullOrWhiteSpace(appDirectory))
            {
                MessageBox.Show("O diretório da aplicação não pôde ser determinado.");
                return;
            }

            // Verificar se os diretórios de entrada e log existem, e criá-los se necessário
            EnsureDirectoryExists(Path.Combine(appDirectory, "INPUT"));
            EnsureDirectoryExists(Path.Combine(appDirectory, "LOG"));

            InitializeBackgroundWorker();
        }

        private void EnsureDirectoryExists(string directoryPath)
        {
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }
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
            if (string.IsNullOrWhiteSpace(textBoxNs.Text) ||
                string.IsNullOrWhiteSpace(textBoxDefeito.Text) ||
                string.IsNullOrWhiteSpace(textBoxEndereco.Text) ||
                string.IsNullOrWhiteSpace(textBoxComplemento.Text) ||
                string.IsNullOrWhiteSpace(textBoxPessoaContato.Text) ||
                string.IsNullOrWhiteSpace(textBoxTelContato.Text) ||
                string.IsNullOrWhiteSpace(textBoxExped.Text))
            {
                MessageBox.Show("Por favor, preencha todos os campos obrigatórios.");
                return;
            }

            string ns = textBoxNs.Text;
            string defeito = textBoxDefeito.Text;
            string endereco = textBoxEndereco.Text;
            string complemento = textBoxComplemento.Text;
            string nomecontato = textBoxPessoaContato.Text;
            string tel = textBoxTelContato.Text;
            string exped = textBoxExped.Text;


            string inputFilePath = Path.Combine(appDirectory, "INPUT", "email_input.txt");
            string logFilePath = Path.Combine(appDirectory, "LOG", "python_log.txt");

            string input = $"\"{ns}\",\"{defeito}\",\"{endereco}\",\"{complemento}\",\"{nomecontato}\",\"{tel}\",\"{exped}\"";
            File.WriteAllText(inputFilePath, input); // Salva os dados no arquivo

            progressBar1.Value = 0;
            progressBar1.Maximum = 100;

            labelDebug.Text = "Iniciando o processo do Python...\n";
            bgWorker.RunWorkerAsync(new string[] { inputFilePath, logFilePath });
        }

        private void BgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var args = (string[])e.Argument;
            string inputFilePath = args[0];
            string logFilePath = args[1];

            var process = new Process
            {
                StartInfo =
                {
                    FileName = "python",  // Caminho do .exe do python
                    Arguments = $"\"{Path.Combine(appDirectory, "script_chamado_positivo.py")}\" \"{inputFilePath}\" \"{logFilePath}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                }
            };

            // Captura a saída do processo
            process.OutputDataReceived += (o, dataArgs) =>
            {
                if (dataArgs.Data != null)
                {
                    bgWorker.ReportProgress(50, $"Saída: {dataArgs.Data}");
                    LogToFile($"Saída: {dataArgs.Data}");
                }
            };

            // Captura os erros do processo
            process.ErrorDataReceived += (o, errorArgs) =>
            {
                if (errorArgs.Data != null)
                {
                    bgWorker.ReportProgress(50, $"Erro: {errorArgs.Data}");
                    LogToFile($"Erro: {errorArgs.Data}");
                }
            };

            try
            {
                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                process.WaitForExit();
                bgWorker.ReportProgress(100, "Processo do Python concluído com sucesso.");
                LogToFile("Processo do Python concluído com sucesso.\n");
            }
            catch (Exception ex)
            {
                bgWorker.ReportProgress(0, $"Erro ao iniciar o processo: {ex.Message}");
                LogToFile($"Erro ao iniciar o processo: {ex.Message}");
            }
        }

        private void BgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            labelDebug.Invoke((MethodInvoker)delegate
            {
                labelDebug.Text += $"{e.UserState}\n";
            });
        }

        private void BgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            labelDebug.Invoke((MethodInvoker)delegate
            {
                labelDebug.Text += "Email enviado com sucesso!\n";
            });
        }

        private void LogToFile(string message)
        {
            try
            {
                File.AppendAllText(Path.Combine(appDirectory, "LOG", "python_log.txt"), $"{DateTime.Now}: {message}\n");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao escrever no log: {ex.Message}");
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxSede.Checked)
            {
                textBoxEndereco.Text = "Endereço sede";
                textBoxComplemento.Text = "Complemento sede";
                textBoxPessoaContato.Text = "Pessoa contato sede";
                textBoxTelContato.Text = "Tel contato sede";
                textBoxExped.Text = "Exepediente sede";
            }
            else
            {
                textBoxEndereco.Clear();
                textBoxComplemento.Clear();
                textBoxPessoaContato.Clear();
                textBoxTelContato.Clear();
                textBoxExped.Clear();
            }
        }
    }
}
