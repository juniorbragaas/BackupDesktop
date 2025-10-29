using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace BackupDesktop
{
    internal class Program
    {
        // ==============================================
        // CONFIGURAÇÕES
        // ==============================================

        public static string Origem = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Desktop");
        public static readonly string LetraHD = ConfigurationManager.AppSettings["LetraHD"];
        public static readonly string PastaBackup = ConfigurationManager.AppSettings["PastaBackup"];
        public static readonly string LogDir = Path.Combine(LetraHD, ConfigurationManager.AppSettings["LogDir"]);
        public static readonly int QtdBackups = int.TryParse(ConfigurationManager.AppSettings["QtdBackups"], out int qtd) ? qtd : 10000;
        public static readonly string Data = DateTime.Now.ToString("yyyy-MM-dd");
        public static readonly string CaminhoLog = Path.Combine(LogDir, $"log_{Data}.txt");
        public static readonly string DestinoBase = Path.Combine(LetraHD, $"{PastaBackup}");

        public static readonly int Desligar = Convert.ToInt32(ConfigurationManager.AppSettings["DesligarAoFinal"]);

        static void Main(string[] args)
        {
            Console.WriteLine("====================================");
            Console.WriteLine("Backup automático da Área de Trabalho");
            Console.WriteLine("Autor: Valdelei Junior Braga");
            Console.WriteLine("====================================\n");

            try
            {
                // Impede desligamento durante o backup
                PreventShutdown(true);

                if (!Directory.Exists(LetraHD))
                {
                    EscreverLog($"HD externo ({LetraHD}) não encontrado. Encerrando...");
                    Console.WriteLine("HD externo não encontrado. Conecte o disco e tente novamente.");
                    return;
                }

                Directory.CreateDirectory(DestinoBase);
                Directory.CreateDirectory(LogDir);

                EscreverLog($"Iniciando backup incremental em {Data}");
                EscreverLog($"Origem: {Origem}");
                EscreverLog($"Destino base: {DestinoBase}");

                string destinoAtual = Path.Combine(DestinoBase, $"Backup_{Data}");
                Directory.CreateDirectory(destinoAtual);

                var backupsExistentes = new DirectoryInfo(DestinoBase)
                    .GetDirectories()
                    .OrderByDescending(d => d.Name)
                    .ToList();

                if (backupsExistentes.Count > 1)
                {
                    var backupAnterior = backupsExistentes.Skip(1).First();
                    EscreverLog($"Backup anterior encontrado: {backupAnterior.Name}");

                    string linkDestino = Path.Combine(DestinoBase, "PreviousBackup");
                    if (Directory.Exists(linkDestino))
                        Directory.Delete(linkDestino, true);

                    CriarLinkSimbolico(linkDestino, backupAnterior.FullName);

                    ExecutarRobocopy(Origem, destinoAtual, CaminhoLog, $"/MIR /XO /FFT /R:1 /W:1 /NP /XD \"{linkDestino}\"");

                    if (Directory.Exists(linkDestino))
                        Directory.Delete(linkDestino, true);
                }
                else
                {
                    EscreverLog("Nenhum backup anterior encontrado — cópia completa inicial.");
                    ExecutarRobocopy(Origem, destinoAtual, CaminhoLog, "/E /R:1 /W:1 /NP");
                }

                EscreverLog($"Backup concluído para: {destinoAtual}");

                // Rotação de backups antigos
                var backupsRestantes = new DirectoryInfo(DestinoBase)
                    .GetDirectories()
                    .OrderByDescending(d => d.CreationTime)
                    .ToList();

                if (backupsRestantes.Count > QtdBackups)
                {
                    var paraExcluir = backupsRestantes.Skip(QtdBackups);
                    foreach (var b in paraExcluir)
                    {
                        EscreverLog($"Removendo backup antigo: {b.FullName}");
                        Directory.Delete(b.FullName, true);
                    }
                }

                EscreverLog($"Rotação concluída — mantidos {QtdBackups} backups.");
                EscreverLog("Backup finalizado com sucesso!");
                Console.WriteLine("\n✅ Backup finalizado com sucesso!");
            }
            catch (Exception ex)
            {
                EscreverLog($"Erro: {ex.Message}");
                Console.WriteLine($"\n❌ Erro: {ex.Message}");
            }
            finally
            {
                // Libera bloqueio e desliga o PC
                PreventShutdown(false);
                if(Desligar == 1)
                {
                    Console.WriteLine("\nDesligando computador em 10 segundos...");
                    System.Threading.Thread.Sleep(10000);
                    DesligarComputador();
                }
                
            }
        }

        // ==============================================
        // FUNÇÕES AUXILIARES
        // ==============================================
        static void EscreverLog(string texto)
        {
            string linha = $"{DateTime.Now:HH:mm:ss} - {texto}";
            Console.WriteLine(linha);
            File.AppendAllText(CaminhoLog, linha + Environment.NewLine);
        }

        static void ExecutarRobocopy(string origem, string destino, string log, string parametrosExtras)
        {
            string argumentos = $"\"{origem}\" \"{destino}\" {parametrosExtras} /LOG+:\"{log}\"";
            ProcessStartInfo psi = new ProcessStartInfo("robocopy", argumentos)
            {
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using var p = Process.Start(psi);
            p.WaitForExit();
        }

        static void CriarLinkSimbolico(string link, string destino)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/c mklink /D \"{link}\" \"{destino}\"",
                CreateNoWindow = true,
                UseShellExecute = false
            })?.WaitForExit();
        }

        static void DesligarComputador()
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "shutdown",
                Arguments = "/s /t 0",
                CreateNoWindow = true,
                UseShellExecute = false
            });
        }

        // Impede o desligamento durante o backup
        [DllImport("kernel32.dll")]
        private static extern uint SetThreadExecutionState(uint esFlags);
        private const uint ES_CONTINUOUS = 0x80000000;
        private const uint ES_SYSTEM_REQUIRED = 0x00000001;

        static void PreventShutdown(bool bloquear)
        {
            if (bloquear)
                SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED);
            else
                SetThreadExecutionState(ES_CONTINUOUS);
        }
    }
}
