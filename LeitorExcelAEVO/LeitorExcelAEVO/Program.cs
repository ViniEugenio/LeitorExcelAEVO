using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace LeitorExcelAEVO
{
    internal class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("------------------------------ Processador de Excel AEVO ------------------------------");
            Console.WriteLine("Selecione a opção desejada:");
            Console.WriteLine(string.Empty);

            var Opcoes = PreencherMenu();
            int OpcaoSelecionada = Convert.ToInt32(Console.ReadLine());
            while (OpcaoSelecionada != 0)
            {

                switch (OpcaoSelecionada)
                {

                    case 1:
                        ProcessarArquivo();
                        break;

                    default:
                        Console.WriteLine("A opção selecionada não existe, por favor selecione outra opção!");
                        break;

                }

                Console.WriteLine();
                Console.WriteLine();

                PreencherMenu();
                OpcaoSelecionada = Convert.ToInt32(Console.ReadLine());

            }

            Console.WriteLine(String.Empty);
            Console.WriteLine("------------------------------ FIM ------------------------------");
            Console.ReadKey();

        }

        private static List<Opcoes> PreencherMenu()
        {

            Console.WriteLine("Opções:");

            List<Opcoes> Opcoes = new List<Opcoes>()
            {

                new Opcoes()
                {
                    Id = 0,
                    Titulo = "Encerrar programa"
                },

                new Opcoes()
                {
                    Id = 1,
                    Titulo = "Processar arquivo"
                }

            };

            foreach (var Opcao in Opcoes)
            {
                Console.WriteLine($"{Opcao.Id} - {Opcao.Titulo}");
            }

            Console.WriteLine(string.Empty);

            return Opcoes;

        }

        private static void ProcessarArquivo()
        {

            Console.WriteLine("--------------- PROCESSANDO ARQUIVO ---------------");
            Console.WriteLine();
            Console.WriteLine();

            string FileName = "script.sql";

            var Dados = ReadFile();
            if (Dados == null)
            {
                Console.WriteLine("Não foi possível ler o arquivo ou o arquivo está em branco, cheque o arquivo e tente novamente!");
            }

            else
            {

                string FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), FileName);
                FileInfo FileInfo = new FileInfo(FilePath);
                bool Exists = FileInfo.Exists;

                int NumeroScript = 0;
                while (Exists)
                {

                    NumeroScript = NumeroScript + 1;
                    FileName = $"script{NumeroScript}.sql";
                    FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), FileName);
                    FileInfo = new FileInfo(FilePath);
                    Exists = FileInfo.Exists;

                }

                using (StreamWriter writer = new StreamWriter(FilePath))
                {

                    foreach (var dado in Dados)
                    {
                        writer.WriteLine($"UPDATE AspNetUsers set UserName = '{dado.UserName}' where Email = '{dado.Email}'");
                    }

                }

            }

            Console.WriteLine($"O arquivo foi salvo na sua área de trabalho com o nome de '{FileName}'");

            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("--------------- ARQUIVO PROCESSADO ---------------");

        }


        private static List<Dados> ReadFile()
        {

            try
            {

                List<Dados> Dados = new List<Dados>();

                using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo("F:\\dados.xlsx")))
                {

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var myWorksheet = xlPackage.Workbook.Worksheets.First();
                    var totalRows = myWorksheet.Dimension.End.Row;
                    var totalColumns = myWorksheet.Dimension.End.Column;

                    for (int rowNum = 2; rowNum <= totalRows; rowNum++)
                    {

                        var UserNameCell = myWorksheet.Cells[rowNum, 2].Value;
                        var EmailCell = myWorksheet.Cells[rowNum, 4].Value;

                        if (UserNameCell != null && EmailCell != null)
                        {
                            Dados.Add(new Dados()
                            {
                                UserName = UserNameCell.ToString(),
                                Email = EmailCell.ToString()
                            });
                        }

                    }

                }

                return Dados;

            }

            catch
            {
                return null;
            }

        }

        private class Dados
        {
            public string UserName { get; set; }
            public string Email { get; set; }
        }

        private class Opcoes
        {
            public int Id { get; set; }
            public string Titulo { get; set; }
        }

    }
}
