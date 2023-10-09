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

            PreencherMenu();
            int OpcaoSelecionada = Convert.ToInt32(Console.ReadLine());
            while (OpcaoSelecionada != 0)
            {

                RealizarPasso(OpcaoSelecionada);

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
                    Titulo = "Executar script da VW"
                }

            };

            foreach (var Opcao in Opcoes)
            {
                Console.WriteLine($"{Opcao.Id} - {Opcao.Titulo}");
            }

            Console.WriteLine(string.Empty);

            return Opcoes;

        }

        private static void RealizarPasso(int passo)
        {

            Console.WriteLine($"--------------- PROCESSANDO {passo}º PASSO ---------------");
            Console.WriteLine();
            Console.WriteLine();

            string FileName = $"limpeza.sql";

            var Dados = ReadFile(passo);
            if (Dados == null)
            {
                Console.WriteLine("Não foi possível ler o arquivo ou o arquivo está em branco, cheque o arquivo e tente novamente!");
            }

            else
            {

                int indiceTrocar = Dados.IndexOf(Dados.Single(dado => dado.Nome == "RAFAEL SANT' CLAIR CORREA"));
                Dados[indiceTrocar].Nome = "RAFAEL SANT'' CLAIR CORREA";
                var teste = Dados.Where(dado => dado.UserNameAntigo == "UMASCAR");

                string FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), FileName);
                FileInfo FileInfo = new FileInfo(FilePath);

                if (FileInfo.Exists)
                {

                    FileInfo.Delete();

                }

                using (StreamWriter writer = new StreamWriter(FilePath))
                {

                    string inUserNamesAntigos = string.Join(",", Dados.Select(dado => "\'" + dado.UserNameAntigo + "\'"));

                    writer.WriteLine($@"

                       create table TabelaTemporaria(
	                        UserName varchar(max)
                        )

                    ");

                    foreach (var dado in Dados)
                    {

                        writer.WriteLine($@"

                          insert into TabelaTemporaria (UserName) values ('{dado.UserNameAntigo}')

                        ");

                    }

                    writer.WriteLine($@"

                            -- Script para limpar duplicações no banco
                            
                            UPDATE AspNetUsers 
                            SET 
	                            [Name] = 'Usuário ' + CAST([Id] AS NVARCHAR(36)),
	                            [UserName] = 'Usuário ' + CAST([Id] AS NVARCHAR(36)), 
	                            [Email] = CAST([Id] AS NVARCHAR(36)) + '@nao.tem.email',
                                Ativo = 0
	                            where (email <> 'admin@aevoinnovate.net') and (email <> 'aevoinnovate@aevo.com.br')
                                and UserName not in (select UserName from TabelaTemporaria)

                            drop table TabelaTemporaria

                    ");

                }

                Console.WriteLine($"O arquivo foi salvo na sua área de trabalho com o nome de '{FileName}'");

                FileName = "atualizacao.sql";
                FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), FileName);
                FileInfo = new FileInfo(FilePath);

                if (FileInfo.Exists)
                {

                    FileInfo.Delete();

                }

                using (StreamWriter writer = new StreamWriter(FilePath))
                {

                    foreach (var dado in Dados)
                    {

                        writer.WriteLine($@"

                            update AspNetUsers
                            set
                            Name = '{dado.Nome}',
                            UserName = '{dado.UserNameNovo}',
                            Email = {(string.IsNullOrEmpty(dado.Email) ? "convert(varchar(max), Id) + '@nao.tem.email'" : "\'" + dado.Email + "\'")},
                            CentroCusto = '{dado.CentroCusto}'
                            where UserName = '{dado.UserNameAntigo}'                            

                        ");

                    }

                }

                FileName = "cadastro.sql";
                FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), FileName);
                FileInfo = new FileInfo(FilePath);

                if (FileInfo.Exists)
                {

                    FileInfo.Delete();

                }

                using (StreamWriter writer = new StreamWriter(FilePath))
                {

                    foreach (var dado in Dados)
                    {

                        writer.WriteLine($@"

                            if not exists(select * from AspNetUsers where UserName = '{dado.UserNameNovo}')

                                insert into AspNetUsers (Id, Name, UserName, Email, CentroCusto, EmailConfirmed, PhoneNumberConfirmed, TwoFactorEnabled,LockoutEnabled, AccessFailedCount, Ativo)
                                values (
                                    newid(),
                                   '{dado.Nome}',
                                   '{dado.UserNameNovo}',
                                   {(string.IsNullOrEmpty(dado.Email) ? "'@nao.tem.email'" : "\'" + dado.Email + "\'")},
                                   '{dado.CentroCusto}',
                                    0,
									0,
									0,
									0,
									0,
									1
                                )

                        ");

                    }

                    writer.WriteLine($@"update AspNetUsers set Email = convert(varchar(max), Id) + '@nao.tem.email' where Email = '@nao.tem.email'");

                }

                Console.WriteLine($"O arquivo foi salvo na sua área de trabalho com o nome de '{FileName}'");

                FileName = "cadastroDepartamento.sql";
                FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), FileName);
                FileInfo = new FileInfo(FilePath);

                if (FileInfo.Exists)
                {

                    FileInfo.Delete();

                }

                using (StreamWriter writer = new StreamWriter(FilePath))
                {

                    foreach (var dado in Dados)
                    {

                        writer.WriteLine($@"

                            if not exists(select * from Departamento where Nome = '{dado.Departamento}')

                                insert into Departamento (Nome, GestorId, Padrao, Ativa)
                                values (
                                    '{dado.Departamento}',
                                    '43C6049D-511D-4BF3-B255-2683E0C6B25A',
                                    0,
                                    1
                                )

                        ");

                    }

                }

                Console.WriteLine($"O arquivo foi salvo na sua área de trabalho com o nome de '{FileName}'");

                FileName = "atualizacaoDepartamento.sql";
                FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), FileName);
                FileInfo = new FileInfo(FilePath);

                if (FileInfo.Exists)
                {

                    FileInfo.Delete();

                }

                using (StreamWriter writer = new StreamWriter(FilePath))
                {

                    foreach (var dado in Dados)
                    {

                        writer.WriteLine($@"

                            update AspNetUsers set DepartamentoId = (select Id from Departamento where Nome = '{dado.Departamento}')
                            where UserName = '{dado.UserNameNovo}'

                        ");

                    }

                }

                Console.WriteLine($"O arquivo foi salvo na sua área de trabalho com o nome de '{FileName}'");


            }

            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine($"--------------- {passo}º PASSO PROCESSADO ---------------");

        }


        private static List<Dado> ReadFile(int passo)
        {

            try
            {

                List<Dado> Dados = new List<Dado>();

                using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo($"D:\\vw.xlsx")))
                {

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var myWorksheet = xlPackage.Workbook.Worksheets.First();
                    var totalRows = myWorksheet.Dimension.End.Row;
                    var totalColumns = myWorksheet.Dimension.End.Column;

                    for (int rowNum = 2; rowNum <= totalRows; rowNum++)
                    {

                        var NomeCell = myWorksheet.Cells[rowNum, 1].Value;
                        var UserNameAntigoCell = myWorksheet.Cells[rowNum, 2].Value;
                        var UserNameNovoCell = myWorksheet.Cells[rowNum, 3].Value;
                        var DepartamentoCell = myWorksheet.Cells[rowNum, 4].Value;
                        var EmailCell = myWorksheet.Cells[rowNum, 5].Value;
                        var SenhaCell = myWorksheet.Cells[rowNum, 6].Value;
                        var CentroCustoCell = myWorksheet.Cells[rowNum, 13].Value;

                        if (NomeCell == null)
                        {
                            break;
                        }

                        Dados.Add(new Dado()
                        {
                            Nome = NomeCell.ToString(),
                            UserNameAntigo = UserNameAntigoCell.ToString(),
                            UserNameNovo = UserNameNovoCell.ToString(),
                            Departamento = DepartamentoCell.ToString(),
                            Email = EmailCell == null ? string.Empty : EmailCell.ToString(),
                            Senha = SenhaCell.ToString(),
                            CentroCusto = CentroCustoCell == null ? string.Empty : CentroCustoCell.ToString()
                        });

                    }

                }

                return Dados;

            }

            catch
            {
                return null;
            }

        }

        public class Dado
        {

            public string Nome { get; set; }
            public string UserNameAntigo { get; set; }
            public string UserNameNovo { get; set; }
            public string Departamento { get; set; }
            public string Email { get; set; }
            public string Senha { get; set; }
            public string CentroCusto { get; set; }

        }

        private class Opcoes
        {
            public int Id { get; set; }
            public string Titulo { get; set; }
        }

    }
}
