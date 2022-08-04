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
                    Titulo = "Processar 1º Passo"
                },

                new Opcoes()
                {
                    Id = 2,
                    Titulo = "Processar 2º Passo"
                },

                new Opcoes()
                {
                    Id = 3,
                    Titulo = "Processar 3º Passo"
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

            string FileName = $"script-{passo}-passo.sql";

            var Dados = ReadFile(passo);
            if (Dados == null)
            {
                Console.WriteLine("Não foi possível ler o arquivo ou o arquivo está em branco, cheque o arquivo e tente novamente!");
            }

            else
            {

                string FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), FileName);
                FileInfo FileInfo = new FileInfo(FilePath);

                if (FileInfo.Exists)
                {

                    FileInfo.Delete();

                }

                using (StreamWriter writer = new StreamWriter(FilePath))
                {

                    var DadosFormatados = Dados.GroupBy(dado => dado.Associacao);

                    if (passo == 1)
                    {

                        foreach (var dado in DadosFormatados)
                        {

                            string CadastrarGrupo = string.Empty;
                            string GestorDepartamento = $"( select Id from AspNetUsers where [UserName] = '{dado.First().NomeUsuario}' )";

                            if (dado.Count() > 1)
                            {

                                CadastrarGrupo = $@"

                                    INSERT INTO AspNetGroups ( Id, Name, Descricao, Ativo, DataAtualizacao, UsarRegraDinamica, ConfiguracaoRegras, Status, GrupoParaImplantacao )
                                    VALUES ( NEWID(), 'Grupo - {dado.Key}', NULL, 1, NULL, 0, NULL, NULL, NULL )

                                    INSERT INTO AspNetUserGroups ( GrupoId, UsuarioId )
                                    select (select Id from AspNetGroups where [Name] = 'Grupo - {dado.Key}'), Id from AspNetUsers where [UserName] in ({string.Join(",", dado.Select(x => $"'{x.NomeUsuario}'").Distinct())})

                                ";

                                GestorDepartamento = $"(select Id from AspNetGroups where [Name] = 'Grupo - {dado.Key}')";

                            }

                            writer.WriteLine($@"                                

                                {CadastrarGrupo}

                                IF (select count(1) from Departamento where Nome = '{dado.Key}') = 0
                                BEGIN

                                    INSERT INTO Departamento ( Nome, GestorId, Padrao, Ativa)
                                    VALUES ( '{dado.Key}', {GestorDepartamento}, 0, 1 )

                                END

                                ELSE
                                BEGIN

                                   UPDATE Departamento set GestorId = {GestorDepartamento} where Nome = '{dado.Key}'

                                END                                                               

                            ");

                            writer.WriteLine(string.Empty);

                        }

                    }

                    if (passo == 2)
                    {

                        foreach (var dado in DadosFormatados)
                        {

                            var NomesUsuarios = dado.Select(x => x.NomeUsuario).Distinct();

                            foreach (var Nome in NomesUsuarios)
                            {

                                writer.WriteLine($@"

                                    IF (select count(1) from Departamento where Nome = '{dado.Key}') = 0
                                    BEGIN

                                        INSERT INTO Departamento ( Nome, GestorId, Padrao, Ativa)
                                        VALUES ( '{dado.Key}', (select GestorId from Departamento where Padrao = 1), 0, 1 )

                                    END

                                    UPDATE AspNetUsers set DepartamentoId = ( select Id from Departamento where Nome = '{dado.Key}' ) where UserName = '{Nome}'

                                    UPDATE Ideia set DepartamentoId = ( select DepartamentoId from AspNetUsers where [UserName] = '{Nome}' ) 
                                    where ElaboradorId = ( select Id from AspNetUsers where [UserName] = '{Nome}' )

                                ");

                            }


                        }

                    }

                    if (passo == 3)
                    {

                        foreach (var dado in DadosFormatados)
                        {

                            writer.WriteLine($@"

                                    IF ( select COUNT(1) from AspNetGroups where [Name] = '{dado.Key}' ) = 0
                                    BEGIN

                                        INSERT INTO AspNetGroups ( Id, Name, Descricao, Ativo, DataAtualizacao, UsarRegraDinamica, ConfiguracaoRegras, Status, GrupoParaImplantacao )
                                        VALUES ( NEWID(), '{dado.Key}', NULL, 1, NULL, 0, NULL, NULL, NULL )

                                    END

                                    INSERT INTO AspNetUserGroups ( GrupoId, UsuarioId )
                                    select (select Id from AspNetGroups where [Name] = '{dado.Key}'), Id from AspNetUsers where [UserName] in ({string.Join(",", dado.Select(x => $"'{x.NomeUsuario}'").Distinct())})
                                    
                                ");

                        }

                    }

                }

            }

            Console.WriteLine($"O arquivo foi salvo na sua área de trabalho com o nome de '{FileName}'");

            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine($"--------------- {passo}º PASSO PROCESSADO ---------------");

        }


        private static List<Dado> ReadFile(int passo)
        {

            try
            {

                int NumeroPlanilha = 0;
                int ColunaNomeUsuario = 2;
                int ColunaAssociado = 3;

                if (passo == 1)
                {

                    NumeroPlanilha = 2;

                }


                else if (passo == 2)
                {

                    NumeroPlanilha = 1;
                    ColunaNomeUsuario = 2;
                    ColunaAssociado = 4;

                }

                else if (passo == 3)
                {
                    NumeroPlanilha = 4;
                }

                List<Dado> Dados = new List<Dado>();

                using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo($"F:\\{NumeroPlanilha}.xlsx")))
                {

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var myWorksheet = xlPackage.Workbook.Worksheets.First();
                    var totalRows = myWorksheet.Dimension.End.Row;
                    var totalColumns = myWorksheet.Dimension.End.Column;

                    for (int rowNum = 2; rowNum <= totalRows; rowNum++)
                    {

                        var UserNameCell = myWorksheet.Cells[rowNum, ColunaNomeUsuario].Value;
                        var AssociadoCell = myWorksheet.Cells[rowNum, ColunaAssociado].Value;

                        if (UserNameCell != null && AssociadoCell != null)
                        {
                            Dados.Add(new Dado()
                            {
                                NomeUsuario = UserNameCell.ToString(),
                                Associacao = AssociadoCell.ToString()
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

        public class Dado
        {

            public string NomeUsuario { get; set; }
            public string Associacao { get; set; }

        }

        private class Opcoes
        {
            public int Id { get; set; }
            public string Titulo { get; set; }
        }

    }
}
