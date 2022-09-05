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
                },

                new Opcoes()
                {
                    Id = 4,
                    Titulo = "Processar 4º Passo"
                },

                new Opcoes()
                {
                    Id = 5,
                    Titulo = "De Para Irani"
                },

                new Opcoes()
                {
                    Id = 6,
                    Titulo = "Gestor do Departamento"
                },

                new Opcoes()
                {
                    Id = 7,
                    Titulo = "De-Para Produção Irani"
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

                            var Usuarios = dado.Select(x => $"'{x.NomeUsuario}'").Distinct();
                            if (Usuarios.Count() > 1)
                            {

                                CadastrarGrupo = $@"

                                    INSERT INTO AspNetGroups ( Id, Name, Descricao, Ativo, DataAtualizacao, UsarRegraDinamica, ConfiguracaoRegras, Status, GrupoParaImplantacao )
                                    VALUES ( NEWID(), 'Grupo - {dado.Key}', NULL, 1, NULL, 0, NULL, NULL, NULL )

                                    INSERT INTO AspNetUserGroups ( GrupoId, UsuarioId )
                                    select (select Id from AspNetGroups where [Name] = 'Grupo - {dado.Key}'), Id from AspNetUsers where [UserName] in ({string.Join(",", Usuarios)})

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
                                    select (select Id from AspNetGroups where [Name] = '{dado.Key}'), Id from AspNetUsers where [UserName] in ({string.Join(",", dado.Select(x => $"'{x.NomeUsuario.Trim()}'").Distinct())})
                                    
                            ");

                        }

                    }

                    if (passo == 4)
                    {

                        foreach (var dado in Dados)
                        {

                            if (dado.NomeUsuario != "37806523")
                            {

                                writer.WriteLine($@"

                                    UPDATE Departamento set GestorId = ( select Id from AspNetUsers where UserName = '{dado.NomeUsuario}' )
                                    where Nome = '{dado.Associacao}'

                                ");

                            }


                        }

                    }

                    if (passo == 5)
                    {

                        var UsuariosAtivos = Dados.Where(dado => dado.Ativo).ToList();
                        foreach (var usuario in UsuariosAtivos)
                        {

                            writer.WriteLine($@"

                                    UPDATE AspNetUsers SET UserName = '{usuario.Associacao}' where UserName = '{usuario.NomeUsuario}'

                            ");

                        }

                    }

                    if (passo == 10)
                    {

                        writer.WriteLine($@"

                            insert into AspNetUserClaims (UserId, ClaimType, ClaimValue)
                            select Id, '5330b15b-1b51-4d48-977a-74a7e72e67f2', 'True' from AspNetUsers 
                            where Email in (
                                {string.Join("\n,", Dados.Select(dado => $"'{dado.Associacao}'"))}
                            )    

                        ");

                    }

                    if (passo == 6)
                    {

                        foreach (var dado in Dados)
                        {

                            writer.WriteLine($@"

                                update Departamento set GestorId = ( select Id from AspNetUsers where UserName = '{dado.NomeUsuario}' )
                                where Nome = '{dado.Associacao}'

                            ");

                        }

                    }

                    if (passo == 7)
                    {

                        foreach (var dado in Dados)
                        {

                            writer.WriteLine($@"

                                update AspNetUsers set UserName = '{dado.Associacao}' where UserName = '{dado.NomeUsuario}'

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
                int Ativo = 6;

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

                else if (passo == 4)
                {
                    NumeroPlanilha = 11;
                    ColunaNomeUsuario = 2;
                    ColunaAssociado = 3;
                }

                else if (passo == 10)
                {
                    NumeroPlanilha = 10;
                    ColunaNomeUsuario = 1;
                    ColunaAssociado = 2;
                }

                else if (passo == 5)
                {
                    ColunaNomeUsuario = 2;
                    ColunaAssociado = 4;
                }

                else if (passo == 6)
                {
                    ColunaNomeUsuario = 2;
                    ColunaAssociado = 3;
                    NumeroPlanilha = 13;
                }

                else if (passo == 7)
                {
                    ColunaNomeUsuario = 2;
                    ColunaAssociado = 3;
                    NumeroPlanilha = 666;
                }

                List<Dado> Dados = new List<Dado>();

                using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo($"F:\\{(passo == 5 ? "irani" : NumeroPlanilha)}.xlsx")))
                {

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var myWorksheet = xlPackage.Workbook.Worksheets.First();
                    var totalRows = myWorksheet.Dimension.End.Row;
                    var totalColumns = myWorksheet.Dimension.End.Column;

                    for (int rowNum = 2; rowNum <= totalRows; rowNum++)
                    {

                        var UserNameCell = myWorksheet.Cells[rowNum, ColunaNomeUsuario].Value;
                        var AssociadoCell = myWorksheet.Cells[rowNum, ColunaAssociado].Value;
                        //var AtivoCell = myWorksheet.Cells[rowNum, Ativo].Value;

                        if (UserNameCell != null && AssociadoCell != null)
                        {

                            Dados.Add(new Dado()
                            {
                                NomeUsuario = UserNameCell.ToString(),
                                Associacao = AssociadoCell.ToString()
                                //Ativo = Convert.ToBoolean(AtivoCell)
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
            public bool Ativo { get; set; }

        }

        private class Opcoes
        {
            public int Id { get; set; }
            public string Titulo { get; set; }
        }

    }
}
