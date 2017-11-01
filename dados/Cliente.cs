using System;
using System.IO;
using NetOffice.ExcelApi;
using util;
namespace dados
{
    public class Cliente{
        public string documento {get; set;}
        public string nome {get; set;}
        //public string DataNascimento {get; set;}
        public string email {get; set;}
        public Endereco endereco {get; set;}

/// <summary>
/// 
/// </summary>
/// <param name="tipoDoc"></param>
        public void iniciarDados(String tipoDoc){
            Console.Write("Nome do Cliente: ");
            this.nome = Console.ReadLine();
            Console.Write("Email do cliente: ");
            this.email = Console.ReadLine();
            Console.Write(tipoDoc + " do cliente: ");
            Validador documento = new Validador();
            this.documento = tipoDoc.Equals("CPF") ? documento.pedirCPF() : documento.pedirCNPJ();
            this.endereco = new Endereco();
            Console.Write("Rua: ");
            this.endereco.rua = Console.ReadLine();
            Console.Write("Número: ");
            this.endereco.numero = Int16.Parse(Console.ReadLine());
            Console.Write("Bairro: ");
            this.endereco.bairro = Console.ReadLine();
        }

        public void salvar(String arquivo){
            Application ex = new Application();
            if(!File.Exists(arquivo) || getUltimaLinha(arquivo) == 1){
                gerarCabecalho(arquivo);
            }
            ex.Workbooks.Open(arquivo);
            int ultimaLinha = getUltimaLinha(arquivo);
            ex.Cells[ultimaLinha, 1].Value = ultimaLinha + 1;
            ex.Cells[ultimaLinha, 2].Value = this.nome;
            ex.Cells[ultimaLinha, 3].Value = this.email;
            ex.Cells[ultimaLinha, 4].Value = this.documento;
            ex.Cells[ultimaLinha, 5].Value = this.endereco.rua;
            ex.Cells[ultimaLinha, 6].Value = this.endereco.numero;
            ex.Cells[ultimaLinha, 7].Value = this.endereco.bairro;
            ex.Cells[ultimaLinha, 8].Value = DateTime.Now;
            ex.ActiveWorkbook.Save();
            ex.Quit();
            ex.Dispose();
        }

        private void gerarCabecalho(String arquivo){
            Application ex = new Application();
            bool existeArquivo = File.Exists(arquivo);
            if(!existeArquivo){
                ex.Workbooks.Add();
            } else {
                ex.Workbooks.Open(arquivo);
            }
            if(!File.Exists(arquivo) || getUltimaLinha(arquivo) == 1){
                ex.Cells[1, 1].Value = "Cód Cliente";
                ex.Cells[1, 2].Value = "Nome";
                ex.Cells[1, 3].Value = "E-mail";
                ex.Cells[1, 4].Value = "Documento";
                ex.Cells[1, 5].Value = "Rua";
                ex.Cells[1, 6].Value = "Número";
                ex.Cells[1, 7].Value = "Bairro";
                ex.Cells[1, 8].Value = "Data";
            }
            if(existeArquivo){
                ex.ActiveWorkbook.Save();
            } else {
                ex.ActiveWorkbook.SaveAs(arquivo);
            }
            ex.Quit();
            ex.Dispose();
        }

        private static int getUltimaLinha(String arquivo){
            int contador = 0;
            Application ex = new Application();
            if(File.Exists(arquivo)){
                ex.Workbooks.Open(arquivo);
                do{
                    contador++;
                } while (ex.Cells[contador,1].Value != null);
                ex.Quit();
                ex.Dispose();
            } else {
                contador = 1;
            }
            return contador;
        }
    }
}