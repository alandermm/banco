using System;
using System.IO;
using System.Collections;
using NetOffice.ExcelApi;
using util;


namespace dados{
    /// <summary>
    /// Classe Programa
    /// </summary>
    class Program{
        /// <summary>
        /// Método Main
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args){
            decisaoMenuPrincipal();
        }

        private static void decisaoMenuPrincipal(){
            Menu programa = new Menu();
            string path = Directory.GetCurrentDirectory() + "\\";
            int opt;
            do{
                opt = programa.mostrarMenuPrincipal();
                switch(opt){
                    case 1: String tipoDoc = new Menu().mostrarMenuTipoCliente();
                            String arquivoCliente = tipoDoc == "CPF" ? path + "PessoasFisicas.xlsx" : path + "PessoasJuridicas.xlsx";
                            String arquivoConta = path + "Contas.xlsx";
                            Conta conta = AbrirConta(tipoDoc);
                            if(conta != null){    
                                conta.titular.salvar(arquivoCliente);
                                conta.salvar(arquivoConta);
                            }
                            break;
                    /*case 2: Depositar/Sacar
                            break;
                    case 3: ObterSaldo(); break;*/
                    case 4: Environment.Exit(0); break;
                }
            } while(opt != 4);
        }

        /// <summary>
        /// Iniciar dados do Cliente
        /// </summary>
        /// <param name="tipoDoc">Esperado tipo do Documento  = CPF ou CNPJ</param>
        public static Cliente iniciarCliente(String tipoDoc){
            Cliente cliente = new Cliente();
            Console.Write(tipoDoc + " do cliente: ");
            Validador documento = new Validador();
            cliente.Documento = tipoDoc.Equals("CPF") ? documento.pedirCPF() : documento.pedirCNPJ();
            Console.Write("Nome do Cliente: ");
            cliente.Nome = Console.ReadLine();
            Console.Write("Email do cliente: ");
            cliente.Email = Console.ReadLine();
            cliente.Endereco = iniciarEndereco();
            return cliente;
        }

        public static Endereco iniciarEndereco(){
            Endereco endereco = new Endereco();
            Console.Write("Rua: ");
            endereco.Rua = Console.ReadLine();
            Console.Write("Número: ");
            endereco.Numero = Int16.Parse(Console.ReadLine());
            Console.Write("Bairro: ");
            endereco.Bairro = Console.ReadLine();
            return endereco;
        }

        public static Conta AbrirConta(String tipoDoc){
            string path = Directory.GetCurrentDirectory() + "\\";
            string arquivo = path + "Contas.xlsx";
            String arquivoCliente = tipoDoc == "CPF" ? path + "PessoasFisicas.xlsx" : path + "PessoasJuridicas.xlsx";
            int ultimaLinha = new Arquivo().getUltimaLinha(arquivoCliente);
            if(!File.Exists(arquivoCliente) || ultimaLinha == 1){
                String[] cabecalho = new String[]{"Documento", "Nome", "E-mail", "Rua", "Número", "Bairro", "Data"};
                new Arquivo().gerarCabecalho(arquivoCliente, cabecalho);
            }
            ultimaLinha = new Arquivo().getUltimaLinha(arquivo);
            if(!File.Exists(arquivo) || ultimaLinha == 1){
                String[] cabecalho = new String[]{"Conta", "Documento", "Nome", "Saldo", "Data abertura"};
                new Arquivo().gerarCabecalho(arquivo, cabecalho);
            }
            int linha = 1;
            Conta conta = new Conta();
            if(File.Exists(arquivo)){
                conta.Titular = iniciarCliente(tipoDoc);
                ultimaLinha = new Arquivo().getUltimaLinha(arquivo);
                Application ex = new Application();
                ex.Workbooks.Open(arquivo);
                string documento = conta.titular.documento;
                while(ex.Cells[linha,1].Value != null && !ex.Cells[linha,2].Value.ToString().Equals(documento)){
                    linha++;
                }
                ex.ActiveWorkbook.Close();
                ex.Quit();
                ex.Dispose();
                if(linha == ultimaLinha ){
                    conta.Numero = ultimaLinha;
                    conta.Saldo = 0;
                    return conta;
                } else {
                    Console.WriteLine("Cliente já cadastrado!");
                    return null;
                }
            } else {
                Console.WriteLine("O arquivo " + arquivo + " não existe");
                return null;
            }
        }
    }
}