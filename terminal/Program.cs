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

        /// <summary>
        /// Método para decisão do menu principal do programa
        /// </summary>
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
                    case 2: String operacao = new Menu().mostrarMenuDepositarSacar();
                            tipoDoc = new Menu().mostrarMenuTipoCliente();
                            arquivoCliente = tipoDoc.Equals("CPF")? path + "PessoasFisicas.xlsx" : path + "PessoasJuridicas.xlsx";
                            arquivoConta = path + "Contas.xlsx";
                            Conta contaCliente = new Conta();
                            string doc = tipoDoc.Equals("CPF") ? new Validador().pedirCPF() : new Validador().pedirCNPJ();
                            venda.cliente = new Pessoa().carregarPessoa(Int64.Parse(doc) , arquivoCliente);
                            Console.Write("Digite o valor para " + operacao + ": ");
                            double valor = double.Parse(Console.ReadLine());



                            break;
                    /*case 3: ObterSaldo(); break;*/
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
            cliente.Documento = tipoDoc.Equals("CPF") ? new Validador().pedirCPF() : new Validador().pedirCNPJ();
            Console.Write("Nome do Cliente: ");
            cliente.Nome = Console.ReadLine();
            Console.Write("Email do cliente: ");
            cliente.Email = Console.ReadLine();
            cliente.Endereco = iniciarEndereco();
            return cliente;
        }

        public Cliente carregarCliente(Int64 doc, String arquivo){
            Application ex = new Application();
            ex.Workbooks.Open(arquivo);
            Cliente cliente = new Cliente();
            int linha = 2;
            while(Int64.Parse(ex.Cells[linha, 1].Value.ToString()) != doc && ex.Cells[linha,1].Value != null ){
                linha++;
            }
            cliente.Documento = ex.Cells[linha, 1].Value.ToString();
            cliente.Nome = ex.Cells[linha, 2].Value.ToString();
            cliente.Email = ex.Cells[linha, 3].Value.ToString();
            cliente.Endereco = new Endereco();
            cliente.Endereco.Rua = ex.Cells[linha, 4].Value.ToString();
            cliente.Endereco.Numero = Int16.Parse(ex.Cells[linha, 5].Value.ToString());
            cliente.Endereco.Bairro = ex.Cells[linha, 6].Value.ToString(); 
            ex.ActiveWorkbook.Close();
            ex.Quit();
            ex.Dispose();
            return cliente;
        }

        /// <summary>
        /// Método para iniciar os dados do Endereço do Cliente
        /// </summary>
        /// <returns>Retorna o objeto endereco</returns>
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

        /// <summary>
        /// Método para Abrir uma conta
        /// </summary>
        /// <param name="tipoDoc">"CPF" para Pessoas Físicas e "CNPJ" para Pessoas Jurídicas</param>
        /// <returns>Retorna o Objeto conta</returns>
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