using System;
using System.IO;
using System.Collections;
using NetOffice.ExcelApi;
using util;
namespace dados
{
    /// <summary>
    /// Classe Cliente
    /// </summary>    
    public class Cliente{
        public string documento {get; set;}
        public string nome {get; set;}
        //public string DataNascimento {get; set;}
        public string email {get; set;}
        public Endereco endereco {get; set;}

        /// <summary>
        /// Iniciar dados do Cliente
        /// </summary>
        /// <param name="tipoDoc">Esperado tipo do Documento  = CPF ou CNPJ</param>
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

        /// <summary>
        /// Salva os dados do clinete no arquivo excel especificado
        /// </summary>
        /// <param name="arquivo">Path + nome do arquivo .xlsx</param>
        public void salvar(String arquivo){
            Application ex = new Application();
            if(!File.Exists(arquivo) || getUltimaLinha(arquivo) == 1){
                gerarCabecalho(arquivo);
            }
            ex.Workbooks.Open(arquivo);
            int ultimaLinha = getUltimaLinha(arquivo);
            ex.Cells[ultimaLinha, 1].Value = this.documento;
            ex.Cells[ultimaLinha, 2].Value = this.nome;
            ex.Cells[ultimaLinha, 3].Value = this.email;
            ex.Cells[ultimaLinha, 4].Value = this.endereco.rua;
            ex.Cells[ultimaLinha, 5].Value = this.endereco.numero;
            ex.Cells[ultimaLinha, 6].Value = this.endereco.bairro;
            ex.Cells[ultimaLinha, 7].Value = DateTime.Now;
            ex.ActiveWorkbook.Save();
            ex.Quit();
            ex.Dispose();
        }

        /// <summary>
        /// Carrega e retorna os dados do cliente
        /// </summary>
        /// <param name="doc">número do documento para identificar o cliente</param>
        /// <param name="arquivo">arquivo excel de cadastro dos clientes</param>
        /// <returns>retorna o objeto cliente</returns>
        public Cliente carregarObjeto(int doc, String arquivo){
        Application ex = new Application();
        ex.Workbooks.Open(arquivo);
        Cliente cliente = new Cliente();
        int linha = 2;
        int campo = 1;
        while(!ex.Cells[linha, 1].Value.ToString().Contains(doc.ToString())){
            linha++;
        }
        foreach(var propriedade in doc.GetType().GetProperties()){
            if(propriedade.PropertyType.IsClass &&  !propriedade.PropertyType.Name.Equals("String")) {
                foreach(var subPropriedade in propriedade.GetType().GetProperties()){
                    subPropriedade.SetValue(doc, ex.Cells[linha, campo].Value);
                    campo++;
                }
            } else {
                propriedade.SetValue(doc, ex.Cells[linha, campo].Value);
                campo++;
            }
        }
        return cliente;
    }

    /// <summary>
    /// Buscar cliente
    /// </summary>
    /// <param name="arquivo">arquivo excel de cadastro dos clientes</param>
    /// <param name="doc">número do documento para identificar o cliente convertido em String</param>
    /// <returns>Retorna Array com os dados do cliente</returns>
    public ArrayList buscarCliente(String arquivo, String doc ){
        if(File.Exists(arquivo)){
            ArrayList codigos = new ArrayList();
            Application ex = new Application();
            ex.Workbooks.Open(arquivo);
            int numCampo = 1;
            String cabecalho = null, resultado = null;
            int linha = 0;
            do{
                linha++;
                if(ex.Cells[linha, numCampo].Value.ToString().Equals(doc)){
                    numCampo = 1;
                    while(!ex.Cells[linha, numCampo].Value.Equals(null)){
                        if(numCampo == 1){
                            codigos.Add(ex.Cells[linha, numCampo].Value);
                        } 
                        resultado += ex.Cells[linha, numCampo].Value.ToString() + " | ";
                        numCampo++;
                    }
                    if(!resultado.Equals(null)){
                        resultado += "\n";
                    }
                }
            } while (ex.Cells[linha,1].Value != null);
            if(!resultado.Equals(null)){
                numCampo = 1;
                while(!ex.Cells[linha, numCampo].Value.Equals(null)){
                    cabecalho += ex.Cells[1, numCampo].Value.ToString() + " | ";
                    numCampo++;
                }
                Console.WriteLine("Resultado(s) encontrado(s): ");
                Console.WriteLine(cabecalho);
                Console.WriteLine(resultado);
                return codigos;
            } else {
                Console.WriteLine("O termo buscado não foi encontrado");
                return null;
            } 
            ex.Quit();
        } else {
            Console.WriteLine("O arquivo " + arquivo + " não foi encontrado!");
            return null;
        }
    }

        /// <summary>
        /// Gera o cabeçalho no arquivo de cadastro
        /// </summary>
        /// <param name="arquivo">Path + nome do arquivo .xlsx utilizado no cadstro</param>
        private void gerarCabecalho(String arquivo){
            Application ex = new Application();
            bool existeArquivo = File.Exists(arquivo);
            if(!existeArquivo){
                ex.Workbooks.Add();
            } else {
                ex.Workbooks.Open(arquivo);
            }
            if(!File.Exists(arquivo) || getUltimaLinha(arquivo) == 1){
                ex.Cells[1, 1].Value = "Documento";
                ex.Cells[1, 2].Value = "Nome";
                ex.Cells[1, 3].Value = "E-mail";
                ex.Cells[1, 4].Value = "Rua";
                ex.Cells[1, 5].Value = "Número";
                ex.Cells[1, 6].Value = "Bairro";
                ex.Cells[1, 7].Value = "Data";
            }
            if(existeArquivo){
                ex.ActiveWorkbook.Save();
            } else {
                ex.ActiveWorkbook.SaveAs(arquivo);
            }
            ex.Quit();
            ex.Dispose();
        }

        /// <summary>
        /// Retorna a ultima linha em branco do arquivo de cadastro
        /// </summary>
        /// <param name="arquivo">Path + nome do arquivo .xlsx utilizado no cadstro</param>
        /// <returns>número da ultima linha em branco</returns>
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