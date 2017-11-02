using System;
using System.IO;
using dados;
namespace dados{
    /// <summary>
    /// Classe Menu - contém os menus do programa
    /// </summary>
    public class Menu{

        /// <summary>
        /// Método para mostrar o menu principal
        /// </summary>
        public void mostrarMenuPrincipal(){
            string path = Directory.GetCurrentDirectory();
            int opt;
            do {
                Console.WriteLine("Escola uma das opções abaixo\n"
                        + "1 - Cadastrar Cliente\n"
                        + @"2 - Depositar\Sacar\n"
                        + "3 - Obter Saldo\n"
                        + "4 - Sair\n"
                );
                Console.Write("Opção: ");
                opt = 4;            
                do{
                    opt = Int16.Parse(Console.ReadLine());
                } while (opt < 1 || opt > 4);
                switch(opt){
                    case 1: 
                            Cliente cliente = new Cliente();
                            string tipoDoc = mostrarMenuTipoCliente();
                            cliente.iniciarDados(tipoDoc);
                            string arquivo;
                            arquivo = tipoDoc == "CPF" ? arquivo = "PessoasFisicas.xlsx" : arquivo = "PessoasJuridicas.xlsx";
                            cliente.salvar( path + arquivo);
                            break;
                    /*case 2: 
                            break;
                    case 3: venderCarro(); break;*/
                }
            } while(opt != 0);
        }

        /// <summary>
        /// Método para mostrar o menu de escolha do tipo de cliente, CPF ou CNPJ
        /// </summary>
        /// <returns>Retorna string com o valor "CPF" ou "CNPJ", dependendo da escolha</returns>
        private string mostrarMenuTipoCliente(){
            string tipoDoc;
            Console.WriteLine("Escolha o tipo do cliente:\n"
                        + "1 - Pessoa Física\n"
                        + "2 - Pessoa Jurídica\n");
            do{
                Console.Write("Opção: ");
                tipoDoc = Console.ReadLine();
            } while( tipoDoc != "1" && tipoDoc != "2");
            if (tipoDoc.Equals("1")){
                return "CPF";
            } else {
                return "CNPJ";
            }
        }
    }
}