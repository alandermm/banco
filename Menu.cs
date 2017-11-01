using System;
using System.IO;
public class Menu{
    public void mostrarMenuPrincipal(){
        string path = Directory.GetCurrentDirectory();
        int opt;
        do {
            Console.WriteLine("Escola uma das opções abaixo\n"
                    + "1 - Cadastrar Cliente\n"
                    + "2 - Cadastrar Carro\n"
                    + "3 - Vender Carro\n"
                    + "4 - Listar Carros Vendidos no dia\n"
                    + "0 - Sair\n");
            Console.Write("Opção: ");
            opt = 0;            
            do{
                opt = Int16.Parse(Console.ReadLine());
            } while (opt < 0 || opt > 6);
            switch(opt){
                case 0: Environment.Exit(0); break;
                case 1: 
                        Pessoa pessoa = new Pessoa();
                        string tipoDoc = mostrarMenuTipoCliente();
                        pessoa.iniciarDados(tipoDoc);
                        Cadastro<Pessoa> cadastroCliente = new Cadastro<Pessoa>();
                        string arquivo;
                        arquivo = tipoDoc == "CPF" ? arquivo = "PessoasFisicas.xlsx" : arquivo = "PessoasJuridicas.xlsx";
                        cadastroCliente.salvar( path + arquivo, pessoa);
                        break;
                case 2: Carro carro = new Carro();
                        carro.iniciarDados();
                        Cadastro<Carro> cadastroCarro = new Cadastro<Carro>();
                        cadastroCarro.salvar(path + "carros.xlsx", carro);
                        break;
                /*case 3: venderCarro(); break;
                case 4: listarCarroVendidoDia(); break;*/
            }
        } while(opt != 0);
    }
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