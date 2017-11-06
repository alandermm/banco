using System;
using System.IO;
using NetOffice.ExcelApi;
namespace dados
{
    /// <summary>
    /// Classe Conta
    /// </summary>
    public class Conta{
        public int numero {get; private set;}
        public int Numero {set => numero = value;}
        public double saldo {get; private set;}
        public double Saldo {set => saldo = value;}
        public Cliente titular {get; private set;}
        public Cliente Titular {set => titular = value;}

        public void salvar(String arquivo){
            Application ex = new Application();
            int ultimaLinha = new Arquivo().getUltimaLinha(arquivo);
            /*if(!File.Exists(arquivo) || ultimaLinha == 1){
                String[] cabecalho = new String[]{"Conta", "Documento", "Nome", "Saldo", "Data abertura"};
                new Arquivo().gerarCabecalho(arquivo, cabecalho);
            }*/
            ex.Workbooks.Open(arquivo);
            ex.Cells[ultimaLinha, 1].Value = this.numero;
            ex.Cells[ultimaLinha, 2].Value = this.titular.documento.ToString();
            ex.Cells[ultimaLinha, 3].Value = this.titular.nome;
            ex.Cells[ultimaLinha, 4].Value = this.saldo;
            ex.Cells[ultimaLinha, 5].Value = DateTime.Now;
            //ex.Cells.AutoFit();
            
            ex.ActiveWorkbook.Save();
            ex.ActiveWorkbook.Close();
            ex.Quit();
            ex.Dispose();
        }

        /*private void gerarCabecalho(String arquivo){
            Application ex = new Application();
            bool existeArquivo = File.Exists(arquivo);
            if(!existeArquivo){
                ex.Workbooks.Add();
            } else {
                ex.Workbooks.Open(arquivo);
            }
            if(!File.Exists(arquivo) || Arquivo.getUltimaLinha(arquivo) == 1){
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
        }*/

        /// <summary>
        /// Método para sacar dinheiro
        /// </summary>
        /// <param name="valor">Valor a ser sacado</param>
        public void Sacar(double valor){
            this.saldo -= valor;
        }

        /// <summary>
        /// Método para depositar dinheiro
        /// </summary>
        /// <param name="valor">Valor a ser depositado</param>
        public void Depositar(double valor){
            this.saldo += valor;
        }

        /// <summary>
        /// Método para obter saldo da conta
        /// </summary>
        /// <returns>Saldo da conta</returns>
        public double MeuSaldo(){
            return this.saldo;
        }

        /*private static int getUltimaLinha(String arquivo){
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
        }*/
    }    
}