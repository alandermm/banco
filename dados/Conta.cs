namespace dados
{
    /// <summary>
    /// Classe Conta
    /// </summary>
    public class Conta{
        public int numero {get; set;}
        public double saldo {get; private set;}
        public Cliente titular {get; set;}

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
    }    
}