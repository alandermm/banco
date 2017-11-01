namespace dados
{
    public class Conta{
        public int numero {get; set;}
        public double saldo {get; private set;}
        public Cliente titular {get; set;}

        public void Sacar(double valor){
            this.saldo -= valor;
        }

        public void Depositar(double valor){
            this.saldo += valor;
        }

        public double MeuSaldo(){
            return this.saldo;
        }
    }    
}