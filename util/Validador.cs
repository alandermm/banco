using System;
using System.Text.RegularExpressions;

namespace util{
    
    /// <summary>
    /// Classe Validador - Utilizada para validar documentos
    /// </summary>
    public class Validador{
        private string doc {get; set;}
        private string primeiroDigito {get; set;}
        private string segundoDigito {get; set;}
        private bool valido {get; set;} = false;
        
        /// <summary>
        /// Método para limpar a entada de dados, restando apenas os números do documento
        /// </summary>
        /// <param name="doc">String com o número do documento</param>
        /// <returns></returns>
        private string limparCaracteresDocumento(string doc){
            return doc.Trim().Replace("/","").Replace("-","").Replace(".","");
        }

        /// <summary>
        /// Método validador do dígito do documento
        /// </summary>
        /// <param name="chave">Chave utilizada para validação do documento</param>
        /// <returns>Retorna o dígito</returns>
        private string validarDigito(int[] chave){
            string tempdoc;
            int soma = 0, resto = 0;
            tempdoc = this.doc.Substring(0, chave.Length);
            for(int i = 0; i < chave.Length; i++){
                soma += ((int)Char.GetNumericValue(tempdoc[i]) * chave[i]);
            }
            resto = soma % 11;
            if(resto < 2){
                return "0";
            } else {
                return (11-resto).ToString();
            } 
        }

        /// <summary>
        /// Método para solicitar o CPF
        /// </summary>
        /// <returns>Retorna o CPF solicitado</returns>
        public string pedirCPF(){
            do{
                this.doc = limparCaracteresDocumento(Console.ReadLine());
                this.validarCPF();
            } while (!this.valido);
            return this.doc;
        }

        /// <summary>
        /// Método para validar CPF
        /// </summary>
        /// <returns>Retorna true se o CPF for válido ou false para inválido</returns>
        private bool validarCPF(){
            Regex rgx = new Regex(@"^\d*$");
            int[] chaveCPF = {10,9,8,7,6,5,4,3,2};
            int[] chaveCPF2 = {11,10,9,8,7,6,5,4,3,2};
            if(this.doc.Length != 11 || !rgx.IsMatch(this.doc)){
                return this.valido;
            }
            this.primeiroDigito = validarDigito(chaveCPF);
            if(this.primeiroDigito != this.doc.Substring(9, 1)){
                Console.WriteLine("CPF inválido!\n");
                return this.valido;
            } else {
                this.segundoDigito = validarDigito(chaveCPF2);
                if(this.doc.EndsWith(this.segundoDigito)){
                    Console.WriteLine("CPF válido!\n");
                    return this.valido = true;
                } else {
                    Console.WriteLine("CPF inválido!\n");
                    return this.valido;
                }
            }
        }

        /// <summary>
        /// Método para solicitar o CNPJ
        /// </summary>
        /// <returns>Retorna o CNPJ solicitado</returns>
        public string pedirCNPJ(){
            do{
                this.doc = limparCaracteresDocumento(Console.ReadLine());
                this.validarCNPJ();
            } while (!this.valido);
            return this.doc;
        }

        /// <summary>
        /// Método para validar CNPJ
        /// </summary>
        /// <returns>Retorna true se o CNPJ for válido ou false para inválido</returns>
        private bool validarCNPJ(){
            Regex rgx = new Regex(@"^\d*$");
            int[] chaveCNPJ = {5,4,3,2,9,8,7,6,5,4,3,2};
            int[] chaveCNPJ2 = {6,5,4,3,2,9,8,7,6,5,4,3,2};
            
            if(this.doc.Length != 14 || !rgx.IsMatch(this.doc)){
                return this.valido;
            }

            this.primeiroDigito = validarDigito(chaveCNPJ);

            if(this.primeiroDigito != this.doc.Substring(12, 1)){
                Console.WriteLine("CNPJ inválido!\n");
                return this.valido;
            }else {
                this.segundoDigito = validarDigito(chaveCNPJ2);
                if(this.doc.EndsWith(this.segundoDigito)){
                    Console.WriteLine("CNPJ válido!\n");
                    return this.valido = true;
                } else {
                    Console.WriteLine("CNPJ inválido!\n");
                    return this.valido;
                }
            }
        }
    }
}