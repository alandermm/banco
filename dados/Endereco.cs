using System;
/// <summary>
/// Classe Endereço
/// </summary>
public class Endereco {
    public String rua { get; private set;}
    public String Rua {set => rua = value;}
    public int numero { get; private set;}
    public int Numero {set => numero = value;}
    public String bairro { get; private set;}
    public String Bairro {set => bairro = value;}
}