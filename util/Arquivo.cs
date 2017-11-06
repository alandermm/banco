using System;
using System.IO;
using NetOffice.ExcelApi;
public class Arquivo{
    public int getUltimaLinha(String arquivo){
        int contador = 0;
        Application ex = new Application();
        if(File.Exists(arquivo)){
            ex.Workbooks.Open(arquivo);
            do{
                contador++;
            } while (ex.Cells[contador,1].Value != null);
            ex.Quit();
            //ex.Dispose();
        } else {
            contador = 1;
        }
        return contador;
    }

    public void gerarCabecalho(String arquivo, String[] cabecalho){
            Application ex = new Application();
            bool existeArquivo = File.Exists(arquivo);
            if(!existeArquivo){
                ex.Workbooks.Add();
            } else {
                ex.Workbooks.Open(arquivo);
            }
            if(!File.Exists(arquivo) || new Arquivo().getUltimaLinha(arquivo) == 1){
                for (int i = 0; i < cabecalho.Length; i++){
                    ex.Cells[1,i+1].Value = cabecalho[i];
                }
            }
            if(existeArquivo){
                ex.ActiveWorkbook.Save();
            } else {
                ex.ActiveWorkbook.SaveAs(arquivo);
            }
            ex.Quit();
            //ex.Dispose();
        }
}