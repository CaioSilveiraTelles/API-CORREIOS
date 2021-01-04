using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel; //importação do pacote CLOSEDXML
using ConsoleApp1.ServiceReference1;
using DotNet.CEP.Search.App;//API CEP

namespace Prova
{

    class Program
    {
        //NECESSITA SER ASYNC POR TER QUE AGUARDAR A RESPOSTA DA API
        public static async Task Main(string[] args)

        {
            Console.WriteLine("Consultando Ceps");
            // Abrir arquivo excel existente

            var tabela = new XLWorkbook(@"C:\Users\caio.telles\Desktop\importar_bairros\bairro-taxa.xlsx"); //coloca o caminho da pasta com o arquivo XLSM do EXCEL
            var planilha = tabela.Worksheet(1);

            //  Console.WriteLine("".PadRight('-'));
            // Console.WriteLine("Bairros".PadRight(35) + "Taxa".PadRight(15) + "Entrega".PadRight(15));
            System.Collections.ArrayList Ceps = new System.Collections.ArrayList();

            //CRIAÇÃO DE VARIAVEIS
            var linha = 2;//CONTADOR DE LINHAS EXCEL

            string ok = "nao";//STRING DE REECONFIRMAÇÃO
            System.Collections.ArrayList CepsValidos = new System.Collections.ArrayList();//CRIAÇÃO DE ARRAY
            var Faixa = planilha.Cell("a" + linha.ToString()).Value.ToString();//CONVERSAO DE STRING FAIXA DOS CEPS
            var FaixaInicial = (planilha.Cell("b" + linha.ToString()).Value.ToString());////CONVERSAO DE STRING FAIXAINICIAL DOS CEPS
            var FaixaFinal = (planilha.Cell("c" + linha.ToString()).Value.ToString());//CONVERSAO DE STRING FAIXAFINAL DOS CEPS
            var lInicio = int.Parse(FaixaInicial);//CONVERSAO PARA USAR NA LOGICA FAIXA INICIAL
            var lfim = int.Parse(FaixaFinal);////CONVERSAO PARA USAR NA LOGICA FAIXA FINAL
            var Lok = int.Parse(FaixaInicial);// VARIAVEL PARA INCRIMENTAR TODOS OS CEPS PROCURADOS
            string Validcep = Lok.ToString();//TRANSFORMAR PARA STRING


            while (true)
            {
                while (true)
                {

                    if (Lok <= lfim) // LINHA DO CEP NAO PODE SER MAIOR QUE A FAIXA FINAL
                    {

                        Lok++;//INCRIMENTA O CEP

                        Validcep = Lok.ToString();
                        Ceps.Add(Validcep);//ADICIONA NO ARRAY
                        ConsultaCEP(ok, linha, Lok, Validcep, CepsValidos);//CHAMA A FUNÇÃO DE CONSULTA API
                    }

                    if (Lok > lfim)//CASO O CEP SEJA MAIOR QUE A FAIXA FINAL PULA PARA A PROXIMA LINHA , OBS NÃO PODE SER VAZIA A LINHA
                    {
                        //ATRIBUIS NOVOS VALORES PARA A VARIAVEL DE FAIXAS E SOMA A LINHA 
                        linha++;
                        Faixa = planilha.Cell("a" + linha.ToString()).Value.ToString();
                        FaixaInicial = (planilha.Cell("b" + linha.ToString()).Value.ToString());
                        FaixaFinal = (planilha.Cell("c" + linha.ToString()).Value.ToString());

                        //VERIFICA SE A LINHA ESTÁ VAZIA 
                        if (string.IsNullOrEmpty(Faixa) ||
                               string.IsNullOrEmpty(FaixaInicial) ||
                            string.IsNullOrEmpty(FaixaFinal)
                           )
                        {
                            //CASO ESTEJA VAZIA CHAMA A FUNÇÃO DE CONSULTA PASSANDO SIM , PARA INICIAR A CRIAÇÃO DO EXCEL
                            Validcep = Lok.ToString();
                            Ceps.Add(Validcep);
                            ok = "sim";//PROVOCA A CRIAÇÃO DO EXCEL
                            ConsultaCEP(ok, linha, Lok, Validcep, CepsValidos);
                            break;

                        }
                        lInicio = int.Parse(FaixaInicial);
                        lfim = int.Parse(FaixaFinal);
                        Lok = int.Parse(FaixaInicial);
                    }

                    break;

                }

            }

        }

        private static void ConsultaCEP(string ok, int linha, int Lok, string Validcep, ArrayList CepsValidos)//RECEBE TODAS AS INFORMAÇÕES VALIDAS
        {
            //FUNÇÃO RESPONSAVEL PELA API DOS CORREIOS 
            if (ok == "nao")//CASO FOR SIM CRIA O EXCEL
            {

                using (var wd = new AtendeClienteClient())//UTILZIA UMA FUNC DA BBLIOTECA PARA A BUSCA DO CEP
                {

                    try
                    {
                        string Search = Validcep;//CEP PROCURADO
                        var resposta = wd.consultaCEP(Search);//FAZ A CONSULTA PELA API
                        var end = wd.consultaCEP(Search).end;//FAZ A CONSULTA PELA API
                        var bairro = wd.consultaCEP(Search).bairro;//FAZ A CONSULTA PELA API
                        var cidade = wd.consultaCEP(Search).cidade;//FAZ A CONSULTA PELA API
                        var uf = wd.consultaCEP(Search).uf;//FAZ A CONSULTA PELA API
                        var cep = wd.consultaCEP(Search).cep;//FAZ A CONSULTA PELA API
                        var complemento = wd.consultaCEP(Search).complemento2;//FAZ A CONSULTA PELA API
                         DateTime thisDay = DateTime.Now;
                        if (resposta != null)
                        {
                            //CASO O RETORNO DA API NÃO SEJA ERRO , ADICIONA CADA INFORMAÇÃO NO ARRAY

                            CepsValidos.Add(end);
                            CepsValidos.Add(bairro);
                            CepsValidos.Add(cidade);
                            CepsValidos.Add(uf);
                            CepsValidos.Add(cep);
                            CepsValidos.Add(complemento);
                            CepsValidos.Add(thisDay);

                        }
                    }
                    catch
                    {
                        Console.WriteLine("CEP INVALIDO 404 Numero = " + Lok.ToString());//RETORNO 404 REST ERRO
                    }
                }
            }
            else
            {
                //CASO OK = SIM CHAMA A VERIFICAÇÃO E CRIAÇÃO DO EXCEL

                VerifyAPI(linha, CepsValidos);
            }

        }

        private static void VerifyAPI(int linha, ArrayList CepsValidos)
        {

            CriaExcel(CepsValidos);//CHAMA A CRIAÇÃO DO EXCEL PASSANDO O ARRAY COMPLETO 
        }

        private static void CriaExcel(ArrayList CepsValidos)
        {

            Console.WriteLine("Criando Excel");
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("planilha 1");
            var nArray = CepsValidos.Count / 7;
            var nCount = 0;

            //titulo

            ws.Cell("B2").Value = "PROVA ";
            var range = ws.Range("B2:I2");
            range.Merge().Style.Font.SetBold().Font.FontSize = 20;

            //Cabeçalho do Relatrio
            ws.Cell("B3").Value = "Logradouro/Nome";
            ws.Cell("C3").Value = "Bairro";
            ws.Cell("D3").Value = "Cidade";
            ws.Cell("E3").Value = "UF";
            ws.Cell("F3").Value = "CEP";
            ws.Cell("G3").Value = "Complemento";
            ws.Cell("H3").Value = "Processamento";

            //CORPO RELATORIO


            var linha = 4;


            for (int i = 0; i < nArray; i++)
            {


                using (var wd = new AtendeClienteClient())
                {

                    // CepsValidos[CepsValidos.Count - 1].end;
                    
                    //criação de cada campo por Array
                    var end = CepsValidos[i + nCount].ToString();
                    var bairro = CepsValidos[i + nCount + 1].ToString();
                    var cidade = CepsValidos[i + nCount + 2].ToString();
                    var uf = CepsValidos[i + nCount + 3].ToString();
                    var cep = CepsValidos[i + nCount + 4].ToString();
                    var complemento = CepsValidos[i + nCount + 5].ToString();
                    var processamento = CepsValidos[i + nCount + 6].ToString();

                    ws.Cell("B" + linha.ToString()).Value = (String.Format(end));
                    ws.Cell("C" + linha.ToString()).Value = (String.Format(bairro));
                    ws.Cell("D" + linha.ToString()).Value = (String.Format(cidade));
                    ws.Cell("E" + linha.ToString()).Value = (String.Format(uf));
                    ws.Cell("F" + linha.ToString()).Value = (String.Format(cep));
                    ws.Cell("G" + linha.ToString()).Value = (String.Format(complemento));
                    ws.Cell("H" + linha.ToString()).Value = (String.Format(processamento));
                    linha++;

                    nCount += 6;//logica para incrementar o array dinamicamente
                    


                }
            }
            linha--;

            wb.SaveAs(@"C:\Users\caio.telles\Desktop\importar_bairros\Resultado.xlsx");//SALVA o resultado no seu disco local
            wb.Dispose();//LIBERA A MEMORIA DO EXCEL
            Console.WriteLine("Excel Criado com Sucesso!");
            Console.ReadKey();

        }
    }

}
