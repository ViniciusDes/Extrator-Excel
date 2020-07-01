
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ConsoleLeitorArquivos
{
    class ExtratorEmp27
    {

        public void LerArquivo()
        {
            //EXTRATOR PARA RELATORIO DE FORNECEDOR (AMERICA SAUDE) 

            var caminho = @"D:\Exe audit\sources\27\America saúde{353203}.xlsx";

            var wb = new XLWorkbook(@caminho);
            var planilha = wb.Worksheet(1);
            var linha = 1;

            var nomeTit = planilha.Cell("B" + linha.ToString()).Value.ToString();
            string nomeDep = planilha.Cell("A" + linha.ToString()).Value.ToString();
            var listaNomeTit = new List<string>();
            var listaNomeDep = new List<string>();
            var listaNomeTitParaDep = new List<string>();
            var listaValorDep = new List<string>();
            var listaValorTit = new List<string>();
            double valor;
            int celVazia = 0;

            while (nomeDep == "" || nomeDep == "Segurado")
            {
                linha++;
                nomeDep = planilha.Cell("A" + linha.ToString()).Value.ToString();
            }
            while (true)
            {
                nomeTit = planilha.Cell("B" + linha.ToString()).Value.ToString();
                nomeDep = planilha.Cell("A" + linha.ToString()).Value.ToString();
                if (nomeDep == "" && nomeTit == "")
                {

                }
                else
                {
                    nomeTit = planilha.Cell("B" + linha.ToString()).Value.ToString();
                    nomeDep = planilha.Cell("A" + linha.ToString()).Value.ToString();
                    valor = Convert.ToDouble(planilha.Cell("G" + linha.ToString()).Value.ToString());

                     


                    if (nomeDep != "" && nomeTit == "")//titular ta na lista de dependente entao adciono a lista de tit
                    {
                        listaNomeTit.Add(nomeDep);//nome dep pode ser o TITULAR
                        listaValorTit.Add(Convert.ToString(valor));
                    }
                    else if(nomeDep != "" && nomeTit != "")
                    {
                        listaNomeDep.Add(nomeDep);
                        listaValorDep.Add(Convert.ToString(valor));
                        listaNomeTitParaDep.Add(nomeTit);
                    }
                    linha++;
                }
                if (nomeDep == "")
                {
                    celVazia++;
                }
                if (celVazia == 50)
                {
                    break;
                }
            }

            var qtdeTit = listaNomeTit.Count();
            var j = 0;
            for(int i = 0; i < qtdeTit; i++)
            {
                Console.WriteLine("----------------------------------------------------------------------------------------------");
                Console.WriteLine(listaNomeTit[i] + " " + listaValorTit[i]);

                while(j < qtdeTit)
                {
                    if(listaNomeTit[i] == "JOAO MARCOS BORGES DA SILVA")
                    {

                    }
                    if(j >= listaNomeTitParaDep.Count())
                    {
                        j++;
                    }
                    else if(listaNomeTitParaDep[j] == listaNomeTit[i])
                    {
                        Console.WriteLine(listaNomeDep[j] + listaNomeTitParaDep[j]);
                       
                        j++;
                    }
                    else
                    {
                        break;
                    }
                }


            }
             
             
        }

        public Tuple<List<string>, List<string> > RetornarDadosExternos()//pega o arquivo externo de coparticipação e gera as lista com os dados para preencher na planilha de analise
        {
            var caminho = @"D:\Exe audit\sources\27\Unimed Cooparticipação{353202}.xlsx";

            var wb = new XLWorkbook(caminho);
            var planilha = wb.Worksheet(1);
            var linha = 1;

 
            var listaCpf = new List<string>();
            var listaValor = new List<string>();
        
            int celVazia = 0;
            int duplicado = 0;
            double valor2 = 0;
            int qtdeItensLista;
            while (true)
            {
                var cpf = planilha.Cell("J" + linha.ToString()).Value.ToString();
                var valorTot = planilha.Cell("AA" + linha.ToString()).Value.ToString();
                if (cpf ==  "")
                {
                    celVazia++;

                }

                if(celVazia == 50)
                {
                    break;
                }

                while (cpf != "CPF" && cpf != "")
                {
                    int linhaOld = linha - 1;
                    var cpfOld = planilha.Cell("J" + linhaOld.ToString()).Value.ToString();


                    cpf = planilha.Cell("J" + linha.ToString()).Value.ToString();
                    valorTot = planilha.Cell("AA" + linha.ToString()).Value.ToString();

                    while(cpf == cpfOld && cpf != "")
                    {
                        duplicado++;
                        double vAux = Convert.ToDouble(planilha.Cell("AA" + linha.ToString()).Value.ToString());
                        valor2 += Convert.ToDouble(vAux);
                        if (duplicado <= 1)
                        {
                            valor2 += Convert.ToDouble(planilha.Cell("AA" + linhaOld.ToString()).Value.ToString());
                        }
                        valorTot = valor2.ToString();
                        linha++;
                        linhaOld = linha - 1;
                        cpfOld = planilha.Cell("J" + linhaOld.ToString()).Value.ToString();
                        cpf = planilha.Cell("J" + linha.ToString()).Value.ToString();


                    }
                    qtdeItensLista = listaCpf.Count();

                    valor2 = 0;
                    cpf = planilha.Cell("J" + linha.ToString()).Value.ToString();

                    if (duplicado > 1)
                    {
                        listaCpf[qtdeItensLista - 1] = cpfOld;
                        listaValor[qtdeItensLista - 1] = valorTot;
                        duplicado = 0;
                        linha--;

                    }
                    else if (cpf != "")
                    {
                        listaCpf.Add(cpf);
                        listaValor.Add(Convert.ToString(valorTot));

                    }


                    linha++;
                  
                }

                linha++;
                
            }
            int qtdeReg = listaValor.Count();
            for(int i = 0; i < qtdeReg; i++)
            {
                Console.WriteLine(listaCpf[i] + " " + listaValor[i]);
            }

            return new Tuple<List<string>, List<string>>(listaCpf, listaValor);
        }

        public bool VerificaRelatorio()
        {
            bool temarquivo = false;
            //verifica se tem o relatorio ja preparado na pasta se nao a funcão que chama essa ira fazer a leitura do relatorio externo e preparar o outro
            //Marca o diretório a ser listado
            DirectoryInfo diretorio = new DirectoryInfo(@"D:\Exe audit\sources\27\");
            //Executa função GetFile(Lista os arquivos desejados de acordo com o parametro)
            FileInfo[] Arquivos = diretorio.GetFiles("*.*");

            //Começamos a listar os arquivos
            var a = Arquivos.Count();
            int contador = 0;
            foreach (FileInfo fileinfo in Arquivos)
            {
                Console.WriteLine(fileinfo.Name);//verifica se tem a planilha pronta a ser analisada se nao coleta as cop. do arquivo externo
                if (fileinfo.Name == "Unimed mensalidades.xlsx")
                {
                    temarquivo = true;

                    return temarquivo;
                }
                else if (contador < a)
                {
                    contador++;
                }
                else
                {
                    temarquivo = false;
                    return temarquivo;
                }
            }

            return temarquivo;
             
        }

        public void LerArquivoUnimed()//A EMPRESA QUE TEM ESSE RELATORIO NAO TA DEFINIDA POIS HA FUNCIONARIOS NA MESMA FILIAL DO OUTRO RELATORIO DA (AMERICA SAUDE)
        {
         

            //PRIMEIRO PASSO VERIFICAR SE HA O RELATORIO JA PREPARADO ANTES DE EXTRAIR 
            var temArq = VerificaRelatorio();
            int contador = 0;
            int qtdeReg = 0;
            if (temArq)
            {
                var caminho = @"D:\Exe audit\sources\27\Unimed mensalidades.xlsx";
                var wb = new XLWorkbook(caminho);
                var planilha = wb.Worksheet(1);
                int linha = 1;

                var nome = planilha.Cell("N" + linha.ToString()).Value.ToString();
                var cpf = planilha.Cell("O" + linha.ToString()).Value.ToString();
                var dependente = planilha.Cell("I" + linha.ToString()).Value.ToString();
                var valorInclu = planilha.Cell("R" + linha.ToString()).Value.ToString();
                var valorMen = planilha.Cell("S" + linha.ToString()).Value.ToString();
                var codFamilia = planilha.Cell("H" + linha.ToString()).Value.ToString();

                var listaNomeDep = new List<string>();
                var listaNomeTit = new List<string>();

                var listaCpfTit = new List<string>();
                var listaCpfDep = new List<string>();

               
                var listaValorTit = new List<string>();
                
                var listaValorDep = new List<string>();
                
                var listacodFamiliaDep = new List<string>();
                var listacodFamiliaTit = new List<string>();
                int celVazia = 0;
                var qtdR = 0;
                while (true)
                {

                    nome = planilha.Cell("N" + linha.ToString()).Value.ToString();
                    if (nome == "")
                    {
                        celVazia++;
                    }
                    if (celVazia == 50)
                    {
                        break;
                    }

                    while (nome != "" && nome != "Nome do Beneficiário") {

                        nome = planilha.Cell("N" + linha.ToString()).Value.ToString();
                        dependente = planilha.Cell("I" + linha.ToString()).Value.ToString();
                        codFamilia = planilha.Cell("H" + linha.ToString()).Value.ToString();
                        valorInclu = planilha.Cell("R" + linha.ToString()).Value.ToString();
                        cpf = planilha.Cell("O" + linha.ToString()).Value.ToString();
                        
                    

                        if(valorInclu == "")
                        {
                            valorInclu = "0";
                        }   
                        valorMen = planilha.Cell("S" + linha.ToString()).Value.ToString();
                        if (valorMen == "")
                        {
                            valorMen = "0";
                        }

                        var valorTot = (Convert.ToDouble(valorMen) + Convert.ToDouble(valorInclu));

                    
                        if (dependente == "TITULAR" && dependente != "G. Dep." )
                        {
                            listaNomeTit.Add(nome);
                            listaValorTit.Add(valorTot.ToString());
                            listacodFamiliaTit.Add(codFamilia);
                            listaCpfTit.Add(cpf);
                            qtdR++;
                        }
                        else if(dependente != "")
                        {
                            listaNomeDep.Add(nome);
                            listaValorDep.Add(valorTot.ToString());
                            listacodFamiliaDep.Add(codFamilia);
                            listaCpfDep.Add(cpf);
                           
                        }
                        linha++;
                    }
                    linha++;
                }
                qtdR = listaCpfDep.Count();
                var qtdRT = listaCpfTit.Count();
                int j = 0;
                for(int i = 0; i < qtdRT; i++)
                {
                    Console.WriteLine("_______________________________________________________________________________________________");
                    Console.WriteLine(listaNomeTit[i] + " " + listaCpfTit[i]);
                    while(j < qtdR)
                    {
                        if(listacodFamiliaDep[j] == listacodFamiliaTit[i])
                        {
                            Console.WriteLine(listaNomeDep[j] + " " + listaCpfDep[j]);
                            j++;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                
            }
            else//se nao tiver o caminho ira pegar o original das mensalidaes e colocar as coparticipações e criar outro arquivo para ser extraido
            {
                var caminho2 = @"D:\Exe audit\sources\27\Unimed mensalidades{353201}.xlsx";//esse arquivo ira ser alterado e criado uma copia com nome correto

                var wb2 = new XLWorkbook(@caminho2);
                var planilha2 = wb2.Worksheet(1);
                var linha2 = 1;
                var cpf = planilha2.Cell("O" + linha2.ToString()).Value.ToString();
                var listasValoreCpf = RetornarDadosExternos();

                planilha2.Cell("W" + 1).Value = "Valor Cop.";
                planilha2.Cell("X" + 1).Value = "Valor Old.";

                while (contador <= qtdeReg)
                {
                    cpf = planilha2.Cell("O" + linha2.ToString()).Value.ToString();

                    if (cpf != "Cpf" && cpf != "" && cpf != "CPF")
                    {
                        foreach(string x in listasValoreCpf.Item1)//x é o cpf se for igual o dessa 
                        {                                         //se for igual atualiza o valor de cop.

                            qtdeReg = listasValoreCpf.Item1.Count();

                            if (x == planilha2.Cell("O" + linha2.ToString()).Value.ToString())
                            {
                                var vAtual = planilha2.Cell("S" + linha2).Value.ToString();
                                var valorAux = Convert.ToDouble(listasValoreCpf.Item2[contador]) + Convert.ToDouble(vAtual);

                                planilha2.Cell("S" + linha2).Value = valorAux;
                                contador++;
                                //wb2.Save();
                            }
                            else if (x != planilha2.Cell("O" + linha2.ToString()).Value.ToString())
                            {
                                int linhaaux = linha2;
                                for (int k = 0; k < qtdeReg; k++)
                                {
                                    linhaaux++;
                                    if (x == planilha2.Cell("O" + linhaaux.ToString()).Value.ToString())
                                    {
                                        var vAtual = planilha2.Cell("S" + linhaaux).Value.ToString();
                                        var valorAux = Convert.ToDouble(listasValoreCpf.Item2[contador]) + Convert.ToDouble(vAtual);

                                        planilha2.Cell("S" + linhaaux).Value = valorAux; 
                                        //wb2.Save();
                                        contador++;
                                    }
                                }
                            }

                            linha2++;
                        }
                        break;
                    }
                    linha2++;
                }
                wb2.SaveAs(@"D:\Exe audit\sources\27\Unimed mensalidades.xlsx");
            }

        }
    }
}
