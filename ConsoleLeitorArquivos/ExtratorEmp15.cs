using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ConsoleLeitorArquivos
{
    class ExtratorEmp15
    {
		public void LerArquivo()
		{

			//SERVE PARA TODOS DO GRUPO EMPRESARIAL 15 !!  

			string caminho = @"D:\Exe audit\sources\15\Unimed - Operar{355116}.xlsx";

			var wb = new XLWorkbook(@caminho);
			var planilha = wb.Worksheet(1);
			var listaNomesDep = new List<string>();
			var listaNomesAgr = new List<string>();
			var listaNomesTit = new List<string>();
			var listaCpfDep = new List<string>();
			var listaCpfAgr = new List<string>();
			var listaCpfTit = new List<string>();
			var listaDataNascTit = new List<string>();
			var listaDataNascDep = new List<string>();
			var listaDataNascAgr = new List<string>();
			var listaDataIncluAgr = new List<string>();
			var listaDataIncluDep = new List<string>();
			var listaDataIncluTit = new List<string>();
			var listaValorDep = new List<string>();
			var listaValorAgr = new List<string>();
			var listaValorTit = new List<string>();
			var listaMatInPlanoDep = new List<string>();
			var listaMatInPlanoAgr = new List<string>();
			var listaMatInPlanoTit = new List<string>();
			var listaDataExcluDep = new List<string>();
			var listaDataExcluTit = new List<string>();
			var linha = 1;
			string cpfTitular = "";
			string dependente;
			var listaCpfTitGrupo = new List<string>();
			//string valorColM = planilha.Cell("M" + linha.ToString()).Value.ToString();//valor


			//planilha.Cell("B" + linha.ToString()).Value.ToString();
			string nome = planilha.Cell("C" + linha.ToString()).Value.ToString();
			while (nome.Length == 0 || nome == "Beneficiário")
			{
				linha++;
				nome = planilha.Cell("C" + linha.ToString()).Value.ToString();

			}
			var qtdLinhas = planilha.LastCellUsed();
			Console.WriteLine(Convert.ToString(qtdLinhas) + "ULTIMA CELULA USADA");

			int celVazia = 0;
			//COlUNAS B = NOME, C = DATANASC, I = DEPENDENTE 
			while (true)
			{
				nome = planilha.Cell("C" + linha.ToString()).Value.ToString();
				var cpf = planilha.Cell("E" + linha.ToString()).Value.ToString().Replace(".", "");
				dependente = planilha.Cell("G" + linha.ToString()).Value.ToString();
				var dataNasc = planilha.Cell("S" + linha.ToString()).Value.ToString();
				var dataInclu = planilha.Cell("K" + linha.ToString()).Value.ToString();
				var valor = planilha.Cell("O" + linha.ToString()).Value.ToString();
				var matPlano = planilha.Cell("B" + linha.ToString()).Value.ToString();
				var matTitu = planilha.Cell("D" + linha.ToString()).Value.ToString();
				var dataExclu = planilha.Cell("L" + linha.ToString()).Value.ToString();
				if(dataExclu == "")
				{
					dataExclu = "none";
				}
			
				if (nome == "")
				{
					celVazia++;
				}
				if (celVazia == 50)
				{
					break;
				}


				//if (dependente == "" && nome.Length > 0 )
				//{
				//	linha++;
				//}
				else
				{
					if (dependente == "T")
					{
						listaNomesTit.Add(nome);
						listaCpfTit.Add(cpf);
						listaDataNascTit.Add(dataNasc);
						listaDataIncluTit.Add(dataInclu);
						listaValorTit.Add(valor);
						listaMatInPlanoTit.Add(matTitu);
						listaDataExcluTit.Add(dataExclu);
						cpfTitular = cpf;
					}
					else if (dependente == "A")
					{
						//estou adc como dependente tbm os agregados
						listaNomesDep.Add(nome);
						listaCpfDep.Add(cpf);
						listaDataIncluDep.Add(dataInclu);
						listaDataNascDep.Add(dataNasc);
						listaMatInPlanoDep.Add(matPlano);
						listaValorDep.Add(valor);
						//listaDataExclu.Add(dataExclu);
						listaMatInPlanoTit.Add(matTitu);


						listaNomesAgr.Add(nome);
						listaCpfAgr.Add(cpf);
						listaDataIncluAgr.Add(dataInclu);
						listaDataNascAgr.Add(dataNasc);
						listaCpfTitGrupo.Add(cpfTitular);
						listaMatInPlanoAgr.Add(matTitu);
						listaDataExcluDep.Add(dataExclu);
						listaValorAgr.Add(valor);
						listaMatInPlanoTit.Add(matTitu);

					}
					else if (dependente == "D")
					{
						listaNomesDep.Add(nome);
						listaCpfDep.Add(cpf);
						listaCpfTitGrupo.Add(cpfTitular);
						listaDataIncluDep.Add(dataInclu);
						listaMatInPlanoDep.Add(matTitu);
						listaDataNascDep.Add(dataNasc);
						listaValorDep.Add(valor);
						listaDataExcluDep.Add(dataExclu);
						listaMatInPlanoTit.Add(matTitu);

					}
					else
					{

					}
					linha++;
				}


			}
			int j = 0;
			int a = listaCpfTitGrupo.Count();
			//var w = listaCpfTit.Count();

			//for (int g = 0; g < a; g++)
			//{
			//	Console.WriteLine(listaCpfTitGrupo[g]);

			//}

			var f = listaCpfTit.Count();
			for(int i = 0; i < f; i++)
			{
				if(i == 63)
				{

				}
				Console.WriteLine("---------------------------------------------------------------------------------------------------------------------------------------");
				Console.WriteLine("\n");
				Console.WriteLine(listaCpfTit[i] + " " + listaNomesTit[i] + " "+ listaMatInPlanoTit[i] + " Data Nasc " + listaDataNascTit[i] + " Data Inclu: " + listaDataIncluTit[i] + " Data Exclu: " + listaDataExcluTit[i] + " Valor: " + listaValorTit[i]);
				{
					while(j < a)
					{
						if (listaCpfTitGrupo[j] != listaCpfTit[i])
						{

							//Console.WriteLine("---------------------------------------------------------------------------------------------------------------------------------------");
							Console.WriteLine("\n");
							//j ++;
							break;
						}
						else
						{
							Console.WriteLine(listaNomesDep[j] + " CPF DEP.: " + listaCpfDep[j] + " CPF TIT: " + listaCpfTitGrupo[j] + " Data Nasc " + listaDataNascDep[j] + " Data Inclu: " + listaDataIncluDep[j] + " Data Exclu: " + listaDataExcluDep[j] + " Valor: " + listaValorDep[j]);
							j++;
						}
					}
				}
			}

			//for (int i = 0; i < a; i++)
			//{
			//	Console.WriteLine(listaNomesTit[i] + " CPF: " + listaCpfTit[i] + " Data Nasc " + listaDataNascTit[i] + " Data Inclu: " + listaDataIncluTit[i] + " Valor: " + listaValorTit[i]);
			//	int b = listaCpfTitGrupo.Count();
			//	while (j < b)
			//	{
			//		if (j > 0 && listaCpfTitGrupo[j] != listaCpfTit[i])
			//		{

			//			Console.WriteLine("\n");
			//			//j ++;
			//			break;
			//		}
			//		else
			//		{
			//			Console.WriteLine(listaNomesDep[j] + " CPF DEP.: " + listaCpfDep[j] + " CPF TIT: " + listaCpfTitGrupo[j] + " Data Nasc " + listaDataNascDep[j] + " Data Inclu: " + listaDataIncluDep[j] + " Valor: " + listaValorDep[j]);
			//			j++;
			//		}
			//	}
			//}

		}

	}
}
