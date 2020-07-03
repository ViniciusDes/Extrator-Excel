
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ConsoleLeitorArquivos
{
	class ExtratorEmp11
	{
		public void LerArquivo()
		{
			string caminho = @"D:\Exe audit\sources\11\RELAÇÃO (1){353001}.xlsx";

			var wb = new XLWorkbook(@caminho);
			var planilha = wb.Worksheet(1);
			var listaNomesDep = new List<string>();
			var listaNomesTit = new List<string>();
			var listaCpfDep = new List<string>();
			var listaCpfTit = new List<string>();
			var listaDataNascTit = new List<string>();
			var listaDataNascDep = new List<string>();
			var listaDataIncluDep = new List<string>();
			var listaDataIncluTit = new List<string>();
			var listaValorDep = new List<string>();
			var listaValorTit = new List<string>();
			var listaMatInPlanoDep = new List<string>();
			var listaMatInPlanoTit = new List<string>();
			var listaDataExcluDep = new List<string>();
			var listaDataExcluTit = new List<string>();

			var listaNomeDepIsolado = new List<string>();
			var listaCpfDepIsolado = new List<string>();
			var listaNomesTitIsolado = new List<string>();
			var listaDataNascDepIsolado = new List<string>();
			var listaDataIncluDepIsolado = new List<string>();
			var listaDataExcluDepIsolado = new List<string>();
			var listaValorDepIsolado = new List<string>();
			var listaMatInPlanoDepIsolado = new List<string>();

			var linha = 1;
			int celVazia = 0;
			var nomeTit = planilha.Cell("J" + linha.ToString()).Value.ToString();

			var dic = new Dictionary<int, List<int>>();



			while (nomeTit == "" || nomeTit == "Nome Titular")
			{
				linha++;
				nomeTit = planilha.Cell("J" + linha.ToString()).Value.ToString();
			}
			while (true)
			{
				nomeTit = planilha.Cell("J" + linha.ToString()).Value.ToString();
				var nomeBenef = planilha.Cell("K" + linha.ToString()).Value.ToString();
				var cpf = planilha.Cell("G" + linha.ToString()).Value.ToString();
				var dependente = planilha.Cell("E" + linha.ToString()).Value.ToString();
				var dataNasc = planilha.Cell("L" + linha.ToString()).Value.ToString();
				var dataInclu = planilha.Cell("M" + linha.ToString()).Value.ToString();
				var dataExclu = planilha.Cell("N" + linha.ToString()).Value.ToString();
				var valor = planilha.Cell("T" + linha.ToString()).Value.ToString();
				var valorFat = planilha.Cell("U" + linha.ToString()).Value.ToString();
				var identTit = planilha.Cell("D" + linha.ToString()).Value.ToString();





				if (nomeTit == "")
				{
					celVazia++;
				}
				if (celVazia == 50)
				{
					break;
				}
				//else if (valor == "0" && valorFat == "0")
				//{

				//}
				else
				{
					if (dependente == "TITULAR")
					{
						if (identTit == "")
						{
							identTit = "0";
						}
						listaNomesTit.Add(nomeTit);
						listaCpfTit.Add(cpf);
						listaDataNascTit.Add(dataNasc);
						listaDataIncluTit.Add(dataInclu);
						listaValorTit.Add(valorFat);
						listaMatInPlanoTit.Add(identTit);
						listaDataExcluTit.Add(dataExclu);
						//pfTitular = cpf;

					}
					else if (dependente == "DEPENDENTE")
					{
						listaNomesDep.Add(nomeBenef);
						listaCpfDep.Add(cpf);
						listaDataIncluDep.Add(dataInclu);
						listaDataNascDep.Add(dataNasc);
						listaMatInPlanoDep.Add(identTit);
						listaValorDep.Add(valorFat);
						listaDataExcluDep.Add(dataExclu);
						if (identTit == "")
						{
							identTit = "0";
						}
						if (!dic.ContainsKey(Convert.ToInt32(identTit)))//se nao tiver a matricula como chave é criado o dicionario
						{
							dic.Add(Convert.ToInt32(identTit), new List<int>());

						}
						dic[Convert.ToInt32(identTit)].Add(listaNomesDep.Count() - 1);

					}
					else
					{

					}
					linha++;
				}
				//listaNomesTit.Sort();
			}

			var qtdTit = listaNomesTit.Count();
			var qtdDep = listaNomesDep.Count();

			//for (int k = 0; k < 100; k++)
			//{
			//	Console.WriteLine(listaNomesTit[k]);
			//}
			var qtd = listaNomesTit.Count();
			var listaDup = new List<int>();
			for (int i = 0; i < qtd; i++)
			{
				//Console.WriteLine("_______________________________________________________________________________________");
				//Console.WriteLine(listaMatInPlanoTit[i] + " " + listaNomesTit[i] + " " + listaCpfTit[i]);
				int j = 0;
				Console.WriteLine("_______________________________________________________________________________________");
				Console.WriteLine("Tit " + listaMatInPlanoTit[i] + " " + listaNomesTit[i] + " " + listaCpfTit[i]);
				var auxMat = Convert.ToInt32(listaMatInPlanoTit[i]);
				if (dic.ContainsKey(auxMat) && !listaDup.Contains(auxMat))
				{
					//Console.WriteLine(x.Key + " " + x.Value.Count());
					listaDup.Add(auxMat);
					foreach (var y in dic[auxMat])//para itens do dicionario
					{
						Console.WriteLine(auxMat + " " + listaNomesDep[y] + " " + listaCpfDep[y] + " " + listaNomesTit[i]);
					}

				}

			}
		}
	}
}
