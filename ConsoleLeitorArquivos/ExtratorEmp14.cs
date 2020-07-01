using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using ClosedXML.Excel;

public class ExtratorEmp14

{

	public ExtratorEmp14()
	{
		
	}
	
	public void LerArquivo()
    {
		string caminho = @"D:\Exe audit\sources\14\Relação faturados Unimed _Frec XLSX PTESTE.xlsx";

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
		var linha = 1;
		string cpfTitular = "";
		string dependente;
		var listaCpfTitGrupo = new List<string>();
		//string valorColM = planilha.Cell("M" + linha.ToString()).Value.ToString();//valor


		//planilha.Cell("B" + linha.ToString()).Value.ToString();
		string nome = planilha.Cell("B" + linha.ToString()).Value.ToString();
		while(nome.Length == 0 || nome == "NOME BENEFICIÁRIO")
		{
			linha++;
			nome = planilha.Cell("B" + linha.ToString()).Value.ToString();
			
		}
		var qtdLinhas = planilha.LastCellUsed();
		Console.WriteLine(Convert.ToString(qtdLinhas) + "ULTIMA CELULA USADA");
		
		int celVazia = 0;
		//COlUNAS B = NOME, C = DATANASC, I = DEPENDENTE 
		while (true)
		{
			nome = planilha.Cell("B" + linha.ToString()).Value.ToString();
			var cpf = planilha.Cell("S" + linha.ToString()).Value.ToString().Replace(".", "");
			dependente = planilha.Cell("I" + linha.ToString()).Value.ToString();
			var dataNasc = planilha.Cell("C" + linha.ToString()).Value.ToString();
			var dataInclu = planilha.Cell("D" + linha.ToString()).Value.ToString();
			var valor = planilha.Cell("P" + linha.ToString()).Value.ToString();
			var matPlano = planilha.Cell("A" + linha.ToString()).Value.ToString();

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
				if (dependente == "TITULAR")
				{
					listaNomesTit.Add(nome);
					listaCpfTit.Add(cpf);
					listaDataNascTit.Add(dataNasc);
					listaDataIncluTit.Add(dataInclu);
					listaValorTit.Add(valor);
					listaMatInPlanoTit.Add(matPlano);
					cpfTitular = cpf;
				}
				else if (dependente == "AGREGADO")
				{
					//estou adc como dependente tbm os agregados
					listaNomesDep.Add(nome);
					listaCpfDep.Add(cpf);
					listaDataIncluDep.Add(dataInclu);
					listaDataNascDep.Add(dataNasc);
					listaMatInPlanoDep.Add(matPlano);
					listaValorDep.Add(valor);


					listaNomesAgr.Add(nome);
					listaCpfAgr.Add(cpf);
					listaDataIncluAgr.Add(dataInclu);
					listaDataNascAgr.Add(dataNasc);
					listaCpfTitGrupo.Add(cpfTitular);
					listaMatInPlanoAgr.Add(matPlano);
					listaValorAgr.Add(valor);


				}
				else if (dependente == "DIRETO")
				{
					listaNomesDep.Add(nome);
					listaCpfDep.Add(cpf);
					listaCpfTitGrupo.Add(cpfTitular);
					listaDataIncluDep.Add(dataInclu);
					listaMatInPlanoDep.Add(matPlano);
					listaDataNascDep.Add(dataNasc);
					listaValorDep.Add(valor);


				}
				else
				{

				}
				linha++;
			}
			

		}
		int j = 0;
		int a = listaCpfTitGrupo.Count();

		for (int i = 0; i < a; i++)
		{
			Console.WriteLine(listaNomesTit[i] + " CPF: " + listaCpfTit[i] + " Data Nasc " + listaDataNascTit[i] + " Data Inclu: " + listaDataIncluTit[i] + " Valor: " + listaValorTit[i]);
			int b = listaCpfTitGrupo.Count();
			while (j < b)
			{
				if (j > 0 && listaCpfTitGrupo[j] != listaCpfTit[i])
				{

					Console.WriteLine("\n");
					//j ++;
					break;
				}
				else
				{
					Console.WriteLine(listaNomesDep[j] + " CPF DEP.: " + listaCpfDep[j] + " CPF TIT: " + listaCpfTitGrupo[j] + " Data Nasc " + listaDataNascDep[j] + " Data Inclu: " + listaDataIncluDep[j] + " Valor: " + listaValorDep[j]);
					j++;
				}
			}
		}
		 
	}
}
