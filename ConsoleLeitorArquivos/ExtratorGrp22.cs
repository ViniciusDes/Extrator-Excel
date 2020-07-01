using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ConsoleLeitorArquivos
{
	class ExtratorGrp22
	{

		

				

		public ExtratorGrp22()
		{
		}
		//public Tuple<List<string>,  Tuple<List<string>>> teste()
		//{
		//	var a = new List<string>();
		//	a.Add("AAAAAA");
		//	a.Add("BBBBBBBBBB");
		//	a.Add("CCCCCCCCCCCCC");
		//	var wb = new List<string>();
		//	wb.Add("DDDDDD");
		//	wb.Add("EEEEEEEEEE");
		//	wb.Add("FFFFFFFFFFFFF");
		//	var c = new List<string>();
		//	var aux = new Tuple<List<string>>(wb);

		//	return new Tuple<List<string>, Tuple<List<string>>>(a, aux);

		//}
		public void LerArquivo()
		{

			  var listaNomesDep = new List<string>();
		      var listaNomesAgr = new List<string>();
			  var listaNomesTit = new List<string>();
			  var listaMatTitGrupo = new List<string>();
			  var listaCpfDep = new List<string>();
			  var listaCpfAgr = new List<string>();
			  var listaCpfTit = new List<string>();
			  var listaDataNascTit = new List<string>();
		  	  var listaDataNascDep = new List<string>();
			  var listaDataNascAgr = new List<string>();
			  var listaDataIncluAgr = new List<string>();
			  var listaDataIncluDep = new List<string>();
			  var listaDataIncluTit = new List<string>();

			  var listaValorTotDep = new List<string>();
			  var listaValorMenDep = new List<string>();
			  var listaValorCopDep = new List<string>();
			  var listaValorOpcDep = new List<string>();
			  var listaTaxaIncluDep = new List<string>();

			  var listaValorTotTit = new List<string>();
			  var listaValorMenTit = new List<string>();
			  var listaValorCopTit = new List<string>();
			  var listaValorOpcTit = new List<string>();
			  var listaTaxaIncluTit = new List<string>();


			  var listaMatInPlanoDep = new List<string>();
			  var listaMatInPlanoAgr = new List<string>();
			  var listaMatInPlanoTit = new List<string>();
			 
			string caminho = @"D:\Exe audit\sources\22\Faturamento Analitico 8540 - 062020 - 21067323 (2){353292}.xlsx";

			var wb = new XLWorkbook(@caminho);
			var planilha = wb.Worksheet(1);
		 
			var linha = 1;
			var parentesco = "";

			string nomeTit = planilha.Cell("C" + linha.ToString()).Value.ToString();
			while (nomeTit.Length == 0 || nomeTit == "Nome Títular")
			{
				linha++;
				nomeTit = planilha.Cell("C" + linha.ToString()).Value.ToString();

			}
			var empresa = "";
			int qtdeTitu = 0;
			int celVazia = 0;
			//COlUNAS B = nomeTit, C = DATANASC, I = DEPENDENTE 
			while (true)
			{
				nomeTit = planilha.Cell("C" + linha.ToString()).Value.ToString();
				var nomeDep = planilha.Cell("E" + linha.ToString()).Value.ToString();

				parentesco = planilha.Cell("F" + linha.ToString()).Value.ToString();

				var matPlanoTit = planilha.Cell("B" + linha.ToString()).Value.ToString();
				var matPlanoDep = planilha.Cell("D" + linha.ToString()).Value.ToString();

				var dataNasc = planilha.Cell("G" + linha.ToString()).Value.ToString();
				var dataInclu = planilha.Cell("H" + linha.ToString()).Value.ToString();

				var valorTot = planilha.Cell("J" + linha.ToString()).Value.ToString();
				var valorMen = planilha.Cell("K" + linha.ToString()).Value.ToString();
				var valorCop = planilha.Cell("L" + linha.ToString()).Value.ToString();
				var valorOpc = planilha.Cell("M" + linha.ToString()).Value.ToString();
				var taxaInclu = planilha.Cell("N" + linha.ToString()).Value.ToString();

				empresa = planilha.Cell("P" + 2.ToString()).Value.ToString();

				if (nomeTit == "")
				{
					celVazia++;
				}
				if (celVazia == 50)
				{
					break;
				}
				else
				{
					if(parentesco == "T")
					{
						listaNomesTit.Add(nomeTit);
						listaMatInPlanoTit.Add(matPlanoTit);
						listaDataNascTit.Add(dataNasc);
						listaDataIncluTit.Add(dataInclu);
						listaValorTotTit.Add(valorTot);
						listaValorMenTit.Add(valorMen);
						listaValorCopTit.Add(valorCop);
						listaValorOpcTit.Add(valorOpc);
						listaTaxaIncluTit.Add(taxaInclu);

						qtdeTitu++;
					}
					if(parentesco != "T" && parentesco != "")
					{
						listaNomesDep.Add(nomeDep);
						listaMatInPlanoDep.Add(matPlanoDep);
						listaMatTitGrupo.Add(matPlanoTit);
						listaDataNascDep.Add(dataNasc);
						listaDataIncluDep.Add(dataInclu);
						listaValorTotDep.Add(valorTot);
						listaValorMenDep.Add(valorMen);
						listaValorCopDep.Add(valorCop);
						listaValorOpcDep.Add(valorOpc);
						listaTaxaIncluDep.Add(taxaInclu);
					}
					else
					{

					}
					linha++;
				}
				
			}
			var a = listaMatTitGrupo.Count();
			int j = 0;
			Console.WriteLine(empresa);
			for (int i = 0; i < qtdeTitu; i++)
			{
				var b = listaMatTitGrupo.Count();

				Console.WriteLine(listaMatInPlanoTit[i] + " " + listaNomesTit[i] + " Data Nasc: " + listaDataNascTit[i] + " Data Inclu: " + listaDataIncluTit[i] + " " +
				listaValorTotTit[i] + " " + listaValorMenTit[i] + " " + listaValorCopTit[i] + " " + listaValorOpcTit[i] + " " + listaTaxaIncluTit[i]
					);
				while (j < b)
				{
					if (j > 0 && listaMatTitGrupo[j] != listaMatInPlanoTit[i])
					{
						Console.WriteLine("\n");
						break;
					}
					else
					{
						Console.WriteLine(listaNomesDep[j] + " " + listaMatInPlanoTit[j] + " Data Nasc: " + listaDataNascDep[j] + " Data Inclu: " + listaDataIncluDep[j] + " " +
						listaValorTotDep[j] + " " + listaValorMenDep[j] + " " + listaValorCopDep[j] + " " + listaValorOpcDep[j] + " " + listaTaxaIncluDep[j]
							);
						j++;
					}
				}
			}

			//return new Tuple<List<string>, List<string> >(listaNomesTit, listaNomesDep);

		}

		public bool VerificarRelatorio()
		{
			bool temarquivo = false;
			//verifica se tem o relatorio ja preparado na pasta se nao a funcão que chama essa ira fazer a leitura do relatorio externo e preparar o outro
			//Marca o diretório a ser listado
			DirectoryInfo diretorio = new DirectoryInfo(@"D:\Exe audit\sources\23\");
			//Executa função GetFile(Lista os arquivos desejados de acordo com o parametro)
			FileInfo[] Arquivos = diretorio.GetFiles("*.*");

			//Começamos a listar os arquivos
			var a = Arquivos.Count();
			int contador = 0;
			foreach (FileInfo fileinfo in Arquivos)
			{
				Console.WriteLine(fileinfo.Name);//verifica se tem a planilha pronta a ser analisada se nao coleta as cop. do arquivo externo
				if (fileinfo.Name == "Faturamento Analitico.xlsx")
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

		public Tuple<List<string>, List<string>> LerArquivoExterno()
		{
			 
			var listaIdUsuarios = new List<string>();
			var listaValCop = new List<string>();
			string caminho = @"D:\Exe audit\sources\23\Relatorio de Fator Moderador_Ref052020 (9){353298}.xlsx";

			var temRel = VerificarRelatorio();

			if (temRel)
			{

			 }
			else
			{
				var wb = new XLWorkbook(@caminho);
				var planilha = wb.Worksheet(1);
		

				int linha = 1;
				int celVazia = 0;
				string identUsuario = planilha.Cell("F" + linha.ToString()).Value.ToString();
				var valor = planilha.Cell("M" + linha.ToString()).Value.ToString();
				string nome = planilha.Cell("G" + linha.ToString()).Value.ToString();

				while (identUsuario.Length == 0 || nome == "Nome Usuário")
				{
					linha++;
					identUsuario = planilha.Cell("F" + linha.ToString()).Value.ToString();
					nome = planilha.Cell("G" + linha.ToString()).Value.ToString();
				}
				double valoradsa = 0;
				int duplicado = 0;
				int qtdeItensLista = 0;


				while (true)
				{

					int linhaOld = linha - 1;
					var identUsuarioOld = planilha.Cell("F" + linhaOld.ToString()).Value.ToString();

					identUsuario = planilha.Cell("F" + linha.ToString()).Value.ToString();
					if (identUsuario != identUsuarioOld && identUsuario != "")
					{
						valor = planilha.Cell("M" + linha.ToString()).Value.ToString();

					}
					while (identUsuario == identUsuarioOld && identUsuario != "")
					{
						duplicado++;
						double vAux = Convert.ToDouble(planilha.Cell("M" + linha.ToString()).Value.ToString());
						valoradsa += Convert.ToDouble(vAux);
						if (duplicado <= 1)
						{
							valoradsa += Convert.ToDouble(planilha.Cell("M" + linhaOld.ToString()).Value.ToString());
						}
						valor = valoradsa.ToString();
						linha++;
						linhaOld = linha - 1;
						identUsuarioOld = planilha.Cell("F" + linhaOld.ToString()).Value.ToString();
						identUsuario = planilha.Cell("F" + linha.ToString()).Value.ToString();

					}
					qtdeItensLista = listaIdUsuarios.Count();

					valoradsa = 0;
					nome = planilha.Cell("G" + linha.ToString()).Value.ToString();



					if (nome == "")
					{
						celVazia++;
					}
					if (celVazia == 50)
					{
						break;
					}
					if (duplicado > 1)
					{
						listaIdUsuarios[qtdeItensLista - 1] = identUsuarioOld;
						listaValCop[qtdeItensLista - 1] = valor;
						duplicado = 0;
						linha--;

					}
					else if (nome != "")
					{
						listaIdUsuarios.Add(identUsuario);
						listaValCop.Add(Convert.ToString(valor));

					}


					linha++;
				}
				var qtd = listaValCop.Count();
				//for (int i = 0; i< qtd; i++)
				//{
				//	Console.WriteLine(listaIdUsuarios[i] + " " + listaValCop[i]);
				//}
			}
			 

			return new Tuple<List<string>, List<string>>(listaIdUsuarios, listaValCop);

		}

		public void LerArquivoEmp23()
			//esse metodo precisa de 2 arquivos pois as coparticipações fica no segundo arquivo
			//e o metodo vai ler e colocar em apenas um arquivo para extrair os dados completos para analise final.
		{
			   


			var listaNomesDep = new List<string>();
			var listaNomesAgr = new List<string>();
			var listaNomesTit = new List<string>();
			var listaMatTitGrupo = new List<string>();
			var listaCpfDep = new List<string>();
			var listaCpfAgr = new List<string>();
			var listaCpfTit = new List<string>();
			var listaDataNascTit = new List<string>();
			var listaDataNascDep = new List<string>();
			var listaDataNascAgr = new List<string>();
			var listaDataIncluAgr = new List<string>();
			var listaDataIncluDep = new List<string>();
			var listaDataIncluTit = new List<string>();

			var listaValorTotDep = new List<string>();
			var listaValorMenDep = new List<string>();
			var listaValorCopDep = new List<string>();
			var listaValorOpcDep = new List<string>();
			var listaTaxaIncluDep = new List<string>();

			var listaValorTotTit = new List<string>();
			var listaValorMenTit = new List<string>();
			var listaValorCopTit = new List<string>();
			var listaValorOpcTit = new List<string>();
			var listaTaxaIncluTit = new List<string>();


			var listaMatInPlanoDep = new List<string>();
			var listaMatInPlanoAgr = new List<string>();
			var listaMatInPlanoTit = new List<string>();
			var temarq = VerificarRelatorio();

			//caminho arquivo CONCRETO vai ser atualizado com os dados das coparticipações de um arquivo externo
			string caminho = @"D:\Exe audit\sources\23\Faturamento Analitico 7837 - 062020 - 21067292{353297}.xlsx";

			var wb = new XLWorkbook(@caminho);
			var planilha = wb.Worksheet(1);

			var linha = 1;
			var parentesco = "";

			string nomeTit = planilha.Cell("C" + linha.ToString()).Value.ToString();
			while (nomeTit.Length == 0 || nomeTit == "Nome Títular")
			{
				linha++;
				nomeTit = planilha.Cell("C" + linha.ToString()).Value.ToString();

			}
			var empresa = "";
			int qtdeTitu = 0;
			int celVazia = 0;
			//COlUNAS B = nomeTit, C = DATANASC, I = DEPENDENTE 

			var aux = LerArquivoExterno();
			
			if (temarq)
			{

				while (true)
				{

					nomeTit = planilha.Cell("C" + linha.ToString()).Value.ToString();
					var nomeDep = planilha.Cell("E" + linha.ToString()).Value.ToString();

					parentesco = planilha.Cell("F" + linha.ToString()).Value.ToString();

					var matPlanoTit = planilha.Cell("B" + linha.ToString()).Value.ToString();
					var matPlanoDep = planilha.Cell("D" + linha.ToString()).Value.ToString();

					var dataNasc = planilha.Cell("G" + linha.ToString()).Value.ToString();
					var dataInclu = planilha.Cell("H" + linha.ToString()).Value.ToString();

					//var valorTot  = planilha.Cell("J" + linha.ToString()).Value.ToString();
					var valorMen = planilha.Cell("K" + linha.ToString()).Value.ToString();
					var valorCop = planilha.Cell("L" + linha.ToString()).Value.ToString();
					var valorOpc = planilha.Cell("M" + linha.ToString()).Value.ToString();
					var taxaInclu = planilha.Cell("N" + linha.ToString()).Value.ToString();
					if (nomeTit != "" || nomeTit == "")
					{
						if (valorCop == "")
						{
							valorCop = "0";
						}		
						
						if (valorMen == "")
						{
							valorMen = "0";
						}
						var identUsuario = planilha.Cell("D" + linha.ToString()).Value.ToString();

						double valorTotal = Convert.ToDouble(valorMen) + Convert.ToDouble(valorCop);

						empresa = planilha.Cell("P" + 2.ToString()).Value.ToString();

						if (nomeTit == "")
						{
							celVazia++;
						}
						if (celVazia == 50)
						{
							break;
						}
						
						else
						{

							if (parentesco == "T")
							{
								listaNomesTit.Add(nomeTit);
								listaMatInPlanoTit.Add(matPlanoTit);
								listaDataNascTit.Add(dataNasc);
								listaDataIncluTit.Add(dataInclu);
								listaValorTotTit.Add(Convert.ToString(valorTotal));
								listaValorMenTit.Add(valorMen);
								listaValorCopTit.Add(valorCop);
								listaValorOpcTit.Add(valorOpc);
								listaTaxaIncluTit.Add(taxaInclu);

								qtdeTitu++;
							}
							if (parentesco != "T" && parentesco != "")
							{
								listaNomesDep.Add(nomeDep);
								listaMatInPlanoDep.Add(matPlanoDep);
								listaMatTitGrupo.Add(matPlanoTit);
								listaDataNascDep.Add(dataNasc);
								listaDataIncluDep.Add(dataInclu);
								listaValorTotDep.Add(Convert.ToString(valorTotal));
								listaValorMenDep.Add(valorMen);
								listaValorCopDep.Add(valorCop);
								listaValorOpcDep.Add(valorOpc);
								listaTaxaIncluDep.Add(taxaInclu);
							}
							else
							{

							}
							linha++;
						}



					}
				}

				
			}
			else
			{
				var identUsuario = planilha.Cell("D" + linha.ToString()).Value.ToString();
				int linha2 = 2;
				int contador = 0;
				int qtdeReg = 0;
				for (int l = 0; l < aux.Item1.Count(); l++)
				{
					Console.WriteLine(aux.Item1[l] + aux.Item2[l]);
				}
				Console.WriteLine("\n");

				while (contador <= qtdeReg)
				{
					//planilha.Cell("L" +linha2).Value = 
					foreach (string x in aux.Item1)//x é a ident de usuario se for igual o da planilha atualiza o valor da coparticipação
					{
						qtdeReg = aux.Item1.Count();

						if (x == planilha.Cell("D" + linha2.ToString()).Value.ToString())
						{
							planilha.Cell("L" + linha2).Value = Convert.ToDouble(aux.Item2[contador]);
							//wb.Save();
							contador++;
						}

						else if (x != planilha.Cell("D" + linha2.ToString()).Value.ToString())
						{
							int linhaaux = linha2;
							for (int k = 0; k < qtdeReg; k++)
							{
								linhaaux++;
								if (x == planilha.Cell("D" + linhaaux.ToString()).Value.ToString())
								{
									planilha.Cell("L" + linhaaux).Value = Convert.ToDouble(aux.Item2[contador]);
									//wb.Save();
									contador++;
								}
							}
						}

						linha2++;
					}
					break;//finalizou a sobrescrita do arquivo com os valores das coparticipações arquivo pronto para comparações
				}
				wb.SaveAs(@"D:\Exe audit\sources\23\Faturamento Analitico.xlsx");

			}

			int j = 0;
			var q = listaMatTitGrupo.Count();
			for (int c = 0; c < listaMatInPlanoTit.Count(); c++)
			{
				Console.WriteLine("---------------------------------------------------------------------------------------------------------------------------------------");
				Console.WriteLine("\n");
				var b = listaMatTitGrupo.Count();
				Console.WriteLine(listaMatInPlanoTit[c] + " " + listaNomesTit[c] + " " + listaValorTotTit[c]);
				while (j < q)
				{
					if (listaMatTitGrupo[j] != listaMatInPlanoTit[c])
					{

						Console.WriteLine("\n");
						//j ++;
						break;
					}
					else
					{
						Console.WriteLine(listaNomesDep[j] + " " + listaMatTitGrupo[j] + " " + listaValorTotDep[j]);
						j++;
					}
				}

			}


			//var a = listaMatTitGrupo.Count();
			//int j = 0;
			//Console.WriteLine(empresa);
			//for (int i = 0; i < qtdeTitu; i++)
			//{
			//	var b = listaMatTitGrupo.Count();

			//	Console.WriteLine(listaMatInPlanoTit[i] + " " + listaNomesTit[i] + " Data Nasc: " + listaDataNascTit[i] + " Data Inclu: " + listaDataIncluTit[i] + " " +
			//	 listaValorMenTit[i] + " " + listaValorCopTit[i] + " " + listaValorOpcTit[i] + " " + listaTaxaIncluTit[i]
			//		);
			//	while (j < b)
			//	{
			//		if (j > 0 && listaMatTitGrupo[j] != listaMatInPlanoTit[i])
			//		{
			//			Console.WriteLine("\n");
			//			break;
			//		}
			//		else
			//		{
			//			Console.WriteLine(listaNomesDep[j] + " " + listaMatInPlanoTit[j] + " Data Nasc: " + listaDataNascDep[j] + " Data Inclu: " + listaDataIncluDep[j] + " " +
			//			listaValorMenDep[j] + " " + listaValorCopDep[j] + " " + listaValorOpcDep[j] + " " + listaTaxaIncluDep[j]
			//				);
			//			j++;
			//		}
			//	}
			//}


		}
	}
}
