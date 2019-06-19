function Preenche_Carta() 
{		
	var tm_total = new Timer();	

	//Laço por operacoes no json
	for( var i = 0; i < operacoes.length; i++ )
	{
		try
		{
			AutoIt.ToolTip( "Processando: " + ( i + 1 )  + " de " + operacoes.length + "registro(s).", 10, 10 );

			//Instancia o objeto Word e abre o contrato selecionado
			wrdObject = WScript.CreateObject("Word.Application");
			wrdObject.Visible = true;	
			wrd = wrdObject.Documents.Open( fileModelo );
			wrd.ActiveWindow.View.ReadingLayout = false;

			var minuta = pathRede + "Habilitação Massiva_" + operacoes[i].Contrato + "_" + operacoes[i].Empresas[0].Nome.replace(/[^\A-Za-z\d ]+/g, "") + ".pdf";

			//Insere os dados Contrato, Razão Social e CNPJ
			wrd.FormFields("Contrato").Result = operacoes[i].Contrato;
			wrdObject.Selection.Range.Font.Bold = true;			

			CriaTabela( operacoes[i], i + 1 );

			//Salva o novo arquivo
			wrd.SaveAs2( minuta, 17, true );
			ss_kill("WINWORD");

			operacoes[i].Status = "OK";	
		}
		catch(e)
		{
			operacoes[i].Status = "NOK";
			operacoes[i].Observacao = "[Word] " + e.message;
			ss_kill("WINWORD");
		}		

		EscreveResultado( operacoes[i] );
		EscreveLog( operacoes[i] );
	}
	
	return tm_total.elapsed();
}
function CriaTabela( obj, indiceContagem )
{
	var cnpj = "";

	for( var j = 0; j < obj.Empresas.length; j++ )
	{
		AutoIt.ToolTip( "Processando: " + indiceContagem + " de " + operacoes.length + " registro(s). Contrato: " + obj.Contrato + " - Empresa " + ( j + 1 ) + " de " + obj.Empresas.length + " para este contrato.", 10, 10 );

		if( j == 0 ){
			wrd.FormFields("Contrato").Select();
			Enter( 2 );
		}		

		//Preenche os campos CNPJ e Razão Social	
		cnpj = MascarasCpfCnpj(String( obj.Empresas[j].CNPJ.substring( obj.Empresas[j].CNPJ.length - 14)));
		wrdObject.Selection.Range.Text = "CNPJ: " + cnpj;		
		wrdObject.Selection.Paragraphs(1).Range.Font.Name = "Calibri";
		wrdObject.Selection.Paragraphs(1).Range.Font.Size = 12;
		wrdObject.Selection.Paragraphs(1).Range.Font.Bold = true;
		wrdObject.Selection.Range.ParagraphFormat.Alignment = 0;		
		Enter( 1 );

		wrdObject.Selection.Range.Text = "Razão Social: " + obj.Empresas[j].Nome;
		wrdObject.Selection.Paragraphs(1).Range.Font.Name = "Calibri";
		wrdObject.Selection.Paragraphs(1).Range.Font.Size = 12;
		wrdObject.Selection.Paragraphs(1).Range.Font.Bold = true;
		wrdObject.Selection.Range.ParagraphFormat.Alignment = 0;
		Enter( 2 );
		
		//Cria o título antes da tabela
		wrdObject.Selection.Range.Text = "Lista de usuários que receberão o QR Code para leitura";
		wrdObject.Selection.Paragraphs(1).Range.Font.Name = "Calibri";
		wrdObject.Selection.Paragraphs(1).Range.Font.Size = 12;
		wrdObject.Selection.Paragraphs(1).Range.Font.Bold = true;
		wrdObject.Selection.Range.ParagraphFormat.Alignment = 1;
		Enter( 2 );

		var newTable = wrd.Tables.Add( wrdObject.Selection.Range, obj.Empresas[j].Usuarios.length + 1, 3 );
		newTable.Range.ParagraphFormat.Alignment = 1;		
		wrdObject.Selection.Range.Font.Size = 11;		

		newTable.Borders( -1 ).LineStyle = 1;
		newTable.Borders( -2 ).LineStyle = 1;
		newTable.Borders( -3 ).LineStyle = 1;
		newTable.Borders( -4 ).LineStyle = 1;
		newTable.Borders( -5 ).LineStyle = 1;
		newTable.Borders( -6 ).LineStyle = 1;

		var indiceTabela = 0;
		var cpf = "";

		for( var k = 0; k < obj.Empresas[j].Usuarios.length; k++ )
		{
			indiceTabela = k + 2;
			cpf = ( obj.Empresas[j].Usuarios[k].CPF == "NULL" || obj.Empresas[j].Usuarios[k].CPF == "" ) ? 
				"" : MascarasCpfCnpj(String(obj.Empresas[j].Usuarios[k].CPF.substring( obj.Empresas[j].Usuarios[k].CPF.length - 11 )));

			//Escreve o cabeçalho e a primeira linha
			if( k == 0 )
			{
				newTable.Cell( indiceTabela - 1, 1 ).SetWidth( 200, true );
				newTable.Cell( indiceTabela - 1, 1 ).Range.ParagraphFormat.Alignment = 1;
				newTable.Cell( indiceTabela - 1, 1 ).Range.Font.Bold = true;
				newTable.Cell( indiceTabela - 1, 1 ).Range.Text = "Nome do usuário";

				newTable.Cell( indiceTabela - 1, 2 ).SetWidth( 100, true );
				newTable.Cell( indiceTabela - 1, 2 ).Range.ParagraphFormat.Alignment = 1;
				newTable.Cell( indiceTabela - 1, 2 ).Range.Font.Bold = true;
				newTable.Cell( indiceTabela - 1, 2 ).Range.Text = "CPF";

				newTable.Cell( indiceTabela - 1, 3 ).SetWidth( 100, true );
				newTable.Cell( indiceTabela - 1, 3 ).Range.ParagraphFormat.Alignment = 1;
				newTable.Cell( indiceTabela - 1, 3 ).Range.Font.Bold = true;
				newTable.Cell( indiceTabela - 1, 3 ).Range.Text = "Assinale os usuários que receberão o QR Code";
			}

			newTable.Cell( indiceTabela, 1 ).SetWidth( 200, true );
			newTable.Cell( indiceTabela, 1 ).Range.ParagraphFormat.Alignment = 0;
			newTable.Cell( indiceTabela, 1 ).Range.Font.Bold = false;
			newTable.Cell( indiceTabela, 1 ).Range.Text = obj.Empresas[j].Usuarios[k].Nome;
			
			newTable.Cell( indiceTabela, 2 ).SetWidth( 100, true );
			newTable.Cell( indiceTabela, 2 ).Range.Font.Bold = false;
			newTable.Cell( indiceTabela, 2 ).Range.Text = cpf;

			newTable.Cell( indiceTabela, 3 ).SetWidth( 100, true );
			newTable.Cell( indiceTabela, 3 ).Range.Font.Bold = false;
			newTable.Cell( indiceTabela, 3 ).Range.Text = "[         ]";						
		}

		newTable.Cell( indiceTabela, 3 ).Select();
		wrdObject.Selection.MoveDown();

		Enter( 5 );

		wrdObject.Selection.Range.Text = "____________________________________";
		wrdObject.Selection.Paragraphs(1).Range.Font.Name = "Calibri";
		wrdObject.Selection.Paragraphs(1).Range.Font.Size = 11;
		wrdObject.Selection.Paragraphs(1).Range.Font.Bold = false;
		wrdObject.Selection.Range.ParagraphFormat.Alignment = 0;
		Enter( 1 );

		wrdObject.Selection.Range.Text = "Representante Legal:";
		wrdObject.Selection.Paragraphs(1).Range.Font.Name = "Calibri";
		wrdObject.Selection.Paragraphs(1).Range.Font.Size = 11;
		wrdObject.Selection.Paragraphs(1).Range.Font.Bold = true;
		wrdObject.Selection.Range.ParagraphFormat.Alignment = 0;
		Enter( 1 );

		wrdObject.Selection.Range.Text = "CPF:";
		wrdObject.Selection.Paragraphs(1).Range.Font.Name = "Calibri";
		wrdObject.Selection.Paragraphs(1).Range.Font.Size = 11;
		wrdObject.Selection.Paragraphs(1).Range.Font.Bold = true;
		wrdObject.Selection.Range.ParagraphFormat.Alignment = 0;	

		if( j < obj.Empresas.length - 1 ){

			Enter( 1 );

			wrdObject.Selection.ParagraphFormat.Borders( -3 ).LineStyle = 1;

			Enter( 2 );
		}	
	}

	//Alinhamento de Tabela
	for( var x = 1; x <= wrd.Tables.Count; x++ ){
		wrd.Tables(x).Rows.Alignment = 1;
		wrd.Tables(x).Rows(1).Cells.VerticalAlignment = 1;		
	}
}
function CriarFields( pField, newName, dado, qtdeEnter, bold)
{
	//Seta a posição no Word para criar o campo, usando pField como referência
	wrd.FormFields(pField).Select();

	for(var e=0; e<qtdeEnter; e++)
		wrdObject.Selection.InsertAfter("\r");

	wrd.FormFields(pField).Select();

	for(var e=0; e<qtdeEnter; e++)
		wrdObject.Selection.MoveDown();
	
	//Cria o campo
	wrd.FormFields.Add(wrdObject.Selection.Range, 70);

	//Procura e renomeia o campo
	for(var i=1; i<=wrd.FormFields.Count; i++)
	{
		if(wrd.FormFields(i).Name == "Text1" || wrd.FormFields(i).Name == "Texto1")
		{
			wrd.FormFields(i).Name = newName;
			wrd.FormFields(i).Result = ( dado != "" ) ? dado : "";
			wrd.FormFields(i).Select();			
			if( bold == null ) wrdObject.Selection.Range.Font.Bold = false;
			else wrdObject.Selection.Range.Font.Bold = true;
			break;
		}
	}
}
function CriarFieldsLado(pField, newName, dado, qtdeTab, bold)
{
	//Seta a posição no Word para criar o campo, usando pField como referência
	wrd.FormFields(pField).Select();

	for(var e=0; e<qtdeTab; e++)
		wrdObject.Selection.InsertAfter(" ");
		
	wrdObject.Selection.MoveRight();
	
	//Cria o campo
	wrd.FormFields.Add(wrdObject.Selection.Range, 70);

	//Procura e renomeia o campo
	for(var i=1; i<=wrd.FormFields.Count; i++)
	{
		if(wrd.FormFields(i).Name == "Text1" || wrd.FormFields(i).Name == "Texto1")
		{
			wrd.FormFields(i).Name = newName;
			wrd.FormFields(i).Result = ( dado != "" ) ? dado : "";
			wrd.FormFields(i).Select();			
			if( bold == null ) wrdObject.Selection.Range.Font.Bold = false;
			else wrdObject.Selection.Range.Font.Bold = true;
			break;
		}
	}
}
function CriarFieldsSolo( newName, dado, qtdeEnter, bold)
{	
	//Cria o campo
	wrd.FormFields.Add(wrdObject.Selection.Range, 70);

	//Procura e renomeia o campo
	for(var i=1; i<=wrd.FormFields.Count; i++)
	{
		if(wrd.FormFields(i).Name == "Text1" || wrd.FormFields(i).Name == "Texto1")
		{
			wrd.FormFields(i).Name = newName;
			wrd.FormFields(i).Result = ( dado != "" ) ? dado : "";
			wrd.FormFields(i).Select();			
			if( bold == null ) wrdObject.Selection.Range.Font.Bold = false;
			else wrdObject.Selection.Range.Font.Bold = true;
			break;
		}
	}
}
function ValidaCPFRepetido( obj, indice )
{
	for( var m = 0; m < obj.Usuarios.length; m++ )
	{
		if( m == indice ) continue;

		if( ( obj.Usuarios[m].CPF == obj.Usuarios[indice].CPF && obj.Usuarios[m].Nome == obj.Usuarios[indice].Nome ) && obj.Usuarios[m].Escrito ) return true;
	}

	return false;
}
function Enter( qtd ){
	for(var e = 0; e < qtd; e++)
		wrdObject.Selection.Paragraphs(1).Range.InsertAfter("\r");

	for(var e = 0; e < qtd; e++)
		wrdObject.Selection.MoveDown();
}	
