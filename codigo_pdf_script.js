function gerarPdf(idVisita = "xxxxx") {

  const nomePastaModelo = 'modelos'; //pasta drive
  const nomeModelo = 'MODELO_WORD';    //nome modelo word

  //novo// pegando id do folder//
  var idPastaModelo = 'ID DA SUA PASTA DRIVE';   //
  var pastaPrincipal = DriveApp.getFolderById(idPastaModelo);
  var pastasModelo = pastaPrincipal.getFoldersByName(nomePastaModelo);
  //novo// pegando id do folder//

  if (!pastasModelo.hasNext()) {
    return ContentService.createTextOutput(`Erro: pasta "${nomePastaModelo}" não encontrada.`);
  }

  //pasta modelo e arquivo//
  const pastaModelo = pastasModelo.next();
  const arquivosModelo = pastaModelo.getFilesByName(nomeModelo);
  //pasta modelo e arquivo//

  //arquivos//
  const debugArquivos = pastaModelo.getFiles();
  while (debugArquivos.hasNext()) {
    const file = debugArquivos.next();
  }
  //arquivos//

  if (!arquivosModelo.hasNext()) {
      return ContentService.createTextOutput(`Erro: Arquivo "${nomeModelo}" não encontrado na pasta "${nomePastaModelo}".`);
  }

  const modelo = arquivosModelo.next();

  //variavel data//
  const agora = new Date();
  const dataFormatada = Utilities.formatDate(agora, Session.getScriptTimeZone(), "dd/MM/yyyy");
  const timestamp = `${String(agora.getDate()).padStart(2, '0')}-${String(agora.getMonth() + 1).padStart(2, '0')}-${agora.getFullYear()}_${String(agora.getHours()).padStart(2, '0')}-${String(agora.getMinutes()).padStart(2, '0')}-${String(agora.getSeconds()).padStart(2, '0')}`;
  //variavel data//


  const nomeArquivo = `ARQUIVO_NOME_${timestamp}`;
  const copia = modelo.makeCopy(nomeArquivo, pastaModelo);

  //body do doc//
  const doc = DocumentApp.openById(copia.getId());
  const body = doc.getBody();
  //body do doc//


  //EXEMPLO QUEBRA DE PAGINA APOS ALGUMA TABELA (COLOCAR INDICE DELA)
  adicionarQuebraDePaginaTabela(15,body);

  //BUSCAR INFORMAÇÕES DO EXCEL//
  var info = buscarInformacoesPorId(idVisita);
  //console.log(info.vetor)

  //PREENCHER CAMPOS NO WORD
  preencherCamposDasTabelas(info.vetor,body);

  gerarGraficoComTabelaEventosPorTipoEMes(info.vetor,body);

  var images  = buscarImagens(idVisita);
  inserirTabelaImagens(body, images.vetor);

  //preecher todos os valores setados previamente//
  if (info != null) {
  for (var chave in info.vetor) {
      if (info.vetor.hasOwnProperty(chave)) {
        var valor = info.vetor[chave];
        // Monta a expressão de busca no seu template
        var campoTemplate = '\\[' + chave + '\\]';
      
        // faz o replace no body
        body.replaceText(campoTemplate, valor ? valor.toString() : '');
      }
    }
  } else {
    Logger.log('ID não encontrado');
  }

  //preecher todos os valores setados previamente//

  doc.saveAndClose();

  const nomeDoc = "Relatório_Completo_"+idVisita+"_"+timestamp;
  const file = DriveApp.getFileById(doc.getId());
  const pdf = file.getAs(MimeType.PDF);
  pdf.setName(nomeDoc + ".pdf");

  var idPastaPrincipal = 'id_pasta_destino';
  var pastaPrincipal = DriveApp.getFolderById(idPastaPrincipal);
  var pastas = pastaPrincipal.getFoldersByName('Relatorios');
  var pastaRelatorios;

  if (pastas.hasNext()) {
    pastaRelatorios = pastas.next();
  } else {
    pastaRelatorios = pastaPrincipal.createFolder('Relatorios');
  }

  var arquivoPdf = pastaRelatorios.createFile(pdf);
  Logger.log('Arquivo salvo em: ' + arquivoPdf.getUrl());
  DriveApp.getFileById(copia.getId()).setTrashed(true);

  return ContentService.createTextOutput(`Relatório gerado com sucesso: ${arquivoPdf.getUrl()}`);
}

// ============================================================================================= //

function buscarInformacoesPorId(idVisita) {

  var planilha = SpreadsheetApp.openById('ID_PLANILHA'); //
  var aba = planilha.getSheetByName('VisitasExecutadas');
  
  var dados = aba.getDataRange().getValues();
  var vetorRef = [];
  var vetorReturn = {};


  for (var i = 0; i < dados.length; i++) {
    
    if (i === 0) {
      // Pega o cabeçalho
      for (var j = 0; j < dados[i].length; j++) {
        vetorRef[j] = dados[i][j];
      }
    } else {
      var idLinha = dados[i][2]; // ajuste para a sua coluna do ID
      
      if (idLinha == idVisita) {

        for (var g = 0; g < dados[i].length; g++) {

          let valorCampoIncompleto = dados[i][g];
          
          let valorCampo = formatarDataSeNecessario(valorCampoIncompleto);
         
          if(valorCampo == null || valorCampo == ""){
            valorCampo = "-";
          }
        
          vetorReturn[vetorRef[g]] = valorCampo;
        }
        
        //console.log(vetorReturn)
        return {
          vetor: vetorReturn,
        };
      }
    }
  }

  return null;
}

// ============================================================================================= //

function buscarImagens(idVisita = "be12b189") {
  var planilha = SpreadsheetApp.openById('ID_PLANILHA');
  var aba = planilha.getSheetByName('Fotos');
  
  var dados = aba.getDataRange().getValues();
  var vetorRef = [];
  var vetorReturn = [];

  for (var i = 0; i < dados.length; i++) {
    if (i === 0) {
      for (var j = 0; j < dados[i].length; j++) {
        vetorRef[j] = dados[i][j];
      }
    }
  }

  for (var i = 1; i < dados.length; i++) { 
    var idLinha = dados[i][1]; 
    if (idLinha == idVisita) {
      var objetoLinha = {};
      for (var j = 0; j < dados[i].length; j++) {
        var valorCampo = dados[i][j];
        objetoLinha[vetorRef[j]] = valorCampo;
      }
      vetorReturn.push(objetoLinha);
    }
  }
  vetorReturn.sort(function(a, b) {
    return Number(a.Ordem) - Number(b.Ordem);
  });
  return { vetor: vetorReturn };
}

// ============================================================================================= //


function inserirTabelaImagens(body, imagens) {
  const fotosPorLinha = 3;
  let linhaImagens = [];
  let linhaTitulos = [];
  let contadorImgs = 0;

  const tablesGeral = body.getTables();
  const tabela = tablesGeral[7]; //
  const numCols = fotosPorLinha;

  let linhaAtual = 0;

  imagens.forEach((item, index) => {
    contadorImgs++;

    const caminhoCompleto = item.Foto;
    const partes = caminhoCompleto.split('/');
    const nomeArquivo = partes[partes.length - 1];

    const pastas = DriveApp.getFolderById('ID_PASTA_IMAGENS');

    const arquivos = pastas.getFilesByName(nomeArquivo);

    let imagemBlob = null;
    if (arquivos.hasNext()) {
      const arquivo = arquivos.next();
      imagemBlob = arquivo.getBlob();
    }

    if (imagemBlob) {
      linhaImagens.push(imagemBlob);
      linhaTitulos.push("Foto " + contadorImgs + " - " + (item.Titulo || ''));

      if (linhaImagens.length === fotosPorLinha || index === imagens.length - 1) {
        // pelo menos duas linhas a mais se necessário
        while (tabela.getNumRows() < linhaAtual + 2) {
          inserirLinhaComFormato(tabela);
        }

        // linha de imagens
        const rowImagens = tabela.getRow(linhaAtual);
        for (let i = 0; i < fotosPorLinha; i++) {
          const cell = rowImagens.getCell(i);
          cell.clear();
          if (linhaImagens[i]) {
            const p = cell.appendParagraph('');
            p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            p.appendInlineImage(linhaImagens[i]).setWidth(200).setHeight(100);
          }
        }

        linhaAtual++;

        // linha de títulos
        const rowTitulos = tabela.getRow(linhaAtual);
        for (let i = 0; i < fotosPorLinha; i++) {
          const cell = rowTitulos.getCell(i);
          cell.clear();
          if (linhaTitulos[i]) {
            cell.appendParagraph(linhaTitulos[i])
              .setFontSize(6)
              .setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
          }
        }

        linhaAtual++;
        linhaImagens = [];
        linhaTitulos = [];
      }
    }
  });
}

// ============================================================================================= //

function atualizarNomeArquivoPorId(idVisita,nomeDoc = "Relatório_Completo_11111_11111111",coluna = "PDF_Rel_Geral_gerado") {

  var planilha = SpreadsheetApp.openById('ID_PLANILHA'); 
  var aba = planilha.getSheetByName('ABA_EXECEL');
  var dados = aba.getDataRange().getValues();

  //variavel data//
  const agora = new Date();
  const dataFormatada = Utilities.formatDate(agora, Session.getScriptTimeZone(), "dd/MM/yyyy");
  const timestamp = `${String(agora.getDate()).padStart(2, '0')}-${String(agora.getMonth() + 1).padStart(2, '0')}-${agora.getFullYear()}_${String(agora.getHours()).padStart(2, '0')}-${String(agora.getMinutes()).padStart(2, '0')}-${String(agora.getSeconds()).padStart(2, '0')}`;
  //variavel data//

  var novoNome = "//CAMINHO//"+nomeDoc+".pdf";
  
  var header = dados[0];
  var indiceId = header.indexOf("ID");
  var indiceNome = header.indexOf(coluna);

  if (indiceId === -1 || indiceNome === -1) {
    Logger.log("Colunas 'ID' ou 'NOME' não encontradas.");
    return;
  }

  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][indiceId]) === idVisita) {
      aba.getRange(i + 1, indiceNome + 1).setValue(novoNome); // i+1 porque o range começa na linha 1
      Logger.log("Nome atualizado na linha " + (i + 1));
      return;
    }
  }

  Logger.log("ID " + idVisita + " não encontrado.");
}

// ============================================================================================= //

///////////////////////////////////////////////////
function formatarDataSeNecessario(valor) {
  // string no formato DD/MM/AAAA
    if (typeof valor === 'string' && /^\d{2}\/\d{2}\/\d{4}$/.test(valor)) {
      return valor;
    }

    //for nulo, vazio ou só espaço
    if (valor == null || String(valor).trim() === "") {
      return valor;
    }

    //for número ou string numérica simples 
    if (typeof valor === 'number' || /^\d+$/.test(valor)) {
      return valor;
    }

    // converter em data
    const data = new Date(valor);
    if (!isNaN(data.getTime())) {
      const dia = String(data.getDate()).padStart(2, '0');
      const mes = String(data.getMonth() + 1).padStart(2, '0');
      const ano = data.getFullYear();
      return `${dia}/${mes}/${ano}`;
    }

    // caso  não for data valida retorna o valor original
    return valor;
}

// ============================================================================================= //

function inserirLinhaComFormato(tabela) {
  const indexModelo = tabela.getNumRows() - 1;
  if (indexModelo < 0) {
    //se a tabela estiver vazia adiciona uma linha
    tabela.appendTableRow();
    for (let i = 0; i < 3; i++) {
      tabela.getRow(0).appendTableCell("");
    }
  } else {
    const linhaModelo = tabela.getRow(indexModelo).copy();
    tabela.insertTableRow(tabela.getNumRows(), linhaModelo);
  }
}

// ============================================================================================= //

function inserirLinhaComFormatoEspecifico(tabela) {
  if (tabela.getNumRows() === 0) {
    // se a tabela estiver vazia adiciona uma linha padrão
    const novaLinha = tabela.appendTableRow();
    for (let i = 0; i < 3; i++) {
      novaLinha.appendTableCell("");
    }
  } else {
    // copia sempre a primeira linha do modelo
    const linhaModelo = tabela.getRow(1).copy();
    tabela.insertTableRow(tabela.getNumRows(), linhaModelo);
  }
}

// ============================================================================================= //

function adicionarQuebraDePaginaTabela(tabelaIndex,body) {
  const tabelas = body.getTables();

  if (tabelas.length > tabelaIndex) {
    const tabela = tabelas[tabelaIndex];
    const posicao = body.getChildIndex(tabela);

    //quebra de página depois da tabela
    body.insertPageBreak(posicao + 1);
  }
}

// ============================================================================================= //








