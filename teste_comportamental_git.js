
// Função para determinar o valor dominante entre dois valores com base em sua frequência
function valorDominante(valor1, valor2, frequenciaValores) {
  // Verificar qual valor tem a maior frequência
  if (frequenciaValores[valor1] > frequenciaValores[valor2]) {
    return valor1;
  } else if (frequenciaValores[valor1] < frequenciaValores[valor2]) {
    return valor2;
  } else {
    
  }
}

// Função para determinar os valores dominantes na ordem específica -  (Valor A ou Valor B, Valor C ou Valor D, Valor E ou Valor F e Valor G ou Valor H)
function valoresDominantes(frequenciaValores) {
  // Array para armazenar os valores dominantes na ordem específica
  var valoresDominantesArray = [];

  // Determinar o valor dominante entre E e I e adicionar ao array
  valoresDominantesArray.push(valorDominante("Valor A", "Valor B", frequenciaValores));

  // Determinar o valor dominante entre S e N e adicionar ao array
  valoresDominantesArray.push(valorDominante("Valor C", "Valor D", frequenciaValores));

  // Determinar o valor dominante entre T e F e adicionar ao array
  valoresDominantesArray.push(valorDominante("Valor E", "Valor F", frequenciaValores));

  // Determinar o valor dominante entre J e P e adicionar ao array
  valoresDominantesArray.push(valorDominante("Valor G", "Valor H", frequenciaValores));

  // Retornar os valores dominantes no formato de array na ordem específica
  return valoresDominantesArray;
}

function enviarEmail() {
// Estrutura de dados que mapeia cada pergunta para as opções de resposta e seus valores correspondentes
  var perguntasRespostas = {
  //exemplo:
  "1. Pergunta 1": {
    "Resposta a": "Valor C",
    "Resposta b": "Valor D"
  },
  
};

 // Inicializar objeto para armazenar contagens
  var contagemValores = {
    "Valor A": 0,
    "Valor B": 0,
    "Valor C": 0,
    "Valor D": 0,
    "Valor E": 0,
    "Valor F": 0,
    "Valor G": 0,
    "Valor H": 0
  };

  // Abrir a planilha - linka com planilha de respostas do google forms 
  var planilha = SpreadsheetApp.openByUrl('URL DA PLANILHA');
  
  // Acessar a aba (sheet) onde estão os dados do formulário
  var aba = planilha.getSheetByName('Respostas ao formulário'); // Substitua 'NOME_DA_ABA' pelo nome correto da sua aba
  
  // Obter a última linha preenchida na planilha
  var ultimaLinha = aba.getLastRow();

  // Obtém a resposta na última linha da planilha para uma determinada pergunta
  function obterResposta(aba, pergunta) {
    // Encontra a coluna correspondente à pergunta
    var colunaPergunta = 0;
    for (var i = inicio_das_respostas; i <= numero_de_colunas_de_respostas; i++) {
      if (aba.getRange(1, i).getValue() === pergunta) {
        colunaPergunta = i;
        break;
      }
    }

    // Verifica se a coluna da pergunta foi encontrada
    if (colunaPergunta > 0) {
      // Obtém a última linha da planilha
      var ultimaLinha = aba.getLastRow();
      // Obtém a resposta na última linha da coluna correspondente à pergunta
      var resposta = aba.getRange(ultimaLinha, colunaPergunta).getValue(); // Corrigido
      return resposta;
    } else {
      // Retorna null se a pergunta não foi encontrada
      return null;
    }
  }

// Itera sobre as perguntas e respostas
for (var pergunta in perguntasRespostas) {
  // Obtém a resposta na última linha da coluna correspondente à pergunta
  var resposta = obterResposta(aba, pergunta); // Corrigido
  // Compara a resposta com o mapeamento e incrementa a contagem
  if (resposta !== null && resposta in perguntasRespostas[pergunta]) { // Verifica se a resposta não é nula
    var valorResposta = perguntasRespostas[pergunta][resposta];
    contagemValores[valorResposta]++;
  }
}

// Exibe a contagem dos valores no console
for (var valor in contagemValores) {
  //Logger.log("Valor " + valor + ": " + contagemValores[valor]);
}


// Determinar os valores dominantes na ordem específica
var valoresDominantesResult = valoresDominantes(contagemValores);

//Logger.log(valoresDominantesResult);

var resultado = valoresDominantesResult.join(""); // União dos elementos do array separados por vírgula e espaço


// Obter a célula na coluna 48 e linha especificada
var celula = aba.getRange(ultimaLinha, 48);

// Atribuir o valor à célula
celula.setValue(resultado);

//LÓGICA DE ENVIAR E-MAIL
// Obter os dados da última linha preenchida
var dados = aba.getRange(ultimaLinha, 1, 1, aba.getLastColumn()).getValues()[0];

// Extrair os campos necessários das respostas
var email = dados[1]; // Supondo que o e-mail esteja na primeira coluna
var nome = dados[2]; // Supondo que o nome esteja na segunda coluna

//case para envio de e-mail -  exemplo
switch (resultado) {
  case "ACEG":
      assunto = "Seu Resultado é ACEG";
      corpoEmail = 'corpo pode ser um html com os dados necessários';
      break;
  
  default:
      console.log("Tipo de resultado não reconhecido.");
}

// Enviar o e-mail
MailApp.sendEmail({
  to: email, 
  subject: assunto, 
  htmlBody: corpoEmail
  });
}