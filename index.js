const { GoogleSpreadsheet } = require('google-spreadsheet');

async function calcularResultados() {
  // Carregue o ID e as credenciais da sua planilha do Google Sheets
  const doc = new GoogleSpreadsheet('SEU_ID_DO_GOOGLE_SHEET');
  await doc.useServiceAccountAuth({
    client_email: 'SEU_EMAIL_DO_CLIENTE',
    private_key: 'SUA_CHAVE_PRIVADA',
  });

  // Carregue a planilha
  await doc.loadInfo();

  // Selecione a primeira planilha (índice 0)
  const planilha = doc.sheetsByIndex[0];

  // Obtenha todas as linhas da planilha
  const linhas = await planilha.getRows();

  // Número total de aulas no semestre
  const totalAulas = 60;

  for (const linha of linhas) {
    // Calcule a média das três provas (P1, P2, P3)
    const media = Math.round((parseFloat(linha.P1) + parseFloat(linha.P2) + parseFloat(linha.P3)) / 3);

    // Calcule a porcentagem de faltas
    const porcentagemFaltas = (parseFloat(linha.Faltas) / totalAulas) * 100;

    // Determine a situação do aluno
    let situacao = '';
    if (porcentagemFaltas > 25) {
      situacao = 'Reprovado por Falta';
    } else if (media < 5) {
      situacao = 'Reprovado por Nota';
    } else if (media >= 5 && media < 7) {
      situacao = 'Exame Final';

      // Calcule a Nota para Aprovação Final (naf)
      const naf = Math.round(2 * (5 - media));

      // Atualize a Nota para Aprovação Final na planilha
      linha['Nota para Aprovação Final'] = naf;
    } else {
      situacao = 'Aprovado';
      // Se o aluno for aprovado, defina a Nota para Aprovação Final como 0
      linha['Nota para Aprovação Final'] = 0;
    }

    // Atualize a coluna Situação na planilha
    linha.Situação = situacao;

    // Salve as alterações na planilha
    await linha.save();
  }

  console.log('Resultados atualizados na planilha.');
}

// Execute a função
calcularResultados();
