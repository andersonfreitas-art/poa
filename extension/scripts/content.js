// scripts/content.js
"use strict";

/**
 * POA - Professor Online Automático
 * Script de conteúdo responsável por manipular o DOM da página da SEDUC/SISEDU.
 */

console.log("POA: Content script ativo.");

// ===================================================================
// LISTENER PRINCIPAL
// ===================================================================
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  let resultado = { type: 'error', message: 'Erro desconhecido.' };

  try {
    switch (request.action) {
      case 'preencher_notas':
        resultado = preencherNotas(request.dados);
        break;
      case 'preencher_frequencia':
        resultado = lancarFaltas(request.dados);
        break;
      case 'verificar_pendencias':
        resultado = verificarPendencias();
        break;
      case 'preencher_sisedu':
        resultado = preencherSisedu(request.dados);
        break;
      default:
        resultado = { type: 'error', message: `Ação não reconhecida: ${request.action}` };
    }
  } catch (e) {
    console.error("POA Error:", e);
    resultado = { type: 'error', message: 'Erro interno no script de automação.' };
  }
  
  sendResponse(resultado);
  return true;
});


// ===================================================================
// FUNÇÕES AUXILIARES
// ===================================================================

function dispararEventosDeMudanca(elemento) {
  ['change', 'input', 'blur'].forEach(evento => {
    elemento.dispatchEvent(new Event(evento, { bubbles: true }));
  });
}

// Remove acentos, espaços extras e coloca em maiúsculo para comparação segura
function normalizarTexto(texto) {
  if (!texto) return "";
  return texto
    .toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // Remove acentos
    .replace(/\s+/g, " ")             // Remove espaços duplos
    .trim()
    .toUpperCase();
}


// ===================================================================
// 1. MÓDULO DE NOTAS
// ===================================================================
function preencherNotas(dadosTexto) {
  if (!dadosTexto) return { type: 'error', message: 'Nenhuma informação recebida.' };

  const notasMap = new Map();
  const linhas = dadosTexto.trim().split("\n");

  linhas.forEach((linha) => {
    const colunas = linha.split("\t");
    if (colunas.length >= 2) {
      const matricula = colunas[0].trim();
      const nota = colunas[1].trim().replace(",", ".");
      if (matricula && nota) {
        notasMap.set(matricula, nota);
      }
    }
  });

  if (notasMap.size === 0) return { type: 'error', message: 'Dados inválidos.' };

  let processados = 0;
  const alunosNaPagina = document.querySelectorAll(".div-card.Aligner");

  alunosNaPagina.forEach((alunoDiv) => {
    const matriculaElement = alunoDiv.querySelector("small[data-item-subtitle]");
    const notaInput = alunoDiv.querySelector("input[data-nota]");

    if (matriculaElement && notaInput) {
      const matriculaPagina = matriculaElement.textContent.trim();
      if (notasMap.has(matriculaPagina)) {
        notaInput.value = notasMap.get(matriculaPagina);
        dispararEventosDeMudanca(notaInput);
        processados++;
        notaInput.style.backgroundColor = "#d4edda";
        notaInput.style.transition = "background 0.5s";
      }
    }
  });

  const mensagens = [];
  if (processados > 0) {
    const snack = document.getElementById("snack");
    if (snack) snack.style.display = "block";
    mensagens.push({ type: 'success', message: `${processados} notas preenchidas!` });
  } else {
    return { type: 'info', message: 'Nenhuma matrícula encontrada.' };
  }

  return mensagens;
}


// ===================================================================
// 2. MÓDULO DE FREQUÊNCIA
// ===================================================================
function lancarFaltas(dadosTexto) {
  const matriculasFaltosos = dadosTexto.split(/[\n,]+/).map(m => m.trim()).filter(m => m);

  if (matriculasFaltosos.length === 0) return { type: 'error', message: 'Área de texto vazia.' };

  let faltasLancadas = 0;
  let naoEncontradas = 0;

  matriculasFaltosos.forEach((matricula) => {
    const matriculaSegura = matricula.replace(/["\\]/g, ''); 
    const alunoPanel = document.querySelector(`.panel[data-aluno="${matriculaSegura}"]`);
    
    if (alunoPanel) {
      const faltaButton = alunoPanel.querySelector(".toggle-off");
      if (faltaButton) {
        faltaButton.click();
        faltasLancadas++;
      }
    } else {
      naoEncontradas++;
    }
  });

  const mensagens = [];
  if (faltasLancadas > 0) mensagens.push({ type: 'success', message: `${faltasLancadas} faltas lançadas!` });
  if (naoEncontradas > 0) mensagens.push({ type: 'info', message: `${naoEncontradas} não encontradas.` });
  if (mensagens.length === 0) return { type: 'info', message: 'Nenhuma ação realizada.' };

  return mensagens;
}


// ===================================================================
// 3. MÓDULO DE PENDÊNCIAS
// ===================================================================
function verificarPendencias() {
  const datasPendentes = new Set();
  const todasAsLinhas = document.querySelectorAll("#table-body tr");

  todasAsLinhas.forEach((linha) => {
    const celulas = linha.querySelectorAll("td");
    if (celulas.length < 6) return;
    const situacao = celulas[5].textContent.trim();
    if (situacao === "Prevista") {
      const dataCompleta = celulas[0].textContent.trim();
      datasPendentes.add(dataCompleta.slice(0, 5));
    }
  });

  const lista = Array.from(datasPendentes);
  if (lista.length === 0) return { type: 'success', message: 'Nenhuma pendência!', dados: null };
  
  return { type: 'info', message: `${lista.length} dias pendentes.`, dados: lista.join(", ") };
}


// ===================================================================
// 4. MÓDULO SISEDU (BUSCA POR NOME)
// ===================================================================
function preencherSisedu(dadosTexto) {
  if (!dadosTexto) return { type: 'error', message: "Cole o gabarito no painel." };

  // 1. Identificar o NOME do aluno na página
  const spans = document.querySelectorAll('.card-header span');
  let nomePagina = null;

  // Procura pelo span que contém "Nome:"
  for (const span of spans) {
    const texto = span.textContent;
    if (texto.includes('Nome:')) {
      // Ex: " Nome: ANA BIANKA SOUZA " -> "ANA BIANKA SOUZA"
      nomePagina = texto.split('Nome:')[1];
      break;
    }
  }

  if (!nomePagina) {
    return { type: 'error', message: "Não foi possível identificar o NOME do aluno na página." };
  }

  // Normaliza o nome da página (sem acentos, sem espaços extras)
  const nomePaginaNorm = normalizarTexto(nomePagina);

  // 2. Buscar as respostas nos dados colados (Formato: NOME [TAB] A,B,C...)
  const linhas = dadosTexto.trim().split('\n');
  let respostasAluno = null;

  for (const linha of linhas) {
    const partes = linha.split('\t');
    if (partes.length >= 2) {
      const nomeLista = partes[0];
      
      // Compara os nomes normalizados
      if (normalizarTexto(nomeLista) === nomePaginaNorm) {
        respostasAluno = partes[1].trim().split(',').map(r => r.trim());
        break;
      }
    }
  }

  if (!respostasAluno) {
    return { type: 'warning', message: `Aluno "${nomePaginaNorm}" não encontrado na lista colada.` };
  }

  // 3. Preencher o gabarito
  const mapaLetras = { 'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5 };
  const linhasTabela = document.querySelectorAll('table tbody tr');
  let marcados = 0;
  let questaoIndex = 0;

  linhasTabela.forEach((tr) => {
    const badge = tr.querySelector('.badge-item');
    if (badge) {
      if (questaoIndex < respostasAluno.length) {
        const letraResposta = respostasAluno[questaoIndex].toUpperCase();
        const indiceColuna = mapaLetras[letraResposta];

        if (indiceColuna) {
          const celulas = tr.querySelectorAll('td');
          if (celulas[indiceColuna]) {
            const radio = celulas[indiceColuna].querySelector('input[type="radio"]');
            if (radio && !radio.disabled) {
              radio.click();
              marcados++;
            }
          }
        }
      }
      questaoIndex++;
    }
  });

  if (marcados > 0) {
    return { type: 'success', message: `SISEDU: ${marcados} respostas marcadas para ${nomePaginaNorm}.` };
  } else {
    return { type: 'warning', message: `Nenhuma resposta válida encontrada para marcar.` };
  }
}