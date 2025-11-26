// scripts/content.js
"use strict";

/**
 * POA - Professor Online Automático
 * Script de conteúdo responsável por manipular o DOM da página da SEDUC/SISEDU.
 * Versão: 2.1 (Mensagens integradas ao Sidepanel)
 */

console.log("POA: Content script ativo.");

// ===================================================================
// LISTENER PRINCIPAL
// ===================================================================
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  let resultado = { type: 'info', message: 'Processando...' };

  try {
    switch (request.action) {
      case 'preencher_notas':
        // Agora capturamos o retorno da função
        resultado = preencherNotas(request.dados);
        break;

      case 'preencher_frequencia':
        // Capturamos o retorno da função (msg de sucesso ou aviso)
        resultado = lancarFaltas(request.dados);
        break;

      case 'verificar_pendencias':
        const pendencias = verificarPendencias();
        sendResponse({ dados: pendencias }); 
        return true; 

      case 'preencher_sisedu':
        resultado = preencherSisedu(request.dados);
        break;

      default:
        console.warn(`POA: Ação não reconhecida: ${request.action}`);
        resultado = { type: 'error', message: 'Ação desconhecida.' };
    }
  } catch (e) {
    console.error("POA Error:", e);
    resultado = { type: 'error', message: "Erro interno: " + e.message };
  }
  
  // Envia a resposta (objeto com type e message) de volta para o Sidepanel
  sendResponse(resultado);
  return true;
});


// ===================================================================
// FUNÇÕES AUXILIARES
// ===================================================================

function dispararEventosDeMudanca(elemento) {
  ['input', 'change', 'blur'].forEach(evento => {
    elemento.dispatchEvent(new Event(evento, { bubbles: true }));
  });
}

function normalizarTexto(texto) {
  if (!texto) return "";
  return texto.toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ").trim().toUpperCase();
}

function getDiaMes(dataStr) {
  if (!dataStr) return "";
  const partes = dataStr.trim().split('/');
  if (partes.length < 2) return dataStr.trim();
  const dia = partes[0].padStart(2, '0');
  const mes = partes[1].padStart(2, '0');
  return `${dia}/${mes}`;
}


// ===================================================================
// 1. MÓDULO DE NOTAS
// ===================================================================
function preencherNotas(dadosTexto) {
  if (!dadosTexto) {
    return { type: 'error', message: "Nenhum dado recebido." };
  }

  const notasMap = new Map();
  const linhas = dadosTexto.trim().split("\n");

  linhas.forEach((linha) => {
    const colunas = linha.split("\t");
    if (colunas.length >= 2) {
      const matricula = colunas[colunas.length - 2].trim();
      const nota = colunas[colunas.length - 1].trim().replace(",", ".");
      if (matricula.match(/^\d+$/)) {
        notasMap.set(matricula, nota);
      }
    }
  });

  let processados = 0;
  const inputs = document.querySelectorAll("input[name^='nota']");

  inputs.forEach((input) => {
    const match = input.name.match(/\[(\d+)\]/);
    if (match) {
      const matricula = match[1];
      if (notasMap.has(matricula)) {
        input.value = notasMap.get(matricula).replace(".", ",");
        dispararEventosDeMudanca(input);
        input.style.backgroundColor = "#e8f5e9";
        processados++;
      }
    }
  });

  if (processados > 0) {
    // Retorna objeto de sucesso para o Sidepanel
    return { type: 'success', message: `${processados} notas preenchidas!` };
  } else {
    return { type: 'info', message: "Nenhuma matrícula correspondente encontrada." };
  }
}


// ===================================================================
// 2. MÓDULO DE FREQUÊNCIA INTELIGENTE
// ===================================================================
function lancarFaltas(dadosTexto) {
  if (!dadosTexto) {
    return { type: 'error', message: "Cole os dados no painel primeiro." };
  }

  // A. Identificar a data da página atual
  const inputData = document.getElementById("data");
  if (!inputData) {
    return { type: 'error', message: "Campo de data não encontrado na página." };
  }
  
  const dataPaginaRaw = inputData.value.trim();
  const diaMesPagina = getDiaMes(dataPaginaRaw);

  console.log(`POA: Data Página: ${diaMesPagina}`);

  // B. Processar o texto colado
  const faltososDoDia = new Set();
  const linhas = dadosTexto.trim().split("\n");

  linhas.forEach(linha => {
    let partes = linha.trim().split(/\t+/); 
    if (partes.length < 2) partes = linha.trim().split(";");
    if (partes.length < 2) partes = linha.trim().split(" ");

    if (partes.length >= 2) {
      const dataLinhaRaw = partes[0].trim();
      const matricula = partes[partes.length - 1].trim(); 

      if (getDiaMes(dataLinhaRaw) === diaMesPagina) {
        faltososDoDia.add(matricula);
      }
    }
  });

  // C. Aplicar na Tela
  let alteracoes = 0;
  const paineisAlunos = document.querySelectorAll("div[data-aluno]");

  paineisAlunos.forEach(painel => {
    const matriculaAluno = painel.getAttribute("data-aluno");
    const checkbox = painel.querySelector("input[type='checkbox']");
    
    if (checkbox) {
      const toggleVisual = painel.querySelector(".toggle");
      const isMarcadoPresente = checkbox.checked; // true = Presença
      const deveFaltar = faltososDoDia.has(matriculaAluno);

      const clicar = () => {
        if (toggleVisual) toggleVisual.click();
        else checkbox.click();
      };

      if (deveFaltar && isMarcadoPresente) {
        clicar(); // Marca FALTA
        alteracoes++;
      } 
      else if (!deveFaltar && !isMarcadoPresente) {
        clicar(); // Corrige para PRESENÇA
      }
    }
  });

  if (alteracoes > 0) {
    return { 
      type: 'success', 
      message: `${alteracoes} faltas lançadas para o dia ${diaMesPagina}.` 
    };
  } else if (faltososDoDia.size > 0) {
    return { 
      type: 'info', 
      message: `Dia ${diaMesPagina}: Alunos já estavam com falta.` 
    };
  } else {
    return { 
      type: 'info', 
      message: `Nenhuma falta encontrada para ${diaMesPagina} na lista.` 
    };
  }
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

  const lista = Array.from(datasPendentes).join(", ");
  return lista || "Nenhuma pendência encontrada.";
}


// ===================================================================
// 4. MÓDULO SISEDU
// ===================================================================
function preencherSisedu(dadosTexto) {
  if (!dadosTexto) return { type: 'error', message: "Cole o gabarito no painel." };

  const spans = document.querySelectorAll('.card-header span');
  let nomePagina = null;

  for (const span of spans) {
    if (span.textContent.includes('Nome:')) {
      nomePagina = span.textContent.split('Nome:')[1];
      break;
    }
  }

  if (!nomePagina) return { type: 'error', message: "Nome do aluno não identificado." };

  const nomePaginaNorm = normalizarTexto(nomePagina);
  const linhas = dadosTexto.trim().split('\n');
  let respostasAluno = null;

  for (const linha of linhas) {
    const partes = linha.split('\t');
    if (partes.length >= 2) {
      if (normalizarTexto(partes[0]) === nomePaginaNorm) {
        respostasAluno = partes[1].trim().split(',').map(r => r.trim());
        break;
      }
    }
  }

  if (!respostasAluno) return { type: 'warning', message: `Aluno "${nomePaginaNorm}" não encontrado na lista.` };

  const mapaLetras = { 'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5 };
  const linhasTabela = document.querySelectorAll('table tbody tr');
  let marcados = 0;
  let questaoIndex = 0;

  linhasTabela.forEach((tr) => {
    if (tr.querySelector('.badge-item') || tr.cells.length > 5) {
      if (questaoIndex < respostasAluno.length) {
        const indiceColuna = mapaLetras[respostasAluno[questaoIndex].toUpperCase()];
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

  if (marcados > 0) return { type: 'success', message: `${marcados} questões marcadas.` };
  else return { type: 'warning', message: `Falha ao marcar questões.` };
}