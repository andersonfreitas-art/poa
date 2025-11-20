// scripts/content.js
"use strict";

/**
 * POA - Professor Online Automático
 * Script de conteúdo responsável por manipular o DOM da página da SEDUC.
 * Recebe comandos do Side Panel, executa ações locais e retorna o status.
 */

console.log("POA: Content script ativo.");

// --- Listener Principal ---
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

      default:
        resultado = { type: 'error', message: `Ação não reconhecida: ${request.action}` };
    }
  } catch (e) {
    console.error("POA Error:", e);
    resultado = { type: 'error', message: 'Erro interno no script de automação.' };
  }

  // Retorna o resultado (que pode ser um Objeto ou um Array de Objetos)
  sendResponse(resultado);
});


// ===================================================================
// FUNÇÕES AUXILIARES
// ===================================================================

function dispararEventosDeMudanca(elemento) {
  ['change', 'input', 'blur'].forEach(evento => {
    elemento.dispatchEvent(new Event(evento, { bubbles: true }));
  });
}


// ===================================================================
// LÓGICA DE AUTOMAÇÃO
// ===================================================================

function preencherNotas(dadosTexto) {
  if (!dadosTexto) return { type: 'error', message: 'Nenhuma informação recebida.' };

  const notasMap = new Map();
  const linhas = dadosTexto.trim().split("\n");

  // 1. Parse
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

  if (notasMap.size === 0) {
    return { type: 'error', message: 'Dados inválidos. Verifique a cópia da planilha.' };
  }

  // 2. Aplicação
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

  // 3. Preparação da Resposta (Múltiplas Mensagens)
  const totalInputados = notasMap.size;
  const naoEncontradas = totalInputados - processados;
  
  const mensagens = [];

  if (processados > 0) {
    // Ativa o snackbar nativo da página como backup
    const snack = document.getElementById("snack");
    if (snack) snack.style.display = "block";

    mensagens.push({ 
      type: 'success', 
      message: `${processados} notas preenchidas com sucesso!` 
    });
  }

  if (naoEncontradas > 0) {
    mensagens.push({ 
      type: 'info', 
      message: `${naoEncontradas} matrículas não foram encontradas.` 
    });
  }

  if (processados === 0 && naoEncontradas === 0) {
     return { type: 'info', message: 'Nenhuma matrícula correspondente encontrada essa turma.' };
  }

  // Retorna o array de mensagens
  return mensagens;
}


function lancarFaltas(dadosTexto) {
  const matriculasFaltosos = dadosTexto
    .split(/[\n,]+/)
    .map((m) => m.trim())
    .filter((m) => m);

  if (matriculasFaltosos.length === 0) {
    return { type: 'error', message: 'A área de texto está vazia.' };
  }

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

  if (faltasLancadas > 0) {
    mensagens.push({ 
      type: 'success', 
      message: `${faltasLancadas} faltas lançadas com sucesso!` 
    });
  }

  if (naoEncontradas > 0) {
    mensagens.push({ 
      type: 'info', 
      message: `${naoEncontradas} matrículas não encontradas nesta turma.` 
    });
  }

  if (mensagens.length === 0) {
    return { type: 'info', message: 'Nenhuma ação realizada.' };
  }

  return mensagens;
}


function verificarPendencias() {
  const datasPendentes = new Set();
  const todasAsLinhas = document.querySelectorAll("#table-body tr");

  todasAsLinhas.forEach((linha) => {
    const celulas = linha.querySelectorAll("td");
    if (celulas.length < 6) return;
    
    const situacao = celulas[5].textContent.trim();
    if (situacao === "Prevista") {
      const dataCompleta = celulas[0].textContent.trim();
      const dataSemAno = dataCompleta.slice(0, 5);
      datasPendentes.add(dataSemAno);
    }
  });

  const lista = Array.from(datasPendentes);
  
  if (lista.length === 0) {
    return { type: 'success', message: 'Nenhuma pendência encontrada. Frequências em dia!', dados: null };
  }
  
  return { 
    type: 'info', 
    message: `${lista.length} dias com frequências pendentes encontrados.`, 
    dados: lista.join(", ") 
  };
}