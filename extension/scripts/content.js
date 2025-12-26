// scripts/content.js
"use strict";

/**
 * POA - Professor Online Automático
 * Script de conteúdo responsável por manipular o DOM da página da SEDUC/SISEDU.
 * Versão: 1.3 (Final - Com retorno automático)
 */

console.log("POA v1.3: Content script ativo.");

// ===================================================================
// LISTENER PRINCIPAL
// ===================================================================
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  let resultado = { type: 'info', message: 'Processando...' };

  try {
    switch (request.action) {
      case 'preencher_notas':
        resultado = preencherNotas(request.dados);
        break;

      case 'preencher_frequencia':
        resultado = lancarFaltas(request.dados);
        break;

      case 'verificar_pendencias':
        const pendencias = verificarPendencias();
        sendResponse({ dados: pendencias });
        return true;

      case 'preencher_aulas':
        // Como é uma função assíncrona longa, precisamos tratar diferente
        processarFilaAulas(request.dados)
          .then(res => sendResponse(res));
        return true; // Mantém o canal aberto para resposta assíncrona

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

  sendResponse(resultado);
  return true;
});


// ===================================================================
// FUNÇÕES AUXILIARES (ATUALIZADA)
// ===================================================================

/**
 * Simula a interação humana completa para garantir que o sistema salve.
 */
function dispararEventosDeMudanca(elemento) {
  // 1. Simula focar no campo
  elemento.dispatchEvent(new Event('focus', { bubbles: true }));

  // 2. Simula a digitação (Input)
  elemento.dispatchEvent(new Event('input', { bubbles: true }));

  // 3. Simula a alteração (Change) - Critico para jQuery
  elemento.dispatchEvent(new Event('change', { bubbles: true }));

  // 4. Simula sair do campo (Blur) - Geralmente onde o salvamento ocorre
  elemento.dispatchEvent(new Event('blur', { bubbles: true }));
}

function normalizarTexto(texto) {
  if (!texto) return "";
  return texto
    .toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
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

  console.group("POA Debug: Processamento de Notas");

  // 1. Mapear Matrículas -> Notas
  linhas.forEach((linha) => {
    const colunas = linha.split("\t").map(c => c.trim()).filter(c => c !== "");

    if (colunas.length >= 2) {
      // Matrícula = Penúltima, Nota = Última
      const notaRaw = colunas[colunas.length - 1];
      const matriculaRaw = colunas[colunas.length - 2];

      const matricula = matriculaRaw.replace(/[^0-9]/g, "");
      // Padroniza a nota internamente com PONTO para cálculos/armazenamento
      const nota = notaRaw.replace(",", ".");

      if (matricula && matricula.length > 0) {
        notasMap.set(matricula, nota);
      }
    }
  });

  console.log(`Total mapeado do Excel: ${notasMap.size}`);

  // 2. Aplicar na Página usando os atributos data-*
  let processados = 0;

  // Seleciona todos os "cartões" de alunos
  const cartoesAlunos = document.querySelectorAll("[data-template-item]");

  console.log(`Alunos encontrados na página: ${cartoesAlunos.length}`);

  cartoesAlunos.forEach((cartao) => {
    // Busca o elemento que contém a matrícula
    const elementoMatricula = cartao.querySelector("[data-item-subtitle]");

    if (elementoMatricula) {
      const matriculaPagina = elementoMatricula.innerText.replace(/[^0-9]/g, "");

      if (notasMap.has(matriculaPagina)) {

        // Busca o input de nota
        const inputNota = cartao.querySelector("input[data-nota]");

        if (inputNota && !inputNota.disabled) {
          let valorNota = notasMap.get(matriculaPagina);

          if (inputNota.type === "number") {
            valorNota = valorNota.replace(",", ".");
          } else {
            valorNota = valorNota.replace(".", ",");
          }

          inputNota.value = valorNota;
          dispararEventosDeMudanca(inputNota);

          // Feedback Visual
          inputNota.style.backgroundColor = "#e8f5e9";
          inputNota.style.borderColor = "#007f4e";
          inputNota.style.transition = "all 0.3s";

          processados++;
        }
      }
    }
  });

  console.log(`Total preenchido: ${processados}`);
  console.groupEnd();

  if (processados > 0) {
    return { type: 'success', message: `${processados} notas preenchidas com sucesso!` };
  } else {
    if (notasMap.size === 0) return { type: 'error', message: "Erro na leitura dos dados. Verifique a cópia." };
    if (cartoesAlunos.length === 0) return { type: 'error', message: "Lista de alunos não encontrada na página." };
    return { type: 'warning', message: "Nenhuma matrícula coincidiu com a página." };
  }
}


// ===================================================================
// 2. MÓDULO DE FREQUÊNCIA
// ===================================================================
function lancarFaltas(dadosTexto) {
  if (!dadosTexto) return { type: 'error', message: "Cole os dados no painel primeiro." };

  const inputData = document.getElementById("data");
  if (!inputData) return { type: 'error', message: "Campo de data não encontrado na página." };

  const diaMesPagina = getDiaMes(inputData.value.trim());
  console.log(`POA: Data Página: ${diaMesPagina}`);

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

  let alteracoes = 0;
  const paineisAlunos = document.querySelectorAll("div[data-aluno]");

  paineisAlunos.forEach(painel => {
    const matriculaAluno = painel.getAttribute("data-aluno");
    const checkbox = painel.querySelector("input[type='checkbox']");

    if (checkbox) {
      const toggleVisual = painel.querySelector(".toggle");
      const isMarcadoPresente = checkbox.checked;
      const deveFaltar = faltososDoDia.has(matriculaAluno);

      const clicar = () => {
        if (toggleVisual) toggleVisual.click();
        else checkbox.click();
      };

      if (deveFaltar && isMarcadoPresente) {
        clicar();
        alteracoes++;
      }
      else if (!deveFaltar && !isMarcadoPresente) {
        clicar();
      }
    }
  });

  if (alteracoes > 0) {
    return { type: 'success', message: `${alteracoes} faltas lançadas para o dia ${diaMesPagina}.` };
  } else if (faltososDoDia.size > 0) {
    return { type: 'info', message: `Dia ${diaMesPagina}: Alunos já estavam com falta.` };
  } else {
    return { type: 'info', message: `Nenhuma falta encontrada para ${diaMesPagina} na lista.` };
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
// 4. MÓDULO SISEDU (v2.5)
// ===================================================================
function preencherSisedu(dadosTexto) {
  if (!dadosTexto) return { type: 'error', message: "Cole o gabarito no painel." };

  const headers = document.querySelectorAll('.card-header span');
  let nomePagina = null;

  for (const el of headers) {
    const texto = el.textContent || "";
    if (texto.includes('Nome:')) {
      let partes = texto.split('Nome:');
      if (partes.length > 1) {
        let nomeBruto = partes[1].trim();
        if (nomeBruto.includes('-')) nomeBruto = nomeBruto.split('-')[0];
        nomePagina = nomeBruto.trim();
        break;
      }
    }
  }

  if (!nomePagina) return { type: 'error', message: "Nome do aluno não identificado." };

  const nomePaginaNorm = normalizarTexto(nomePagina);
  console.log(`POA SISEDU: Aluno na página: "${nomePaginaNorm}"`);

  const linhas = dadosTexto.trim().split('\n');
  let respostasAluno = null;

  for (const linha of linhas) {
    const partes = linha.split('\t');
    if (partes.length >= 2) {
      const nomeLista = partes[0];
      if (normalizarTexto(nomeLista) === nomePaginaNorm) {
        respostasAluno = partes[1].trim().split(',').map(r => r.trim());
        break;
      }
    }
  }

  if (!respostasAluno) {
    return { type: 'warning', message: `Aluno "${nomePagina}" não encontrado na lista.` };
  }

  const mapaLetras = { 'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5 };
  const linhasTabela = document.querySelectorAll('table tbody tr');
  let marcados = 0;
  let questaoIndex = 0;

  linhasTabela.forEach((tr) => {
    const possuiRadios = tr.querySelector('input[type="radio"]');

    if (possuiRadios) {
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
    return { type: 'success', message: `SISEDU: ${marcados} respostas marcadas para ${nomePagina}.` };
  } else {
    return { type: 'warning', message: `Falha ao marcar. Verifique se o gabarito corresponde.` };
  }
}

// ===================================================================
// 5. MÓDULO DE REGISTRO DE AULAS (LÓGICA SIMPLIFICADA + VOLTAR)
// ===================================================================

async function processarFilaAulas(dadosTexto) {
  const linhas = dadosTexto.trim().split("\n");
  let sucesso = 0;
  let erro = 0;

  for (const linha of linhas) {
    if (!linha.trim()) continue;

    // Divide dados
    let cols = linha.split("\t");
    if (cols.length < 3) cols = linha.split(";");
    
    if (cols.length >= 3) {
      const data = cols[0].trim();
      const conteudo = cols[1].trim();
      const subconteudo = cols.slice(2).join(" ").trim();

      try {
        console.log(`Processando: ${data} | ${conteudo}`);
        await passoAPassoAula(data, conteudo, subconteudo);
        sucesso++;
      } catch (err) {
        console.error("Erro na aula:", err);
        erro++;
      }
    }
  }

  // --- NOVO: Lógica de retorno automático ---
  if (sucesso > 0) {
    console.log("Todas as aulas foram lançadas. Voltando para a listagem...");
    const btnVoltar = document.getElementById('btn-cancel'); // ID identificado no HTML: id="btn-cancel"
    if (btnVoltar) {
      btnVoltar.click();
    }
  }

  return { 
    type: sucesso > 0 ? 'success' : 'error', 
    message: `Fim! ${sucesso} registrados. ${erro} erros.` 
  };
}

// Função que executa EXATAMENTE os passos solicitados
function passoAPassoAula(data, conteudo, subconteudo) {
  return new Promise(async (resolve, reject) => {
    
    // Elementos principais
    const elData = document.getElementById('data');
    const elRadioAvulso = document.querySelector('input[type="radio"][value="avulso"]');
    const elConteudo = document.getElementById('conteudo');
    const elSub = document.getElementById('subconteudo');
    const btnSave = document.getElementById('btn-save');

    if (!elData || !elRadioAvulso || !btnSave) {
      reject("Elementos do formulário não encontrados.");
      return;
    }

    // 0. Limpa alertas visuais anteriores
    document.querySelectorAll('.alert-success').forEach(a => a.style.display = 'none');

    // PASSO 1: Escreve a Data Correta
    elData.value = '';
    elData.value = data;
    dispararEventosDeMudanca(elData);

    // PASSO 2: Clica na opção Avulso
    elRadioAvulso.click();

    // PEQUENA PAUSA TÉCNICA (Seus 1000ms solicitados)
    await new Promise(r => setTimeout(r, 1000));

    // Validação de segurança
    if (elConteudo.offsetParent === null) {
      reject("Campos de conteúdo não apareceram após clicar em Avulso.");
      return;
    }

    // PASSO 3: Apaga e Escreve Conteúdo
    elConteudo.value = '';
    dispararEventosDeMudanca(elConteudo);
    elConteudo.value = conteudo;
    dispararEventosDeMudanca(elConteudo);

    // PASSO 4: Apaga e Escreve Subconteúdo
    elSub.value = '';
    dispararEventosDeMudanca(elSub);
    elSub.value = subconteudo;
    dispararEventosDeMudanca(elSub);

    // PASSO 5: Clica no botão adicionar
    btnSave.click();

    // PASSO 6: Aguarda mensagem de confirmação (Polling)
    let tentativas = 0;
    const checkInterval = setInterval(() => {
      tentativas++;
      
      const sucesso = Array.from(document.querySelectorAll('.alert-success'))
                           .some(el => el.style.display !== 'none' && el.offsetParent !== null);
      
      const loading = document.getElementById('pleaseWaitDialog');
      const isLoading = loading && loading.style.display !== 'none';

      if (sucesso && !isLoading) {
        clearInterval(checkInterval);
        
        // PASSO 7: Espera 1000ms (Sua pausa final)
        setTimeout(() => {
          resolve(true); 
        }, 1000);
      } 
      
      else if (tentativas > 60) {
        clearInterval(checkInterval);
        reject("Timeout esperando confirmação.");
      }

      const erroSite = document.querySelector('.alert-danger');
      if (erroSite && erroSite.style.display !== 'none' && erroSite.innerText.length > 5) {
        clearInterval(checkInterval);
        reject("Site recusou: " + erroSite.innerText);
      }

    }, 1000); // Seus 1000ms de intervalo
  });
}