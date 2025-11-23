// scripts/sidepanel.js

/**
 * POA - Professor Online Automático
 * Script do Painel Lateral
 */

document.addEventListener('DOMContentLoaded', () => {
  const textArea = document.getElementById('data-area');

  // --- Botões de Clipboard ---
  configurarBotao('btn-copy-clip', () => {
    if (!textArea.value) {
      mostrarToast("A área de texto está vazia.", "info");
      return;
    }
    copiarParaAreaTransferencia(textArea.value, textArea);
    mostrarToast("Conteúdo copiado!", "success");
  });

  configurarBotao('btn-paste-clip', async () => {
    try {
      const texto = await navigator.clipboard.readText();
      if (texto) {
        textArea.value = texto;
        mostrarToast("Conteúdo colado!", "success");
        textArea.focus();
        textArea.dispatchEvent(new Event('input')); 
      } else {
        mostrarToast("Área de transferência vazia.", "info");
      }
    } catch (err) {
      console.warn('Erro de clipboard:', err);
      mostrarToast("Erro ao colar. Use Ctrl+V.", "error");
      textArea.focus();
    }
  });

  configurarBotao('btn-clear-clip', () => {
    if (!textArea.value) return;
    textArea.value = '';
    textArea.focus();
    mostrarToast("Área de texto limpa.", "info");
  });


  // Inicializa verificação de LEDs
  atualizarStatusLeds();

  // --- Botões Principais ---

  // 1. Lançar Notas
  configurarBotao('btn-lancar-notas', async () => {
    const dados = textArea.value;
    if (!dados) {
      mostrarToast("Cole os dados das notas primeiro.", "error");
      textArea.focus(); return;
    }
    mostrarToast("Processando notas...", "info");
    const resposta = await enviarComando('preencher_notas', dados);
    processarResposta(resposta);
    if (verificarSucesso(resposta)) textArea.value = '';
  });

  // 2. Lançar Frequência
  configurarBotao('btn-lancar-frequencia', async () => {
    const dados = textArea.value;
    if (!dados) {
      mostrarToast("Cole as matrículas primeiro.", "error");
      textArea.focus(); return;
    }
    mostrarToast("Processando frequência...", "info");
    const resposta = await enviarComando('preencher_frequencia', dados);
    processarResposta(resposta);
    if (verificarSucesso(resposta)) textArea.value = '';
  });

  // 3. Verificar Pendências
  configurarBotao('btn-verificar-pendencias', async () => {
    mostrarToast("Analisando pendências...", "info");
    const resposta = await enviarComando('verificar_pendencias', null);
    if (resposta && resposta.dados) {
      copiarParaAreaTransferencia(resposta.dados, textArea);
      mostrarToast("Datas copiadas!", "success");
    } else {
      processarResposta(resposta);
    }
  });

  // 4. Botão SISEDU
  configurarBotao('btn-lancar-gabaritos-sisedu', async () => {
    const dados = textArea.value;
    if (!dados) {
      mostrarToast("Cole o gabarito primeiro.", "error");
      textArea.focus(); return;
    }
    mostrarToast("Identificando aluno...", "info");
    const resposta = await enviarComando('preencher_sisedu', dados);
    processarResposta(resposta);
  });
});


// ===================================================================
// FUNÇÕES AUXILIARES
// ===================================================================

function configurarBotao(id, callback) {
  const btn = document.getElementById(id);
  if (btn) btn.addEventListener('click', callback);
}

function copiarParaAreaTransferencia(texto, elementoInput) {
  elementoInput.value = texto;
  elementoInput.select();
  document.execCommand('copy');
}

function verificarSucesso(resposta) {
  if (!resposta) return false;
  if (Array.isArray(resposta)) return resposta.some(item => item.type === 'success');
  return resposta.type === 'success';
}

function processarResposta(resposta) {
  if (!resposta) return;
  if (Array.isArray(resposta)) {
    resposta.forEach((item, index) => {
      setTimeout(() => { if (item.message) mostrarToast(item.message, item.type); }, index * 300); 
    });
  } else if (resposta.message) {
    mostrarToast(resposta.message, resposta.type);
  }
}

function mostrarToast(mensagem, tipo = 'info') {
  const container = document.getElementById('toast-container');
  if (!container) return;
  const ICONS = { success: '✅ ', error: '⚠️ ', info: 'ℹ️ ' };
  const toast = document.createElement('div');
  toast.className = `toast ${tipo}`;
  toast.textContent = (ICONS[tipo] || '') + mensagem;
  container.appendChild(toast);
  void toast.offsetWidth; 
  toast.classList.add('show');
  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => { if (container.contains(toast)) container.removeChild(toast); }, 300);
  }, 4000);
}

async function enviarComando(acao, dados) {
  try {
    const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
    if (!tab || !tab.id) throw new Error("Nenhuma aba ativa.");
    return await chrome.tabs.sendMessage(tab.id, { action: acao, dados: dados });
  } catch (error) {
    console.warn("Erro comunicação:", error);
    mostrarToast("Não é possível usar esta função nesta página.", "error");
    return null;
  }
}

// ===================================================================
// GERENCIAMENTO DE LEDS (URL DETECTOR)
// ===================================================================

chrome.tabs.onActivated.addListener(() => atualizarStatusLeds());
chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (changeInfo.url || changeInfo.status === 'complete') atualizarStatusLeds(); 
});

async function atualizarStatusLeds() {
  try {
    const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
    if (!tab || !tab.url) return;
    
    const url = tab.url;
    
    // Mapeamento de URLs
    const statusMap = {
      'led-notas': url.includes('avaliacao_nota'),
      'led-frequencia': url.includes('frequencia_chamada'),
      'led-pendencias': url.includes('relatorio_aulas_dadas'),
      'led-sisedu': url.includes('cadastrar_gabarito') 
    };

    for (const [id, ativo] of Object.entries(statusMap)) {
      toggleLed(id, ativo);
    }
  } catch (error) {
    console.log("Erro ao atualizar LEDs:", error);
  }
}

// *** ESTA FUNÇÃO ESTAVA FALTANDO E É A CAUSA DO PROBLEMA ***
function toggleLed(id, ativo) {
  const led = document.getElementById(id);
  if (led) {
    if (ativo) led.classList.add('active');
    else led.classList.remove('active');
  }
}