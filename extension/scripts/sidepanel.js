// scripts/sidepanel.js

/**
 * POA - Professor Online Automático
 * Script do Painel Lateral: Gerencia a interação do usuário, feedback visual (Toasts/LEDs)
 * e comunicação com o Content Script injetado na página ativa.
 */

document.addEventListener('DOMContentLoaded', () => {
  const textArea = document.getElementById('data-area');

  // --- Botões de Clipboard e Ferramentas ---
  
  // Botão Copiar
  configurarBotao('btn-copy-clip', () => {
    if (!textArea.value) {
      mostrarToast("A área de texto está vazia.", "info");
      return;
    }
    copiarParaAreaTransferencia(textArea.value, textArea);
    mostrarToast("Conteúdo copiado!", "success");
  });

  // Botão Colar
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

  // Botão Limpar (Manual)
  configurarBotao('btn-clear-clip', () => {
    if (!textArea.value) {
      return; // Nada a limpar
    }
    textArea.value = '';
    textArea.focus();
    mostrarToast("Área de texto limpa.", "info");
  });


  // Inicializa verificação de LEDs ao abrir o painel
  atualizarStatusLeds();

  // --- Listeners de Botões Principais ---

  // 1. Lançar Notas
  configurarBotao('btn-lancar-notas', async () => {
    const dados = textArea.value;
    if (!dados) {
      mostrarToast("Primeiramente, cole as informações na área de texto acima.", "error");
      textArea.focus();
      return;
    }
    mostrarToast("Processando notas...", "info");
    
    const resposta = await enviarComando('preencher_notas', dados);
    processarResposta(resposta);
    
    // Auto-limpeza se houver sucesso
    if (verificarSucesso(resposta)) {
      textArea.value = '';
    }
  });

  // 2. Lançar Frequência
  configurarBotao('btn-lancar-frequencia', async () => {
    const dados = textArea.value;
    if (!dados) {
      mostrarToast("Primeiramente, cole as informações na área de texto acima.", "error");
      textArea.focus();
      return;
    }
    mostrarToast("Processando frequência...", "info");
    
    const resposta = await enviarComando('preencher_frequencia', dados);
    processarResposta(resposta);

    // Auto-limpeza se houver sucesso
    if (verificarSucesso(resposta)) {
      textArea.value = '';
    }
  });

  // 3. Verificar Pendências
  configurarBotao('btn-verificar-pendencias', async () => {
    mostrarToast("Analisando frequências pendentes...", "info");
    const resposta = await enviarComando('verificar_pendencias', null);
    
    if (resposta && resposta.dados) {
      copiarParaAreaTransferencia(resposta.dados, textArea);
      mostrarToast("Datas copiadas para a área de transferência.", "success");
    } else {
      processarResposta(resposta);
    }
  });

  // 4. Botão SISEDU (Placeholder)
  configurarBotao('btn-lancar-nota-sisedu', () => {
    mostrarToast("Essa funcionalidade será disponibilizada em breve!", "info");
  });
});


// ===================================================================
// FUNÇÕES DE INTERFACE E UTILITÁRIOS
// ===================================================================

function configurarBotao(id, callback) {
  const btn = document.getElementById(id);
  if (btn) {
    btn.addEventListener('click', callback);
  }
}

function copiarParaAreaTransferencia(texto, elementoInput) {
  elementoInput.value = texto;
  elementoInput.select();
  document.execCommand('copy');
}

/**
 * Verifica se a resposta contém alguma indicação de sucesso
 * para autorizar a limpeza automática da área de texto.
 */
function verificarSucesso(resposta) {
  if (!resposta) return false;
  
  // Se for array (múltiplas mensagens), procura por pelo menos um sucesso
  if (Array.isArray(resposta)) {
    return resposta.some(item => item.type === 'success');
  }
  
  // Se for objeto único
  return resposta.type === 'success';
}

function processarResposta(resposta) {
  if (!resposta) return;

  // Se for um Array (várias mensagens)
  if (Array.isArray(resposta)) {
    resposta.forEach((item, index) => {
      setTimeout(() => {
        if (item.message) mostrarToast(item.message, item.type);
      }, index * 300); 
    });
  } 
  // Se for um Objeto único
  else if (resposta.message) {
    mostrarToast(resposta.message, resposta.type);
  }
}

function mostrarToast(mensagem, tipo = 'info') {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const ICONS = {
    success: '✅ ',
    error: '⚠️ ',
    info: 'ℹ️ '
  };

  const toast = document.createElement('div');
  toast.className = `toast ${tipo}`;
  toast.textContent = (ICONS[tipo] || '') + mensagem;
  
  container.appendChild(toast);
  
  void toast.offsetWidth; 
  toast.classList.add('show');

  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => {
      if (container.contains(toast)) container.removeChild(toast);
    }, 300);
  }, 4000);
}


// ===================================================================
// COMUNICAÇÃO E MONITORAMENTO
// ===================================================================

async function enviarComando(acao, dados) {
  try {
    const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
    
    if (!tab || !tab.id) {
      throw new Error("Nenhuma aba ativa encontrada.");
    }

    const response = await chrome.tabs.sendMessage(tab.id, { action: acao, dados: dados });
    return response;

  } catch (error) {
    console.warn("POA Communication Error:", error);
    mostrarToast("A página em que você está não permite usar esta função.", "error");
    return null;
  }
}

chrome.tabs.onActivated.addListener(() => atualizarStatusLeds());
chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (changeInfo.url || changeInfo.status === 'complete') { 
    atualizarStatusLeds(); 
  }
});

async function atualizarStatusLeds() {
  try {
    const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
    if (!tab || !tab.url) return;
    
    const url = tab.url;
    const statusMap = {
      'led-notas': url.includes('avaliacao_nota'),
      'led-frequencia': url.includes('frequencia_chamada'),
      'led-pendencias': url.includes('relatorio_aulas_dadas')
    };

    for (const [id, ativo] of Object.entries(statusMap)) {
      toggleLed(id, ativo);
    }
  } catch (error) {}
}

function toggleLed(id, ativo) {
  const led = document.getElementById(id);
  if (led) {
    if (ativo) led.classList.add('active');
    else led.classList.remove('active');
  }
}