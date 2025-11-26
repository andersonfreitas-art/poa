// scripts/sidepanel.js

/**
 * POA - Professor Online Automático
 * Script do Painel Lateral
 * Versão: 2.2 (Com Validação Rigorosa de URL/Contexto)
 */

// Mapeamento rigoroso: Qual trecho de URL é necessário para cada função?
const REGRAS_CONTEXTO = {
  'notas': 'avaliacao_nota',         // Lançar Notas
  'frequencia': 'frequencia_chamada', // Lançar Frequência
  'pendencias': 'relatorio_aulas_dadas', // Pendências
  'sisedu': 'cadastrar_gabarito'     // SISEDU
};

document.addEventListener('DOMContentLoaded', () => {
  const textArea = document.getElementById('data-area');

  // --- Botões de Clipboard (Mantidos) ---
  configurarBotao('btn-copy-clip', () => {
    if (!textArea.value) { mostrarToast("A área de texto está vazia.", "info"); return; }
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
      } else { mostrarToast("Área de transferência vazia.", "info"); }
    } catch (err) {
      console.warn('Erro clipboard:', err);
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


  // Inicializa verificação de LEDs e Listener de mudanças de aba
  atualizarStatusLeds();
  // Monitora mudanças para atualizar os LEDs em tempo real
  chrome.tabs.onActivated.addListener(() => atualizarStatusLeds());
  chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
    if (changeInfo.url || changeInfo.status === 'complete') atualizarStatusLeds(); 
  });


  // --- Botões Principais (Agora com Validação) ---

  // 1. Lançar Notas
  configurarBotao('btn-lancar-notas', async () => {
    // 1. Validação de Página
    if (!await validarPaginaCorreta('notas')) return;

    // 2. Validação de Dados
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
    if (!await validarPaginaCorreta('frequencia')) return;

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
    if (!await validarPaginaCorreta('pendencias')) return;

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
    if (!await validarPaginaCorreta('sisedu')) return;

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
// LÓGICA DE VALIDAÇÃO E COMUNICAÇÃO (CORRIGIDA)
// ===================================================================

/**
 * Verifica se a aba ativa contém a URL correta para a ação desejada.
 * Se não tiver, exibe erro e retorna FALSE.
 */
async function validarPaginaCorreta(tipoAcao) {
  try {
    const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
    
    if (!tab || !tab.url) {
      mostrarToast("Nenhuma aba ativa detectada.", "error");
      return false;
    }

    const trechoObrigatorio = REGRAS_CONTEXTO[tipoAcao];
    if (!tab.url.includes(trechoObrigatorio)) {
      mostrarToast("Recurso não disponível nesta página.", "error");
      return false; // BLOQUEIA A AÇÃO
    }

    return true; // PERMITE A AÇÃO
  } catch (e) {
    console.error(e);
    return false;
  }
}

async function enviarComando(acao, dados) {
  try {
    const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
    
    // Envia a mensagem. Se o content.js não estiver lá, isso vai cair no catch.
    const resposta = await chrome.tabs.sendMessage(tab.id, { action: acao, dados: dados });
    return resposta;

  } catch (error) {
    console.warn("Erro de comunicação:", error);
    // Este erro geralmente acontece se o content script não carregou (página errada ou recarregamento necessário)
    mostrarToast("Erro de conexão. Recarregue a página (F5).", "error");
    return null;
  }
}

async function atualizarStatusLeds() {
  try {
    const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
    if (!tab || !tab.url) return;
    
    const url = tab.url;
    
    // Atualiza visualmente os LEDs
    toggleLed('led-notas', url.includes(REGRAS_CONTEXTO['notas']));
    toggleLed('led-frequencia', url.includes(REGRAS_CONTEXTO['frequencia']));
    toggleLed('led-pendencias', url.includes(REGRAS_CONTEXTO['pendencias']));
    toggleLed('led-sisedu', url.includes(REGRAS_CONTEXTO['sisedu']));

  } catch (error) {
    console.log("Erro ao atualizar LEDs:", error);
  }
}

function toggleLed(id, ativo) {
  const led = document.getElementById(id);
  if (led) {
    if (ativo) led.classList.add('active');
    else led.classList.remove('active');
  }
}

// ===================================================================
// UTILITÁRIOS VISUAIS
// ===================================================================

function configuringBotao(id, callback) { // Ops, corrigindo nome no helper abaixo
  const btn = document.getElementById(id);
  if (btn) btn.addEventListener('click', callback);
}
// Alias para manter compatibilidade caso tenha usado o nome acima
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
  const ICONS = { success: '✅ ', error: '⛔ ', info: 'ℹ️ ', warning: '⚠️ ' };
  
  const toast = document.createElement('div');
  toast.className = `toast ${tipo}`;
  toast.textContent = (ICONS[tipo] || '') + mensagem;
  
  container.appendChild(toast);
  
  // Reflow para animação
  void toast.offsetWidth; 
  toast.classList.add('show');

  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => { if (container.contains(toast)) container.removeChild(toast); }, 300);
  }, 4000);
}