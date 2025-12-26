// scripts/sidepanel.js

/**
 * POA - Professor Online Automático
 * Script do Painel Lateral
 * Versão: 1.3
 */

const REGRAS_CONTEXTO = {
  'notas': 'avaliacao_nota',
  'frequencia': 'frequencia_chamada',
  'pendencias': 'relatorio_aulas_dadas',
  'aulas': 'registro_aula_item',
  'sisedu': 'cadastrar_gabarito'
};

document.addEventListener('DOMContentLoaded', () => {
  const textArea = document.getElementById('data-area');

  // ===================================================================
  // 1. GERENCIAMENTO DA ÁREA DE TRANSFERÊNCIA
  // ===================================================================

  // Botão COPIAR
  configurarBotao('btn-copiar', () => {
    if (!textArea.value) { mostrarToast("A área de texto está vazia.", "info"); return; }
    copiarParaAreaTransferencia(textArea.value, textArea);
    mostrarToast("Conteúdo copiado!", "success");
  });

  // Botão COLAR
  configurarBotao('btn-colar', async () => {
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
      console.warn('Erro clipboard:', err);
      mostrarToast("Erro ao colar. Tente Ctrl+V.", "error");
      textArea.focus();
    }
  });

  // Botão LIMPAR
  configurarBotao('btn-limpar', () => {
    if (!textArea.value) return;
    textArea.value = '';
    textArea.focus();
    mostrarToast("Área de texto limpa.", "info");
  });


  // ===================================================================
  // 2. MONITORAMENTO DE STATUS (LEDs)
  // ===================================================================
  atualizarStatusLeds();
  
  chrome.tabs.onActivated.addListener(() => atualizarStatusLeds());
  chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
    if (changeInfo.url || changeInfo.status === 'complete') atualizarStatusLeds(); 
  });


  // ===================================================================
  // 3. BOTÕES DE AÇÃO PRINCIPAL
  // ===================================================================

  // A. Lançar Notas
  configurarBotao('btn-lancar-notas', async () => {
    if (!await validarPaginaCorreta('notas')) return;

    const dados = textArea.value;
    if (!dados) {
      mostrarToast("Cole os dados das notas primeiro.", "error");
      textArea.focus(); return;
    }

    mostrarToast("Processando notas...", "info");
    const resposta = await enviarComando('preencher_notas', dados);
    processarResposta(resposta);
    
    // Limpa a área após sucesso (Notas geralmente é 1x por turma)
    if (verificarSucesso(resposta)) textArea.value = ''; 
  });

  // B. Lançar Frequência
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
    
    // NÃO LIMPA a área (permite lançar múltiplos dias)
  });

  // C. Verificar Pendências
  configurarBotao('btn-verificar-pendencias', async () => {
    if (!await validarPaginaCorreta('pendencias')) return;

    mostrarToast("Analisando pendências...", "info");
    const resposta = await enviarComando('verificar_pendencias', null);
    
    if (resposta && resposta.dados) {
      copiarParaAreaTransferencia(resposta.dados, textArea);
      mostrarToast("Datas copiadas para a caixa de texto!", "success");
    } else {
      processarResposta(resposta);
    }
  });

  // D. Botão SISEDU
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

  // E. Lançar Aulas em Lote
  configurarBotao('btn-lancar-aulas', async () => {
    if (!await validarPaginaCorreta('aulas')) return;

    const dados = textArea.value;
    if (!dados) {
      mostrarToast("Cole os dados (Data | Conteúdo | Subconteúdo).", "error");
      textArea.focus(); return;
    }

    mostrarToast("Iniciando lançamento em lote...", "info");
    
    // Feedback visual imediato que o processo começou
    const btn = document.getElementById('btn-lancar-aulas');
    const textoOriginal = btn.innerHTML;
    btn.innerHTML = '⏳ Processando...';
    btn.disabled = true;

    const resposta = await enviarComando('preencher_aulas', dados);
    
    // Restaura botão
    btn.innerHTML = textoOriginal;
    btn.disabled = false;

    processarResposta(resposta);
  });


  // ===================================================================
  // 4. MÓDULO DE TUTORIAL (DRIVER.JS)
  // ===================================================================
  
  const driver = window.driver.js.driver;
  const driverObj = driver({
    showProgress: true,
    animate: true,
    allowClose: true,
    doneBtnText: "Vamos lá!",
    nextBtnText: "Próximo →",
    prevBtnText: "← Anterior",
    popoverClass: 'driver-popover-poa',
    onDestroyStarted: () => {
      localStorage.setItem('poa_tutorial_visto', 'true');
      driverObj.destroy();
    }
  });
  
  const passosTour = [
    { 
      element: '#header-logo', 
      popover: { 
        title: 'Bem-vindo ao POA!', 
        description: 'Agora que sua extensão está instalada, vamos ver como funciona!', 
        side: 'bottom', align: 'center' 
      } 
    },
    { 
      element: '#acesso-rapido-drive-poa', 
      popover: { 
        title: 'Use nossas planilhas', 
        description: 'Se sua escola nunca usou o POA, baixe nossas planilhas modelo para facilitar o uso com nossa extensão.<br><br> Se já usa, ótimo! Continue com suas próprias planilhas.', 
        side: 'bottom', align: 'center' 
      } 
    },
    { 
      element: '#led-notas', 
      popover: { 
        title: 'Fique de olho nos LEDs!', 
        description: 'Nossos LEDs indicam se você está na página correta para cada ação (lançar notas, frequência, etc).<br><br>🟢 LED verde = tudo certo para inserir os dados.<br>⚪ LED cinza = página não compatível com o POA.', 
        side: 'bottom' 
      } 
    },
    { 
      element: '#transferencia-container', 
      popover: { 
        title: '1. Copie e Cole', 
        description: 'Copie os dados de uma planilha POA e cole dentro desta área.<br><br>Você pode usar os botões para facilitar o processo ou os atalhos Ctrl+C / Ctrl+V.', 
        side: 'bottom' 
      } 
    },
    { 
      element: '#acoes-container', 
      popover: { 
        title: '2. Execute a Ação', 
        description: 'Com os dados colados na Área de Transferência e o LED verde, clique na ação desejada para que o preenchimento automático seja realizado.', 
        side: 'top' 
      } 
    },
    { 
      element: '#social-container', 
      popover: {
        title: 'Contribua', 
        description: 'Gostou do POA? Contriua com ideias e sugestões de ferramentas e melhorias!<br><br>Seu feedback é muito importante para nós.', 
        side: 'top' 
      } 
    },
    { 
      element: '#btn-ajuda', 
      popover: { 
        title: 'Ajuda', 
        description: 'Clique neste botão sempre que precisar rever este tutorial ou tiver dúvidas sobre o funcionamento da nossa extensão.', 
        side: 'left' 
      } 
    }
  ];

  function iniciarTutorial() {
    const stepsValidos = passosTour.filter(passo => document.querySelector(passo.element));
    if (stepsValidos.length > 0) {
      driverObj.setSteps(stepsValidos);
      driverObj.drive();
    }
  }

  // Gatilho Automático (Primeira Vez)
  const jaViu = localStorage.getItem('poa_tutorial_visto');
  if (!jaViu) {
    setTimeout(iniciarTutorial, 800);
  }

  // Gatilho Manual (Botão Ajuda no Header)
  configurarBotao('btn-ajuda', (e) => {
    e.preventDefault();
    iniciarTutorial();
  });

});


// ===================================================================
// LÓGICA DE VALIDAÇÃO E COMUNICAÇÃO
// ===================================================================

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
      return false;
    }
    return true;
  } catch (e) {
    console.error(e);
    return false;
  }
}

async function enviarComando(acao, dados) {
  try {
    const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
    const resposta = await chrome.tabs.sendMessage(tab.id, { action: acao, dados: dados });
    return resposta;
  } catch (error) {
    console.warn("Erro de comunicação:", error);
    mostrarToast("Erro de conexão. Recarregue a página (F5).", "error");
    return null;
  }
}

async function atualizarStatusLeds() {
  try {
    const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
    if (!tab || !tab.url) return;
    
    const url = tab.url;
    toggleLed('led-notas', url.includes(REGRAS_CONTEXTO['notas']));
    toggleLed('led-frequencia', url.includes(REGRAS_CONTEXTO['frequencia']));
    toggleLed('led-pendencias', url.includes(REGRAS_CONTEXTO['pendencias']));
    toggleLed('led-aulas', url.includes(REGRAS_CONTEXTO['aulas']));
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
  
  void toast.offsetWidth; // Force Reflow
  toast.classList.add('show');

  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => { if (container.contains(toast)) container.removeChild(toast); }, 300);
  }, 4000);
}