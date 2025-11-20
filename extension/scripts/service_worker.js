// scripts/service_worker.js

/**
 * Configura o comportamento da extensão.
 * Define que o clique no ícone (Action) deve abrir o Painel Lateral (Side Panel).
 */
chrome.sidePanel
  .setPanelBehavior({ openPanelOnActionClick: true })
  .catch((error) => console.error("POA: Erro ao configurar sidePanel:", error));