// ==UserScript==
// @name         Professor Online Automático (POA)
// @namespace    andersonfreitas-art
// @version      1.2
// @description  Automatiza o preenchimento de notas no portal Professor Online da SEDUC-CE a partir de uma lista copiada do Google Planilhas.
// @author       Anderson Freitas (anderson.silva7@prof.ce.gov.br)
// @match        https://professor.seduc.ce.gov.br/avaliacao_nota?ci_avaliacao=*
// @grant        none
// @icon         none
// ==/UserScript==

(function() {
    'use strict';

    // Cria e insere um botão na página para iniciar o processo.
    const parentElement = document.querySelector('.box-shadow.gray-column.col-xs-12 h3');
    if (parentElement) {
        const button = document.createElement('button');
        button.textContent = '🚀 POA';
        button.className = 'btn btn-primary pull-right';
        button.style.marginLeft = '15px';
        button.onclick = preencherNotas;
        parentElement.appendChild(button);
    }

    function preencherNotas() {
        // 1. Pede ao usuário para colar os dados da planilha.
        const pastedData = prompt(
            'COMO USAR:\n\n' +
            '1. Na sua planilha (Google Sheets ou Excel), selecione as DUAS colunas: MATRÍCULA e NOTA.\n' +
            '2. Copie os dados (Ctrl+C).\n' +
            '3. Cole os dados nesta caixa de diálogo (Ctrl+V) e clique em OK.'
        );

        if (!pastedData) {
            alert('Nenhum dado foi colado. Operação cancelada.');
            return;
        }

        // 2. Processa os dados colados.
        const notasMap = new Map();
        const linhas = pastedData.trim().split('\n');
        let processados = 0;

        linhas.forEach(linha => {
            const colunas = linha.split('\t'); // '\t' é o separador de tabulação ao copiar de planilhas
            if (colunas.length >= 2) {
                const matricula = colunas[0].trim();
                const nota = colunas[1].trim().replace(',', '.'); // Garante que a nota use ponto decimal
                if (matricula && nota) {
                    notasMap.set(matricula, nota);
                }
            }
        });

        if (notasMap.size === 0) {
            alert('Não foi possível encontrar dados válidos (matrícula e nota) nos dados colados. Verifique se copiou as colunas corretas.');
            return;
        }

        // 3. Encontra os alunos na página e preenche as notas.
        const alunosNaPagina = document.querySelectorAll('.div-card.Aligner');
        alunosNaPagina.forEach(alunoDiv => {
            const matriculaElement = alunoDiv.querySelector('small[data-item-subtitle]');
            const notaInput = alunoDiv.querySelector('input[data-nota]');

            if (matriculaElement && notaInput) {
                const matriculaPagina = matriculaElement.textContent.trim();
                if (notasMap.has(matriculaPagina)) {
                    const notaParaLancar = notasMap.get(matriculaPagina);
                    notaInput.value = notaParaLancar;

                    // Simula um evento de "change" para que a página reconheça a alteração
                    notaInput.dispatchEvent(new Event('change', { bubbles: true }));

                    processados++;
                    // Adiciona um feedback visual de que a nota foi preenchida
                    notaInput.style.backgroundColor = '#d4edda'; // Verde claro
                }
            }
        });

        // 4. Exibe um resumo da operação.
        alert(
            `Processo Concluído!\n\n` +
            `- ${processados} notas foram preenchidas na página.\n` +
            `- ${notasMap.size - processados} matrículas da sua lista não foram encontradas na turma atual.\n\n` +
            '**IMPORTANTE:** Verifique visualmente se as notas estão corretas antes de clicar em "Salvar"!'
        );

        // Mostra o botão de salvar da página, caso ele esteja oculto
        const snack = document.getElementById('snack');
        if (snack) {
            snack.style.display = 'block';
        }
    }
})();