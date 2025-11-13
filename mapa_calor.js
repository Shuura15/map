const RUAS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15];
const CAMINHO_ARQUIVO_FIXO = 'Mapa.xlsx';

function interpolarCor(cor1, cor2, fator ) {
    const r = Math.round(cor1.r + (cor2.r - cor1.r) * fator);
    const g = Math.round(cor1.g + (cor2.g - cor1.g) * fator);
    const b = Math.round(cor1.b + (cor2.b - cor1.b) * fator);
    return `rgb(${r}, ${g}, ${b})`;
}

const corMinimo = { r: 255, g: 255, b: 255 };
const corBaixo = { r: 255, g: 255, b: 0 };
const corMedio = { r: 255, g: 165, b: 0 };
const corAlto = { r: 255, g: 69, b: 0 };
const corMaximo = { r: 255, g: 0, b: 0 };

function calcularCor(valor, min, max) {
    if (valor === undefined || valor === null || valor === 0) {
        return 'rgba(200, 200, 200, 0.3)';
    }

    const range = max - min;
    const normalized = range === 0 ? 0.5 : (valor - min) / range;

    if (normalized < 0.25) {
        const fator = normalized / 0.25;
        return interpolarCor(corMinimo, corBaixo, fator);
    } else if (normalized < 0.5) {
        const fator = (normalized - 0.25) / 0.25;
        return interpolarCor(corBaixo, corMedio, fator);
    } else if (normalized < 0.75) {
        const fator = (normalized - 0.5) / 0.25;
        return interpolarCor(corMedio, corAlto, fator);
    } else {
        const fator = (normalized - 0.75) / 0.25;
        return interpolarCor(corAlto, corMaximo, fator);
    }
}

function getEstruturaPredios() {
    const estrutura = {};
    for (let rua of RUAS) {
        let prediosImpar = [];
        let prediosPar = [];

        if (rua === 1) {
            for (let i = 41; i <= 69; i += 2) {
                prediosImpar.push(i);
            }
            prediosPar = Array.from({ length: 5 }, (_, i) => 48 + i * 2);

        } else if (rua === 2) {
            for (let i = 3; i <= 46; i += 2) prediosImpar.push(i);
            for (let i = 2; i <= 46; i += 2) prediosPar.push(i);

        } else if (rua === 12) {
            for (let i = 1; i <= 46; i += 2) prediosImpar.push(i);
            for (let i = 2; i <= 46; i += 2) prediosPar.push(i);


        } else if (rua === 13) {
            for (let i = 1; i <= 46; i += 2) prediosImpar.push(i);
            for (let i = 2; i <= 46; i += 2) prediosPar.push(i);
        
        } else {
            for (let i = 1; i <= 46; i += 2) prediosImpar.push(i);
            for (let i = 2; i <= 46; i += 2) prediosPar.push(i);
        }
        estrutura[rua] = { Impar: prediosImpar, Par: prediosPar };
    }
    return estrutura;
}

let dadosGlobais = {
    filtros: { tipo_movimento: [], picking: [], curva: [] },
    dados_completos: [] 
};
let filtrosAtivos = { tipoMovimento: '', picking: '', curva: '', dataInicial: '', dataFinal: ''};
let predioSelecionado = null; 

function formatarNumero(numero) {
    let numStr = numero.toString().replace('.', ',');
    return numStr.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

// ===================================================================
//                FUNÇÕES HELPER
// ===================================================================
function converterData(dataValor) {
    if (dataValor === 'N/A' || !dataValor) {
        return null;
    }
    
    const dataValorStr = String(dataValor); 
    const dataNum = parseFloat(dataValorStr); 

    if (!isNaN(dataNum) && dataNum > 25569 && !dataValorStr.includes('/') && !dataValorStr.includes('-')) {
        const utcTimestamp = Math.round((dataNum - 25569) * 86400 * 1000);
        const timezoneOffsetInMs = new Date(utcTimestamp).getTimezoneOffset() * 60 * 1000; 
        return new Date(utcTimestamp + timezoneOffsetInMs);
    
    } else if (dataValorStr.includes('/')) { 
        const partes = dataValorStr.split(' ')[0].split('/');
        if (partes.length === 3) {
            const [dia, mes, ano] = partes;
            return new Date(`${mes}/${dia}/${ano}`); // MM/DD/YYYY
        }
    }
    
    const data = new Date(dataValor); 
    return !isNaN(data) ? data : null;
}

function formatarDataString(data) {
    if (!data || isNaN(data.getTime())) {
        return "N/A";
    }
    const dia = String(data.getDate()).padStart(2, '0');
    const mes = String(data.getMonth() + 1).padStart(2, '0');
    const ano = data.getFullYear();
    return `${dia}/${mes}/${ano}`;
}

function getDadosFiltrados() {
    const { tipoMovimento, picking, curva, dataInicial, dataFinal } = filtrosAtivos;

    return dadosGlobais.dados_completos.filter(item => {
        const matchTipo = !tipoMovimento || item['Tipo de movimento'] === tipoMovimento;
        const matchPicking = !picking || item['Picking?'] === picking;
        const matchCurva = !curva || item['Curva'] === curva;
        
        let matchData = true;
        if (dataInicial || dataFinal) {
            if (!item.DataObj) { 
                matchData = false;
            } else {
                if (dataInicial && item.DataObj < dataInicial) {
                    matchData = false;
                }
                if (dataFinal && item.DataObj > dataFinal) {
                    matchData = false;
                }
            }
        }
        
        return matchTipo && matchPicking && matchCurva && matchData;
    });
}

// ===================================================================
//                FIM DAS FUNÇÕES HELPER
// ===================================================================

function renderizarMapa() {
    const container = document.getElementById('mapaContainer');
    container.innerHTML = '';
    const estrutura = getEstruturaPredios();
    
    fecharDetalhes(); 

    const dadosFiltrados = getDadosFiltrados();

    const contagensPorLocal = {};
    let totalFiltrado = 0;

    for (const item of dadosFiltrados) {
        contagensPorLocal[item.Localizacao] = (contagensPorLocal[item.Localizacao] || 0) + 1; 
        totalFiltrado++;
    }

    let minFiltrado = 0;
    let maxFiltrado = 0;
    let contagemPosicoes = 0;
    const valoresContagem = Object.values(contagensPorLocal);

    if (valoresContagem.length > 0) {
        minFiltrado = Infinity;
        maxFiltrado = -Infinity;
        for (const contagem of valoresContagem) {
            minFiltrado = Math.min(minFiltrado, contagem);
            maxFiltrado = Math.max(maxFiltrado, contagem);
            contagemPosicoes++;
        }
    }

    for (let rua of RUAS) {
        const ruaDiv = document.createElement('div');
        ruaDiv.className = 'rua';
        const ruaTitulo = document.createElement('div');
        ruaTitulo.className = 'rua-titulo';
        ruaTitulo.textContent = `Rua ${rua}`;
        ruaDiv.appendChild(ruaTitulo);
        const ruaContent = document.createElement('div');
        ruaContent.className = 'rua-content';

        if (estrutura[rua].Impar.length > 0) {
            const ladoDiv = document.createElement('div');
            ladoDiv.className = 'lado';
            const prediosDiv = document.createElement('div');
            prediosDiv.className = 'predios';
            const prediosImpar = estrutura[rua].Impar;
            for (let predio of prediosImpar) {
                const localizacao = `${rua}-${predio}-Impar`;
                const contagem = contagensPorLocal[localizacao] || 0;
                const predioDiv = renderizarPredio(predio, contagem, minFiltrado, maxFiltrado, localizacao);
                prediosDiv.appendChild(predioDiv);
            }
            ladoDiv.appendChild(prediosDiv);
            ruaContent.appendChild(ladoDiv);
        }

        if (estrutura[rua].Par.length > 0) {
            const ladoDiv = document.createElement('div');
            ladoDiv.className = 'lado';
            const prediosDiv = document.createElement('div');
            prediosDiv.className = 'predios';
            const prediosPar = estrutura[rua].Par;
            for (let predio of prediosPar) {
                const localizacao = `${rua}-${predio}-Par`;
                const contagem = contagensPorLocal[localizacao] || 0;
                const predioDiv = renderizarPredio(predio, contagem, minFiltrado, maxFiltrado, localizacao);
                prediosDiv.appendChild(predioDiv);
            }
            ladoDiv.appendChild(prediosDiv);
            ruaContent.appendChild(ladoDiv);
        }
        ruaDiv.appendChild(ruaContent);
        container.appendChild(ruaDiv);
    }

    document.getElementById('minValue').textContent = contagemPosicoes > 0 ? formatarNumero(minFiltrado) : '-';
    document.getElementById('maxValue').textContent = contagemPosicoes > 0 ? formatarNumero(maxFiltrado) : '-';
    const media = contagemPosicoes > 0 ? (totalFiltrado / contagemPosicoes).toFixed(2) : '0';
    document.getElementById('avgValue').textContent = contagemPosicoes > 0 ? formatarNumero(media) : '-';

    document.getElementById('totalPositions').textContent = formatarNumero(contagemPosicoes);
    document.getElementById('totalMovements').textContent = totalFiltrado > 0 ? formatarNumero(totalFiltrado) : (contagemPosicoes > 0 ? '0' : '-');
}


function isPAR(rua, predio) {
    return rua === 4 && predio >= 25 && predio <= 44;
}

function renderizarPredio(predio, contagem, min, max, localizacao) {
    const predioDiv = document.createElement('div');
    predioDiv.className = 'predio';

    const [rua, predioNum] = localizacao.split('-').map(s => parseInt(s));
    if (isPAR(rua, predioNum)) {
        predioDiv.classList.add('par');
    }

    if (contagem === 0) {
        predioDiv.classList.add('sem-dados');
        predioDiv.textContent = '-';
    } else {
        const parStatus = isPAR(rua, predioNum) ? ' (P.A.R.)' : '';
        const cor = calcularCor(contagem, min, max);
        predioDiv.style.backgroundColor = cor;
        predioDiv.textContent = formatarNumero(contagem);

        const tooltip = document.createElement('div');
        tooltip.className = 'predio-tooltip';
        tooltip.textContent = `Prédio - ${predio}${parStatus}. Clique para ver detalhes.`;
        predioDiv.appendChild(tooltip);

        predioDiv.addEventListener('click', (e) => {
            e.stopPropagation(); 
            
            if (predioSelecionado) {
                predioSelecionado.classList.remove('selecionado');
            }
            
            predioSelecionado = predioDiv;
            predioSelecionado.classList.add('selecionado');

            mostrarDetalhesPredio(localizacao);
        });
    }
    return predioDiv;
}

// ===================================================================
//      FUNÇÃO 'mostrarDetalhesPredio' (COM ORDENAÇÃO AUTOMÁTICA)
// ===================================================================

function fecharDetalhes() {
    const detalhesContainer = document.getElementById('detalhesContainer');
    detalhesContainer.style.display = 'none';
    detalhesContainer.innerHTML = '';
    
    if (predioSelecionado) {
        predioSelecionado.classList.remove('selecionado');
        predioSelecionado = null;
    }
}

function mostrarDetalhesPredio(localizacao) {
    // =================================================================
    // !! ALTERAÇÃO 1: Adicionado teste de console !!
    // =================================================================
    console.log('--- EXECUTANDO FUNÇÃO DE DETALHES (VERSÃO ORDENAÇÃO AUTOMÁTICA v2) ---');

    const detalhesContainer = document.getElementById('detalhesContainer');
    
    const dadosFiltrados = getDadosFiltrados();
    const dadosPredio = dadosFiltrados.filter(item => item.Localizacao === localizacao);
    
    const pivotData = {};
    const todosTiposMovSet = new Set(); 

    for (const item of dadosPredio) {
        const tipoMov = item['Tipo de movimento'];
        const endereco = `${item['Nível']}-${item['Apartamento']}`; 

        todosTiposMovSet.add(tipoMov); 
        
        if (!pivotData[endereco]) {
            pivotData[endereco] = {}; 
            pivotData[endereco]['Total'] = 0;
        }
        
        pivotData[endereco][tipoMov] = (pivotData[endereco][tipoMov] || 0) + 1;
        pivotData[endereco]['Total']++;
    }


    // 4. Monta o HTML
    const [rua, predioNum, lado] = localizacao.split('-');
    let html = `
        <h3>
            Detalhes: Rua ${rua} - Prédio ${predioNum} (${lado})
            <button id="fecharDetalhesBtn">Fechar (X)</button>
        </h3>
        
        <div id="detalhes-enderecos">
            <h4>Movimentos por Endereço</h4>
            <div class="tabela-detalhes-container">
                <table class="detalhes">
                    <thead>
                        <tr>
    `;
    
    // Cabeçalho da Tabela Pivot
    html += `<th>Endereço (Nível-Apto)</th>`;
    const headersMov = Array.from(todosTiposMovSet).sort(); 
    for (const tipo of headersMov) {
        html += `<th>${tipo}</th>`;
    }
    html += `<th style="font-weight: bold;">Total</th>`;
    html += `
                        </tr>
                    </thead>
                    <tbody>
    `;
    
    // =================================================================
    // !! ALTERAÇÃO 2: Método de ordenação modificado !!
    // =================================================================
    
    // 1. Converte o objeto pivotData em um array de [chave, valor]
    // Ex: [ ["1-101", {Total: 5, ...}], ["1-102", {Total: 10, ...}] ]
    const entradasPivot = Object.entries(pivotData);

    // 2. Ordena o array com base no valor 'Total' (que está em item[1].Total)
    entradasPivot.sort(([, dadosA], [, dadosB]) => {
        const totalA = dadosA['Total'] || 0;
        const totalB = dadosB['Total'] || 0;
        return totalB - totalA; // Ordena do maior para o menor
    });
    // =================================================================

     if (entradasPivot.length > 0) {
        // =================================================================
        // !! ALTERAÇÃO 3: Loop modificado para usar o array ordenado !!
        // =================================================================
        for (const [endereco, rowData] of entradasPivot) {
            html += `<tr><td>${endereco}</td>`;
            
            for (const tipo of headersMov) {
                const contagem = rowData[tipo] || 0; 
                html += `<td>${formatarNumero(contagem)}</td>`;
            }
            html += `<td style="font-weight: bold;">${formatarNumero(rowData['Total'])}</td>`;
            html += `</tr>`;
        }
    } else {
        // +2 para "Endereço" e "Total"
        html += `<tr><td colspan="${headersMov.length + 2}">Nenhum endereço encontrado.</td></tr>`;
    }

    html += `
                    </tbody>
                </table>
            </div>
        </div>
    `;

    detalhesContainer.innerHTML = html;
    detalhesContainer.style.display = 'block';
    
    document.getElementById('fecharDetalhesBtn').addEventListener('click', fecharDetalhes);
    
    detalhesContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
}
// ===================================================================
//                FIM DA FUNÇÃO DE DETALHES
// ===================================================================


async function garantirXLSX() {
    if (typeof XLSX === 'undefined') {
        try {
            console.log('Carregando biblioteca XLSX...');
            const script = document.createElement('script');
            script.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
            document.head.appendChild(script );
            await new Promise((resolve, reject) => {
                script.onload = resolve;
                script.onerror = reject;
            });
            console.log('XLSX carregada.');
        } catch (error) {
            console.error('Falha ao carregar a biblioteca XLSX.', error);
            throw new Error('Não foi possível carregar a dependência (XLSX).');
        }
    }
}

async function carregarExcelFixo(url) {
    try {
        document.getElementById('loading').classList.add('active');
        document.getElementById('errorMessage').classList.remove('active');
        await garantirXLSX();
        const response = await fetch(url);
        if (!response.ok) {
            throw new Error(`Falha ao buscar arquivo: ${response.status} ${response.statusText}`);
        }
        const data = await response.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        processarDados(jsonData);
        document.getElementById('loading').classList.remove('active');
    } catch (error) {
        mostrarErro(`Erro ao carregar arquivo fixo (${url}): ${error.message}`);
        document.getElementById('loading').classList.remove('active');
    }
}

// ===================================================================
//                FUNÇÃO 'processarDados'
// ===================================================================

function processarDados(dados) {
    const tiposMovimento = new Set();
    const pickings = new Set();
    const curvas = new Set();
    const dadosCompletos = [];

    const estrutura = getEstruturaPredios();
    const validLocations = new Set();
    for (const rua in estrutura) {
        for (const lado in estrutura[rua]) {
            for (const predio of estrutura[rua][lado]) {
                validLocations.add(`${rua}-${predio}-${lado}`);
            }
        }
    }

    for (let row of dados) {
        const rua = parseInt(row['Rua']);
        const predio = parseInt(row['Prédio']);
        
        const tipoMov = row['Tipo de movimento'] ?? 'N/A';
        const picking = row['Picking?'] ?? 'N/A';
        const curva = row['Curva'] ?? 'N/A';
        const dataValor = String(row['Data'] ?? 'N/A'); 
        
        const produto = row['Produto'] ?? 'N/A';
        const nivel = row['Nível'] ?? 'N/A'; // Corrigido para 'Nível'
        const apartamento = row['Apartamento'] ?? 'N/A';

        if (isNaN(rua) || isNaN(predio)) {
            continue; 
        }

        const lado = predio % 2 === 0 ? 'Par' : 'Impar';
        const localizacao = `${rua}-${predio}-${lado}`;

        if (validLocations.has(localizacao)) {
            dadosCompletos.push({
                Localizacao: localizacao,
                'Tipo de movimento': tipoMov,
                'Picking?': picking,
                'Curva': curva,
                'Data': dataValor, 
                'DataObj': converterData(dataValor),
                'Produto': produto,
                'Nível': nivel, 
                'Apartamento': apartamento
            });
            
            tiposMovimento.add(tipoMov);
            pickings.add(picking);
            curvas.add(curva);
        }
    }

    dadosGlobais.dados_completos = dadosCompletos;
    dadosGlobais.filtros.tipo_movimento = Array.from(tiposMovimento).sort();
    dadosGlobais.filtros.picking = Array.from(pickings).sort();
    dadosGlobais.filtros.curva = Array.from(curvas).sort();

    console.log(`Processamento concluído. ${dadosCompletos.length} linhas válidas carregadas.`);
    
    atualizarFiltrosUI();
    renderizarMapa();
}
// ===================================================================
//                FIM DA FUNÇÃO 'processarDados'
// ===================================================================


function atualizarFiltrosUI() { 
    let dataInicialPicker;
    let dataFinalPicker;

    const flatpickrConfig = {
        dateFormat: "d/m/Y",
        locale: "pt",
        allowInput: true
    };

    dataInicialPicker = flatpickr("#dataInicial", {
        ...flatpickrConfig,
        onClose: function(selectedDates, dateStr, instance) {
            if (selectedDates.length > 0) {
                filtrosAtivos.dataInicial = new Date(selectedDates[0].setHours(0, 0, 0, 0));
                if (dataFinalPicker) {
                    dataFinalPicker.set("minDate", selectedDates[0]);
                }
            } else {
                filtrosAtivos.dataInicial = null;
                if (dataFinalPicker) {
                    dataFinalPicker.set("minDate", null); 
                }
            }
        }
    });

    dataFinalPicker = flatpickr("#dataFinal", {
        ...flatpickrConfig,
        onClose: function(selectedDates, dateStr, instance) {
            if (selectedDates.length > 0) {
                filtrosAtivos.dataFinal = new Date(selectedDates[0].setHours(23, 59, 59, 999));
                if (dataInicialPicker) {
                    dataInicialPicker.set("maxDate", selectedDates[0]);
                }
            } else {
                filtrosAtivos.dataFinal = null;
                if (dataInicialPicker) {
                    dataInicialPicker.set("maxDate", null);
                }
            }
        }
    });

    const tipoMovSelect = document.getElementById('tipoMovimento');
    const pickingSelect = document.getElementById('picking');
    const curvaSelect = document.getElementById('curva');
    [tipoMovSelect, pickingSelect, curvaSelect].forEach(sel => {
        while (sel.options.length > 1) sel.remove(1);
    });
    dadosGlobais.filtros.tipo_movimento.forEach(val => tipoMovSelect.add(new Option(val, val)));
    dadosGlobais.filtros.picking.forEach(val => pickingSelect.add(new Option(val, val)));
    dadosGlobais.filtros.curva.forEach(val => curvaSelect.add(new Option(val, val)));
}

function mostrarErro(mensagem) {
    const errorDiv = document.getElementById('errorMessage');
    errorDiv.textContent = mensagem;
    errorDiv.classList.add('active');
}

document.getElementById('aplicarFiltros').addEventListener('click', () => {
    filtrosAtivos.tipoMovimento = document.getElementById('tipoMovimento').value;
    filtrosAtivos.picking = document.getElementById('picking').value;
    filtrosAtivos.curva = document.getElementById('curva').value;
    
    renderizarMapa();
});

document.getElementById('resetarFiltros').addEventListener('click', () => {
    document.getElementById('tipoMovimento').value = '';
    document.getElementById('picking').value = '';
    document.getElementById('curva').value = '';
    filtrosAtivos = { tipoMovimento: '', picking: '', curva: '', dataInicial: null, dataFinal: null };

    const fpInicial = document.querySelector("#dataInicial")._flatpickr;
    if (fpInicial) {
        fpInicial.clear();
        fpInicial.set("maxDate", null);
    }

    const fpFinal = document.querySelector("#dataFinal")._flatpickr;
    if (fpFinal) {
        fpFinal.clear();
        fpFinal.set("minDate", null);
    }

    renderizarMapa(); 
});

document.addEventListener('DOMContentLoaded', () => {
    renderizarMapa(); 
    console.log('Mapa de Calor carregado. Iniciando carregamento do arquivo fixo...');
    carregarExcelFixo(CAMINHO_ARQUIVO_FIXO);

    document.querySelector('.container').addEventListener('click', (e) => {
        if (!e.target.closest('.predio') && !e.target.closest('.detalhes-container')) {
            fecharDetalhes();
        }
    });
});