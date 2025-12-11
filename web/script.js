let dadosEnf = [], dadosUti = [];
let editorEnf = [], editorUti = [];
let filaImpressao = [];
let setorAtual = 'ENF';
let abaEditorAtual = 'ENF';

// Inicialização
window.addEventListener('pywebviewready', carregarDados);
setTimeout(() => { if (dadosEnf.length === 0) carregarDados(); }, 1500);

async function recarregarDadosComFeedback() {
    await carregarDados();
    alert("✅ Dados atualizados!");
}

function escaparTexto(texto) {
    if (texto === null || texto === undefined) return "";
    return String(texto).replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// --- NAVEGAÇÃO ---
function mudarAba(aba) {
    document.querySelectorAll('.tab-content').forEach(d => d.style.display = 'none');
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));

    let btnIndex = (aba === 'enf') ? 0 : (aba === 'uti') ? 1 : 2;
    document.querySelectorAll('.nav-btn')[btnIndex].classList.add('active');

    if (aba === 'editor') {
        document.getElementById('tab-editor').style.display = 'block';
        document.getElementById('painelBusca').style.display = 'none';
        abaEditorAtual = 'ENF';
        renderizarEditor();
    } else {
        setorAtual = (aba === 'enf') ? 'ENF' : 'UTI';
        document.getElementById('tab-dashboard').style.display = 'block';
        let titulo = (setorAtual === 'ENF') ? 'Enfermarias' : 'UTI - HRMSS';
        document.getElementById('tituloSetor').innerText = titulo;

        document.getElementById('painelBusca').style.display = 'block';
        renderizarLista(setorAtual === 'ENF' ? dadosEnf : dadosUti);

        filaImpressao = []; atualizarFila();
    }
}

// --- CARREGAMENTO ---
async function carregarDados() {
    if (window.pywebview && window.pywebview.api) {
        try {
            let res = await pywebview.api.carregar_dados_excel();

            if (res.sucesso) {
                dadosEnf = res.dados_enf;
                dadosUti = res.dados_uti;
                editorEnf = res.editor_enf;
                editorUti = res.editor_uti;

                if (document.getElementById('tab-editor').style.display === 'block') {
                    renderizarEditor();
                } else {
                    renderizarLista(setorAtual === 'ENF' ? dadosEnf : dadosUti);
                }
            } else {
                alert("Erro ao ler Excel: " + res.erro);
            }
        } catch (e) {
            console.error(e);
        }
    }
}

// --- LISTA LATERAL ---
function renderizarLista(lista) {
    const div = document.getElementById("listaPacientes");
    div.innerHTML = "";
    if (!lista || lista.length === 0) {
        div.innerHTML = "<p style='text-align:center; padding:20px; color:#aaa'>Nenhum paciente encontrado.</p>";
        return;
    }
    lista.forEach(p => {
        let item = document.createElement("div");
        item.className = "patient-item";
        item.onclick = () => adicionarFila(p);

        let local = setorAtual === 'ENF' ? p['ENFERMARIA'] : 'UTI';
        let dieta = p['DIETA'] ? p['DIETA'] : '---';

        item.innerHTML = `<h4>${p['LEITO']} - ${p['NOME DO PACIENTE']}</h4><p>${local} | ${dieta}</p>`;
        div.appendChild(item);
    });
}

function filtrarLista() {
    let termo = document.getElementById("inputBusca").value.toLowerCase();
    let listaBase = (setorAtual === 'ENF') ? dadosEnf : dadosUti;

    let filtrados = listaBase.filter(p => {
        let nome = String(p['NOME DO PACIENTE']).toLowerCase();
        let leito = String(p['LEITO']).toLowerCase();
        return nome.includes(termo) || leito.includes(termo);
    });
    renderizarLista(filtrados);
}

// --- EDITOR ---
function renderizarEditor() {
    const container = document.getElementById("editorControls");
    const clsEnf = abaEditorAtual === 'ENF' ? 'btn-primary' : 'btn-secondary';
    const clsUti = abaEditorAtual === 'UTI' ? 'btn-primary' : 'btn-secondary';

    container.innerHTML = `
        <button class="btn ${clsEnf}" onclick="trocarEditor('ENF')">Enfermaria (${editorEnf.length})</button>
        <button class="btn ${clsUti}" onclick="trocarEditor('UTI')">UTI (${editorUti.length})</button>
    `;

    const dados = (abaEditorAtual === 'ENF') ? editorEnf : editorUti;

    const tbody = document.getElementById("corpoTabelaEditor");
    const thead = document.querySelector("#tabelaEditor thead");
    tbody.innerHTML = "";
    thead.innerHTML = "";

    let trHead = document.createElement("tr");
    if (abaEditorAtual === 'ENF') {
        trHead.innerHTML = `<th>ENFERMARIA</th><th>LEITO</th><th>NOME DO PACIENTE</th><th>DIETA</th><th>OBSERVAÇÕES</th><th style="width:50px">X</th>`;
    } else {
        trHead.innerHTML = `<th>LEITO</th><th>NOME DO PACIENTE</th><th>DIETA</th><th>OBSERVAÇÕES</th><th style="width:50px">X</th>`;
    }
    thead.appendChild(trHead);

    dados.forEach(row => criarLinhaEditor(tbody, row));
}

function criarLinhaEditor(tbody, row = {}) {
    let tr = document.createElement("tr");
    let html = "";
    const val = (k) => escaparTexto(row[k]);

    if (abaEditorAtual === 'ENF') {
        html += `<td><input class="edit-enf" value="${val('ENFERMARIA')}"></td>`;
    }
    html += `
        <td><input class="edit-leito" value="${val('LEITO')}"></td>
        <td><input class="edit-nome" value="${val('NOME DO PACIENTE')}"></td>
        <td><input class="edit-dieta" value="${val('DIETA')}"></td>
        <td><input class="edit-obs" value="${val('OBSERVAÇÕES')}"></td>
        <td style="text-align:center">
            <button class="btn-remove" tabindex="-1" onclick="this.closest('tr').remove()" title="Excluir">
                <span class="material-icons" style="font-size:18px">delete</span>
            </button>
        </td>
    `;
    tr.innerHTML = html;
    tbody.appendChild(tr);
}

function trocarEditor(tipo) {
    salvarEstadoTemporario();
    abaEditorAtual = tipo;
    renderizarEditor();
}

function salvarEstadoTemporario() {
    const linhas = document.querySelectorAll("#corpoTabelaEditor tr");
    let novosDados = [];

    linhas.forEach(tr => {
        let obj = {};
        if (abaEditorAtual === 'ENF') {
            obj['ENFERMARIA'] = tr.querySelector(".edit-enf") ? tr.querySelector(".edit-enf").value : "";
        }
        obj['LEITO'] = tr.querySelector(".edit-leito").value;
        obj['NOME DO PACIENTE'] = tr.querySelector(".edit-nome").value;
        obj['DIETA'] = tr.querySelector(".edit-dieta").value;
        obj['OBSERVAÇÕES'] = tr.querySelector(".edit-obs").value;

        if (obj['LEITO'] || obj['NOME DO PACIENTE']) {
            novosDados.push(obj);
        }
    });

    if (abaEditorAtual === 'ENF') editorEnf = novosDados;
    else editorUti = novosDados;
}

function adicionarLinhaVazia() {
    const tbody = document.getElementById("corpoTabelaEditor");
    criarLinhaEditor(tbody);
    setTimeout(() => {
        tbody.lastElementChild.scrollIntoView({ behavior: 'smooth', block: 'center' });
        let inputs = tbody.lastElementChild.querySelectorAll('input');
        if (inputs.length > 0) inputs[0].focus();
    }, 100);
}

async function salvarExcel() {
    salvarEstadoTemporario();
    if (confirm(`Salvar alterações?\n\nEnfermaria: ${editorEnf.length} registros\nUTI: ${editorUti.length} registros`)) {
        let res = await pywebview.api.salvar_dados_excel(editorEnf, editorUti);
        if (res.sucesso) {
            alert("✅ Salvo com sucesso!");
            carregarDados();
        } else {
            alert("❌ " + res.msg);
        }
    }
}

function adicionarFila(p) { filaImpressao.push(p); atualizarFila(); }
function limparFila() { filaImpressao = []; atualizarFila(); }
function adicionarTodos() {
    let lista = (setorAtual === 'ENF') ? dadosEnf : dadosUti;
    filaImpressao = [...lista]; atualizarFila();
}
function atualizarFila() {
    document.getElementById("contadorFila").innerText = filaImpressao.length + " etiquetas";
    const ul = document.getElementById("listaFila");
    ul.innerHTML = "";
    if (filaImpressao.length === 0) ul.innerHTML = '<li class="empty-msg">Fila vazia.</li>';
    else filaImpressao.forEach((p, i) => {
        ul.innerHTML += `<li ondblclick="filaImpressao.splice(${i},1);atualizarFila()">✅ ${p['LEITO']} - ${p['NOME DO PACIENTE']}</li>`;
    });
}

async function imprimirFila() {
    if (filaImpressao.length === 0) { alert("Adicione pacientes à fila primeiro."); return; }
    let msg = await pywebview.api.imprimir_etiquetas(filaImpressao);
    if (msg !== "Cancelado.") alert(msg);
}

// --- CORREÇÃO AQUI (ADICIONADO OS ALERTAS) ---
async function gerarRelatorioSimples() {
    let msg;
    if (setorAtual === 'ENF') {
        msg = await pywebview.api.gerar_relatorio_enf('simples');
    } else {
        msg = await pywebview.api.gerar_relatorio_uti('simples');
    }

    if (msg !== "Cancelado.") alert(msg);
}

async function gerarMapaGeral() {
    let msg;
    if (setorAtual === 'ENF') {
        msg = await pywebview.api.gerar_relatorio_enf('geral');
    } else {
        msg = await pywebview.api.gerar_relatorio_uti('geral');
    }

    if (msg !== "Cancelado.") alert(msg);
}