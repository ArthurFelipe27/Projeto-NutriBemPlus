let todosPacientes = [];
let dadosCompletos = [];
let filaImpressao = [];

// Inicialização
window.addEventListener('pywebviewready', function () {
    carregarExcel();
});

// Fallback
setTimeout(() => {
    if (todosPacientes.length === 0) carregarExcel();
}, 1000);

// --- FUNÇÃO DO BOTÃO DE RECARREGAR ---
async function recarregarDadosComFeedback() {
    await carregarExcel();
    alert("✅ Dados atualizados com sucesso!");
}

// --- NAVEGAÇÃO ---
function mudarAba(abaNome) {
    document.getElementById('tab-dashboard').style.display = 'none';
    document.getElementById('tab-editor').style.display = 'none';
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));

    if (abaNome === 'dashboard') {
        document.getElementById('tab-dashboard').style.display = 'block';
        document.querySelectorAll('.nav-btn')[0].classList.add('active');
        document.getElementById('painelBusca').style.display = 'block';
        carregarExcel(); // Garante dados frescos ao voltar pro painel
    } else {
        document.getElementById('tab-editor').style.display = 'block';
        document.querySelectorAll('.nav-btn')[1].classList.add('active');
        document.getElementById('painelBusca').style.display = 'none';
        renderizarEditor();
    }
}

// --- CARREGAMENTO ---
async function carregarExcel() {
    if (window.pywebview && window.pywebview.api) {
        try {
            // Limpa fila antiga para evitar conflito de dados
            // (Opcional: se quiser manter a fila, remova a linha abaixo)
            filaImpressao = [];
            atualizarVisualFila();

            let resultado = await pywebview.api.carregar_dados_excel();

            if (resultado.sucesso) {
                todosPacientes = resultado.dados;
                dadosCompletos = resultado.dados_editor;

                renderizarLista(todosPacientes);

                // Se estiver na aba editor, atualiza ela também
                if (document.getElementById('tab-editor').style.display === 'block') {
                    renderizarEditor();
                }
            } else {
                alert(resultado.erro);
            }
        } catch (error) {
            console.error("Erro comunicação Python:", error);
        }
    }
}

// --- RENDERIZAÇÃO LISTA LATERAL ---
function renderizarLista(lista) {
    const divLista = document.getElementById("listaPacientes");
    divLista.innerHTML = "";

    if (lista.length === 0) {
        divLista.innerHTML = "<p style='text-align:center; color:#777;'>Nenhum paciente encontrado.</p>";
        return;
    }

    lista.forEach(p => {
        let div = document.createElement("div");
        div.className = "patient-item";
        div.onclick = () => adicionarFila(p);

        let dieta = p['DIETA'] ? p['DIETA'] : '---';

        div.innerHTML = `
            <h4>${p['LEITO']} - ${p['NOME DO PACIENTE']}</h4>
            <p>${p['ENFERMARIA']} | ${dieta}</p>
        `;
        divLista.appendChild(div);
    });
}

function filtrarLista() {
    let termo = document.getElementById("inputBusca").value.toLowerCase();

    let filtrados = todosPacientes.filter(p => {
        let nome = String(p['NOME DO PACIENTE']).toLowerCase();
        let leito = String(p['LEITO']).toLowerCase();
        return nome.includes(termo) || leito.includes(termo);
    });

    renderizarLista(filtrados);
}

// --- EDITOR ---
function renderizarEditor() {
    const tbody = document.getElementById("corpoTabelaEditor");
    tbody.innerHTML = "";

    dadosCompletos.forEach(row => {
        criarLinhaEditor(tbody, row);
    });
}

function criarLinhaEditor(tbody, dados = {}) {
    let tr = document.createElement("tr");

    tr.innerHTML = `
        <td><input type="text" class="edit-enf" value="${dados['ENFERMARIA'] || ''}"></td>
        <td><input type="text" class="edit-leito" value="${dados['LEITO'] || ''}"></td>
        <td><input type="text" class="edit-nome" value="${dados['NOME DO PACIENTE'] || ''}"></td>
        <td><input type="text" class="edit-dieta" value="${dados['DIETA'] || ''}"></td>
        <td><input type="text" class="edit-obs" value="${dados['OBSERVAÇÕES'] || ''}"></td>
        <td style="text-align:center;">
            <button class="btn-remove" onclick="this.closest('tr').remove()" title="Excluir">
                <span class="material-icons">delete</span>
            </button>
        </td>
    `;
    tbody.appendChild(tr);
}

function adicionarLinhaVazia() {
    const tbody = document.getElementById("corpoTabelaEditor");
    criarLinhaEditor(tbody);
    tbody.lastElementChild.scrollIntoView({ behavior: 'smooth' });
}

async function salvarAlteracoesExcel() {
    const linhas = document.querySelectorAll("#corpoTabelaEditor tr");
    let novosDados = [];

    linhas.forEach(tr => {
        let linhaObj = {
            'ENFERMARIA': tr.querySelector(".edit-enf").value,
            'LEITO': tr.querySelector(".edit-leito").value,
            'NOME DO PACIENTE': tr.querySelector(".edit-nome").value,
            'DIETA': tr.querySelector(".edit-dieta").value,
            'OBSERVAÇÕES': tr.querySelector(".edit-obs").value
        };
        novosDados.push(linhaObj);
    });

    if (confirm("Salvar alterações no Excel?")) {
        let resposta = await pywebview.api.salvar_dados_excel(novosDados);
        if (resposta.sucesso) {
            alert("✅ " + resposta.msg);
            carregarExcel();
        } else {
            alert("❌ " + resposta.msg);
        }
    }
}

// --- FILA ---
function adicionarFila(paciente) {
    filaImpressao.push(paciente);
    atualizarVisualFila();
}

function adicionarTodos() {
    if (confirm(`Adicionar ${todosPacientes.length} pacientes?`)) {
        filaImpressao = [...todosPacientes];
        atualizarVisualFila();
    }
}

function limparFila() {
    filaImpressao = [];
    atualizarVisualFila();
}

function atualizarVisualFila() {
    const ul = document.getElementById("listaFila");
    const contador = document.getElementById("contadorFila");
    ul.innerHTML = "";

    if (filaImpressao.length === 0) {
        ul.innerHTML = '<li class="empty-msg">Fila vazia.</li>';
    } else {
        filaImpressao.forEach((p, index) => {
            let li = document.createElement("li");
            li.innerHTML = `✅ <b>${p['LEITO']}</b> - ${p['NOME DO PACIENTE']}`;
            li.ondblclick = () => { filaImpressao.splice(index, 1); atualizarVisualFila(); };
            li.title = "Clique duplo para remover";
            ul.appendChild(li);
        });
    }
    contador.innerText = `${filaImpressao.length} etiquetas`;
}

// --- AÇÕES ---
async function imprimirFila() {
    if (filaImpressao.length === 0) {
        alert("Fila vazia!"); return;
    }
    let msg = await pywebview.api.imprimir_etiquetas(filaImpressao);
    if (msg !== "Cancelado.") alert(msg);
}

async function gerarRelatorioSimples() {
    let msg = await pywebview.api.gerar_relatorio_simples();
    if (msg !== "Cancelado.") alert(msg);
}

async function gerarMapaGeral() {
    let msg = await pywebview.api.gerar_mapa_geral();
    if (msg !== "Cancelado.") alert(msg);
}