// Tela do ERP fictício. JavaScript puro, sem dependências.
(function () {
  "use strict";

  var cfg = { hoje: "2026-10-14", contas: [], formas: [] };
  var estado = { pagina: 1, total: 0, linhas: [], aberto: null };
  var $ = function (id) { return document.getElementById(id); };

  function brl(v) {
    return "R$ " + Number(v || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  function num(texto) {
    var t = String(texto || "").replace(/\./g, "").replace(",", ".").trim();
    var n = parseFloat(t);
    return isNaN(n) ? 0 : n;
  }
  function virgula(v) { return Number(v).toFixed(2).replace(".", ","); }
  function dataBr(iso) { return iso ? iso.slice(8, 10) + "/" + iso.slice(5, 7) + "/" + iso.slice(0, 4) : "—"; }
  function esc(t) { return String(t == null ? "" : t).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }

  function postar(url, corpo) {
    return fetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(corpo) })
      .then(function (r) { return r.json().then(function (j) { return { ok: r.ok, corpo: j }; }); });
  }

  // ------------------------------------------------------------ listagem ---
  function periodo() {
    var p = $("filtro-periodo").value, a = cfg.hoje.slice(0, 7);
    if (p === "todo") return [null, null];
    var ano = Number(cfg.hoje.slice(0, 4)), mes = Number(cfg.hoje.slice(5, 7));
    if (p === "passado") { mes -= 1; if (mes === 0) { mes = 12; ano -= 1; } }
    var m = (mes < 10 ? "0" : "") + mes, ult = new Date(ano, mes, 0).getDate();
    return [ano + "-" + m + "-01", ano + "-" + m + "-" + ult];
  }

  function carregar() {
    var per = periodo();
    return postar("/api/transacoes/listar", {
      document_number: $("filtro-doc").value.trim(), page: estado.pagina,
      page_size: Number($("itens-pagina").value), start_date: per[0], end_date: per[1]
    }).then(function (r) {
      estado.linhas = r.corpo.data || [];
      estado.total = r.corpo.total_count || 0;
      desenhar();
    });
  }

  function situacao(t) {
    var mapa = { PAID: ["Pago", "pago"], PENDING: ["Em aberto", "aberto"], OVERDUE: ["Vencido", "vencido"], PARTIAL: ["Parcial", "parcial"] };
    var s = mapa[t.status] || [t.status, ""];
    return '<span class="selo ' + s[1] + '">' + s[0] + "</span>";
  }

  function desenhar() {
    var corpo = estado.linhas.map(function (t) {
      var aberto = t.status === "PENDING" || t.status === "OVERDUE";
      return "<tr data-code=\"" + t.code + "\">" +
        "<td class=\"id\">" + t.code + "</td><td>" + esc(t.document_number) + "</td><td>" + esc(t.party_name) + "</td>" +
        "<td class=\"desc\">" + esc(t.description) + "</td><td class=\"num\">" + brl(t.amount) + "</td>" +
        "<td>" + dataBr(t.due_date) + "</td><td>" + situacao(t) + "</td><td>" + dataBr(t.payment_date) + "</td>" +
        "<td class=\"acoes\">" +
        (aberto ? "<button class=\"icone baixa\" aria-label=\"Dar baixa\" title=\"Dar baixa\">✓</button>" : "") +
        "<button class=\"icone lixo\" aria-label=\"Excluir lançamento\" title=\"Excluir (desativado na demonstração)\" data-excluir>🗑</button>" +
        "</td></tr>";
    }).join("");
    $("linhas").innerHTML = corpo || "<tr><td colspan=\"9\" class=\"vazio\">Nenhum lançamento encontrado.</td></tr>";
    var porPag = Number($("itens-pagina").value), paginas = Math.max(1, Math.ceil(estado.total / porPag));
    $("resumo").textContent = estado.total + " lançamento" + (estado.total === 1 ? "" : "s");
    $("pagina").textContent = "Página " + estado.pagina + " de " + paginas;
    $("anterior").disabled = estado.pagina <= 1;
    $("proxima").disabled = estado.pagina >= paginas;
  }

  var espera = null;
  function agendar() {
    clearTimeout(espera);
    espera = setTimeout(function () { estado.pagina = 1; carregar(); }, 200);
  }

  // ------------------------------------------------------------- gaveta ---
  function fechar() { $("fundo").hidden = true; estado.aberto = null; fecharLista(); }

  function abrir(t) {
    estado.aberto = t;
    $("resumo-lanc").innerHTML =
      "<div><strong>Descrição:</strong> " + esc(t.description) + "</div>" +
      "<div><strong>Valor Total:</strong> " + brl(t.amount) + "</div>" +
      "<div><strong>Total pago:</strong> " + brl(t.paid_amount) + "</div>" +
      "<div><strong>Saldo pendente:</strong> " + brl(t.remaining_amount) + "</div>" +
      "<div><strong>Tipo:</strong> Despesa</div>";
    $("v-pagamento").value = virgula(t.remaining_amount);
    ["v-desconto", "v-juros", "v-multa", "v-tarifa", "v-outras"].forEach(function (id) { $(id).value = "0,00"; });
    $("v-data").value = cfg.hoje;
    $("periodo-aviso").hidden = true;
    definirCombo("cb-conta", "Selecione"); definirCombo("cb-forma", "Selecione");
    recalcular();
    $("fundo").hidden = false;
  }

  function recalcular() {
    var t = estado.aberto; if (!t) return;
    var valor = num($("v-pagamento").value);
    var liquido = valor - num($("v-desconto").value) + num($("v-juros").value) + num($("v-multa").value) +
      num($("v-tarifa").value) + num($("v-outras").value);
    $("v-liquido").value = virgula(liquido);
    $("saldo-atual").textContent = "Saldo atual: " + brl(t.remaining_amount);
    $("saldo-apos").textContent = "Saldo após este registro: " + brl(Math.max(0, t.remaining_amount - valor));
    $("liquido-nota").textContent = "Este será o valor considerado no pagamento: " + brl(liquido);
    var ok = valor > 0 && liquido > 0 && $("cb-conta").dataset.valor && $("cb-forma").dataset.valor && $("periodo-aviso").hidden;
    $("confirmar").disabled = !ok;
  }

  function verificarPeriodo() {
    return fetch("/api/periodo-fechado?data=" + encodeURIComponent($("v-data").value))
      .then(function (r) { return r.json(); })
      .then(function (j) { $("periodo-aviso").hidden = !j.closed; recalcular(); });
  }

  // ----------------------------------------------------- listas suspensas ---
  var listaAtual = null;
  function fecharLista() {
    if (listaAtual) { listaAtual.remove(); listaAtual = null; }
    ["cb-conta", "cb-forma"].forEach(function (id) { $(id).setAttribute("aria-expanded", "false"); });
  }
  function definirCombo(id, texto, valor) {
    $(id).textContent = texto; $(id).dataset.valor = valor || "";
  }
  function abrirLista(botao, opcoes) {
    fecharLista();
    var r = botao.getBoundingClientRect();
    var lista = document.createElement("div");
    lista.className = "lista"; lista.setAttribute("role", "listbox");
    lista.style.left = r.left + "px"; lista.style.top = (r.bottom + 4) + "px"; lista.style.width = r.width + "px";
    opcoes.forEach(function (o) {
      var item = document.createElement("div");
      item.setAttribute("role", "option"); item.textContent = o; item.tabIndex = -1;
      item.addEventListener("click", function () { definirCombo(botao.id, o, o); fecharLista(); recalcular(); });
      lista.appendChild(item);
    });
    document.body.appendChild(lista); listaAtual = lista;
    // sem espaço embaixo (janela pequena): abre para cima
    var alt = lista.offsetHeight;
    if (r.bottom + alt + 8 > window.innerHeight && r.top - alt - 4 > 8) {
      lista.style.top = (r.top - alt - 4) + "px";
    }
    botao.setAttribute("aria-expanded", "true");
  }

  function aviso(texto) {
    var el = $("aviso-topo"); el.textContent = texto; el.hidden = false;
    setTimeout(function () { el.hidden = true; }, 2600);
  }

  // ----------------------------------------------------------- eventos ---
  document.addEventListener("DOMContentLoaded", function () {
    fetch("/api/config").then(function (r) { return r.json(); }).then(function (c) {
      cfg = c; carregar();
    });
    $("filtro-doc").addEventListener("input", agendar);
    $("filtro-periodo").addEventListener("change", function () { estado.pagina = 1; carregar(); });
    $("itens-pagina").addEventListener("change", function () { estado.pagina = 1; carregar(); });
    $("anterior").addEventListener("click", function () { if (estado.pagina > 1) { estado.pagina--; carregar(); } });
    $("proxima").addEventListener("click", function () { estado.pagina++; carregar(); });

    $("linhas").addEventListener("click", function (e) {
      var btn = e.target.closest("button"); if (!btn) return;
      var tr = btn.closest("tr"), code = Number(tr.dataset.code);
      var t = estado.linhas.filter(function (x) { return x.code === code; })[0];
      if (btn.hasAttribute("data-excluir")) { aviso("Exclusão desativada na demonstração."); return; }
      if (t) abrir(t);
    });

    ["v-pagamento", "v-desconto", "v-juros", "v-multa", "v-tarifa", "v-outras"].forEach(function (id) {
      $(id).addEventListener("input", recalcular);
    });
    $("v-data").addEventListener("change", verificarPeriodo);
    $("cb-conta").addEventListener("click", function (e) { e.stopPropagation(); abrirLista($("cb-conta"), cfg.contas); });
    $("cb-forma").addEventListener("click", function (e) { e.stopPropagation(); abrirLista($("cb-forma"), cfg.formas); });
    document.addEventListener("click", function (e) { if (listaAtual && !listaAtual.contains(e.target)) fecharLista(); });
    document.addEventListener("keydown", function (e) { if (e.key === "Escape") { if (listaAtual) fecharLista(); else fechar(); } });
    $("cancelar").addEventListener("click", fechar);

    $("confirmar").addEventListener("click", function () {
      var t = estado.aberto; if (!t) return;
      $("confirmar").disabled = true;
      postar("/api/transacoes/pagar", {
        code: t.code, amount: num($("v-pagamento").value), discount: num($("v-desconto").value),
        interest: num($("v-juros").value), fine: num($("v-multa").value), fee: num($("v-tarifa").value),
        other: num($("v-outras").value), date: $("v-data").value,
        account: $("cb-conta").dataset.valor, method: $("cb-forma").dataset.valor
      }).then(function (r) {
        if (r.ok) { fechar(); aviso("Pagamento registrado."); carregar(); }
        else { aviso("Não foi possível pagar: " + (r.corpo.erro || "erro")); recalcular(); }
      });
    });
  });
})();
