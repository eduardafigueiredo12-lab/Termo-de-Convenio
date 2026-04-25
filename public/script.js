let sociosEncontrados = [];

function el(id){ return document.getElementById(id); }
function somenteNumeros(v){ return String(v || "").replace(/\D/g, ""); }

function mostrarLoading(titulo = "Gerando documento", texto = "Aguarde alguns instantes. Não feche esta página."){
  const overlay = el("loadingOverlay");
  if (!overlay) return;
  const title = el("loadingTitle");
  const text = el("loadingText");
  if (title) title.textContent = titulo;
  if (text) text.textContent = texto;
  overlay.classList.remove("hidden");
}

function esconderLoading(){
  const overlay = el("loadingOverlay");
  if (overlay) overlay.classList.add("hidden");
}


function atualizarCampoOutroCurso(){
  const box = el("outroCursoBox");
  const input = el("outro_curso");
  if (!box || !input) return;
  const mostrar = el("curso").value === "Outro";
  box.classList.toggle("hidden", !mostrar);
  input.required = mostrar;
  if (!mostrar) input.value = "";
}

el("curso").addEventListener("change", atualizarCampoOutroCurso);
document.addEventListener("DOMContentLoaded", atualizarCampoOutroCurso);

function preencher(dados){
  const mapa = ["razao_social","cep","endereco","numero","complemento","bairro","cidade","estado","telefone"];
  mapa.forEach(k => { if (dados[k]) el(k).value = dados[k]; });
  if (dados.email && !el("site").value) el("site").value = dados.email;

  sociosEncontrados = dados.socios || [];
  const sel = el("socios");
  sel.innerHTML = "";

  if (sociosEncontrados.length) {
    el("sociosBox").classList.remove("hidden");
    sociosEncontrados.forEach((s, i) => {
      const opt = document.createElement("option");
      opt.value = i;
      opt.textContent = `${s.nome} - ${s.cargo}`;
      sel.appendChild(opt);
    });
    selecionarSocio(0);
  } else {
    el("sociosBox").classList.add("hidden");
  }
}

function selecionarSocio(i){
  const s = sociosEncontrados[Number(i)];
  if (!s) return;
  el("representante").value = s.nome || "";
  el("cargo").value = s.cargo || "Sócio/Administrador";
}

el("socios").addEventListener("change", e => selecionarSocio(e.target.value));

el("outraPessoa").addEventListener("change", e => {
  el("avisoCartorio").classList.toggle("hidden", !e.target.checked);
  if (e.target.checked) {
    el("representante").value = "";
    el("cargo").value = "";
    el("representante").focus();
  } else if (sociosEncontrados.length) {
    selecionarSocio(el("socios").value || 0);
  }
});

el("buscar").addEventListener("click", async () => {
  const cnpj = somenteNumeros(el("cnpj").value);
  const msg = el("msg");
  if (cnpj.length !== 14) {
    msg.textContent = "Informe um CNPJ válido.";
    return;
  }
  msg.textContent = "Consultando CNPJ...";
  mostrarLoading("Consultando CNPJ", "Buscando os dados da unidade concedente. Aguarde...");
  try {
    const r = await fetch(`/api/cnpj/${cnpj}`);
    const dados = await r.json();
    if (!r.ok) throw new Error(dados.erro || "Erro ao consultar CNPJ.");
    preencher(dados);
    msg.textContent = "Dados localizados. Confira as informações antes de baixar o termo.";
  } catch (e) {
    msg.textContent = e.message || "Não foi possível consultar o CNPJ. Preencha manualmente.";
  } finally {
    esconderLoading();
  }
});

el("form").addEventListener("submit", async (e) => {
  e.preventDefault();

  const ids = [
    "tipo_estagio","curso","outro_curso","cnpj","razao_social","alvara","estimativa_vagas",
    "endereco","numero","complemento","bairro","cep","cidade","estado",
    "telefone","site","representante","cargo","email_assinatura"
  ];

  const dados = {};
  ids.forEach(id => dados[id] = el(id)?.value || "");

  mostrarLoading("Gerando documento", "Estamos preparando o Termo de Convênio. Aguarde alguns instantes.");

  try {
    const resp = await fetch("/api/gerar", {
      method:"POST",
      headers:{ "Content-Type":"application/json" },
      body: JSON.stringify(dados)
    });

    if (!resp.ok) {
      const erro = await resp.json().catch(()=>({erro:"Erro ao gerar documento."}));
      alert(erro.erro || "Erro ao gerar documento.");
      return;
    }

    const blob = await resp.blob();
    const cd = resp.headers.get("Content-Disposition") || "";
    let filename = "TERMO DE CONVÊNIO.docx";
    const match = cd.match(/filename\*=UTF-8''([^;]+)/);
    if (match) filename = decodeURIComponent(match[1]);
    const a = document.createElement("a");
    const url = URL.createObjectURL(blob);
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  } catch (e) {
    alert("Erro ao gerar documento. Tente novamente em alguns instantes.");
  } finally {
    esconderLoading();
  }
});
