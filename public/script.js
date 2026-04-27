let sociosEncontrados = [];

function el(id){ return document.getElementById(id); }
function somenteNumeros(v){ return String(v || "").replace(/\D/g, ""); }
function tipoUnidadeSelecionado(){
  return document.querySelector("input[name='tipo_unidade']:checked")?.value || "cnpj";
}

function mascararCpf(v){
  const n = somenteNumeros(v).slice(0, 11);
  return n
    .replace(/(\d{3})(\d)/, "$1.$2")
    .replace(/(\d{3})(\d)/, "$1.$2")
    .replace(/(\d{3})(\d{1,2})$/, "$1-$2");
}

function mascararCnpj(v){
  const n = somenteNumeros(v).slice(0, 14);
  return n
    .replace(/^(\d{2})(\d)/, "$1.$2")
    .replace(/^(\d{2})\.(\d{3})(\d)/, "$1.$2.$3")
    .replace(/\.(\d{3})(\d)/, ".$1/$2")
    .replace(/(\d{4})(\d)/, "$1-$2");
}

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

function atualizarTipoUnidade(){
  const profissional = tipoUnidadeSelecionado() === "cpf";
  const cnpjBox = el("cnpjBox");
  const cpfBox = el("cpfBox");
  const cnpj = el("cnpj");
  const cpf = el("cpf");
  const buscar = el("buscar");
  const razaoLabel = el("razaoLabel");
  const alvaraLabel = el("alvaraLabel");
  const msg = el("msg");

  cnpjBox?.classList.toggle("hidden", profissional);
  cpfBox?.classList.toggle("hidden", !profissional);
  if (cnpj) cnpj.required = !profissional;
  if (cpf) cpf.required = profissional;
  if (buscar) buscar.disabled = profissional;
  if (razaoLabel) razaoLabel.textContent = profissional ? "Nome completo do profissional *" : "Razão social *";
  if (alvaraLabel) alvaraLabel.textContent = profissional ? "Registro profissional / alvará (se houver)" : "Alvará de Funcionamento/Sanitário";
  if (msg) msg.textContent = "";

  if (profissional) {
    sociosEncontrados = [];
    el("sociosBox")?.classList.add("hidden");
    el("avisoCartorio")?.classList.add("hidden");
    if (el("outraPessoa")) el("outraPessoa").checked = false;
    if (!el("cargo").value || el("cargo").value === "Sócio/Administrador") {
      el("cargo").value = "Profissional autônomo";
    }
  } else if (el("cargo").value === "Profissional autônomo") {
    el("cargo").value = "";
  }
}

document.querySelectorAll("input[name='tipo_unidade']").forEach(input => {
  input.addEventListener("change", atualizarTipoUnidade);
});
document.addEventListener("DOMContentLoaded", atualizarTipoUnidade);

el("cnpj").addEventListener("input", e => {
  e.target.value = mascararCnpj(e.target.value);
});

el("cpf").addEventListener("input", e => {
  e.target.value = mascararCpf(e.target.value);
});

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
  if (tipoUnidadeSelecionado() !== "cnpj") return;
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

  const tipoUnidade = tipoUnidadeSelecionado();
  if (tipoUnidade === "cpf" && somenteNumeros(el("cpf").value).length !== 11) {
    alert("Informe um CPF válido.");
    el("cpf").focus();
    return;
  }

  const ids = [
    "tipo_estagio","curso","outro_curso","cnpj","cpf","razao_social","alvara","estimativa_vagas",
    "endereco","numero","complemento","bairro","cep","cidade","estado",
    "telefone","site","representante","cargo","email_assinatura"
  ];

  const dados = {};
  ids.forEach(id => dados[id] = el(id)?.value || "");
  dados.tipo_unidade = tipoUnidade;
  if (tipoUnidade === "cpf") {
    dados.cnpj = dados.cpf;
  }

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
