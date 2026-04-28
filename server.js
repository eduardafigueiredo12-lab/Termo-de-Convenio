
const express = require("express");
const cors = require("cors");
const path = require("path");
const fs = require("fs");
const crypto = require("crypto");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const app = express();
app.use(cors());
app.use(express.json({ limit: "2mb" }));
app.use(express.static(path.join(__dirname, "public")));

const PORT = process.env.PORT || 3000;
const SENHA_PROTECAO_DOCUMENTO = process.env.DOCX_PROTECTION_PASSWORD || "convenios";
const PROTECAO_SPIN_COUNT = 100000;

const cursosObrigatorios = [
  "Biomedicina",
  "Radiologia",
  "Fisioterapia",
  "Fonoaudiologia",
  "Farmácia",
  "Terapia Ocupacional",
  "Técnicas Oftálmicas",
  "Pós em Procedimentos Injetáveis",
  "Nutrição",
  "Nutrição - Educação Física"
];

function apenasNumeros(v) {
  return String(v || "").replace(/\D/g, "");
}

function limparNomeArquivo(nome) {
  return String(nome || "LOCAL")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function dataExtensoHoje() {
  const meses = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];
  const d = new Date();
  return `${d.getDate()} de ${meses[d.getMonth()]} de ${d.getFullYear()}`;
}

function normalizarTexto(v) {
  return String(v || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

const cursosDoTermo = [
  "Biomedicina",
  "Farmácia",
  "Fonoaudiologia",
  "Fisioterapia",
  "Nutrição",
  "Terapia Ocupacional",
  "Radiologia",
  "Técnicas Oftálmicas",
  "Engenharias",
  "Licenciaturas",
  "Técnico em Enfermagem",
  "Técnico em Transações Imobiliárias",
  "Pós-Graduação em Biomedicina Estética"
];

function cursoExisteNaLista(curso) {
  const c = normalizarTexto(curso);
  return cursosDoTermo.some(nome => normalizarTexto(nome) === c);
}

function checkCurso(curso, nome) {
  return normalizarTexto(curso) === normalizarTexto(nome) ? "X" : " ";
}

function montarContatoEmpresa(d) {
  return [
    d.responsavel_estagios ? `Responsável pelo estágio ou setor responsável: ${d.responsavel_estagios}` : "",
    d.contato_responsavel ? `Telefone de contato direto: ${d.contato_responsavel}` : "",
    d.site ? `E-mail: ${d.site}` : ""
  ].filter(Boolean).join("\n");
}

function corrigirCompatibilidadeDocumento(xml) {
  const rootEnd = xml.indexOf(">");
  if (rootEnd === -1) return xml;

  const root = xml.slice(0, rootEnd + 1);
  const ignorable = root.match(/\s([\w.-]+:Ignorable)="[^"]*"/);
  if (!ignorable) return xml;

  const namespaces = [...root.matchAll(/\sxmlns:([\w.-]+)="([^"]+)"/g)];
  const prefixesIgnoraveis = namespaces
    .filter(([, , uri]) => (
      uri.includes("schemas.microsoft.com/office/word/") ||
      uri.includes("schemas.microsoft.com/office/drawing/") ||
      uri.includes("schemas.microsoft.com/office/powerpoint/")
    ))
    .map(([, prefix]) => prefix);

  const novoRoot = prefixesIgnoraveis.length
    ? root.replace(/\s[\w.-]+:Ignorable="[^"]*"/, ` ${ignorable[1]}="${prefixesIgnoraveis.join(" ")}"`)
    : root.replace(/\s[\w.-]+:Ignorable="[^"]*"/, "");

  return `${novoRoot}${xml.slice(rootEnd + 1)}`;
}

function removerAtributosInvalidosTblLook(xml) {
  const atributos = "firstRow|lastRow|firstColumn|lastColumn|noHBand|noVBand";
  const padrao = new RegExp(`(<w:tblLook\\b[^>]*?)\\s+w:(${atributos})="[^"]*"`, "g");
  let anterior;
  do {
    anterior = xml;
    xml = xml.replace(padrao, "$1");
  } while (xml !== anterior);
  return xml;
}

function corrigirSettings(settings) {
  return settings.replace(
    /<w:compatSetting\b(?=[^>]*w:name="useWord2013TrackBottomHyphenation")[^>]*\/>/g,
    ""
  );
}

function gerarHashProtecaoDocumento(senha, salt) {
  let hash = crypto
    .createHash("sha1")
    .update(Buffer.concat([Buffer.from(String(senha || ""), "utf16le"), salt]))
    .digest();

  for (let i = 0; i < PROTECAO_SPIN_COUNT; i++) {
    const contador = Buffer.alloc(4);
    contador.writeUInt32LE(i, 0);
    hash = crypto.createHash("sha1").update(Buffer.concat([hash, contador])).digest();
  }

  return hash.toString("base64");
}

function aplicarProtecaoSomenteLeitura(settings) {
  const salt = crypto.randomBytes(16);
  const protection = [
    '<w:documentProtection',
    ' w:edit="readOnly"',
    ' w:enforcement="1"',
    ' w:formatting="0"',
    ' w:cryptProviderType="rsaFull"',
    ' w:cryptAlgorithmClass="hash"',
    ' w:cryptAlgorithmType="typeAny"',
    ' w:cryptAlgorithmSid="4"',
    ` w:cryptSpinCount="${PROTECAO_SPIN_COUNT}"`,
    ` w:hash="${gerarHashProtecaoDocumento(SENHA_PROTECAO_DOCUMENTO, salt)}"`,
    ` w:salt="${salt.toString("base64")}"`,
    '/>'
  ].join("");

  settings = settings.replace(/<w:documentProtection\b[^>]*\/>/g, "");
  settings = settings.replace(/<w:documentProtection\b[\s\S]*?<\/w:documentProtection>/g, "");

  const inserirAntes = [
    /<w:defaultTabStop\b[^>]*\/>/,
    /<w:hyphenationZone\b[^>]*\/>/,
    /<w:characterSpacingControl\b[^>]*\/>/,
    /<w:compat\b[\s\S]*?<\/w:compat>/,
    /<w:rsids\b[\s\S]*?<\/w:rsids>/
  ];

  for (const padrao of inserirAntes) {
    if (padrao.test(settings)) {
      return settings.replace(padrao, `${protection}$&`);
    }
  }

  return settings.replace(/(<w:settings\b[^>]*>)/, `$1${protection}`);
}

function prepararDocumentoWord(buffer) {
  const zip = new PizZip(buffer);

  let settings = zip.file("word/settings.xml")
    ? zip.file("word/settings.xml").asText()
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:settings>`;
  settings = corrigirSettings(settings);

  // Aplica restricao de edicao do Word com senha para desbloqueio.
  settings = aplicarProtecaoSomenteLeitura(settings);
  zip.file("word/settings.xml", settings);

  // Remove marcas de permissão herdadas dos modelos e corrige compatibilidade OpenXML.
  let xml = zip.file("word/document.xml").asText();
  xml = corrigirCompatibilidadeDocumento(xml);
  xml = removerAtributosInvalidosTblLook(xml);
  xml = xml.replace(/<w:permStart\b[^>]*\/>/g, "");
  xml = xml.replace(/<w:permEnd\b[^>]*\/>/g, "");
  zip.file("word/document.xml", xml);

  return zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}

async function fetchJson(url) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 12000);
  try {
    const resp = await fetch(url, { signal: controller.signal, headers: { "accept": "application/json" } });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    return await resp.json();
  } finally {
    clearTimeout(timer);
  }
}

function mapSocios(dados) {
  const qsa = dados.qsa || dados.socios || [];
  if (!Array.isArray(qsa)) return [];
  return qsa.map(s => ({
    nome: s.nome_socio || s.nome || s.nomeSocio || "",
    cargo: s.qualificacao_socio || s.qualificacao || s.cargo || "Sócio/Administrador"
  })).filter(s => s.nome);
}

function mapDadosEmpresa(dados) {
  const logradouro = [dados.descricao_tipo_de_logradouro, dados.logradouro].filter(Boolean).join(" ").trim() || dados.logradouro || "";
  return {
    razao_social: dados.razao_social || dados.nome || "",
    nome_fantasia: dados.nome_fantasia || dados.fantasia || "",
    cnpj: dados.cnpj || "",
    cep: dados.cep || "",
    endereco: logradouro,
    numero: dados.numero || "",
    complemento: dados.complemento || "",
    bairro: dados.bairro || "",
    cidade: dados.municipio || dados.cidade || "",
    estado: dados.uf || dados.estado || "",
    telefone: [dados.ddd_telefone_1, dados.ddd_telefone_2].filter(Boolean)[0] || "",
    email: dados.email || "",
    socios: mapSocios(dados)
  };
}

app.get("/api/cnpj/:cnpj", async (req, res) => {
  const cnpj = apenasNumeros(req.params.cnpj);
  if (cnpj.length !== 14) return res.status(400).json({ erro: "CNPJ inválido." });

  const urls = [
    `https://minhareceita.org/${cnpj}`,
    `https://brasilapi.com.br/api/cnpj/v1/${cnpj}`
  ];

  for (const url of urls) {
    try {
      const dados = await fetchJson(url);
      return res.json(mapDadosEmpresa(dados));
    } catch (e) {
      // tenta próxima API
    }
  }
  return res.status(502).json({ erro: "Não foi possível consultar o CNPJ no momento. Preencha manualmente." });
});

app.post("/api/gerar", (req, res) => {
  try {
    const d = req.body || {};
    const tipo = d.tipo_estagio;

    let modelo = "simples.docx";
    if (tipo === "Estágio remunerado") {
      modelo = "remunerado.docx";
    } else if (tipo === "Estágio obrigatório" && cursosObrigatorios.includes(d.curso)) {
      modelo = "contrapartidas.docx";
    } else {
      modelo = "simples.docx";
    }

    const templatePath = path.join(__dirname, "templates", modelo);
    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      nullGetter: () => ""
    });

    const curso = d.curso === "Outro" ? (d.outro_curso || "Outro") : (d.curso || "");
    const outros = cursoExisteNaLista(curso) ? "" : curso;
    const data = {
      razao_social: d.razao_social || "",
      cnpj: d.cnpj || "",
      alvara: d.alvara || "",
      area_atuacao: "",
      outros: outros || "",
      estimativa_vagas: d.estimativa_vagas || "",
      endereco: d.endereco || "",
      numero: d.numero || "",
      complemento: d.complemento || "",
      bairro: d.bairro || "",
      cep: d.cep || "",
      cidade: d.cidade || "",
      estado: d.estado || "",
      telefone: d.contato_responsavel || d.telefone || "",
      site: montarContatoEmpresa(d),
      responsavel_estagios: d.responsavel_estagios || "",
      contato_responsavel: d.contato_responsavel || "",
      representante: d.representante || "",
      cargo: d.cargo || "",
      email_assinatura: d.email_assinatura || "",
      data_extenso: dataExtensoHoje(),

      chk_biomedicina: checkCurso(curso, "Biomedicina"),
      chk_farmacia: checkCurso(curso, "Farmácia"),
      chk_fonoaudiologia: checkCurso(curso, "Fonoaudiologia"),
      chk_fisioterapia: checkCurso(curso, "Fisioterapia"),
      chk_nutricao: checkCurso(curso, "Nutrição"),
      chk_terapia_ocupacional: checkCurso(curso, "Terapia Ocupacional"),
      chk_radiologia: checkCurso(curso, "Radiologia"),
      chk_tecnicas_oftalmicas: checkCurso(curso, "Técnicas Oftálmicas"),
      chk_engenharias: checkCurso(curso, "Engenharias"),
      chk_licenciaturas: checkCurso(curso, "Licenciaturas"),
      chk_tecnico_enfermagem: checkCurso(curso, "Técnico em Enfermagem"),
      chk_tecnico_transacoes_imobiliarias: checkCurso(curso, "Técnico em Transações Imobiliárias"),
      chk_pos_biomedicina_estetica: checkCurso(curso, "Pós-Graduação em Biomedicina Estética")
    };

    doc.render(data);

    const buffer = doc.getZip().generate({
      type: "nodebuffer",
      compression: "DEFLATE"
    });

    const documentoWord = prepararDocumentoWord(buffer);

    const nomeLocal = limparNomeArquivo(d.razao_social || d.nome_fantasia || "LOCAL");
    const filename = `${nomeLocal} - TERMO DE CONVENIO.docx`;

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(filename)}"; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(documentoWord);
  } catch (e) {
    console.error(e);
    res.status(500).json({ erro: "Erro ao gerar arquivo.", detalhe: e.message });
  }
});

app.listen(PORT, () => console.log(`Servidor rodando em http://localhost:${PORT}`));
