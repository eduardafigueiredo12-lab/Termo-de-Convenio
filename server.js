
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

const HIGH_ORDER_WORDS = [
  [0xE1, 0xF0], [0x1D, 0x0F], [0xCC, 0x9C], [0x84, 0xC0], [0x11, 0x0C],
  [0x0E, 0x10], [0xF1, 0xCE], [0x31, 0x3E], [0x18, 0x72], [0xE1, 0x39],
  [0xD4, 0x0F], [0x84, 0xF9], [0x28, 0x0C], [0xA9, 0x6A], [0x4E, 0xC3]
];

const ENCRYPTION_MATRIX = [
  [[0xAE,0xFC],[0x4D,0xD9],[0x9B,0xB2],[0x27,0x45],[0x4E,0x8A],[0x9D,0x14],[0x2A,0x09]],
  [[0x7B,0x61],[0xF6,0xC2],[0xFD,0xA5],[0xEB,0x6B],[0xC6,0xF7],[0x9D,0xCF],[0x2B,0xBF]],
  [[0x45,0x63],[0x8A,0xC6],[0x05,0xAD],[0x0B,0x5A],[0x16,0xB4],[0x2D,0x68],[0x5A,0xD0]],
  [[0x03,0x75],[0x06,0xEA],[0x0D,0xD4],[0x1B,0xA8],[0x37,0x50],[0x6E,0xA0],[0xDD,0x40]],
  [[0xD8,0x49],[0xA0,0xB3],[0x51,0x47],[0xA2,0x8E],[0x55,0x3D],[0xAA,0x7A],[0x44,0xD5]],
  [[0x6F,0x45],[0xDE,0x8A],[0xAD,0x35],[0x4A,0x4B],[0x94,0x96],[0x39,0x0D],[0x72,0x1A]],
  [[0xEB,0x23],[0xC6,0x67],[0x9C,0xEF],[0x29,0xFF],[0x53,0xFE],[0xA7,0xFC],[0x5F,0xD9]],
  [[0x47,0xD3],[0x8F,0xA6],[0x0F,0x6D],[0x1E,0xDA],[0x3D,0xB4],[0x7B,0x68],[0xF6,0xD0]],
  [[0xB8,0x61],[0x60,0xE3],[0xC1,0xC6],[0x93,0xAD],[0x37,0x7B],[0x6E,0xF6],[0xDD,0xEC]],
  [[0x45,0xA0],[0x8B,0x40],[0x06,0xA1],[0x0D,0x42],[0x1A,0x84],[0x35,0x08],[0x6A,0x10]],
  [[0xAA,0x51],[0x44,0x83],[0x89,0x06],[0x02,0x2D],[0x04,0x5A],[0x08,0xB4],[0x11,0x68]],
  [[0x76,0xB4],[0xED,0x68],[0xCA,0xF1],[0x85,0xC3],[0x1B,0xA7],[0x37,0x4E],[0x6E,0x9C]],
  [[0x37,0x30],[0x6E,0x60],[0xDC,0xC0],[0xA9,0xA1],[0x43,0x63],[0x86,0xC6],[0x1D,0xAD]],
  [[0x33,0x31],[0x66,0x62],[0xCC,0xC4],[0x89,0xA9],[0x03,0x73],[0x06,0xE6],[0x0D,0xCC]],
  [[0x10,0x21],[0x20,0x42],[0x40,0x84],[0x81,0x08],[0x12,0x31],[0x24,0x62],[0x48,0xC4]]
];

function senhaParaBytesLegados(senha) {
  const texto = String(senha || "").slice(0, 15);
  const bytes = Buffer.from(texto, "utf16le");
  const resultado = [];
  for (let i = 0; i < texto.length; i++) {
    resultado.push(bytes[i * 2] || bytes[i * 2 + 1] || 0);
  }
  return resultado;
}

function prepararSenhaParaProtecao(senha) {
  const bytes = senhaParaBytesLegados(senha);
  const tamanho = bytes.length;
  if (!tamanho) return Buffer.from([0, 0, 0, 0]);

  const high = HIGH_ORDER_WORDS[tamanho - 1].slice();
  for (let i = 0; i < tamanho; i++) {
    const linha = 15 - tamanho + i;
    for (let bit = 0; bit < 7; bit++) {
      if (bytes[i] & (1 << bit)) {
        high[0] ^= ENCRYPTION_MATRIX[linha][bit][0];
        high[1] ^= ENCRYPTION_MATRIX[linha][bit][1];
      }
    }
  }

  let low = 0;
  for (let i = tamanho - 1; i >= 0; i--) {
    low = ((((low >> 14) & 0x0001) | ((low << 1) & 0x7fff)) ^ bytes[i]) & 0xffff;
  }
  low = ((((low >> 14) & 0x0001) | ((low << 1) & 0x7fff)) ^ tamanho ^ 0xce4b) & 0xffff;

  return Buffer.from([low & 0xff, (low >> 8) & 0xff, high[0], high[1]]);
}

function gerarHashProtecaoDocumento(senha, salt) {
  let hash = crypto
    .createHash("sha512")
    .update(Buffer.concat([salt, prepararSenhaParaProtecao(senha)]))
    .digest();

  for (let i = 0; i < PROTECAO_SPIN_COUNT; i++) {
    const contador = Buffer.alloc(4);
    contador.writeUInt32LE(i, 0);
    hash = crypto.createHash("sha512").update(Buffer.concat([hash, contador])).digest();
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
    ' w:cryptProviderType="rsaAES"',
    ' w:cryptAlgorithmClass="hash"',
    ' w:cryptAlgorithmType="typeAny"',
    ' w:cryptAlgorithmSid="14"',
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
