
const express = require("express");
const cors = require("cors");
const path = require("path");
const fs = require("fs");
const crypto = require("crypto");
const https = require("https");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const PDFDocument = require("pdfkit");

const app = express();
const PORT = process.env.PORT || 3000;
const isProduction = process.env.NODE_ENV === "production";
const JSON_BODY_LIMIT = process.env.JSON_BODY_LIMIT || "100kb";
const MAX_EXTERNAL_JSON_BYTES = numeroInteiroPositivo(process.env.MAX_EXTERNAL_JSON_BYTES, 256 * 1024);
const SENHA_PROTECAO_WORD = (() => {
  const senha = process.env.WORD_PROTECTION_PASSWORD;
  if (!senha) {
    if (isProduction) throw new Error("[SEGURANÇA] WORD_PROTECTION_PASSWORD não definida. Defina a variável de ambiente antes de iniciar em produção.");
    console.warn("[AVISO] WORD_PROTECTION_PASSWORD não definida. Usando senha de desenvolvimento — não use em produção.");
    return "dev-local-convenios-nao-usar-producao";
  }
  return senha;
})();
const WORD_PROTECTION_SPIN_COUNT = 100000;

if (process.env.TRUST_PROXY) {
  const val = process.env.TRUST_PROXY;
  if (val === "true") {
    const msg = "TRUST_PROXY=true confia em todos os proxies e permite spoofing de IP no rate limiter. Use um número (ex: TRUST_PROXY=1 para um único proxy reverso).";
    if (isProduction) throw new Error(`[SEGURANÇA] ${msg}`);
    console.warn(`[AVISO SEGURANÇA] ${msg}`);
  }
  app.set("trust proxy", isNaN(Number(val)) ? val : Number(val));
}

app.disable("x-powered-by");
app.use(aplicarHeadersSeguranca);
app.use(rejeitarOrigemNaoPermitida);
app.use(cors(criarCorsOptions));
app.use("/api", express.json({ limit: JSON_BODY_LIMIT, strict: true, type: "application/json" }));
app.use(tratarErroJson);
app.use(express.static(path.join(__dirname, "public"), {
  dotfiles: "ignore",
  index: "index.html"
}));

function numeroInteiroPositivo(valor, fallback) {
  const n = Number(valor);
  return Number.isInteger(n) && n > 0 ? n : fallback;
}

function origemHostIgualAoRequest(origin, req) {
  try {
    return new URL(origin).host === req.get("host");
  } catch (e) {
    return false;
  }
}

function origemLocalDesenvolvimento(origin) {
  if (isProduction) return false;
  return origin === "null" || /^https?:\/\/(localhost|127\.0\.0\.1|\[::1\])(?::\d+)?$/i.test(origin);
}

function origensPermitidasConfiguradas() {
  return String(process.env.CORS_ORIGINS || process.env.CORS_ORIGIN || "")
    .split(",")
    .map(origin => origin.trim())
    .filter(Boolean);
}

function origemPermitida(origin, req) {
  if (!origin) return true;
  if (origemHostIgualAoRequest(origin, req)) return true;
  if (origemLocalDesenvolvimento(origin)) return true;
  return origensPermitidasConfiguradas().includes(origin);
}

function rejeitarOrigemNaoPermitida(req, res, next) {
  const origin = req.get("Origin");
  if (!origemPermitida(origin, req)) {
    return res.status(403).json({ erro: "Origem não permitida." });
  }
  next();
}

function criarCorsOptions(req, callback) {
  const origin = req.get("Origin");
  callback(null, {
    origin: origin && origemPermitida(origin, req) ? origin : false,
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type"],
    exposedHeaders: ["Content-Disposition"],
    maxAge: 600
  });
}

function aplicarHeadersSeguranca(req, res, next) {
  const csp = [
    "default-src 'self'",
    "base-uri 'self'",
    "object-src 'none'",
    "script-src 'self'",
    "style-src 'self' https://fonts.googleapis.com",
    "font-src 'self' https://fonts.gstatic.com",
    "img-src 'self' data:",
    "connect-src 'self'",
    "form-action 'self'",
    "frame-ancestors 'none'"
  ];

  if (isProduction) csp.push("upgrade-insecure-requests");

  res.setHeader("Content-Security-Policy", csp.join("; "));
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.setHeader("X-Frame-Options", "DENY");
  res.setHeader("Referrer-Policy", "strict-origin-when-cross-origin");
  res.setHeader("Permissions-Policy", "camera=(), microphone=(), geolocation=(), payment=()");
  res.setHeader("Cross-Origin-Opener-Policy", "same-origin");
  res.setHeader("Cross-Origin-Resource-Policy", "same-origin");
  if (isProduction) {
    res.setHeader("Strict-Transport-Security", "max-age=31536000; includeSubDomains");
  }
  next();
}

function tratarErroJson(err, req, res, next) {
  if (!err) return next();
  if (err.type === "entity.too.large") {
    return res.status(413).json({ erro: "Payload muito grande." });
  }
  if (err instanceof SyntaxError && "body" in err) {
    return res.status(400).json({ erro: "JSON inválido." });
  }
  next(err);
}

function criarRateLimiter({ janelaMs, maximo, escopo }) {
  const acessos = new Map();

  return (req, res, next) => {
    const agora = Date.now();
    const chave = `${escopo}:${req.ip}`;
    const registro = acessos.get(chave);

    if (!registro || registro.expiraEm <= agora) {
      acessos.set(chave, { total: 1, expiraEm: agora + janelaMs });
      return next();
    }

    if (registro.total >= maximo) {
      const segundos = Math.max(1, Math.ceil((registro.expiraEm - agora) / 1000));
      res.setHeader("Retry-After", String(segundos));
      return res.status(429).json({ erro: "Muitas requisições. Tente novamente em instantes." });
    }

    registro.total += 1;
    if (acessos.size > 10000) {
      for (const [key, value] of acessos) {
        if (value.expiraEm <= agora) acessos.delete(key);
      }
    }
    next();
  };
}

const limitarConsultaCnpj = criarRateLimiter({
  janelaMs: 60 * 1000,
  maximo: numeroInteiroPositivo(process.env.CNPJ_RATE_LIMIT_MAX, 30),
  escopo: "cnpj"
});

const limitarGeracaoDocumento = criarRateLimiter({
  janelaMs: 10 * 60 * 1000,
  maximo: numeroInteiroPositivo(process.env.DOCUMENT_RATE_LIMIT_MAX, 20),
  escopo: "documento"
});

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

function cpfValido(cpf) {
  const n = apenasNumeros(cpf);
  if (n.length !== 11 || /^(\d)\1{10}$/.test(n)) return false;
  let soma = 0;
  for (let i = 0; i < 9; i++) soma += Number(n[i]) * (10 - i);
  let resto = (soma * 10) % 11;
  if (resto >= 10) resto = 0;
  if (resto !== Number(n[9])) return false;
  soma = 0;
  for (let i = 0; i < 10; i++) soma += Number(n[i]) * (11 - i);
  resto = (soma * 10) % 11;
  if (resto >= 10) resto = 0;
  return resto === Number(n[10]);
}

function cnpjValido(cnpj) {
  const n = apenasNumeros(cnpj);
  if (n.length !== 14 || /^(\d)\1{13}$/.test(n)) return false;
  const pesos1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];
  const pesos2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];
  let soma = pesos1.reduce((acc, p, i) => acc + Number(n[i]) * p, 0);
  let resto = soma % 11;
  if ((resto < 2 ? 0 : 11 - resto) !== Number(n[12])) return false;
  soma = pesos2.reduce((acc, p, i) => acc + Number(n[i]) * p, 0);
  resto = soma % 11;
  return (resto < 2 ? 0 : 11 - resto) === Number(n[13]);
}

function limparNomeArquivo(nome) {
  const limpo = String(nome || "LOCAL")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[\\/:*?"<>|]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase()
    .slice(0, 80);
  return limpo || "LOCAL";
}

function gerarHashProtecaoWord(senha, salt, spinCount = WORD_PROTECTION_SPIN_COUNT) {
  let hash = crypto
    .createHash("sha512")
    .update(Buffer.concat([salt, Buffer.from(String(senha), "utf16le")]))
    .digest();

  for (let i = 0; i < spinCount; i++) {
    const contador = Buffer.alloc(4);
    contador.writeUInt32LE(i, 0);
    hash = crypto
      .createHash("sha512")
      .update(Buffer.concat([contador, hash]))
      .digest();
  }

  return hash.toString("base64");
}

function aplicarRestricaoEdicaoWord(zip, senha = SENHA_PROTECAO_WORD) {
  const settingsFile = zip.file("word/settings.xml");
  if (!settingsFile) throw new Error("word/settings.xml não encontrado no DOCX.");

  const salt = crypto.randomBytes(16);
  const protecao = `<w:documentProtection w:edit="readOnly" w:enforcement="1" w:algorithmName="SHA-512" w:spinCount="${WORD_PROTECTION_SPIN_COUNT}" w:hashValue="${gerarHashProtecaoWord(senha, salt)}" w:saltValue="${salt.toString("base64")}"/>`;
  let settingsXml = settingsFile.asText();

  if (/<w:documentProtection\b[\s\S]*?\/>/.test(settingsXml)) {
    settingsXml = settingsXml.replace(/<w:documentProtection\b[\s\S]*?\/>/, protecao);
  } else if (/<w:documentProtection\b[\s\S]*?<\/w:documentProtection>/.test(settingsXml)) {
    settingsXml = settingsXml.replace(/<w:documentProtection\b[\s\S]*?<\/w:documentProtection>/, protecao);
  } else {
    settingsXml = settingsXml.replace("</w:settings>", `${protecao}</w:settings>`);
  }

  zip.file("word/settings.xml", settingsXml);
}

function textoDocx(v) {
  return String(v || "")
    .replace(/\r\n?/g, "\n")
    .replace(/[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD]/g, "")
    .trim();
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

function cursoEhObrigatorio(curso) {
  const c = normalizarTexto(curso);
  return cursosObrigatorios.some(nome => normalizarTexto(nome) === c);
}

function cursosSelecionados(d) {
  if (Array.isArray(d.cursos) && d.cursos.length) return d.cursos;
  return d.curso ? [d.curso] : [];
}

function checkCurso(cursoOuCursos, nome) {
  const cursos = Array.isArray(cursoOuCursos) ? cursoOuCursos : [cursoOuCursos];
  return cursos.some(curso => normalizarTexto(curso) === normalizarTexto(nome)) ? "X" : " ";
}

const TIPOS_ESTAGIO_CANONICOS = new Map([
  "Estágio obrigatório",
  "Estágio remunerado"
].map(valor => [normalizarTexto(valor), valor]));

const CURSOS_CANONICOS = new Map([
  "Biomedicina",
  "Farmácia",
  "Fonoaudiologia",
  "Fisioterapia",
  "Nutrição",
  "Terapia Ocupacional",
  "Radiologia",
  "Técnicas Oftálmicas",
  "Pós em Procedimentos Injetáveis",
  "Engenharias",
  "Licenciaturas",
  "Técnico em Enfermagem",
  "Técnico em Transações Imobiliárias",
  "Pós-Graduação em Biomedicina Estética",
  "Outro"
].map(valor => [normalizarTexto(valor), valor]));

const CAMPOS_FORMULARIO = [
  "tipo_unidade",
  "tipo_estagio",
  "curso",
  "outro_curso",
  "cnpj",
  "cpf",
  "razao_social",
  "alvara",
  "estimativa_vagas",
  "endereco",
  "numero",
  "complemento",
  "bairro",
  "cep",
  "cidade",
  "estado",
  "telefone_receita",
  "site",
  "responsavel_estagios",
  "contato_responsavel",
  "representante",
  "cargo",
  "email_assinatura"
];

const LIMITES_CAMPOS = {
  tipo_unidade: 10,
  tipo_estagio: 40,
  curso: 80,
  outro_curso: 100,
  cnpj: 20,
  cpf: 20,
  razao_social: 160,
  alvara: 100,
  estimativa_vagas: 40,
  endereco: 160,
  numero: 30,
  complemento: 120,
  bairro: 100,
  cep: 20,
  cidade: 100,
  estado: 40,
  telefone_receita: 50,
  site: 160,
  responsavel_estagios: 120,
  contato_responsavel: 50,
  representante: 120,
  cargo: 100,
  email_assinatura: 160
};

const CAMPOS_OBRIGATORIOS = [
  "tipo_estagio",
  "curso",
  "razao_social",
  "endereco",
  "numero",
  "bairro",
  "cep",
  "cidade",
  "estado",
  "responsavel_estagios",
  "contato_responsavel",
  "representante",
  "cargo",
  "email_assinatura"
];

function erroValidacao(mensagem) {
  const erro = new Error(mensagem);
  erro.statusCode = 400;
  return erro;
}

function textoEntrada(campo, valor) {
  if (Array.isArray(valor) || (valor && typeof valor === "object")) {
    throw erroValidacao(`Campo inválido: ${campo}.`);
  }

  const texto = textoDocx(valor).replace(/[<>]/g, "");
  const limite = LIMITES_CAMPOS[campo] || 160;
  if (texto.length > limite) {
    throw erroValidacao(`Campo muito longo: ${campo}.`);
  }
  return texto;
}

function valorCanonico(mapa, valor, campo) {
  const canonico = mapa.get(normalizarTexto(valor));
  if (!canonico) throw erroValidacao(`Valor inválido para ${campo}.`);
  return canonico;
}

function cursosCanonicosFormulario(body) {
  const valor = Object.prototype.hasOwnProperty.call(body, "cursos") ? body.cursos : body.curso;
  const entradas = Array.isArray(valor) ? valor : [valor];
  const cursos = [];
  const vistos = new Set();

  for (const entrada of entradas) {
    const texto = textoEntrada("curso", entrada);
    if (!texto) continue;
    const canonico = valorCanonico(CURSOS_CANONICOS, texto, "curso");
    const chave = normalizarTexto(canonico);
    if (!vistos.has(chave)) {
      cursos.push(canonico);
      vistos.add(chave);
    }
  }

  if (!cursos.length) throw erroValidacao("Campo obrigatório ausente: curso.");
  return cursos;
}

function emailValido(email) {
  return /^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/.test(email);
}

function validarDadosFormulario(body) {
  if (!body || typeof body !== "object" || Array.isArray(body)) {
    throw erroValidacao("Payload inválido.");
  }

  const dados = {};
  for (const campo of CAMPOS_FORMULARIO) {
    if (campo === "curso") continue;
    dados[campo] = textoEntrada(campo, body[campo]);
  }

  dados.tipo_unidade = dados.tipo_unidade === "cpf" ? "cpf" : "cnpj";
  dados.tipo_estagio = valorCanonico(TIPOS_ESTAGIO_CANONICOS, dados.tipo_estagio, "tipo de estágio");
  dados.cursos = cursosCanonicosFormulario(body);
  dados.curso = dados.cursos[0];

  for (const campo of CAMPOS_OBRIGATORIOS) {
    if (!dados[campo]) throw erroValidacao(`Campo obrigatório ausente: ${campo}.`);
  }

  if (dados.cursos.includes("Outro")) {
    if (!dados.outro_curso) throw erroValidacao("Informe o curso.");
  } else {
    dados.outro_curso = "";
  }

  if (dados.tipo_unidade === "cpf") {
    const cpfInformado = dados.cpf || dados.cnpj;
    if (!cpfValido(cpfInformado)) {
      throw erroValidacao("CPF inválido.");
    }
    dados.cpf = cpfInformado;
    dados.cnpj = cpfInformado;
  } else if (!cnpjValido(dados.cnpj)) {
    throw erroValidacao("CNPJ inválido.");
  }

  if (!emailValido(dados.email_assinatura)) {
    throw erroValidacao("E-mail para assinatura inválido.");
  }

  if (dados.site && !emailValido(dados.site)) {
    throw erroValidacao("E-mail de contato inválido.");
  }

  dados.telefone = dados.telefone_receita;
  return dados;
}

function responderErroAplicacao(res, erro, mensagemPublica) {
  if (erro && erro.statusCode === 400) {
    return res.status(400).json({ erro: erro.message });
  }
  console.error(erro);
  return res.status(500).json({ erro: mensagemPublica });
}

function montarContatoEmpresa(d) {
  return [
    d.responsavel_estagios ? `Responsável pelo estágio ou setor responsável: ${d.responsavel_estagios}` : "",
    d.contato_responsavel ? `Telefone de contato direto: ${d.contato_responsavel}` : "",
    d.site ? `E-mail: ${d.site}` : ""
  ].filter(Boolean).join("\n");
}

function selecionarModelo(d) {
  const tipo = d.tipo_estagio;
  const cursos = cursosSelecionados(d);
  if (tipo === "Estágio remunerado") return "remunerado.docx";
  if (tipo === "Estágio obrigatório" && cursos.length > 0 && cursos.every(cursoEhObrigatorio)) return "contrapartidas.docx";
  return "simples.docx";
}

function dadosTermo(d) {
  const cursos = cursosSelecionados(d)
    .map(curso => textoDocx(curso === "Outro" ? (d.outro_curso || "Outro") : curso))
    .filter(Boolean);
  const outros = cursos.filter(curso => !cursoExisteNaLista(curso)).join("\n");
  return {
    razao_social: textoDocx(d.razao_social),
    cnpj: textoDocx(d.cnpj),
    alvara: textoDocx(d.alvara),
    area_atuacao: "",
    outros: textoDocx(outros),
    estimativa_vagas: textoDocx(d.estimativa_vagas),
    endereco: textoDocx(d.endereco),
    numero: textoDocx(d.numero),
    complemento: textoDocx(d.complemento),
    bairro: textoDocx(d.bairro),
    cep: textoDocx(d.cep),
    cidade: textoDocx(d.cidade),
    estado: textoDocx(d.estado),
    telefone: textoDocx(d.telefone_receita || d.telefone),
    site: textoDocx(montarContatoEmpresa(d)),
    responsavel_estagios: textoDocx(d.responsavel_estagios),
    contato_responsavel: textoDocx(d.contato_responsavel),
    representante: textoDocx(d.representante),
    cargo: textoDocx(d.cargo),
    email_assinatura: textoDocx(d.email_assinatura),
    data_extenso: dataExtensoHoje(),

    chk_biomedicina: checkCurso(cursos, "Biomedicina"),
    chk_farmacia: checkCurso(cursos, "Farmácia"),
    chk_fonoaudiologia: checkCurso(cursos, "Fonoaudiologia"),
    chk_fisioterapia: checkCurso(cursos, "Fisioterapia"),
    chk_nutricao: checkCurso(cursos, "Nutrição"),
    chk_terapia_ocupacional: checkCurso(cursos, "Terapia Ocupacional"),
    chk_radiologia: checkCurso(cursos, "Radiologia"),
    chk_tecnicas_oftalmicas: checkCurso(cursos, "Técnicas Oftálmicas"),
    chk_engenharias: checkCurso(cursos, "Engenharias"),
    chk_licenciaturas: checkCurso(cursos, "Licenciaturas"),
    chk_tecnico_enfermagem: checkCurso(cursos, "Técnico em Enfermagem"),
    chk_tecnico_transacoes_imobiliarias: checkCurso(cursos, "Técnico em Transações Imobiliárias"),
    chk_pos_biomedicina_estetica: checkCurso(cursos, "Pós-Graduação em Biomedicina Estética")
  };
}

function decodificarXmlTexto(texto) {
  return String(texto || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'");
}

const paragrafosTemplateCache = new Map();

function extrairParagrafosTemplate(modelo) {
  if (paragrafosTemplateCache.has(modelo)) return paragrafosTemplateCache.get(modelo);

  const templatePath = path.join(__dirname, "templates", modelo);
  const zip = new PizZip(fs.readFileSync(templatePath));
  const xml = zip.file("word/document.xml").asText();
  const paragrafos = [...xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)]
    .map(match => {
      const partes = [];
      for (const texto of match[0].matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)) {
        partes.push(decodificarXmlTexto(texto[1]));
      }
      return partes.join("").replace(/\s+/g, " ").trim();
    })
    .filter(Boolean);

  paragrafosTemplateCache.set(modelo, paragrafos);
  return paragrafos;
}

function preencherTextoTemplate(texto, data) {
  return texto.replace(/\{([^{}]+)\}/g, (_, chave) => data[chave] ?? "");
}

function textoParaPdf(texto) {
  return textoDocx(texto)
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u2013\u2014]/g, "-")
    .replace(/\u00A0/g, " ");
}

function estiloParagrafo(texto, indice) {
  const limpo = texto.trim();
  const maiusculo = limpo.toUpperCase() === limpo && /[A-ZÁÉÍÓÚÂÊÔÃÕÇ]/.test(limpo);
  if (indice === 0 || limpo.startsWith("TERMO DE CONVÊNIO")) return { size: 12, bold: true, align: "center", gap: 10 };
  if (maiusculo && limpo.length <= 80) return { size: 10.5, bold: true, align: "left", gap: 5 };
  if (/^CLÁUSULA|^\d+\.\d+|^§/.test(limpo)) return { size: 9.5, bold: /^CLÁUSULA/.test(limpo), align: "justify", gap: 4 };
  return { size: 9.2, bold: false, align: "justify", gap: 4 };
}

function valorCampo(valor, fallback = "-") {
  return textoParaPdf(valor || fallback);
}

function ehModeloObrigatorio(modelo) {
  return modelo === "simples.docx" || modelo === "contrapartidas.docx";
}

function buscarIndiceConteudoLegal(modelo, paragrafos) {
  if (modelo === "remunerado.docx") {
    const i = paragrafos.findIndex(p => p.startsWith("A Unidade Concedente"));
    return i >= 0 ? i : 0;
  }

  if (modelo === "contrapartidas.docx") {
    const i = paragrafos.findIndex(p => p === "CONTRAPARTIDAS");
    return i >= 0 ? i : 0;
  }

  const i = paragrafos.findIndex(p => p === "DO OBJETO");
  return i >= 0 ? i : 0;
}

function adicionarCabecalhoPdf(doc) {
  const logoPath = path.join(__dirname, "public", "logo.png");
  if (fs.existsSync(logoPath)) {
    doc.image(logoPath, doc.page.margins.left, 24, { width: 118 });
  }

  doc
    .font("Helvetica-Bold")
    .fontSize(8.5)
    .fillColor("#333333")
    .text("Centro Universitário UniFatecie", doc.page.margins.left, 34, {
      align: "right",
      width: doc.page.width - doc.page.margins.left - doc.page.margins.right
    });

  doc.moveTo(doc.page.margins.left, 50)
    .lineTo(doc.page.width - doc.page.margins.right, 50)
    .strokeColor("#d99a00")
    .lineWidth(1)
    .stroke();

  doc.y = 72;
}

function garantirEspaco(doc, altura) {
  if (doc.y + altura > doc.page.height - doc.page.margins.bottom) {
    doc.addPage();
  }
}

function adicionarTituloSecao(doc, titulo) {
  garantirEspaco(doc, 28);
  doc.moveDown(0.5);
  doc
    .font("Helvetica-Bold")
    .fontSize(10.5)
    .fillColor("#111111")
    .text(textoParaPdf(titulo).toUpperCase(), {
      continued: false
    });
  doc.moveTo(doc.page.margins.left, doc.y + 3)
    .lineTo(doc.page.width - doc.page.margins.right, doc.y + 3)
    .strokeColor("#e5e5e5")
    .lineWidth(0.7)
    .stroke();
  doc.moveDown(0.7);
}

function adicionarBoxInstitucional(doc, texto) {
  const x = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const y = doc.y;
  const alturaTexto = doc.heightOfString(texto, { width: width - 22, align: "justify", lineGap: 1.4 });
  const height = alturaTexto + 18;
  garantirEspaco(doc, height + 10);

  doc.roundedRect(x, y, width, height, 5).fillAndStroke("#fafafa", "#e5e5e5");
  doc
    .font("Helvetica")
    .fontSize(8.8)
    .fillColor("#222222")
    .text(texto, x + 11, y + 9, {
      width: width - 22,
      align: "justify",
      lineGap: 1.4
    });
  doc.y = y + height + 10;
}

function adicionarLinhaTabela(doc, pares, larguraRotulo = 92) {
  const x = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const gap = 8;
  const colWidth = pares.length === 1 ? width : (width - gap) / 2;
  const alturas = pares.map(([rotulo, valor]) => {
    const valorWidth = colWidth - larguraRotulo - 18;
    return Math.max(
      34,
      doc.heightOfString(valorCampo(valor), { width: valorWidth, lineGap: 1 }) + 18
    );
  });
  const height = Math.max(...alturas);
  garantirEspaco(doc, height + 8);

  pares.forEach(([rotulo, valor], indice) => {
    const cellX = x + indice * (colWidth + gap);
    const y = doc.y;
    doc.roundedRect(cellX, y, colWidth, height, 4).fillAndStroke("#ffffff", "#d9d9d9");
    doc
      .font("Helvetica-Bold")
      .fontSize(7.8)
      .fillColor("#555555")
      .text(textoParaPdf(rotulo).toUpperCase(), cellX + 9, y + 9, {
        width: larguraRotulo
      });
    doc
      .font("Helvetica")
      .fontSize(8.8)
      .fillColor("#111111")
      .text(valorCampo(valor), cellX + larguraRotulo + 9, y + 8, {
        width: colWidth - larguraRotulo - 18,
        lineGap: 1
      });
  });

  doc.y += height + 8;
}

function adicionarDadosConcedente(doc, data, modelo) {
  adicionarTituloSecao(doc, "Unidade Concedente - Conveniada");
  adicionarLinhaTabela(doc, [["Razão social", data.razao_social]]);
  adicionarLinhaTabela(doc, [["CNPJ/CPF", data.cnpj], ["Alvará", data.alvara]]);
  adicionarLinhaTabela(doc, [["Endereço", data.endereco], ["Número", data.numero]]);
  adicionarLinhaTabela(doc, [["Complemento", data.complemento], ["Bairro", data.bairro]]);
  adicionarLinhaTabela(doc, [["CEP", data.cep], ["Cidade/UF", `${valorCampo(data.cidade, "")} / ${valorCampo(data.estado, "")}`]]);

  if (ehModeloObrigatorio(modelo)) {
    adicionarLinhaTabela(doc, [["Estimativa de vagas", data.estimativa_vagas], ["Telefone", data.telefone]]);
  } else {
    adicionarLinhaTabela(doc, [["Telefone", data.telefone]]);
  }

  adicionarLinhaTabela(doc, [["Responsável/Contato", data.site]], 118);
  adicionarLinhaTabela(doc, [["Representante", data.representante], ["Cargo", data.cargo]]);
  adicionarLinhaTabela(doc, [["E-mail para assinatura digital", data.email_assinatura]], 150);
}

function adicionarCursos(doc, data, modelo) {
  if (!ehModeloObrigatorio(modelo)) return;

  const cursos = [
    ["Biomedicina", data.chk_biomedicina],
    ["Farmácia", data.chk_farmacia],
    ["Fonoaudiologia", data.chk_fonoaudiologia],
    ["Fisioterapia", data.chk_fisioterapia],
    ["Nutrição", data.chk_nutricao],
    ["Terapia Ocupacional", data.chk_terapia_ocupacional],
    ["Radiologia", data.chk_radiologia],
    ["Técnicas Oftálmicas", data.chk_tecnicas_oftalmicas],
    ["Engenharias", data.chk_engenharias],
    ["Licenciaturas", data.chk_licenciaturas],
    ["Técnico em Enfermagem", data.chk_tecnico_enfermagem],
    ["Técnico em Transações Imobiliárias", data.chk_tecnico_transacoes_imobiliarias],
    ["Pós-Graduação em Biomedicina Estética", data.chk_pos_biomedicina_estetica]
  ].filter(([nome]) => modelo !== "contrapartidas.docx" || !["Engenharias", "Licenciaturas", "Técnico em Enfermagem", "Técnico em Transações Imobiliárias", "Pós-Graduação em Biomedicina Estética"].includes(nome));

  adicionarTituloSecao(doc, "Área de atuação / curso");

  const x = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const gap = 8;
  const colWidth = (width - gap) / 2;
  const rowHeight = 24;

  cursos.forEach(([nome, marcado], indice) => {
    if (indice % 2 === 0) garantirEspaco(doc, rowHeight + 8);
    const col = indice % 2;
    const cellX = x + col * (colWidth + gap);
    const y = doc.y;
    doc.roundedRect(cellX, y, colWidth, rowHeight, 4).fillAndStroke(marcado.trim() === "X" ? "#fff8df" : "#ffffff", "#d9d9d9");
    doc
      .font(marcado.trim() === "X" ? "Helvetica-Bold" : "Helvetica")
      .fontSize(8.4)
      .fillColor("#111111")
      .text(`[${marcado.trim() === "X" ? "X" : " "}] ${textoParaPdf(nome)}`, cellX + 8, y + 7, {
        width: colWidth - 16
      });
    if (col === 1 || indice === cursos.length - 1) doc.y += rowHeight + 8;
  });

  if (data.outros) adicionarLinhaTabela(doc, [["Outros", data.outros]]);
}

function adicionarParagrafoLegal(doc, texto, indice) {
  const paragrafo = textoParaPdf(texto);
  const estilo = estiloParagrafo(paragrafo, indice);
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const height = doc.heightOfString(paragrafo, {
    width,
    align: estilo.align,
    lineGap: 1.5
  }) + estilo.gap + 2;
  garantirEspaco(doc, Math.min(height, doc.page.height - doc.page.margins.top - doc.page.margins.bottom));

  doc
    .font(estilo.bold ? "Helvetica-Bold" : "Helvetica")
    .fontSize(estilo.size)
    .fillColor("#111111")
    .text(paragrafo, {
      align: estilo.align,
      lineGap: 1.5,
      paragraphGap: estilo.gap
    });
}

function adicionarAssinaturas(doc) {
  garantirEspaco(doc, 120);
  const x = doc.page.margins.left;
  const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const gap = 26;
  const colWidth = (width - gap) / 2;
  const y = doc.y + 24;

  [
    ["CONVENIADA - LOCAL", "(Assinatura)"],
    ["CONVENENTE - UNIFATECIE", "(Assinatura)"]
  ].forEach(([titulo, subtitulo], indice) => {
    const cellX = x + indice * (colWidth + gap);
    doc.moveTo(cellX, y).lineTo(cellX + colWidth, y).strokeColor("#111111").lineWidth(0.8).stroke();
    doc
      .font("Helvetica-Bold")
      .fontSize(8.5)
      .fillColor("#111111")
      .text(titulo, cellX, y + 8, { width: colWidth, align: "center" });
    doc
      .font("Helvetica")
      .fontSize(8)
      .fillColor("#555555")
      .text(subtitulo, cellX, y + 22, { width: colWidth, align: "center" });
  });

  doc.y = y + 46;
}

function gerarPdfTermo(d) {
  return new Promise((resolve, reject) => {
    const modelo = selecionarModelo(d);
    const data = dadosTermo(d);
    const paragrafosTemplate = extrairParagrafosTemplate(modelo);
    const titulo = textoParaPdf(preencherTextoTemplate(paragrafosTemplate[0] || "TERMO DE CONVÊNIO", data));
    const indiceConteudo = buscarIndiceConteudoLegal(modelo, paragrafosTemplate);
    const paragrafosLegais = paragrafosTemplate
      .slice(indiceConteudo)
      .map(p => textoParaPdf(preencherTextoTemplate(p, data)))
      .filter(Boolean)
      .filter(p => !/^_{8,}$/.test(p) && p !== "(Assinatura)" && !/^CONVENIADA/.test(p) && !/^CONVENENTE/.test(p))
      .filter(p => p !== "INSTITUIÇÃO DE ENSINO - CONVENENTE" && !p.startsWith("CENTRO UNIVERSITÁRIO UNIFATECIE"));
    const instituicao = "CENTRO UNIVERSITÁRIO UNIFATECIE, pessoa jurídica de direito privado, Instituição de Ensino Superior, com sede na cidade de PARANAVAÍ-PR, inscrita no CNPJ sob o número 07.724.708/0001-34, com sede à BR 376, nº1000, Km 102, Distrito Industrial, CEP 87.720-140, na cidade de Paranavaí-PR, neste ato representado pelo Prof. Me. Gilmar de Oliveira, portador do RG nº 17.773.010-9, inscrito no CPF sob nº.067.346.438-59, residente e domiciliado na cidade de PARANAVAÍ-PR, doravante denominado UNIFATECIE.";
    const chunks = [];
    const doc = new PDFDocument({
      size: "A4",
      margins: { top: 58, right: 56, bottom: 58, left: 56 },
      bufferPages: true,
      info: {
        Title: "Termo de Convênio UniFatecie",
        Author: "Centro Universitário UniFatecie",
        Subject: "Termo de Convênio"
      }
    });

    doc.on("data", chunk => chunks.push(chunk));
    doc.on("error", reject);
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("pageAdded", () => adicionarCabecalhoPdf(doc));

    adicionarCabecalhoPdf(doc);
    doc
      .font("Helvetica-Bold")
      .fontSize(13)
      .fillColor("#111111")
      .text(titulo, {
        align: "center",
        lineGap: 2,
        paragraphGap: 8
      });

    adicionarTituloSecao(doc, "Instituição de Ensino - Convenente");
    adicionarBoxInstitucional(doc, instituicao);
    adicionarDadosConcedente(doc, data, modelo);
    adicionarCursos(doc, data, modelo);
    adicionarTituloSecao(doc, "Termos e condições");
    paragrafosLegais.forEach((paragrafo, indice) => adicionarParagrafoLegal(doc, paragrafo, indice));
    adicionarAssinaturas(doc);

    const range = doc.bufferedPageRange();
    for (let i = range.start; i < range.start + range.count; i++) {
      doc.switchToPage(i);
      doc
        .font("Helvetica")
        .fontSize(8)
        .fillColor("#666666")
        .text(`Página ${i + 1} de ${range.count}`, doc.page.margins.left, doc.page.height - 38, {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          align: "center"
        });
    }

    doc.end();
  });
}

function fetchJson(url, nomeProvedor) {
  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    if (parsed.protocol !== "https:") {
      return reject(new Error(`${nomeProvedor}: apenas HTTPS é permitido para consultas externas.`));
    }
    const client = https;
    const req = client.request(parsed, {
      method: "GET",
      timeout: 12000,
      headers: {
        "accept": "application/json",
        "user-agent": "GeradorConvenioUniFatecie/1.0"
      }
    }, resp => {
      let body = "";
      let bytesRecebidos = 0;
      resp.setEncoding("utf8");
      resp.on("data", chunk => {
        bytesRecebidos += Buffer.byteLength(chunk, "utf8");
        if (bytesRecebidos > MAX_EXTERNAL_JSON_BYTES) {
          req.destroy(new Error(`${nomeProvedor}: resposta excedeu o limite permitido.`));
          return;
        }
        body += chunk;
      });
      resp.on("end", () => {
        let json;
        try {
          json = body ? JSON.parse(body) : {};
        } catch (e) {
          return reject(new Error(`${nomeProvedor}: resposta não é JSON válido (${resp.statusCode}).`));
        }

        const mensagemApi = json.message || json.erro || json.error || json.status;
        if (resp.statusCode < 200 || resp.statusCode >= 300) {
          return reject(new Error(`${nomeProvedor}: HTTP ${resp.statusCode}${mensagemApi ? ` - ${mensagemApi}` : ""}`));
        }

        if (json.status && String(json.status).toUpperCase() === "ERROR") {
          return reject(new Error(`${nomeProvedor}: ${json.message || "consulta recusada pela API"}`));
        }

        resolve(json);
      });
    });

    req.on("timeout", () => req.destroy(new Error(`${nomeProvedor}: tempo limite excedido.`)));
    req.on("error", reject);
    req.end();
  });
}

function mapSocios(dados) {
  const qsa = dados.qsa || dados.socios || [];
  if (!Array.isArray(qsa)) return [];
  return qsa.map(s => ({
    nome: s.nome_socio || s.nome || s.nomeSocio || "",
    cargo: s.qualificacao_socio || s.qualificacao || s.qual || s.cargo || "Sócio/Administrador"
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
    telefone: [dados.ddd_telefone_1, dados.ddd_telefone_2, dados.telefone].filter(Boolean)[0] || "",
    telefone_receita: [dados.ddd_telefone_1, dados.ddd_telefone_2, dados.telefone].filter(Boolean)[0] || "",
    email: dados.email || "",
    socios: mapSocios(dados)
  };
}

app.get("/api/cnpj/:cnpj", limitarConsultaCnpj, async (req, res) => {
  const cnpj = apenasNumeros(req.params.cnpj);
  if (cnpj.length !== 14) return res.status(400).json({ erro: "CNPJ inválido." });

  const provedores = [
    { nome: "Minha Receita", url: `https://minhareceita.org/${cnpj}` },
    { nome: "BrasilAPI", url: `https://brasilapi.com.br/api/cnpj/v1/${cnpj}` },
    { nome: "ReceitaWS", url: `https://receitaws.com.br/v1/cnpj/${cnpj}` }
  ];

  const cnpjMascarado = cnpj.slice(0, 2) + "****" + cnpj.slice(-4);
  const erros = [];
  for (const provedor of provedores) {
    try {
      const dados = await fetchJson(provedor.url, provedor.nome);
      const empresa = mapDadosEmpresa(dados);
      if (!empresa.razao_social) {
        throw new Error(`${provedor.nome}: resposta sem razão social.`);
      }
      return res.json({ ...empresa, fonte: provedor.nome });
    } catch (e) {
      const detalhe = e.message || String(e);
      erros.push(detalhe);
      console.warn(`[CNPJ] Falha no provedor ${provedor.nome} para ${cnpjMascarado}: ${detalhe}`);
    }
  }
  console.warn(`[CNPJ] Todos os provedores falharam para ${cnpjMascarado}: ${erros.join(" | ")}`);
  return res.status(502).json({
    erro: "Não foi possível consultar o CNPJ no momento. Preencha manualmente."
  });
});

app.post("/api/gerar-pdf", limitarGeracaoDocumento, async (req, res) => {
  try {
    const d = validarDadosFormulario(req.body);
    const pdf = await gerarPdfTermo(d);
    const nomeLocal = limparNomeArquivo(d.razao_social || d.nome_fantasia || "LOCAL");
    const filename = `${nomeLocal} - TERMO DE CONVENIO.pdf`;

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(filename)}"; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(pdf);
  } catch (e) {
    responderErroAplicacao(res, e, "Erro ao gerar PDF.");
  }
});

app.post("/api/gerar", limitarGeracaoDocumento, (req, res) => {
  try {
    const d = validarDadosFormulario(req.body);
    const modelo = selecionarModelo(d);

    const templatePath = path.join(__dirname, "templates", modelo);
    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      nullGetter: () => ""
    });

    const data = dadosTermo(d);

    doc.render(data);
    aplicarRestricaoEdicaoWord(doc.getZip());

    const documentoWord = doc.getZip().generate({
      type: "nodebuffer",
      compression: "DEFLATE"
    });

    const nomeLocal = limparNomeArquivo(d.razao_social || d.nome_fantasia || "LOCAL");
    const filename = `${nomeLocal} - TERMO DE CONVENIO.docx`;

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(filename)}"; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(documentoWord);
  } catch (e) {
    responderErroAplicacao(res, e, "Erro ao gerar arquivo.");
  }
});

if (require.main === module) {
  app.listen(PORT, () => console.log(`Servidor rodando em http://localhost:${PORT}`));
}

module.exports = app;
