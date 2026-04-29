
const express = require("express");
const cors = require("cors");
const path = require("path");
const fs = require("fs");
const http = require("http");
const https = require("https");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const PDFDocument = require("pdfkit");

const app = express();
app.use(cors());
app.use(express.json({ limit: "2mb" }));
app.use(express.static(path.join(__dirname, "public")));

const PORT = process.env.PORT || 3000;

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

function selecionarModelo(d) {
  const tipo = d.tipo_estagio;
  if (tipo === "Estágio remunerado") return "remunerado.docx";
  if (tipo === "Estágio obrigatório" && cursosObrigatorios.includes(d.curso)) return "contrapartidas.docx";
  return "simples.docx";
}

function dadosTermo(d) {
  const curso = textoDocx(d.curso === "Outro" ? (d.outro_curso || "Outro") : (d.curso || ""));
  const outros = cursoExisteNaLista(curso) ? "" : curso;
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
    telefone: textoDocx(d.contato_responsavel || d.telefone),
    site: textoDocx(montarContatoEmpresa(d)),
    responsavel_estagios: textoDocx(d.responsavel_estagios),
    contato_responsavel: textoDocx(d.contato_responsavel),
    representante: textoDocx(d.representante),
    cargo: textoDocx(d.cargo),
    email_assinatura: textoDocx(d.email_assinatura),
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
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")
    .replace(/[–—]/g, "-")
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

function gerarPdfTermo(d) {
  return new Promise((resolve, reject) => {
    const modelo = selecionarModelo(d);
    const data = dadosTermo(d);
    const paragrafos = extrairParagrafosTemplate(modelo).map(p => textoParaPdf(preencherTextoTemplate(p, data))).filter(Boolean);
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
    doc.moveTo(doc.page.margins.left, 50).lineTo(doc.page.width - doc.page.margins.right, 50).strokeColor("#d99a00").lineWidth(1).stroke();
    doc.y = 72;

    paragrafos.forEach((paragrafo, indice) => {
      const estilo = estiloParagrafo(paragrafo, indice);
      doc
        .font(estilo.bold ? "Helvetica-Bold" : "Helvetica")
        .fontSize(estilo.size)
        .fillColor("#111111")
        .text(paragrafo, {
          align: estilo.align,
          lineGap: 1.5,
          paragraphGap: estilo.gap
        });
    });

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
    const client = parsed.protocol === "http:" ? http : https;
    const req = client.request(parsed, {
      method: "GET",
      timeout: 12000,
      headers: {
        "accept": "application/json",
        "user-agent": "GeradorConvenioUniFatecie/1.0"
      }
    }, resp => {
      let body = "";
      resp.setEncoding("utf8");
      resp.on("data", chunk => { body += chunk; });
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
    email: dados.email || "",
    socios: mapSocios(dados)
  };
}

app.get("/api/cnpj/:cnpj", async (req, res) => {
  const cnpj = apenasNumeros(req.params.cnpj);
  if (cnpj.length !== 14) return res.status(400).json({ erro: "CNPJ inválido." });

  const provedores = [
    { nome: "Minha Receita", url: `https://minhareceita.org/${cnpj}` },
    { nome: "BrasilAPI", url: `https://brasilapi.com.br/api/cnpj/v1/${cnpj}` },
    { nome: "ReceitaWS", url: `https://receitaws.com.br/v1/cnpj/${cnpj}` }
  ];

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
      console.warn(`[CNPJ] Falha no provedor ${provedor.nome} para ${cnpj}: ${detalhe}`);
    }
  }
  return res.status(502).json({
    erro: "Não foi possível consultar o CNPJ no momento. Preencha manualmente.",
    detalhe: erros.join(" | ")
  });
});

app.post("/api/gerar-pdf", async (req, res) => {
  try {
    const d = req.body || {};
    const pdf = await gerarPdfTermo(d);
    const nomeLocal = limparNomeArquivo(d.razao_social || d.nome_fantasia || "LOCAL");
    const filename = `${nomeLocal} - TERMO DE CONVENIO.pdf`;

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(filename)}"; filename*=UTF-8''${encodeURIComponent(filename)}`);
    res.send(pdf);
  } catch (e) {
    console.error(e);
    res.status(500).json({ erro: "Erro ao gerar PDF.", detalhe: e.message });
  }
});

app.post("/api/gerar", (req, res) => {
  try {
    const d = req.body || {};
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
    console.error(e);
    res.status(500).json({ erro: "Erro ao gerar arquivo.", detalhe: e.message });
  }
});

if (require.main === module) {
  app.listen(PORT, () => console.log(`Servidor rodando em http://localhost:${PORT}`));
}

module.exports = app;
