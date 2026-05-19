const fs = require("fs");
const path = require("path");
const http = require("http");
const crypto = require("crypto");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const app = require("../server");

const SENHA_TESTE_WORD = process.env.WORD_PROTECTION_PASSWORD || "dev-local-convenios-nao-usar-producao";

const dadosBase = {
  tipo_estagio: "Estágio obrigatório",
  curso: "Engenharias",
  outro_curso: "",
  cnpj: "05.433.048/0001-07",
  cpf: "",
  razao_social: "INCOPOSTES INDUSTRIA E COMERCIO DE POSTES LTDA & CIA",
  alvara: "12345",
  estimativa_vagas: "5",
  endereco: "RODOVIA BR - 376",
  numero: "S/N",
  complemento: "KM 111 LOTE 08 09 03 E 04",
  bairro: "DISTRITO INDUSTRIAL (SUMARE)",
  cep: "87.720-140",
  cidade: "PARANAVAÍ",
  estado: "PR",
  telefone_receita: "(44) 99999-0000",
  site: "contabilidade.laisi@incopostes.com.br",
  responsavel_estagios: "Responsável Teste",
  contato_responsavel: "(44) 3045-1500",
  representante: "VILMAR JOSE MARQUES",
  cargo: "Sócio-Administrador",
  email_assinatura: "assinatura@example.com",
  tipo_unidade: "cnpj"
};

function iniciarServidor() {
  return new Promise(resolve => {
    const server = http.createServer(app);
    server.listen(0, "127.0.0.1", () => {
      const { port } = server.address();
      resolve({ server, baseUrl: `http://127.0.0.1:${port}` });
    });
  });
}

function atributoXml(xml, atributo) {
  const match = xml.match(new RegExp(`${atributo}="([^"]+)"`));
  return match ? match[1] : "";
}

function hashProtecaoWord(senha, salt, spinCount) {
  let hash = crypto
    .createHash("sha512")
    .update(Buffer.concat([salt, Buffer.from(String(senha), "utf16le")]))
    .digest();

  for (let i = 0; i < spinCount; i++) {
    const contador = Buffer.alloc(4);
    contador.writeUInt32LE(i, 0);
    hash = crypto
      .createHash("sha512")
      .update(Buffer.concat([hash, contador]))
      .digest();
  }

  return hash.toString("base64");
}

function validarDocx(buffer, nomeArquivo) {
  const zip = new PizZip(buffer);
  const obrigatorios = [
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
    "word/_rels/document.xml.rels"
  ];

  for (const entrada of obrigatorios) {
    if (!zip.file(entrada)) throw new Error(`${nomeArquivo}: entrada ausente no DOCX: ${entrada}`);
  }

  new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

  const xml = zip.file("word/document.xml").asText();
  const settingsXml = zip.file("word/settings.xml").asText();
  const tagsNaoPreenchidas = xml.match(/\{[a-zA-Z0-9_]+\}/g);
  if (tagsNaoPreenchidas) {
    throw new Error(`${nomeArquivo}: marcador não preenchido: ${tagsNaoPreenchidas[0]}`);
  }

  if (!xml.includes("INCOPOSTES")) {
    throw new Error(`${nomeArquivo}: dados do formulário não foram inseridos no documento.`);
  }

  if (nomeArquivo.includes("contrapartidas") && !xml.includes("CONTRAPARTIDAS")) {
    throw new Error(`${nomeArquivo}: modelo com contrapartidas nao foi utilizado.`);
  }

  if (nomeArquivo.includes("simples") && xml.includes("CONTRAPARTIDAS")) {
    throw new Error(`${nomeArquivo}: modelo sem contrapartidas nao foi utilizado.`);
  }

  if (!/<w:documentProtection\b/.test(settingsXml) || !/w:edit="readOnly"/.test(settingsXml) || !/w:enforcement="1"/.test(settingsXml)) {
    throw new Error(`${nomeArquivo}: restrição de edição do Word não foi aplicada.`);
  }

  if (!/w:cryptAlgorithmSid="14"/.test(settingsXml) || !/w:cryptSpinCount="/.test(settingsXml) || !/w:hash="/.test(settingsXml) || !/w:salt="/.test(settingsXml)) {
    throw new Error(`${nomeArquivo}: proteção por senha não foi configurada no DOCX.`);
  }

  const spinCount = Number(atributoXml(settingsXml, "w:cryptSpinCount"));
  const saltValue = atributoXml(settingsXml, "w:salt");
  const hashValue = atributoXml(settingsXml, "w:hash");
  const hashEsperado = hashProtecaoWord(SENHA_TESTE_WORD, Buffer.from(saltValue, "base64"), spinCount);
  if (hashValue !== hashEsperado) {
    throw new Error(`${nomeArquivo}: senha do Word nao corresponde ao hash gerado.`);
  }
}

async function gerarDocx(baseUrl, dados, nomeArquivo) {
  const resp = await fetch(`${baseUrl}/api/gerar`, {
    method: "POST",
    headers: { "content-type": "application/json; charset=utf-8" },
    body: JSON.stringify(dados)
  });
  const buffer = Buffer.from(await resp.arrayBuffer());
  const textoErro = buffer.toString("utf8");

  if (!resp.ok) throw new Error(`DOCX retornou HTTP ${resp.status}: ${textoErro.slice(0, 200)}`);
  if (!String(resp.headers.get("content-type")).includes("application/vnd.openxmlformats-officedocument.wordprocessingml.document")) {
    throw new Error(`Content-Type inválido: ${resp.headers.get("content-type")}`);
  }

  validarDocx(buffer, nomeArquivo);

  const saidaDir = path.join(__dirname, "..", "tmp", "generated-docx-smoke");
  fs.mkdirSync(saidaDir, { recursive: true });
  fs.writeFileSync(path.join(saidaDir, nomeArquivo), buffer);

  return buffer.length;
}

(async () => {
  const { server, baseUrl } = await iniciarServidor();
  try {
    const cenarios = [
      {
        nome: "simples.docx",
        dados: { ...dadosBase, tipo_estagio: "Estágio obrigatório", curso: "Engenharias" }
      },
      {
        nome: "contrapartidas.docx",
        dados: { ...dadosBase, tipo_estagio: "Estágio obrigatório", curso: "Biomedicina" }
      },
      {
        nome: "contrapartidas-multiplos-obrigatorios.docx",
        dados: { ...dadosBase, tipo_estagio: "Estágio obrigatório", curso: ["Biomedicina", "Farmácia"] }
      },
      {
        nome: "simples-misto.docx",
        dados: { ...dadosBase, tipo_estagio: "Estágio obrigatório", curso: ["Biomedicina", "Engenharias"] }
      },
      {
        nome: "remunerado.docx",
        dados: { ...dadosBase, tipo_estagio: "Estágio remunerado", curso: "Engenharias" }
      }
    ];

    for (const cenario of cenarios) {
      const bytes = await gerarDocx(baseUrl, cenario.dados, cenario.nome);
      console.log(`DOCX OK: ${cenario.nome} (${bytes} bytes)`);
    }
  } finally {
    server.close();
  }
})().catch(error => {
  console.error(error);
  process.exit(1);
});
