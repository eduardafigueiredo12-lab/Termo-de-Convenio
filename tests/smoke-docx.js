const fs = require("fs");
const path = require("path");
const http = require("http");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const app = require("../server");

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
  const tagsNaoPreenchidas = xml.match(/\{[a-zA-Z0-9_]+\}/g);
  if (tagsNaoPreenchidas) {
    throw new Error(`${nomeArquivo}: marcador não preenchido: ${tagsNaoPreenchidas[0]}`);
  }

  if (!xml.includes("INCOPOSTES")) {
    throw new Error(`${nomeArquivo}: dados do formulário não foram inseridos no documento.`);
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
