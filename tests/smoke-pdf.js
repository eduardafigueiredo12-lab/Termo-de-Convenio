const http = require("http");
const app = require("../server");

const dadosBase = {
  tipo_estagio: "Estágio remunerado",
  curso: "Engenharias",
  outro_curso: "",
  cnpj: "05.433.048/0001-07",
  cpf: "",
  razao_social: "INCOPOSTES INDUSTRIA E COMERCIO DE POSTES LTDA & CIA",
  alvara: "",
  estimativa_vagas: "5",
  endereco: "RODOVIA BR - 376",
  numero: "S/N",
  complemento: "KM 111 LOTE 08 09 03 E 04",
  bairro: "DISTRITO INDUSTRIAL (SUMARE)",
  cep: "87.720-140",
  cidade: "PARANAVAI",
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

async function testarPdf(baseUrl, dados) {
  const resp = await fetch(`${baseUrl}/api/gerar-pdf`, {
    method: "POST",
    headers: { "content-type": "application/json; charset=utf-8" },
    body: JSON.stringify(dados)
  });
  const buffer = Buffer.from(await resp.arrayBuffer());
  const texto = buffer.toString("latin1");

  if (!resp.ok) throw new Error(`PDF retornou HTTP ${resp.status}: ${texto.slice(0, 200)}`);
  if (!String(resp.headers.get("content-type")).includes("application/pdf")) {
    throw new Error(`Content-Type inválido: ${resp.headers.get("content-type")}`);
  }
  if (!texto.startsWith("%PDF-")) throw new Error("Arquivo gerado não começa com assinatura %PDF.");
  if (!texto.includes("%%EOF")) throw new Error("Arquivo gerado não contém marcador final %%EOF.");
  if ((texto.match(/\/Type\s*\/Page\b/g) || []).length < 1) {
    throw new Error("PDF gerado não contém páginas válidas.");
  }
  if (buffer.length < 5000) throw new Error(`PDF gerado parece pequeno demais (${buffer.length} bytes).`);

  return buffer.length;
}

(async () => {
  const { server, baseUrl } = await iniciarServidor();
  try {
    const cenarios = [
      { tipo_estagio: "Estágio obrigatório", curso: "Biomedicina" },
      { tipo_estagio: "Estágio obrigatório", curso: "Licenciaturas" },
      { tipo_estagio: "Estágio remunerado", curso: "Engenharias" }
    ];

    for (const cenario of cenarios) {
      const bytes = await testarPdf(baseUrl, { ...dadosBase, ...cenario });
      console.log(`PDF OK: ${cenario.tipo_estagio} / ${cenario.curso} (${bytes} bytes)`);
    }
  } finally {
    server.close();
  }
})().catch(error => {
  console.error(error);
  process.exit(1);
});
