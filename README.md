Gerador de Termo de Convênio UniFatecie - v21

Ajustes desta versão:
- Marca o curso com X entre parênteses quando o curso existir na lista do termo.
- Preenche "Outros" somente quando o curso não existir na lista.
- Proteção parcial do Word com senha: convenios.
- Apenas a tabela inicial da unidade concedente fica editável.
- Cláusulas, data e assinaturas ficam protegidas.

Como rodar:
npm.cmd install
npm.cmd start

Acesse:
http://localhost:3000


## Versão v22
- Download em Word (.docx).
- Documento inteiro protegido contra edição.
- Senha de proteção configurável por `WORD_PROTECTION_PASSWORD`.
- O sistema preenche o documento antes de aplicar o bloqueio.


Versão com tela de carregamento institucional UniFatecie durante consulta de CNPJ e geração do termo.

## Configurações de segurança
- `CORS_ORIGINS`: lista de origens permitidas, separadas por vírgula.
- `WORD_PROTECTION_PASSWORD`: senha usada para aplicar a proteção de edição no Word.
- `JSON_BODY_LIMIT`: limite do corpo JSON da API. Padrão: `100kb`.
- `CNPJ_RATE_LIMIT_MAX`: limite de consultas de CNPJ por minuto/IP. Padrão: `30`.
- `DOCUMENT_RATE_LIMIT_MAX`: limite de gerações de documento a cada 10 minutos/IP. Padrão: `20`.
- `TRUST_PROXY`: habilite quando a aplicação estiver atrás de proxy reverso confiável.

