Automação de Follow-up de Pedidos (Google Apps Script)
 O Problema
No ambiente industrial (Metalurgia), o controle de entrega de pedidos era feito manualmente. Eu precisava verificar planilha por planilha e enviar e-mails individuais para cada fornecedor cobrando prazos, o que gerava um alto custo de tempo e risco de erros humanos.

 A Solução
Desenvolvi um script em Google Apps Script (JavaScript) que automatiza 100% esse processo:

Inteligência de Agrupamento: O script identifica pedidos diferentes para o mesmo e-mail e os agrupa em uma única mensagem HTML profissional.

Verificação de Duplicidade: Antes de enviar, o script faz uma busca no Gmail para confirmar se o e-mail já foi enviado, evitando spam para o fornecedor.

Tratamento de Strings: Limpeza automática de e-mails mal digitados (remove espaços, trata pontos e vírgulas).

Assinatura Dinâmica: Inserção de imagem de assinatura via DriveApp.

 Tecnologias:
Google Apps Script

Gmail API / Drive API

JavaScript (ES6+)
