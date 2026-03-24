const ASSINATURA_IMG_ID = "exempleID";
const EMAIL_CC_FIXO = "exemple@exemple.com.br";

function enviarEmailsAgrupados() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return;

  const headers = data[0].map(h => h.toString().trim().toLowerCase());
  const colPedido = headers.indexOf("pedido");
  const colEmail = headers.indexOf("email");
  const colEnviado = headers.indexOf("jaenviado");

  if (colPedido === -1 || colEmail === -1 || colEnviado === -1) {
    throw new Error("Colunas 'Pedido', 'email' ou 'JaEnviado' não encontradas.");
  }

  const enviosPendentes = {};

  for (let i = 1; i < data.length; i++) {
    const pedido = data[i][colPedido];
    const emailRaw = data[i][colEmail];
    const statusEnviado = data[i][colEnviado];

    if (!pedido || !emailRaw) continue;

    if (statusEnviado && statusEnviado.toString().trim().toUpperCase() === "SIM") continue;

    const query = `in:sent subject:"Pedidos (${pedido})"`;
    const threads = GmailApp.search(query, 0, 1);

    if (threads.length > 0) {
      console.log(`Pedido ${pedido} já encontrado nos enviados. Marcando na planilha...`);
      sheet.getRange(i + 1, colEnviado + 1).setValue("SIM");
      continue; 
    }

    const emailTratado = processarEmails(emailRaw);
    if (!emailTratado) continue;

    if (!enviosPendentes[emailTratado]) {
      enviosPendentes[emailTratado] = [];
    }

    const pedidoJaAdicionado = enviosPendentes[emailTratado].some(item => item.id === pedido);

    if (!pedidoJaAdicionado) {
      enviosPendentes[emailTratado].push({ id: pedido, rowIndex: i + 1 });
    } else {
      sheet.getRange(i + 1, colEnviado + 1).setValue("SIM");
    }
  }

  if (Object.keys(enviosPendentes).length === 0) {
    console.log("Nada novo para enviar.");
    return;
  }

  let imgBlob;
  try {
    imgBlob = DriveApp.getFileById(ASSINATURA_IMG_ID).getBlob();
  } catch (e) {
    console.error("Erro ao buscar imagem da assinatura: " + e.message);
  }

  for (const emailDestino in enviosPendentes) {
    const itens = enviosPendentes[emailDestino];
    const numerosPedidos = itens.map(item => item.id);
    const pedidosTexto = numerosPedidos.join(", ");
    
    const listaHtml = numerosPedidos.map(p => `<li>Pedido: <b>${p}</b></li>`).join("");

    const htmlBody = `
      Olá,<br><br>
      Solicito que me informe a data de entrega dos produtos referentes aos seguintes pedidos:<br>
      <ul>${listaHtml}</ul>
      Aguardamos retorno.
      <br><br>
      <img src="cid:assinatura" style="max-width:400px;">
    `;

    try {
      GmailApp.sendEmail(
        emailDestino,
        `Pedidos (${pedidosTexto}) - Data de entrega`,
        "",
        {
          cc: EMAIL_CC_FIXO,
          htmlBody: htmlBody,
          inlineImages: imgBlob ? { assinatura: imgBlob } : {}
        }
      );

      itens.forEach(item => {
        sheet.getRange(item.rowIndex, colEnviado + 1).setValue("SIM");
      });
      console.log(`E-mail enviado para: ${emailDestino} com os pedidos: ${pedidosTexto}`);
    } catch (e) {
      console.error(`Erro no envio para ${emailDestino}: ${e.message}`);
    }
  }
}

function processarEmails(emailRaw) {
  if (!emailRaw) return null;
  const emails = emailRaw.toString()
    .replace(/\s+E\s+/gi, ",")
    .replace(/;/g, ",")
    .replace(/\s+/g, "")
    .split(",")
    .filter(e => e.includes("@"));
    
  return emails.length > 0 ? [...new Set(emails)].join(",") : null;
}
