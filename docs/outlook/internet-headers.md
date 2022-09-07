---
title: Obter e definir cabeçalhos da Internet
description: Como obter e definir cabeçalhos da Internet em uma mensagem em um suplemento do Outlook.
ms.date: 08/30/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1f8e4af70b24a96b8d00acc7ea4101acf53e2b71
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616025"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a>Obter e definir cabeçalhos da Internet em uma mensagem em um suplemento do Outlook

## <a name="background"></a>Histórico

Um requisito comum no desenvolvimento de suplementos do Outlook é armazenar propriedades personalizadas associadas a um suplemento em níveis diferentes. No momento, as propriedades personalizadas são armazenadas no nível do item ou da caixa de correio.

- Nível de item – para propriedades que se aplicam a um item específico, use o [objeto CustomProperties](/javascript/api/outlook/office.customproperties) . Por exemplo, armazene um código de cliente associado à pessoa que enviou o email.
- Nível da caixa de correio – para propriedades que se aplicam a todos os itens de email na caixa de correio do usuário, use o objeto [RoamingSettings](/javascript/api/outlook/office.roamingsettings) . Por exemplo, armazene a preferência de um usuário para mostrar a temperatura em uma escala específica.

Os dois tipos de propriedades não são preservados depois que o item sai do exchange server, portanto, os destinatários de email não podem obter nenhuma propriedade definida no item. Portanto, os desenvolvedores não podem acessar essas configurações ou outras propriedades MIME (Multipurpose Internet Mail Extensions) para habilitar melhores cenários de leitura.

Embora haja uma maneira de definir os cabeçalhos da Internet por meio de solicitações dos Serviços Web do Exchange (EWS), em alguns cenários, fazer uma solicitação EWS não funcionará. Por exemplo, no modo Redigir na área de trabalho do Outlook, a ID do item não é sincronizada no `saveAsync` modo armazenado em cache.

> [!TIP]
> Para saber mais sobre como usar essas opções, consulte [Obter e definir metadados de suplemento para um suplemento do Outlook](metadata-for-an-outlook-add-in.md).

## <a name="purpose-of-the-internet-headers-api"></a>Finalidade da API de cabeçalhos da Internet

Introduzidas [no conjunto de requisitos 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8), as APIs de cabeçalhos da Internet permitem que os desenvolvedores:

- Carimbar informações em um email que persiste depois que ele sair do Exchange em todos os clientes.
- Leia informações sobre um email que persistia depois que o email deixou o Exchange em todos os clientes em cenários de leitura de email.
- Acesse todo o cabeçalho MIME do email.

![Diagrama de cabeçalhos da Internet. Texto: o usuário 1 envia email. O suplemento gerencia cabeçalhos personalizados da Internet enquanto o usuário está redigindo emails. O usuário 2 recebe o email. O suplemento obtém cabeçalhos da Internet de emails recebidos e, em seguida, analisa e usa cabeçalhos personalizados.](../images/outlook-internet-headers.png)

## <a name="set-internet-headers-while-composing-a-message"></a>Definir cabeçalhos da Internet ao redigir uma mensagem

Use a [propriedade item.internetHeaders](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-internetheaders-member) para gerenciar os cabeçalhos de Internet personalizados que você coloca na mensagem atual no modo Redigir.

### <a name="set-get-and-remove-custom-internet-headers-example"></a>Exemplo de definir, obter e remover cabeçalhos personalizados da Internet

O exemplo a seguir mostra como definir, obter e remover cabeçalhos personalizados da Internet.

```js
// Set custom internet headers.
function setCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.setAsync(
    { "x-preferred-fruit": "orange", "x-preferred-vegetable": "broccoli", "x-best-vegetable": "spinach" },
    setCallback
  );
}

function setCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Successfully set headers");
  } else {
    console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
  }
}

// Get custom internet headers.
function getSelectedCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.getAsync(
    ["x-preferred-fruit", "x-preferred-vegetable", "x-best-vegetable", "x-nonexistent-header"],
    getCallback
  );
}

function getCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Selected headers: " + JSON.stringify(asyncResult.value));
  } else {
    console.log("Error getting selected headers: " + JSON.stringify(asyncResult.error));
  }
}

// Remove custom internet headers.
function removeSelectedCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.removeAsync(
    ["x-best-vegetable", "x-nonexistent-header"],
    removeCallback);
}

function removeCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Successfully removed selected headers");
  } else {
    console.log("Error removing selected headers: " + JSON.stringify(asyncResult.error));
  }
}

setCustomHeaders();
getSelectedCustomHeaders();
removeSelectedCustomHeaders();
getSelectedCustomHeaders();

/* Sample output:
Successfully set headers
Selected headers: {"x-best-vegetable":"spinach","x-preferred-fruit":"orange","x-preferred-vegetable":"broccoli"}
Successfully removed selected headers
Selected headers: {"x-preferred-fruit":"orange","x-preferred-vegetable":"broccoli"}
*/
```

## <a name="get-internet-headers-while-reading-a-message"></a>Obter cabeçalhos da Internet durante a leitura de uma mensagem

Chame [item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#outlook-office-messageread-getallinternetheadersasync-member(1)) para obter cabeçalhos da Internet na mensagem atual no modo de leitura.

### <a name="get-sender-preferences-from-current-mime-headers-example"></a>Obter preferências de remetente do exemplo de cabeçalhos MIME atuais

Com base no exemplo da seção anterior, o código a seguir mostra como obter as preferências do remetente dos cabeçalhos MIME do email atual.

```js
Office.context.mailbox.item.getAllInternetHeadersAsync(getCallback);

function getCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Sender's preferred fruit: " + asyncResult.value.match(/x-preferred-fruit:.*/gim)[0].slice(19));
    console.log("Sender's preferred vegetable: " + asyncResult.value.match(/x-preferred-vegetable:.*/gim)[0].slice(23));
  } else {
    console.log("Error getting preferences from header: " + JSON.stringify(asyncResult.error));
  }
}

/* Sample output:
Sender's preferred fruit: orange
Sender's preferred vegetable: broccoli
*/
```

> [!IMPORTANT]
> Este exemplo funciona para casos simples. Para obter uma recuperação de informações mais complexa (por exemplo, cabeçalhos de várias instâncias ou valores dobrados, conforme descrito em [RFC 2822](https://tools.ietf.org/html/rfc2822)), tente usar uma biblioteca de análise MIME apropriada.

## <a name="recommended-practices"></a>Práticas recomendadas

Atualmente, os cabeçalhos da Internet são um recurso finito na caixa de correio de um usuário. Quando a cota estiver esgotada, você não poderá criar mais cabeçalhos de Internet nessa caixa de correio, o que pode resultar em comportamento inesperado de clientes que dependem disso para funcionar.

Aplique as diretrizes a seguir ao criar cabeçalhos da Internet em seu suplemento.

- Crie o número mínimo de cabeçalhos necessários. A cota de cabeçalho baseia-se no tamanho total dos cabeçalhos aplicados a uma mensagem. No Exchange Online, o limite de cabeçalho é limitado a 256 KB, enquanto em um ambiente do Exchange local, o limite é determinado pelo administrador da sua organização. Para obter mais informações sobre limites de cabeçalho, [consulte Exchange Online de](/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#message-limits) mensagem e [Exchange Server de mensagem](/exchange/mail-flow/message-size-limits).
- Nomeie os cabeçalhos para que você possa reutilizar e atualizar seus valores mais tarde. Dessa forma, evite nomear cabeçalhos de maneira variável (por exemplo, com base na entrada do usuário, carimbo de data/hora etc.).

## <a name="see-also"></a>Confira também

- [Obter e definir metadados de suplemento para um suplemento do Outlook](metadata-for-an-outlook-add-in.md)
