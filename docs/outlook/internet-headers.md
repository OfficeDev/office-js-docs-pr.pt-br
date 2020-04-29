---
title: Obter e definir cabeçalhos de Internet
description: Como obter e definir cabeçalhos da Internet em uma mensagem em um suplemento do Outlook.
ms.date: 04/28/2020
localization_priority: Normal
ms.openlocfilehash: 1b6bdbbe77998ce92ea1b1b43874a32a30aa160a
ms.sourcegitcommit: 0fdb78cefa669b727b817614a4147a46d249a0ed
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/28/2020
ms.locfileid: "43930285"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a>Obter e definir cabeçalhos de Internet em uma mensagem em um suplemento do Outlook

## <a name="background"></a>Segundo plano

Um requisito comum no desenvolvimento de suplementos do Outlook é armazenar propriedades personalizadas associadas a um suplemento em diferentes níveis. No momento, as propriedades personalizadas são armazenadas no nível do item ou da caixa de correio.

- Item Level – para propriedades que se aplicam a um item específico, use o objeto [CustomProperties](/javascript/api/outlook/office.customproperties) . Por exemplo, armazene um código de cliente associado à pessoa que enviou o email.
- Nível de caixa de correio – para propriedades que se aplicam a todos os itens de email da caixa de correio do usuário, use o objeto [RoamingSettings](/javascript/api/outlook/office.roamingsettings) . Por exemplo, armazene a preferência de um usuário para mostrar a temperatura em uma determinada escala.

Os dois tipos de propriedades não são preservados depois que o item deixa o servidor do Exchange para que os destinatários de email não possam obter nenhuma propriedade definida no item. Portanto, os desenvolvedores não podem acessar essas configurações ou outras propriedades de MIME para permitir melhores cenários de leitura.

Embora haja uma maneira de definir os cabeçalhos da Internet por meio de solicitações EWS, em alguns cenários, a solicitação do EWS não funcionará. Por exemplo, no modo de redação na área de trabalho do Outlook, a ID do `saveAsync` item não é sincronizada no modo em cache.

> [!TIP]
> Confira [obter e definir metadados de suplemento para um suplemento do Outlook](metadata-for-an-outlook-add-in.md) para saber mais sobre como usar essas opções.

## <a name="purpose-of-the-internet-headers-api"></a>Propósito da API de cabeçalhos de Internet

Introduzido no [conjunto de requisitos 1,8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md), as APIs de cabeçalhos da Internet permitem que os desenvolvedores:

- Informações de carimbo em um email que persiste depois de deixar o Exchange entre todos os clientes.
- Leia as informações em um email que persistiram depois que o email deixou o Exchange entre todos os clientes em cenários de leitura de email.
- Acessar o cabeçalho MIME inteiro do email.

![Diagrama de cabeçalhos de Internet. Text: o usuário 1 envia email. O suplemento gerencia cabeçalhos de Internet personalizados enquanto o usuário está redigindo email. O usuário 2 recebe o email. O suplemento Obtém cabeçalhos de Internet de emails recebidos e, em seguida, analisa e usa cabeçalhos personalizados.](../images/outlook-internet-headers.png)

## <a name="set-internet-headers-while-composing-a-message"></a>Definir cabeçalhos de Internet ao redigir uma mensagem

Tente usar a propriedade [Item. internetheaders:](/javascript/api/outlook/office.messagecompose#internetheaders) para gerenciar os cabeçalhos de Internet personalizados que você coloca na mensagem atual no modo de composição.

### <a name="set-get-and-remove-custom-headers-example"></a>Exemplo dos cabeçalhos set, Get e remove customes

O exemplo a seguir mostra como definir, obter e remover cabeçalhos personalizados.

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

## <a name="get-internet-headers-while-reading-a-message"></a>Obter cabeçalhos de Internet ao ler uma mensagem

Tente chamar [Item. getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-) para obter cabeçalhos da Internet na mensagem atual no modo de leitura.

### <a name="get-sender-preferences-from-current-mime-headers-example"></a>Exemplo de obter as preferências de remetente dos cabeçalhos MIME atuais

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
> Este exemplo funciona para casos simples. Para obter recuperação de informações mais complexas (por exemplo, cabeçalhos de várias instâncias ou valores dobrados conforme descrito na [RFC 2822](https://tools.ietf.org/html/rfc2822)), tente usar uma biblioteca de análise de MIME apropriada.

## <a name="recommended-practices"></a>Práticas recomendadas

No momento, os cabeçalhos da Internet são um recurso finito da caixa de correio de um usuário. Quando a cota estiver esgotada, você não poderá criar mais cabeçalhos de Internet nessa caixa de correio, o que pode resultar em um comportamento inesperado dos clientes que dependem disso para funcionar.

Aplique as seguintes diretrizes ao criar cabeçalhos de Internet no suplemento.

- Crie o número mínimo de cabeçalhos necessários.
- Cabeçalhos de nome para que você possa reutilizar e atualizar seus valores posteriormente. Como tal, evite nomes de cabeçalhos de forma variável (por exemplo, com base na entrada do usuário, carimbo de data/hora, etc.).

## <a name="see-also"></a>Confira também

- [Obter e definir metadados de suplemento para um suplemento do Outlook](metadata-for-an-outlook-add-in.md)
