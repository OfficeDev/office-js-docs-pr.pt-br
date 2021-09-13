---
title: Obter e definir os headers da Internet
description: Como obter e definir os headers da Internet em uma mensagem em um Outlook de um complemento.
ms.date: 04/28/2020
ms.localizationpriority: medium
ms.openlocfilehash: 9784ef16c70e273e6bd1c242ffe91d97aa5d40ed
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149001"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a>Obter e definir os headers da Internet em uma mensagem em um Outlook de um Outlook de dados

## <a name="background"></a>Histórico

Um requisito comum no Outlook desenvolvimento de complementos é armazenar propriedades personalizadas associadas a um complemento em diferentes níveis. Atualmente, as propriedades personalizadas são armazenadas no nível do item ou da caixa de correio.

- Nível do item - Para propriedades que se aplicam a um item específico, use o [objeto CustomProperties.](/javascript/api/outlook/office.customproperties) Por exemplo, armazene um código de cliente associado à pessoa que enviou o email.
- Nível da caixa de correio - Para propriedades que se aplicam a todos os itens de email na caixa de correio do usuário, use o [objeto RoamingSettings.](/javascript/api/outlook/office.roamingsettings) Por exemplo, armazene a preferência de um usuário para mostrar a temperatura em uma escala específica.

Ambos os tipos de propriedades não são preservados depois que o item deixa o servidor Exchange para que os destinatários de email não possam obter nenhuma propriedade definida no item. Portanto, os desenvolvedores não podem acessar essas configurações ou outras propriedades MIME para habilitar cenários de leitura melhores.

Embora haja uma maneira de definir os headers da Internet por meio de solicitações EWS, em alguns cenários fazer uma solicitação EWS não funcionará. Por exemplo, no modo Redação Outlook área de trabalho, a id do item não é sincronizada  `saveAsync`   no modo em cache.

> [!TIP]
> Consulte [Obter e definir metadados](metadata-for-an-outlook-add-in.md) do Outlook de um Outlook para saber mais sobre como usar essas opções.

## <a name="purpose-of-the-internet-headers-api"></a>Finalidade da API de headers da Internet

Introduzido no [conjunto de requisitos 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md), as APIs de headers da Internet permitem que os desenvolvedores:

- Carimbo de informações em um email que persiste depois que ele Exchange em todos os clientes.
- Leia informações sobre um email que persistia depois que o email saiu Exchange todos os clientes em cenários de leitura de email.
- Acesse todo o header MIME do email.

![Diagrama de headers da Internet. Texto: o usuário 1 envia emails. O add-in gerencia os headers personalizados da Internet enquanto o usuário está compondo emails. O usuário 2 recebe o email. O complemento obtém os headers da Internet de emails recebidos e, em seguida, analisados e usa os headers personalizados.](../images/outlook-internet-headers.png)

## <a name="set-internet-headers-while-composing-a-message"></a>Definir os headers da Internet ao compor uma mensagem

Tente usar a [propriedade item.internetHeaders](/javascript/api/outlook/office.messagecompose#internetHeaders) para gerenciar os headers personalizados da Internet que você coloca na mensagem atual no modo Redação.

### <a name="set-get-and-remove-custom-headers-example"></a>Definir, obter e remover exemplo de headers personalizados

O exemplo a seguir mostra como definir, obter e remover os headers personalizados.

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

## <a name="get-internet-headers-while-reading-a-message"></a>Obter os headers da Internet durante a leitura de uma mensagem

Tente chamar [item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getAllInternetHeadersAsync_options__callback_) para obter os headers da Internet na mensagem atual no modo de leitura.

### <a name="get-sender-preferences-from-current-mime-headers-example"></a>Obter preferências de remetente do exemplo de headers MIME atuais

Com base no exemplo da seção anterior, o código a seguir mostra como obter as preferências do remetente a partir dos headers MIME do email atual.

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
> Este exemplo funciona para casos simples. Para recuperação de informações mais complexas (por exemplo, headers de várias instâncias ou valores dobrados conforme descrito em [RFC 2822](https://tools.ietf.org/html/rfc2822)), tente usar uma biblioteca mime-parsing apropriada.

## <a name="recommended-practices"></a>Práticas recomendadas

Atualmente, os headers da Internet são um recurso finito na caixa de correio de um usuário. Quando a cota estiver esgotada, você não poderá criar mais nenhum headers da Internet nessa caixa de correio, o que pode resultar em comportamento inesperado de clientes que dependem disso para funcionar.

Aplique as seguintes diretrizes ao criar os headers da Internet no seu complemento.

- Crie o número mínimo de headers necessário.
- Nomeia os headers para que você possa reutilizar e atualizar seus valores posteriormente. Dessa forma, evite nomear os headers de forma variável (por exemplo, com base na entrada do usuário, no timestamp, etc.).

## <a name="see-also"></a>Confira também

- [Obter e definir metadados de suplemento para um suplemento do Outlook](metadata-for-an-outlook-add-in.md)
