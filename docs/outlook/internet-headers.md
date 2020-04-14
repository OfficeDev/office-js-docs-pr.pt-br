---
title: Obter e definir cabeçalhos de Internet
description: Como obter e definir cabeçalhos da Internet em uma mensagem em um suplemento do Outlook.
ms.date: 04/10/2020
localization_priority: Normal
ms.openlocfilehash: 488a4414580296da59eef3eb703e1c8da7e7d7c2
ms.sourcegitcommit: 231e23d72e04e0536480d6b16df95113f1eff738
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/13/2020
ms.locfileid: "43238210"
---
# <a name="get-and-set-internet-headers-on-a-message-in-an-outlook-add-in"></a><span data-ttu-id="d01aa-103">Obter e definir cabeçalhos de Internet em uma mensagem em um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="d01aa-103">Get and set internet headers on a message in an Outlook add-in</span></span>

## <a name="background"></a><span data-ttu-id="d01aa-104">Segundo plano</span><span class="sxs-lookup"><span data-stu-id="d01aa-104">Background</span></span>

<span data-ttu-id="d01aa-105">Um requisito comum no desenvolvimento de suplementos do Outlook é armazenar propriedades personalizadas associadas a um suplemento em diferentes níveis.</span><span class="sxs-lookup"><span data-stu-id="d01aa-105">A common requirement in Outlook add-ins development is to store custom properties associated with an add-in at different levels.</span></span> <span data-ttu-id="d01aa-106">No momento, as propriedades personalizadas são armazenadas no nível do item ou da caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="d01aa-106">At present, custom properties are stored at the item or mailbox level.</span></span>

- <span data-ttu-id="d01aa-107">Item Level – para propriedades que se aplicam a um item específico, use o objeto [CustomProperties](/javascript/api/outlook/office.customproperties) .</span><span class="sxs-lookup"><span data-stu-id="d01aa-107">Item level - For properties that apply to a specific item, use the [CustomProperties](/javascript/api/outlook/office.customproperties) object.</span></span> <span data-ttu-id="d01aa-108">Por exemplo, armazene um código de cliente associado à pessoa que enviou o email.</span><span class="sxs-lookup"><span data-stu-id="d01aa-108">For example, store a customer code associated with the person who sent the email.</span></span>
- <span data-ttu-id="d01aa-109">Nível de caixa de correio – para propriedades que se aplicam a todos os itens de email da caixa de correio do usuário, use o objeto [RoamingSettings](/javascript/api/outlook/office.roamingsettings) .</span><span class="sxs-lookup"><span data-stu-id="d01aa-109">Mailbox level - For properties that apply to all the mail items in the user's mailbox, use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object.</span></span> <span data-ttu-id="d01aa-110">Por exemplo, armazene a preferência de um usuário para mostrar a temperatura em uma determinada escala.</span><span class="sxs-lookup"><span data-stu-id="d01aa-110">For example, store a user's preference to show the temperature in a particular scale.</span></span>

<span data-ttu-id="d01aa-111">Os dois tipos de propriedades não são preservados depois que o item deixa o servidor do Exchange para que os destinatários de email não possam obter nenhuma propriedade definida no item.</span><span class="sxs-lookup"><span data-stu-id="d01aa-111">Both types of properties are not preserved after the item leaves the Exchange server so the email recipients can't get any properties set on the item.</span></span> <span data-ttu-id="d01aa-112">Portanto, os desenvolvedores não podem acessar essas configurações ou outras propriedades de MIME para permitir melhores cenários de leitura.</span><span class="sxs-lookup"><span data-stu-id="d01aa-112">Therefore, developers can't access those settings or other MIME properties to enable better read scenarios.</span></span>

<span data-ttu-id="d01aa-113">Embora haja uma maneira de definir os cabeçalhos da Internet por meio de solicitações EWS, em alguns cenários, a solicitação do EWS não funcionará.</span><span class="sxs-lookup"><span data-stu-id="d01aa-113">While there's a way for you to set the internet headers through EWS requests, in some scenarios making an EWS request won't work.</span></span> <span data-ttu-id="d01aa-114">Por exemplo, no modo de redação na área de trabalho do Outlook, a ID do `saveAsync` item não é sincronizada no modo em cache.</span><span class="sxs-lookup"><span data-stu-id="d01aa-114">For example, in Compose mode on Outlook desktop, the item id isn't synced on `saveAsync` in cached mode.</span></span>

> [!TIP]
> <span data-ttu-id="d01aa-115">Confira [obter e definir metadados de suplemento para um suplemento do Outlook](metadata-for-an-outlook-add-in.md) para saber mais sobre como usar essas opções.</span><span class="sxs-lookup"><span data-stu-id="d01aa-115">See [Get and set add-in metadata for an Outlook add-in](metadata-for-an-outlook-add-in.md) to learn more about using these options.</span></span>

## <a name="purpose-of-the-internet-headers-api"></a><span data-ttu-id="d01aa-116">Propósito da API de cabeçalhos de Internet</span><span class="sxs-lookup"><span data-stu-id="d01aa-116">Purpose of the internet headers API</span></span>

<span data-ttu-id="d01aa-117">Introduzido no conjunto de requisitos 1,8, as APIs de cabeçalhos da Internet permitem que os desenvolvedores:</span><span class="sxs-lookup"><span data-stu-id="d01aa-117">Introduced in requirement set 1.8, the internet headers APIs enable developers to:</span></span>

- <span data-ttu-id="d01aa-118">Informações de carimbo em um email que persiste depois de deixar o Exchange entre todos os clientes.</span><span class="sxs-lookup"><span data-stu-id="d01aa-118">Stamp information on an email that persists after it leaves Exchange across all clients.</span></span>
- <span data-ttu-id="d01aa-119">Leia as informações em um email que persistiram depois que o email deixou o Exchange entre todos os clientes em cenários de leitura de email.</span><span class="sxs-lookup"><span data-stu-id="d01aa-119">Read information on an email that persisted after the email left Exchange across all clients in mail read scenarios.</span></span>
- <span data-ttu-id="d01aa-120">Acessar o cabeçalho MIME inteiro do email.</span><span class="sxs-lookup"><span data-stu-id="d01aa-120">Access the entire MIME header of the email.</span></span>

![<span data-ttu-id="d01aa-121">Diagrama de cabeçalhos de Internet.</span><span class="sxs-lookup"><span data-stu-id="d01aa-121">Diagram of internet headers.</span></span> <span data-ttu-id="d01aa-122">Text: o usuário 1 envia email.</span><span class="sxs-lookup"><span data-stu-id="d01aa-122">Text: User 1 sends email.</span></span> <span data-ttu-id="d01aa-123">O suplemento gerencia cabeçalhos de Internet personalizados enquanto o usuário está redigindo email.</span><span class="sxs-lookup"><span data-stu-id="d01aa-123">Add-in manages custom internet headers while user is composing email.</span></span> <span data-ttu-id="d01aa-124">O usuário 2 recebe o email.</span><span class="sxs-lookup"><span data-stu-id="d01aa-124">User 2 receives the email.</span></span> <span data-ttu-id="d01aa-125">O suplemento Obtém cabeçalhos de Internet de emails recebidos e, em seguida, analisa e usa cabeçalhos personalizados.</span><span class="sxs-lookup"><span data-stu-id="d01aa-125">Add-in gets internet headers from received email then parses and uses custom headers.</span></span> ](../images/outlook-internet-headers.png)

## <a name="set-internet-headers-while-composing-a-message"></a><span data-ttu-id="d01aa-126">Definir cabeçalhos de Internet ao redigir uma mensagem</span><span class="sxs-lookup"><span data-stu-id="d01aa-126">Set internet headers while composing a message</span></span>

<span data-ttu-id="d01aa-127">Tente usar a propriedade [Item. internetheaders:](/javascript/api/outlook/office.messagecompose#internetheaders) para gerenciar os cabeçalhos de Internet personalizados que você coloca na mensagem atual no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="d01aa-127">Try using the [item.internetHeaders](/javascript/api/outlook/office.messagecompose#internetheaders) property to manage the custom internet headers you place on the current message in Compose mode.</span></span>

### <a name="set-get-and-remove-custom-headers-example"></a><span data-ttu-id="d01aa-128">Exemplo dos cabeçalhos set, Get e remove customes</span><span class="sxs-lookup"><span data-stu-id="d01aa-128">Set, get, and remove custom headers example</span></span>

<span data-ttu-id="d01aa-129">O exemplo a seguir mostra como definir, obter e remover cabeçalhos personalizados.</span><span class="sxs-lookup"><span data-stu-id="d01aa-129">The following example shows how to set, get, and remove custom headers.</span></span>

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

## <a name="get-internet-headers-while-reading-a-message"></a><span data-ttu-id="d01aa-130">Obter cabeçalhos de Internet ao ler uma mensagem</span><span class="sxs-lookup"><span data-stu-id="d01aa-130">Get internet headers while reading a message</span></span>

<span data-ttu-id="d01aa-131">Tente chamar [Item. getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-) para obter cabeçalhos da Internet na mensagem atual no modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="d01aa-131">Try calling [item.getAllInternetHeadersAsync](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-) to get internet headers on the current message in Read mode.</span></span>

### <a name="get-sender-preferences-from-current-mime-headers-example"></a><span data-ttu-id="d01aa-132">Exemplo de obter as preferências de remetente dos cabeçalhos MIME atuais</span><span class="sxs-lookup"><span data-stu-id="d01aa-132">Get sender preferences from current MIME headers example</span></span>

<span data-ttu-id="d01aa-133">Com base no exemplo da seção anterior, o código a seguir mostra como obter as preferências do remetente dos cabeçalhos MIME do email atual.</span><span class="sxs-lookup"><span data-stu-id="d01aa-133">Building on the example from the previous section, the following code shows how to get the sender's preferences from the current email's MIME headers.</span></span>

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
> <span data-ttu-id="d01aa-134">Este exemplo funciona para casos simples.</span><span class="sxs-lookup"><span data-stu-id="d01aa-134">This sample works for simple cases.</span></span> <span data-ttu-id="d01aa-135">Para recuperação de informações mais complexas (por exemplo, cabeçalhos de várias instâncias ou valores dobrados conforme descrito na [RFC 2822](https://tools.ietf.org/html/rfc2822)), tente usar uma biblioteca de análise de MIME apropriada.</span><span class="sxs-lookup"><span data-stu-id="d01aa-135">For more complex information retrieval (e.g., multi-instance headers or folded values as described in [RFC 2822](https://tools.ietf.org/html/rfc2822)), try using an appropriate MIME-parsing library.</span></span>

## <a name="see-also"></a><span data-ttu-id="d01aa-136">Confira também</span><span class="sxs-lookup"><span data-stu-id="d01aa-136">See also</span></span>

- [<span data-ttu-id="d01aa-137">Obter e definir metadados de suplemento para um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="d01aa-137">Get and set add-in metadata for an Outlook add-in</span></span>](metadata-for-an-outlook-add-in.md)
