---
title: Fazendo referência à biblioteca da API JavaScript do Office
description: Saiba como fazer referência à biblioteca da API JavaScript do Office e definições de tipo no suplemento.
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 9f7753b24e0a5861778b09ea93fecdc26fd2ca96
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325154"
---
# <a name="referencing-the-office-javascript-api-library"></a><span data-ttu-id="3a90b-103">Fazendo referência à biblioteca da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="3a90b-103">Referencing the Office JavaScript API library</span></span>

<span data-ttu-id="3a90b-104">A biblioteca da [API JavaScript do Office](../reference/javascript-api-for-office.md) fornece as APIs que o suplemento pode usar para interagir com o host do Office.</span><span class="sxs-lookup"><span data-stu-id="3a90b-104">The [Office JavaScript API](../reference/javascript-api-for-office.md) library provides the APIs that your add-in can use to interact with the Office host.</span></span> <span data-ttu-id="3a90b-105">A maneira mais simples de fazer referência à biblioteca é usar a CDN (rede de distribuição de conteúdo) adicionando a `<script>` seguinte marca dentro `<head>` da seção da página HTML:</span><span class="sxs-lookup"><span data-stu-id="3a90b-105">The simplest way to reference the library is to use the content delivery network (CDN) by adding the following `<script>` tag within the `<head>` section of your HTML page:</span></span>  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

<span data-ttu-id="3a90b-106">Isso baixará e armazenará em cache os arquivos da API JavaScript do Office na primeira vez em que seu suplemento for carregado para garantir que ele esteja usando a implementação mais atualizada do Office. js e seus arquivos associados para a versão especificada.</span><span class="sxs-lookup"><span data-stu-id="3a90b-106">This will download and cache the Office JavaScript API files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3a90b-107">Você deve fazer referência à API JavaScript do Office de `<head>` dentro da seção da página para garantir que a API seja totalmente inicializada antes de qualquer elemento body.</span><span class="sxs-lookup"><span data-stu-id="3a90b-107">You must reference the Office JavaScript API from inside the `<head>` section of the page to ensure that the API is fully initialized prior to any body elements.</span></span> <span data-ttu-id="3a90b-108">Os hosts do Office requerem que os suplementos inicializem até 5 segundos depois da ativação.</span><span class="sxs-lookup"><span data-stu-id="3a90b-108">Office hosts require that add-ins initialize within 5 seconds of activation.</span></span> <span data-ttu-id="3a90b-109">Se seu suplemento não ativar dentro deste limite, ele será declarado sem resposta e uma mensagem de erro será exibida ao usuário.</span><span class="sxs-lookup"><span data-stu-id="3a90b-109">If your add-in doesn't activate within this threshold, it will be declared unresponsive and an error message will be displayed to the user.</span></span>

## <a name="api-versioning-and-backward-compatibility"></a><span data-ttu-id="3a90b-110">Versão da API e compatibilidade com versões anteriores</span><span class="sxs-lookup"><span data-stu-id="3a90b-110">API versioning and backward compatibility</span></span>

<span data-ttu-id="3a90b-111">No trecho de código HTML anterior, `/1/` o na frente `office.js` da URL de CDN especifica a versão incremental mais recente na versão 1 do Office. js.</span><span class="sxs-lookup"><span data-stu-id="3a90b-111">In the previous HTML snippet, the `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js.</span></span> <span data-ttu-id="3a90b-112">Como a API JavaScript do Office mantém a compatibilidade com versões anteriores, a versão mais recente continuará a dar suporte a membros da API que foram introduzidos anteriormente na versão 1.</span><span class="sxs-lookup"><span data-stu-id="3a90b-112">Because the Office JavaScript API maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1.</span></span> <span data-ttu-id="3a90b-113">Se você precisar atualizar um projeto existente, confira [atualizar a versão da API JavaScript do Office e dos arquivos de esquema de manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span><span class="sxs-lookup"><span data-stu-id="3a90b-113">If you need to upgrade an existing project, see [Update the version of your Office JavaScript API and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="3a90b-p104">Caso planeje publicar seu Suplemento do Office no AppSource, você deve usar esta referência da CDN. As referências locais são adequadas somente para cenários internos, de depuração e de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="3a90b-p104">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!NOTE]
> <span data-ttu-id="3a90b-116">Para usar as APIs de visualização, consulte a versão prévia da biblioteca da API JavaScript do Office na `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`CDN:.</span><span class="sxs-lookup"><span data-stu-id="3a90b-116">To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span></span>

## <a name="enabling-intellisense-for-a-typescript-project"></a><span data-ttu-id="3a90b-117">Habilitando o IntelliSense para um projeto TypeScript</span><span class="sxs-lookup"><span data-stu-id="3a90b-117">Enabling Intellisense for a TypeScript project</span></span>

<span data-ttu-id="3a90b-118">Além de fazer referência à API JavaScript do Office, conforme descrito anteriormente, você também pode habilitar o IntelliSense para o projeto de suplemento do TypeScript usando as definições de tipo do [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span><span class="sxs-lookup"><span data-stu-id="3a90b-118">In addition to referencing the Office JavaScript API as described previously, you can also enable Intellisense for TypeScript add-in project by using the type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span></span> <span data-ttu-id="3a90b-119">Para fazer isso, execute o seguinte comando em um prompt do sistema habilitado para nós (ou janela do git bash) da raiz da pasta do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="3a90b-119">To do so, run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder.</span></span> <span data-ttu-id="3a90b-120">Você deve ter o [Node.js](https://nodejs.org) instalado (que inclui o npm).</span><span class="sxs-lookup"><span data-stu-id="3a90b-120">You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>

```command&nbsp;line
npm install --save-dev @types/office-js
```

> [!NOTE]
> <span data-ttu-id="3a90b-121">Para habilitar o IntelliSense para APIs de visualização, use as definições de tipo de visualização do [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js-preview) executando o seguinte comando na raiz da pasta do seu projeto:</span><span class="sxs-lookup"><span data-stu-id="3a90b-121">To enable Intellisense for preview APIs, use the preview type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js-preview) by running the following command in the root of your project folder:</span></span> 
>
> `npm install --save-dev @types/office-js-preview`

## <a name="see-also"></a><span data-ttu-id="3a90b-122">Confira também</span><span class="sxs-lookup"><span data-stu-id="3a90b-122">See also</span></span>

- [<span data-ttu-id="3a90b-123">Noções básicas sobre a API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="3a90b-123">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="3a90b-124">API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="3a90b-124">Office JavaScript API</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
