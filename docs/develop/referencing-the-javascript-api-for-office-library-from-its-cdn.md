---
title: Fazendo referência à biblioteca da API JavaScript do Office
description: Saiba como fazer referência à biblioteca Office da API JavaScript e as definições de tipo no seu complemento.
ms.date: 02/18/2021
localization_priority: Normal
ms.openlocfilehash: 04f97412c07cb39f5b2f753c3ce14e56e87c3de5
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349753"
---
# <a name="referencing-the-office-javascript-api-library"></a><span data-ttu-id="318c7-103">Fazendo referência à biblioteca da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="318c7-103">Referencing the Office JavaScript API library</span></span>

<span data-ttu-id="318c7-104">A [Office api JavaScript](../reference/javascript-api-for-office.md) fornece as APIs que o seu complemento pode usar para interagir com o Office aplicativo.</span><span class="sxs-lookup"><span data-stu-id="318c7-104">The [Office JavaScript API](../reference/javascript-api-for-office.md) library provides the APIs that your add-in can use to interact with the Office application.</span></span> <span data-ttu-id="318c7-105">A maneira mais simples de fazer referência à biblioteca é usar a rede de distribuição de conteúdo (CDN) adicionando a marca a seguir na seção `<script>` `<head>` de sua página HTML.</span><span class="sxs-lookup"><span data-stu-id="318c7-105">The simplest way to reference the library is to use the content delivery network (CDN) by adding the following `<script>` tag within the `<head>` section of your HTML page.</span></span>

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

<span data-ttu-id="318c7-106">Isso baixará e armazenará em cache os arquivos da API JavaScript Office primeira vez que o seu complemento for carregado para garantir que ele está usando a implementação mais atualizada do Office.js e seus arquivos associados para a versão especificada.</span><span class="sxs-lookup"><span data-stu-id="318c7-106">This will download and cache the Office JavaScript API files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="318c7-107">Você deve fazer referência Office API JavaScript de dentro da seção da página para garantir que a API seja totalmente inicializada antes `<head>` de qualquer elemento do corpo.</span><span class="sxs-lookup"><span data-stu-id="318c7-107">You must reference the Office JavaScript API from inside the `<head>` section of the page to ensure that the API is fully initialized prior to any body elements.</span></span>

## <a name="api-versioning-and-backward-compatibility"></a><span data-ttu-id="318c7-108">Versão da API e compatibilidade com versões versões</span><span class="sxs-lookup"><span data-stu-id="318c7-108">API versioning and backward compatibility</span></span>

<span data-ttu-id="318c7-109">No trecho HTML anterior, o na frente da URL CDN especifica a versão incremental mais recente na versão 1 do `/1/` `office.js` Office.js.</span><span class="sxs-lookup"><span data-stu-id="318c7-109">In the previous HTML snippet, the `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js.</span></span> <span data-ttu-id="318c7-110">Como a api Office JavaScript mantém a compatibilidade com versões anteriores, a versão mais recente continuará a dar suporte a membros da API introduzidos anteriormente na versão 1.</span><span class="sxs-lookup"><span data-stu-id="318c7-110">Because the Office JavaScript API maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1.</span></span> <span data-ttu-id="318c7-111">Se você precisar atualizar um projeto existente, consulte [Update the version of your Office JAVAScript API and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span><span class="sxs-lookup"><span data-stu-id="318c7-111">If you need to upgrade an existing project, see [Update the version of your Office JavaScript API and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="318c7-p103">Caso planeje publicar seu Suplemento do Office no AppSource, você deve usar esta referência da CDN. As referências locais são adequadas somente para cenários internos, de depuração e de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="318c7-p103">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!NOTE]
> <span data-ttu-id="318c7-114">Para usar APIs de visualização, faça referência à versão de visualização da biblioteca da API JavaScript do Office na CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span><span class="sxs-lookup"><span data-stu-id="318c7-114">To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span></span>

## <a name="enabling-intellisense-for-a-typescript-project"></a><span data-ttu-id="318c7-115">Habil IntelliSense para um projeto TypeScript</span><span class="sxs-lookup"><span data-stu-id="318c7-115">Enabling IntelliSense for a TypeScript project</span></span>

<span data-ttu-id="318c7-116">Além de fazer referência à API JavaScript Office como descrito anteriormente, você também pode habilitar o IntelliSense para o projeto de add-in TypeScript usando as definições de tipo de [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span><span class="sxs-lookup"><span data-stu-id="318c7-116">In addition to referencing the Office JavaScript API as described previously, you can also enable IntelliSense for TypeScript add-in project by using the type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span></span> <span data-ttu-id="318c7-117">Para fazer isso, execute o seguinte comando em um prompt de sistema habilitado para nó (ou janela git bash) na raiz da pasta do projeto.</span><span class="sxs-lookup"><span data-stu-id="318c7-117">To do so, run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder.</span></span> <span data-ttu-id="318c7-118">Você deve ter o [Node.js](https://nodejs.org) instalado (que inclui o npm).</span><span class="sxs-lookup"><span data-stu-id="318c7-118">You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a><span data-ttu-id="318c7-119">APIs de visualização</span><span class="sxs-lookup"><span data-stu-id="318c7-119">Preview APIs</span></span>

<span data-ttu-id="318c7-120">As novas APIs JavaScript são introduzidas pela primeira vez em "visualização" e, posteriormente, tornam-se parte de um conjunto de requisitos numerados específico depois que ocorrem testes suficientes e os comentários do usuário são necessários.</span><span class="sxs-lookup"><span data-stu-id="318c7-120">New JavaScript APIs are first introduced in "preview" and later become part of a specific numbered requirement set after sufficient testing occurs and user feedback is required.</span></span>

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a><span data-ttu-id="318c7-121">Confira também</span><span class="sxs-lookup"><span data-stu-id="318c7-121">See also</span></span>

- [<span data-ttu-id="318c7-122">Entendendo a API de JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="318c7-122">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="318c7-123">API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="318c7-123">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
