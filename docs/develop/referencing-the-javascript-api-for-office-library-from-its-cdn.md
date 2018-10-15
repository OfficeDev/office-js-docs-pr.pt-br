---
title: Referenciando a biblioteca da API JavaScript do Office a partir de sua rede de distribuição de conteúdo (CDN)
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 0ad589ee98342ee72259cddc0957277e9018f186
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505416"
---
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a><span data-ttu-id="d3b93-102">Referenciando a biblioteca da API JavaScript do Office a partir de sua rede de distribuição de conteúdo (CDN)</span><span class="sxs-lookup"><span data-stu-id="d3b93-102">Referencing the JavaScript API for Office library from its content delivery network (CDN)</span></span>

> [!NOTE]
> <span data-ttu-id="d3b93-p101">Além das etapas descritas neste artigo, se você quiser usar TypeScript e depois obter o Intellisense, você precisará executar o comando a seguir em um prompt de sistema habilitado para Node (ou janela git bash) a partir da raiz da pasta do seu projeto. Você precisa ter [Node.js](https://nodejs.org) instalado (incluindo npm).</span><span class="sxs-lookup"><span data-stu-id="d3b93-p101">In addition to the steps described in this article, if you want to use TypeScript, then to get Intellisense you will need run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder. You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>
> 
> ```
> npm install --save-dev @types/office-js
> ```

<span data-ttu-id="d3b93-105">A biblioteca da [API JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) consiste no arquivo Office.js e nos arquivos associados .js específicos do aplicativo de host, como Excel-15.js e Outlook-15.js.</span><span class="sxs-lookup"><span data-stu-id="d3b93-105">The [JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js.</span></span> 


<span data-ttu-id="d3b93-106">A maneira mais simples de fazer referência à API é usar nossa CDN adicionando o seguinte `<script>` à tag `<head>` da sua página:</span><span class="sxs-lookup"><span data-stu-id="d3b93-106">The simplest way to reference the API is to use our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

<span data-ttu-id="d3b93-p102">O `/1/` antes de `office.js` na URL da CDN especifica a versão incremental mais recente na versão 1 do Office .js. Como a API JavaScript para Office mantém a compatibilidade com versões anteriores, a última versão continuará a dar suporte a membros da API que foram introduzidos anteriormente na versão 1. Se você precisar atualizar um projeto existente, confira [Atualizar a versão da API JavaScript para Office e os arquivos de esquema de manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span><span class="sxs-lookup"><span data-stu-id="d3b93-p102">The  `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js. Because the JavaScript API for Office maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1. If you need to upgrade an existing project, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="d3b93-p103">Caso planeje publicar seu Suplemento do Office no AppSource, você deve usar esta referência da CDN. As referências locais são adequadas somente para cenários internos, de depuração e de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="d3b93-p103">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!IMPORTANT]
>  <span data-ttu-id="d3b93-p104">Ao desenvolver um suplemento para qualquer aplicativo host do Office, faça referência à API JavaScript para Office de dentro da seção `<head>` da página. Isso garante que a API seja totalmente inicializada antes de qualquer elemento de corpo. Os hosts do Office requerem que os suplementos inicializem até 5 segundos depois da ativação. Se seu suplemento não ativar dentro deste limite, ele será declarado sem resposta e uma mensagem de erro será exibida ao usuário.</span><span class="sxs-lookup"><span data-stu-id="d3b93-p104">When you develop an add-in for any Office host application, reference the JavaScript API for Office from inside the `<head>` section of the page. This ensures that the API is fully initialized prior to any body elements. Office hosts require that add-ins initialize within 5 seconds of activation. If your add-in doesn't activate within this threshold, it will be declared unresponsive and an error message will be displayed to the user.</span></span>       

## <a name="see-also"></a><span data-ttu-id="d3b93-116">Veja também</span><span class="sxs-lookup"><span data-stu-id="d3b93-116">See also</span></span>

- [<span data-ttu-id="d3b93-117">Noções básicas da API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="d3b93-117">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)    
- [<span data-ttu-id="d3b93-118">API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="d3b93-118">JavaScript API for Office</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js)
    
