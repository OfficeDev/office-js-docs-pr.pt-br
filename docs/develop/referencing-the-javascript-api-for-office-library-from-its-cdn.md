---
title: Fazer referência à biblioteca da API JavaScript para Office de sua CDN (rede de distribuição de conteúdo)
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 6d435c0e3a0383d03f28e095c56b62e967394e22
ms.sourcegitcommit: 7ecc1dc24bf7488b53117d7a83ad60e952a6f7aa
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/23/2018
ms.locfileid: "19437133"
---
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a><span data-ttu-id="13c96-102">Fazer referência à biblioteca da API JavaScript para Office de sua CDN (rede de distribuição de conteúdo)</span><span class="sxs-lookup"><span data-stu-id="13c96-102">Referencing the JavaScript API for Office library from its content delivery network (CDN)</span></span>


<span data-ttu-id="13c96-103">A biblioteca da [API JavaScript para Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) consiste no arquivo Office.js e nos arquivos .js específicos do aplicativo de host associado, como Excel-15.js e Outlook-15.js.</span><span class="sxs-lookup"><span data-stu-id="13c96-103">The [JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js.</span></span> 


<span data-ttu-id="13c96-104">A maneira mais simples de fazer referência à API é usar nossa CDN adicionando o seguinte `<script>` à marca `<head>` da sua página:</span><span class="sxs-lookup"><span data-stu-id="13c96-104">The simplest way to reference the API is to use our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

<span data-ttu-id="13c96-p101">O `/1/` antes de `office.js` da URL da CDN especifica a versão incremental mais recente na versão 1 do Office .js. Como a API JavaScript para Office mantém a compatibilidade com versões anteriores, a última versão continuará a dar suporte a membros da API que foram introduzidos anteriormente na versão 1. Se você precisar atualizar um projeto existente, confira [Atualizar a versão da API JavaScript para Office e os arquivos de esquema de manifesto](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span><span class="sxs-lookup"><span data-stu-id="13c96-p101">The  `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js. Because the JavaScript API for Office maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1. If you need to upgrade an existing project, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="13c96-p102">Caso planeje publicar seu Suplemento do Office no AppSource, você deve usar esta referência da CDN. As referências locais são adequadas somente para cenários internos, de depuração e de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="13c96-p102">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!IMPORTANT]
>  <span data-ttu-id="13c96-p103">Ao desenvolver um suplemento para qualquer aplicativo host do Office, faça referência à API JavaScript para Office de dentro da seção `<head>` da página. Isso garante que a API seja totalmente inicializada antes de qualquer elemento de corpo. Os hosts do Office requerem que os suplementos inicializem até 5 segundos depois da ativação. Se seu suplemento não ativar dentro deste limite, ele será declarado sem resposta e uma mensagem de erro será exibida ao usuário.</span><span class="sxs-lookup"><span data-stu-id="13c96-p103">When you develop an add-in for any Office host application, reference the JavaScript API for Office from inside the `<head>` section of the page. This ensures that the API is fully initialized prior to any body elements. Office hosts require that add-ins initialize within 5 seconds of activation. If your add-in doesn't activate within this threshold, it will be declared unresponsive and an error message will be displayed to the user.</span></span>       

## <a name="see-also"></a><span data-ttu-id="13c96-114">Veja também</span><span class="sxs-lookup"><span data-stu-id="13c96-114">See also</span></span>

- [<span data-ttu-id="13c96-115">Noções básicas da API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="13c96-115">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)    
- [<span data-ttu-id="13c96-116">API JavaScript para Office</span><span class="sxs-lookup"><span data-stu-id="13c96-116">JavaScript API for Office</span></span>](https://dev.office.com/reference/add-ins/javascript-api-for-office)
    
