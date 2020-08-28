---
title: Como lidar com limitações de política de mesma origem nos Suplementos do Office
description: Saiba como acomodar limitações de política de mesma origem com JSONP, CORS, IFRAMEs e outras técnicas.
ms.date: 10/17/2019
localization_priority: Normal
ms.openlocfilehash: e50292c30d77856c896f892c930038c1e19d7af7
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293335"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a><span data-ttu-id="26fc7-103">Como lidar com limitações de política de mesma origem nos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="26fc7-103">Addressing same-origin policy limitations in Office Add-ins</span></span>

<span data-ttu-id="26fc7-p101">A política de mesma origem imposta pelo navegador impede que um script carregado de um domínio obtenha ou manipule propriedades de uma página da Web de outro domínio. Isso significa que, por padrão, o domínio de uma URL solicitada deve ser igual ao domínio da página da Web atual. Por exemplo, esta política impedirá que uma página da Web de um domínio faça chamadas de serviços Web [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) para um domínio diferente do qual ela está hospedada.</span><span class="sxs-lookup"><span data-stu-id="26fc7-p101">The same-origin policy enforced by the browser prevents a script loaded from one domain from getting or manipulating properties of a webpage from another domain. This means that, by default, the domain of a requested URL must be the same as the domain of the current webpage. For example, this policy will prevent a webpage in one domain from making [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) web-service calls to a domain other than the one where it is hosted.</span></span>

<span data-ttu-id="26fc7-107">Como os Suplementos do Office estão hospedados em um controle do navegador, a política de mesma origem também se aplica a script em execução em suas páginas da Web.</span><span class="sxs-lookup"><span data-stu-id="26fc7-107">Because Office Add-ins are hosted in a browser control, the same-origin policy applies to script running in their web pages as well.</span></span>

<span data-ttu-id="26fc7-108">A política de mesma origem pode ser um deficiente em várias situações, como quando um aplicativo web hospeda conteúdo e as APIs em vários subdomínios desnecessários.</span><span class="sxs-lookup"><span data-stu-id="26fc7-108">The same-origin policy can be an unnecessary handicap in many situations, such as when a web application hosts content and APIs across multiple subdomains.</span></span> <span data-ttu-id="26fc7-109">Há algumas técnicas comuns para superar com segurança a imposição da política de mesma origem.</span><span class="sxs-lookup"><span data-stu-id="26fc7-109">There are a few common techniques for securely overcoming same-origin policy enforcement.</span></span> <span data-ttu-id="26fc7-110">Este artigo pode fornecer somente uma breve introdução de alguns deles.</span><span class="sxs-lookup"><span data-stu-id="26fc7-110">This article can only provide the briefest introduction to some of them.</span></span> <span data-ttu-id="26fc7-111">Use os links fornecidos para começar a usar a pesquisa destas técnicas.</span><span class="sxs-lookup"><span data-stu-id="26fc7-111">Please use the links provided to get started in your research of these techniques.</span></span>

## <a name="use-jsonp-for-anonymous-access"></a><span data-ttu-id="26fc7-112">Use JSONP para acesso anônimo</span><span class="sxs-lookup"><span data-stu-id="26fc7-112">Use JSONP for anonymous access</span></span>

<span data-ttu-id="26fc7-113">Uma maneira de superar essa limitação da política de mesma origem é usar o [JSONP](https://www.w3schools.com/js/js_json_jsonp.asp) para fornecer um proxy para o serviço da Web.</span><span class="sxs-lookup"><span data-stu-id="26fc7-113">One way to overcome same-origin policy limitations is to use [JSONP](https://www.w3schools.com/js/js_json_jsonp.asp) to provide a proxy for the web service.</span></span> <span data-ttu-id="26fc7-114">Faça isso incluindo uma marca `script` com um atributo `src` que aponta para alguns scripts hospedados em qualquer domínio.</span><span class="sxs-lookup"><span data-stu-id="26fc7-114">You do this by including a `script` tag with a `src` attribute that points to some script hosted on any domain.</span></span> <span data-ttu-id="26fc7-115">Você pode criar as marcas `script`, criar dinamicamente a URL para apontar para o atributo `src` e passar parâmetros para a URL por meio de parâmetros de consulta de URI.</span><span class="sxs-lookup"><span data-stu-id="26fc7-115">You can programmatically create the `script` tags, dynamically create the URL to point the `src` attribute to, and then pass parameters to the URL via URI query parameters.</span></span> <span data-ttu-id="26fc7-116">Os provedores de serviços Web criam e hospedam o código JavaScript em URLs específicas e retornam scripts diferentes, dependendo dos parâmetros de consulta de URI.</span><span class="sxs-lookup"><span data-stu-id="26fc7-116">Web service providers create and host JavaScript code at specific URLs, and return different scripts depending on the URI query parameters.</span></span> <span data-ttu-id="26fc7-117">Em seguida, esses scripts serão executados onde estiverem inseridos e funcionarão como esperado.</span><span class="sxs-lookup"><span data-stu-id="26fc7-117">These scripts then execute where they are inserted and work as expected.</span></span>

<span data-ttu-id="26fc7-118">Veja a seguir um exemplo de JSONP que usa uma técnica que funcionará em qualquer Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="26fc7-118">The following is an example of JSONP that uses a technique that will work in any Office Add-in.</span></span>

```js
// Dynamically create an HTML SCRIPT element that obtains the details for the specified video.
function loadVideoDetails(videoIndex) {
    // Dynamically create a new HTML SCRIPT element in the webpage.
    var script = document.createElement("script");
    // Specify the URL to retrieve the indicated video from a feed of a current list of videos,
    // as the value of the src attribute of the SCRIPT element. 
    script.setAttribute("src", "https://gdata.youtube.com/feeds/api/videos/" + 
        videos[videoIndex].Id + "?alt=json-in-script&amp;callback=videoDetailsLoaded");
    // Insert the SCRIPT element at the end of the HEAD section.
    document.getElementsByTagName('head')[0].appendChild(script);
}

```


## <a name="implement-server-side-code-using-a-token-based-authorization-scheme"></a><span data-ttu-id="26fc7-119">Implemente o código servidor usando um esquema de autorização do token</span><span class="sxs-lookup"><span data-stu-id="26fc7-119">Implement server-side code using a token-based authorization scheme</span></span>

<span data-ttu-id="26fc7-120">Outra maneira de resolver limitações de política de mesma origem é fornecer o código no servidor que usa fluxos [OAuth 2.0](https://oauth.net/2/) para habilitar um domínio a obter acesso autorizado a recursos hospedado em outro domínio.</span><span class="sxs-lookup"><span data-stu-id="26fc7-120">Another way to address same-origin policy limitations is to provide server-side code that uses [OAuth 2.0](https://oauth.net/2/) flows to enable one domain to get authorized access to resources hosted on another.</span></span> 


## <a name="use-cross-origin-resource-sharing-cors"></a><span data-ttu-id="26fc7-121">Use o CORS (compartilhamento de recursos entre origens)</span><span class="sxs-lookup"><span data-stu-id="26fc7-121">Use cross-origin resource sharing (CORS)</span></span>


<span data-ttu-id="26fc7-122">Para obter um exemplo de como usar o compartilhamento de recursos entre origens do [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), veja a seção "CORS (Compartilhamento de Recursos entre Origens)" de [Novos Truques no XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).</span><span class="sxs-lookup"><span data-stu-id="26fc7-122">For an example of using the cross-origin resource sharing feature of [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), see the "Cross Origin Resource Sharing (CORS)" section of [New Tricks in XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).</span></span>


## <a name="build-your-own-proxy-using-iframe-and-post-message-cross-window-messaging"></a><span data-ttu-id="26fc7-123">Criar seu próprio proxy usando IFRAME e PUBLICAR MENSAGENS (Mensagens entre Janelas)</span><span class="sxs-lookup"><span data-stu-id="26fc7-123">Build your own proxy using IFRAME and POST MESSAGE (Cross-Window Messaging)</span></span>


<span data-ttu-id="26fc7-124">Confira um exemplo de como criar seu próprio proxy usando IFRAME e PUBLICAR MENSAGEM em [Mensagens entre janelas](http://ejohn.org/blog/cross-window-messaging/).</span><span class="sxs-lookup"><span data-stu-id="26fc7-124">For an example of how to build your own proxy using IFRAME and POST MESSAGE, see [Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/).</span></span>


## <a name="see-also"></a><span data-ttu-id="26fc7-125">Confira também</span><span class="sxs-lookup"><span data-stu-id="26fc7-125">See also</span></span>

- [<span data-ttu-id="26fc7-126">Privacidade e segurança para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="26fc7-126">Privacy and security for Office Add-ins</span></span>](../concepts/privacy-and-security.md)
    
