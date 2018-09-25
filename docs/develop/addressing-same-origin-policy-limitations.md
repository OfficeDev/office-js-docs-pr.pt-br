---
title: Como lidar com limitações de política de mesma origem nos Suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 054a01d554c529579917218361bcb8aeebb04c3c
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004879"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a><span data-ttu-id="228e8-102">Como lidar com limitações de política de mesma origem nos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="228e8-102">Addressing same-origin policy limitations in Office Add-ins</span></span>


<span data-ttu-id="228e8-p101">A política de mesma origem imposta pelo navegador impede que um script carregado de um domínio obtenha ou manipule propriedades de uma página da Web de outro domínio. Isso significa que, por padrão, o domínio de uma URL solicitada deve ser igual ao domínio da página da Web atual. Por exemplo, esta política impedirá que uma página da Web de um domínio faça chamadas de serviços Web [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) para um domínio diferente do qual ela está hospedada.</span><span class="sxs-lookup"><span data-stu-id="228e8-p101">The same-origin policy enforced by the browser prevents a script loaded from one domain from getting or manipulating properties of a webpage from another domain. This means that, by default, the domain of a requested URL must be the same as the domain of the current webpage. For example, this policy will prevent a webpage in one domain from making [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) web-service calls to a domain other than the one where it is hosted.</span></span>

<span data-ttu-id="228e8-106">Como os Suplementos do Office estão hospedados em um controle do navegador, a política de mesma origem também se aplica a script em execução em suas páginas da Web.</span><span class="sxs-lookup"><span data-stu-id="228e8-106">Because Office Add-ins are hosted in a browser control, the same-origin policy applies to script running in their web pages as well.</span></span>

<span data-ttu-id="228e8-107">Para superar a aplicação da política de mesma origem ao desenvolver suplementos, você pode:</span><span class="sxs-lookup"><span data-stu-id="228e8-107">To overcome same-origin policy enforcement when you develop add-ins, you can:</span></span>

- <span data-ttu-id="228e8-108">Usar JSON/P para acesso anônimo.</span><span class="sxs-lookup"><span data-stu-id="228e8-108">Use JSON/P for anonymous access.</span></span> 
    
- <span data-ttu-id="228e8-109">Implementar o script do lado do servidor usando um esquema de autenticação baseado no token.</span><span class="sxs-lookup"><span data-stu-id="228e8-109">Implement server-side script using a token-based authentication scheme.</span></span>
    
- <span data-ttu-id="228e8-110">Usar o CORS (compartilhamento de recursos entre origens).</span><span class="sxs-lookup"><span data-stu-id="228e8-110">Using cross-origin resource sharing (CORS).</span></span>
    
- <span data-ttu-id="228e8-111">Crie seu próprio proxy usando IFRAME e PUBLICAR MENSAGEM.</span><span class="sxs-lookup"><span data-stu-id="228e8-111">Build your own proxy using IFRAME and POST MESSAGE.</span></span>
    

## <a name="using-jsonp-for-anonymous-access"></a><span data-ttu-id="228e8-112">Usar JSON/P para acesso anônimo</span><span class="sxs-lookup"><span data-stu-id="228e8-112">Using JSON/P for anonymous access</span></span>


<span data-ttu-id="228e8-p102">Uma maneira de superar essa limitação é usar o JSON/P para fornecer um proxy para o serviço da Web. Faça isso incluindo uma marca `script` com um atributo `src` que aponta para alguns scripts hospedados em qualquer domínio. Você pode criar programaticamente as marcações `script`, criar dinamicamente a URL para a qual apontar o atributo `src`, em seguida, passar parâmetros para a URL por meio dos parâmetros da consulta URI. Os provedores de serviços Web criam e hospedam o código JavaScript em URLs específicas e retornam scripts diferentes, dependendo dos parâmetros de consulta da URI. Em seguida, esses scripts serão executados onde estiverem inseridos e funcionarão como esperado.</span><span class="sxs-lookup"><span data-stu-id="228e8-p102">One way to overcome this limitation is to use JSON/P to provide a proxy for the web service. You do this by including a `script` tag with a `src` attribute that points to some script hosted on any domain. You can programmatically create the `script` tags, dynamically create the URL to point the `src` attribute to, and then pass parameters to the URL via URI query parameters. Web service providers create and host JavaScript code at specific URLs, and return different scripts depending on the URI query parameters. These scripts then execute where they are inserted and work as expected.</span></span>

<span data-ttu-id="228e8-118">Veja a seguir um exemplo de JSON/P que usa uma técnica que funcionará em qualquer Suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="228e8-118">The following is an example of JSON/P that uses a technique that will work in any Office Add-in.</span></span>

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


## <a name="implementing-server-side-script-using-a-token-based-authentication-scheme"></a><span data-ttu-id="228e8-119">Implementar o script do lado do servidor usando um esquema de autenticação com base no token</span><span class="sxs-lookup"><span data-stu-id="228e8-119">Implementing server-side script using a token-based authentication scheme</span></span>


<span data-ttu-id="228e8-120">Outra maneira para resolver as limitações de política de mesma origem é implementar a página da Web do suplemento como uma página ASP que usa o OAuth ou armazena em cache credenciais em cookies.</span><span class="sxs-lookup"><span data-stu-id="228e8-120">Another way to address same-origin policy limitations is to implement the add-in's webpage as an ASP page that uses OAuth or caches credentials in cookies.</span></span>

<span data-ttu-id="228e8-121">Para obter um exemplo de código do lado do servidor que mostre como usar o objeto `Cookie` em `System.Net` para obter e definir valores de cookies, confira a propriedade [Value](https://docs.microsoft.com/dotnet/api/system.net.cookie.value?view=netframework-4.7.2).</span><span class="sxs-lookup"><span data-stu-id="228e8-121">For an example of server-side code that shows how to use the  `Cookie` object in `System.Net` to get and set cookie values, see the [Value](https://docs.microsoft.com/dotnet/api/system.net.cookie.value?view=netframework-4.7.2) property.</span></span>


## <a name="using-cross-origin-resource-sharing-cors"></a><span data-ttu-id="228e8-122">Usar o CORS (compartilhamento de recursos entre origens)</span><span class="sxs-lookup"><span data-stu-id="228e8-122">Using cross-origin resource sharing (CORS)</span></span>


<span data-ttu-id="228e8-123">Para obter um exemplo de como usar o compartilhamento de recursos entre origens do [XmlHttpRequest2](http://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), veja a seção "CORS (Compartilhamento de Recursos entre Origens)" de [Novos Truques no XMLHttpRequest2](http://www.html5rocks.com/en/tutorials/file/xhr2/).</span><span class="sxs-lookup"><span data-stu-id="228e8-123">For an example of using the cross-origin resource sharing feature of [XmlHttpRequest2](http://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), see the "Cross Origin Resource Sharing (CORS)" section of [New Tricks in XMLHttpRequest2](http://www.html5rocks.com/en/tutorials/file/xhr2/).</span></span>


## <a name="building-your-own-proxy-using-iframe-and-post-message"></a><span data-ttu-id="228e8-124">Criar seu próprio proxy usando IFRAME e PUBLICAR MENSAGEM</span><span class="sxs-lookup"><span data-stu-id="228e8-124">Building your own proxy using IFRAME and POST MESSAGE</span></span>


<span data-ttu-id="228e8-125">Confira um exemplo de como criar seu próprio proxy usando IFRAME e PUBLICAR MENSAGEM em [Mensagens entre janelas](http://ejohn.org/blog/cross-window-messaging/).</span><span class="sxs-lookup"><span data-stu-id="228e8-125">For an example of how to build your own proxy using IFRAME and POST MESSAGE, see [Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/).</span></span>


## <a name="see-also"></a><span data-ttu-id="228e8-126">Veja também</span><span class="sxs-lookup"><span data-stu-id="228e8-126">See also</span></span>

- [<span data-ttu-id="228e8-127">Privacidade e segurança para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="228e8-127">Privacy and security for Office Add-ins</span></span>](../concepts/privacy-and-security.md)
    
