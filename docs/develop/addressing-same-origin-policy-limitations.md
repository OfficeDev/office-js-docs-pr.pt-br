---
title: Como lidar com limitações de política de mesma origem nos Suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: e5aa329eb3f073f3544d8446683debed3239fd00
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270597"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a>Como lidar com limitações de política de mesma origem nos Suplementos do Office


A política de mesma origem imposta pelo navegador impede que um script carregado de um domínio obtenha ou manipule propriedades de uma página da Web de outro domínio. Isso significa que, por padrão, o domínio de uma URL solicitada deve ser igual ao domínio da página da Web atual. Por exemplo, esta política impedirá que uma página da Web de um domínio faça chamadas de serviços Web [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) para um domínio diferente do qual ela está hospedada.

Como os Suplementos do Office estão hospedados em um controle do navegador, a política de mesma origem também se aplica a script em execução em suas páginas da Web.

Para superar a aplicação da política de mesma origem ao desenvolver suplementos, você pode:

- Usar JSON/P para acesso anônimo. 
    
- Implementar o script do lado do servidor usando um esquema de autenticação baseado no token.
    
- Usar o CORS (compartilhamento de recursos entre origens).
    
- Crie seu próprio proxy usando IFRAME e PUBLICAR MENSAGEM.
    

## <a name="using-jsonp-for-anonymous-access"></a>Usar JSON/P para acesso anônimo


Uma maneira de superar essa limitação é usar o JSON/P para fornecer um proxy para o serviço da Web. Faça isso incluindo uma marca `script` com um atributo `src` que aponta para alguns scripts hospedados em qualquer domínio. Você pode criar programaticamente as marcações `script`, criar dinamicamente a URL para a qual apontar o atributo `src`, em seguida, passar parâmetros para a URL por meio dos parâmetros da consulta URI. Os provedores de serviços Web criam e hospedam o código JavaScript em URLs específicas e retornam scripts diferentes, dependendo dos parâmetros de consulta da URI. Em seguida, esses scripts serão executados onde estiverem inseridos e funcionarão como esperado.

Veja a seguir um exemplo de JSON/P que usa uma técnica que funcionará em qualquer Suplemento do Office.

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


## <a name="implementing-server-side-script-using-a-token-based-authentication-scheme"></a>Implementar o script do lado do servidor usando um esquema de autenticação com base no token


Outra maneira para resolver as limitações de política de mesma origem é implementar a página da Web do suplemento como uma página ASP que usa o OAuth ou armazena em cache credenciais em cookies.

Para obter um exemplo de código do lado do servidor que mostre como usar o objeto `Cookie` em `System.Net` para obter e definir valores de cookies, confira a propriedade [Value](https://docs.microsoft.com/dotnet/api/system.net.cookie.value?view=netframework-4.7.2).


## <a name="using-cross-origin-resource-sharing-cors"></a>Usar o CORS (compartilhamento de recursos entre origens)


Para obter um exemplo de como usar o compartilhamento de recursos entre origens do [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), veja a seção "CORS (Compartilhamento de Recursos entre Origens)" de [Novos Truques no XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).


## <a name="building-your-own-proxy-using-iframe-and-post-message"></a>Criar seu próprio proxy usando IFRAME e PUBLICAR MENSAGEM


Confira um exemplo de como criar seu próprio proxy usando IFRAME e PUBLICAR MENSAGEM em [Mensagens entre janelas](http://ejohn.org/blog/cross-window-messaging/).


## <a name="see-also"></a>Veja também

- [Privacidade e segurança para Suplementos do Office](../concepts/privacy-and-security.md)
    
