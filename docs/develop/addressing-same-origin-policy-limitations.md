---
title: Como lidar com limitações de política de mesma origem nos Suplementos do Office
description: ''
ms.date: 02/08/2019
localization_priority: Priority
ms.openlocfilehash: 52af2eef2881b48feb141182233bc194ae406aa0
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/13/2019
ms.locfileid: "29981989"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a>Como lidar com limitações de política de mesma origem nos Suplementos do Office

A política de mesma origem imposta pelo navegador impede que um script carregado de um domínio obtenha ou manipule propriedades de uma página da Web de outro domínio. Isso significa que, por padrão, o domínio de uma URL solicitada deve ser igual ao domínio da página da Web atual. Por exemplo, esta política impedirá que uma página da Web de um domínio faça chamadas de serviços Web [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) para um domínio diferente do qual ela está hospedada.

Como os Suplementos do Office estão hospedados em um controle do navegador, a política de mesma origem também se aplica a script em execução em suas páginas da Web.

A política de mesma origem pode ser um deficiente em várias situações, como quando um aplicativo web hospeda conteúdo e as APIs em vários subdomínios desnecessários. Há algumas técnicas comuns para superar com segurança a imposição da política de mesma origem. Este artigo pode fornecer somente uma breve introdução de alguns deles. Use os links fornecidos para começar a usar a pesquisa destas técnicas.

## <a name="use-jsonp-for-anonymous-access"></a>Usar JSON/P para acesso anônimo.

Uma maneira de superar essa limitação da política de mesma origem é usar o [JSON/P](https://www.w3schools.com/js/js_json_jsonp.asp) para fornecer um proxy para o serviço da Web. Faça isso incluindo uma marca `script` com um atributo `src` que aponta para alguns scripts hospedados em qualquer domínio. Você pode criar as marcas `script`, criar dinamicamente a URL para apontar para o atributo `src` e passar parâmetros para a URL por meio de parâmetros de consulta de URI. Os provedores de serviços Web criam e hospedam o código JavaScript em URLs específicas e retornam scripts diferentes, dependendo dos parâmetros de consulta de URI. Em seguida, esses scripts serão executados onde estiverem inseridos e funcionarão como esperado.

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


## <a name="implement-server-side-code-using-a-token-based-authorization-scheme"></a>Implemente o código servidor usando um esquema de autorização do token

Outra maneira de resolver limitações de política de mesma origem é fornecer o código no servidor que usa fluxos [OAuth 2.0](https://oauth.net/2/) para habilitar um domínio a obter acesso autorizado a recursos hospedado em outro domínio. 


## <a name="use-cross-origin-resource-sharing-cors"></a>Use o CORS (compartilhamento de recursos entre origens)


Para obter um exemplo de como usar o compartilhamento de recursos entre origens do [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), veja a seção "CORS (Compartilhamento de Recursos entre Origens)" de [Novos Truques no XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).


## <a name="build-your-own-proxy-using-iframe-and-post-message-cross-window-messaging"></a>Criar seu próprio proxy usando IFRAME e PUBLICAR MENSAGENS (Mensagens entre Janelas)


Confira um exemplo de como criar seu próprio proxy usando IFRAME e PUBLICAR MENSAGEM em [Mensagens entre janelas](http://ejohn.org/blog/cross-window-messaging/).


## <a name="see-also"></a>Confira também

- [Privacidade e segurança para Suplementos do Office](../concepts/privacy-and-security.md)
    
