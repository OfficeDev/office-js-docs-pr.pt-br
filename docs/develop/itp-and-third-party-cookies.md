---
title: Desenvolva seu Complemento do Office para trabalhar com ITP ao usar cookies de terceiros
description: Como trabalhar com ITP e Os Complementos do Office ao usar cookies de terceiros
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: 468147e923bb27638e45879104db75b99d014986
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917090"
---
# <a name="develop-your-office-add-in-to-work-with-itp-when-using-third-party-cookies"></a>Desenvolva seu Complemento do Office para trabalhar com ITP ao usar cookies de terceiros

Se o seu Add-in do Office exigir cookies de terceiros, esses cookies serão bloqueados se a PREVENÇÃO de Controle Inteligente (ITP) for usada pelo tempo de execução do navegador que carregou o seu complemento. Você pode estar usando cookies de terceiros para autenticar usuários ou para outros cenários, como armazenar configurações.

Se o seu Site e o Seu Add-in do Office devem depender de cookies de terceiros, use as etapas a seguir para trabalhar com ITP:

1. Configurar a [Autorização OAuth 2.0](https://tools.ietf.org/html/rfc6749)para que o domínio de autenticação (no seu caso, o terceiro que espera cookies) encaminhe um token de autorização para seu   site. Use o token para estabelecer uma sessão de logon de primeira parte com um cookie Secure e [HttpOnly](https://developer.mozilla.org/en-US/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies)definido pelo servidor.
2. Use a [API de Acesso de](https://webkit.org/blog/8124/introducing-storage-access-api/)Armazenamento para que o terceiro possa solicitar permissão para obter acesso aos cookies de primeira   parte. As versões atuais do Office no Mac e no Office na Web suportam essa API.
    > [!NOTE]
    > Se você estiver usando cookies para fins diferentes da autenticação, considere usar `localStorage` para seu cenário.

O exemplo de código a seguir mostra como usar a API de Acesso para Armazenamento:

```javascript
function displayLoginButton() {
  var button = createLoginButton();
  button.addEventListener("click", function(ev) {
    document.requestStorageAccess().then(function() {
      authenticateWithCookies(); 
    }).catch(function() {
      // User must have previously interacted with this domain loaded in a top frame
      // Also you should have previously written a cookie when domain was loaded in the top frame
      console.error("User cancelled or requirements were not met.");
    });
  });
}

if (document.hasStorageAccess) { 
  document.hasStorageAccess().then(function(hasStorageAccess) { 
    if (!hasStorageAccess) { 
      displayLoginButton(); 
    } else { 
      authenticateWithCookies(); 
    } 
  }); 
} else { 
    authenticateWithCookies(); 
} 
```

## <a name="about-itp-and-third-party-cookies"></a>Sobre ITP e cookies de terceiros

Cookies de terceiros são cookies carregados em um iframe, onde o domínio é diferente do quadro de nível superior. A ITP pode afetar cenários complexos de autenticação, onde uma caixa de diálogo pop-up é usada para inserir credenciais e, em seguida, o acesso a cookies é necessário por um iframe de um complemento para concluir o fluxo de autenticação. A ITP também pode afetar cenários de autenticação silenciosa, onde você já usou uma caixa de diálogo pop-up para autenticar, mas o uso subsequente do complemento tenta autenticar por meio de um iframe oculto.

Ao desenvolver os Complementos do Office no Mac, o acesso a cookies de terceiros é bloqueado pelo MacOS Big Sur SDK. Isso porque a ITP WKWebView está habilitada por padrão no navegador Safari e o WKWebView bloqueia todos os cookies de terceiros. O Office no Mac versão 16.44 ou posterior é integrado ao MacOS Big Sur SDK.

No navegador Safari, os usuários finais podem alternar a caixa de seleção Impedir rastreamento entre **sites** em **Privacidade** de Preferência para desativar  >   a ITP. No entanto, a ITP não pode ser desligada para o controle WKWebView incorporado.

## <a name="see-also"></a>Confira também

- [Manipular a ITP no Safari e em outros navegadores onde cookies de terceiros são bloqueados](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [Prevenção de Rastreamento no WebKit](https://webkit.org/tracking-prevention/)
- ["Área de Privacidade" do Chrome](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [Apresentando a API de Acesso ao Armazenamento](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)