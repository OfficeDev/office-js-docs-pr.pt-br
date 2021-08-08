---
title: Desenvolva seu Office de usuário para trabalhar com a ITP ao usar cookies de terceiros
description: Como trabalhar com a ITP e Office de complementos ao usar cookies de terceiros
ms.date: 07/8/2021
localization_priority: Normal
ms.openlocfilehash: 0a638f699e7b596bba30dcd12ec57d6da209a4a6a89ad987ef3fcbb8532e5f8c
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57080514"
---
# <a name="develop-your-office-add-in-to-work-with-itp-when-using-third-party-cookies"></a>Desenvolva seu Office de usuário para trabalhar com a ITP ao usar cookies de terceiros

Se o seu Office de usuário exigir cookies de terceiros, esses cookies serão bloqueados se a PREVENÇÃO de Controle Inteligente (ITP) for usada pelo tempo de execução do navegador que carregou o seu complemento. Você pode estar usando cookies de terceiros para autenticar usuários ou para outros cenários, como armazenar configurações.

Se o Office e o site devem depender de cookies de terceiros, use as etapas a seguir para trabalhar com ITP.

1. Configurar a [Autorização OAuth 2.0](https://tools.ietf.org/html/rfc6749)para que o domínio de autenticação (no seu caso, o terceiro que espera cookies) encaminhe um token de autorização para seu   site. Use o token para estabelecer uma sessão de logon de primeira parte com um cookie Secure e [HttpOnly](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies)definido pelo servidor.
2. Use a [Armazenamento de acesso](https://webkit.org/blog/8124/introducing-storage-access-api/)para que o terceiro possa solicitar permissão para obter acesso aos cookies de primeira   parte. As versões atuais do Office no Mac e Office na Web ambas suportam essa API.
    > [!NOTE]
    > Se você estiver usando cookies para fins diferentes da autenticação, considere usar `localStorage` para seu cenário.

O exemplo de código a seguir mostra como usar a API Armazenamento Access.

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

Ao desenvolver Office no Mac, o acesso a cookies de terceiros é bloqueado pelo SDK do MacOS Big Sur. Isso porque a ITP WKWebView está habilitada por padrão no navegador Safari e o WKWebView bloqueia todos os cookies de terceiros. Office no Mac versão 16.44 ou posterior é integrado ao MacOS Big Sur SDK.

No navegador Safari, os usuários finais podem alternar a caixa de seleção Impedir rastreamento entre **sites** em **Privacidade** de Preferência para desativar  >   a ITP. No entanto, a ITP não pode ser desligada para o controle WKWebView incorporado.

## <a name="see-also"></a>Confira também

- [Manipular a ITP no Safari e em outros navegadores onde cookies de terceiros são bloqueados](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [Prevenção de Rastreamento no WebKit](https://webkit.org/tracking-prevention/)
- ["Área de Privacidade" do Chrome](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [Apresentando a API Armazenamento Access](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)