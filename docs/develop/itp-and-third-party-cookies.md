---
title: Desenvolver seu Suplemento do Office para trabalhar com o ITP ao usar cookies de terceiros
description: Como trabalhar com suplementos do ITP e do Office ao usar cookies de terceiros
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: b01051fa39441fddb2453b0bd95a0629ebf3ef65
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423087"
---
# <a name="develop-your-office-add-in-to-work-with-itp-when-using-third-party-cookies"></a>Desenvolver seu Suplemento do Office para trabalhar com o ITP ao usar cookies de terceiros

Se o suplemento do Office exigir cookies de terceiros, esses cookies serão bloqueados se o [Runtime](../testing/runtimes.md) que carregou seu suplemento usar a PREVENÇÃO de Rastreamento Inteligente (ITP). Você pode estar usando cookies de terceiros para autenticar usuários ou para outros cenários, como armazenar configurações.

Se o Suplemento do Office e o site precisarem contar com cookies de terceiros, use as etapas a seguir para trabalhar com o ITP.

1. Configure a [Autorização do OAuth 2.0](https://tools.ietf.org/html/rfc6749) para que o domínio de autenticação (no seu caso, o terceiro que espera cookies) encaminhe um token de autorização para seu site. Use o token para estabelecer uma sessão de logon de terceiros com um cookie Secure e [HttpOnly definido pelo servidor](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies).
1. Use a API  [de Acesso ao](https://webkit.org/blog/8124/introducing-storage-access-api/)Armazenamento para que o terceiro possa solicitar permissão para obter acesso aos cookies de terceiros. As versões atuais do Office no Mac e Office na Web dão suporte a essa API.
    > [!NOTE]
    > Se você estiver usando cookies para fins diferentes de autenticação, considere usar `localStorage` para seu cenário.

O exemplo de código a seguir mostra como usar a API de Acesso ao Armazenamento.

```javascript
function displayLoginButton() {
  const button = createLoginButton();
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

## <a name="about-itp-and-third-party-cookies"></a>Sobre o ITP e cookies de terceiros

Cookies de terceiros são cookies carregados em um iframe, em que o domínio é diferente do quadro de nível superior. O ITP pode afetar cenários de autenticação complexos, em que uma caixa de diálogo pop-up é usada para inserir credenciais e, em seguida, o acesso a cookie é necessário por um iframe de suplemento para concluir o fluxo de autenticação. O ITP também pode afetar cenários de autenticação silenciosa, em que você já usou uma caixa de diálogo pop-up para autenticar, mas o uso subsequente do suplemento tenta autenticar por meio de um iframe oculto.

Ao desenvolver Suplementos do Office no Mac, o acesso a cookies de terceiros é bloqueado pelo SDK do MacOS Big Sur. Isso ocorre porque o ITP do WKWebView está habilitado por padrão no navegador Safari e o WKWebView bloqueia todos os cookies de terceiros. O Office no Mac versão 16.44 ou posterior é integrado ao SDK do MacOS Big Sur.

No navegador Safari, os usuários finais podem alternar a caixa de seleção Impedir acompanhamento entre **sites**  >  em Privacidade de Preferência para desativar o ITP. No entanto, o ITP não pode ser desativado para o controle WKWebView inserido.

## <a name="see-also"></a>Confira também

- [Manipular o ITP no Safari e em outros navegadores em que os cookies de terceiros estão bloqueados](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [Prevenção de rastreamento no WebKit](https://webkit.org/tracking-prevention/)
- ["Área Restrita de Privacidade" do Chrome](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [Introdução à API de Acesso ao Armazenamento](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)
