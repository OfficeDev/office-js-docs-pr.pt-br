---
title: Desenvolva seu Office de usuário para trabalhar com a ITP ao usar cookies de terceiros
description: Como trabalhar com a ITP e Office de complementos ao usar cookies de terceiros
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: dbc23e4ead0abc94ffa173ffc22919342c4fca6d
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349858"
---
# <a name="develop-your-office-add-in-to-work-with-itp-when-using-third-party-cookies"></a><span data-ttu-id="d588d-103">Desenvolva seu Office de usuário para trabalhar com a ITP ao usar cookies de terceiros</span><span class="sxs-lookup"><span data-stu-id="d588d-103">Develop your Office Add-in to work with ITP when using third-party cookies</span></span>

<span data-ttu-id="d588d-104">Se o seu Office de usuário exigir cookies de terceiros, esses cookies serão bloqueados se a PREVENÇÃO de Controle Inteligente (ITP) for usada pelo tempo de execução do navegador que carregou o seu complemento.</span><span class="sxs-lookup"><span data-stu-id="d588d-104">If your Office Add-in requires third-party cookies, those cookies are blocked if Intelligent Tracking Prevention (ITP) is used by the browser runtime that loaded your add-in.</span></span> <span data-ttu-id="d588d-105">Você pode estar usando cookies de terceiros para autenticar usuários ou para outros cenários, como armazenar configurações.</span><span class="sxs-lookup"><span data-stu-id="d588d-105">You may be using third-party cookies to authenticate users, or for other scenarios, such as storing settings.</span></span>

<span data-ttu-id="d588d-106">Se o Office e o site devem depender de cookies de terceiros, use as etapas a seguir para trabalhar com ITP:</span><span class="sxs-lookup"><span data-stu-id="d588d-106">If your Office Add-in and website must rely on third-party cookies, use the following steps to work with ITP:</span></span>

1. <span data-ttu-id="d588d-107">Configurar a [Autorização OAuth 2.0](https://tools.ietf.org/html/rfc6749)para que o domínio de autenticação (no seu caso, o terceiro que espera cookies) encaminhe um token de autorização para seu   site.</span><span class="sxs-lookup"><span data-stu-id="d588d-107">Set up [OAuth 2.0 Authorization](https://tools.ietf.org/html/rfc6749) so that the authenticating domain (in your case, the third-party that expects cookies) forwards an authorization token to your website.</span></span> <span data-ttu-id="d588d-108">Use o token para estabelecer uma sessão de logon de primeira parte com um cookie Secure e [HttpOnly](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies)definido pelo servidor.</span><span class="sxs-lookup"><span data-stu-id="d588d-108">Use the token to establish a first-party login session with a server-set Secure and [HttpOnly cookie](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies).</span></span>
2. <span data-ttu-id="d588d-109">Use a [Armazenamento de acesso](https://webkit.org/blog/8124/introducing-storage-access-api/)para que o terceiro possa solicitar permissão para obter acesso aos cookies de primeira   parte.</span><span class="sxs-lookup"><span data-stu-id="d588d-109">Use the [Storage Access API](https://webkit.org/blog/8124/introducing-storage-access-api/) so that the third-party can request permission to get access to its first-party cookies.</span></span> <span data-ttu-id="d588d-110">As versões atuais do Office no Mac e Office na Web ambas suportam essa API.</span><span class="sxs-lookup"><span data-stu-id="d588d-110">Current versions of Office on Mac and Office on the web both support this API.</span></span>
    > [!NOTE]
    > <span data-ttu-id="d588d-111">Se você estiver usando cookies para fins diferentes da autenticação, considere usar `localStorage` para seu cenário.</span><span class="sxs-lookup"><span data-stu-id="d588d-111">If you're using cookies for purposes other than authentication, then consider using `localStorage` for your scenario.</span></span>

<span data-ttu-id="d588d-112">O exemplo de código a seguir mostra como usar a API Armazenamento Access.</span><span class="sxs-lookup"><span data-stu-id="d588d-112">The following code sample shows how to use the Storage Access API.</span></span>

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

## <a name="about-itp-and-third-party-cookies"></a><span data-ttu-id="d588d-113">Sobre ITP e cookies de terceiros</span><span class="sxs-lookup"><span data-stu-id="d588d-113">About ITP and third-party cookies</span></span>

<span data-ttu-id="d588d-114">Cookies de terceiros são cookies carregados em um iframe, onde o domínio é diferente do quadro de nível superior.</span><span class="sxs-lookup"><span data-stu-id="d588d-114">Third-party cookies are cookies that are loaded in an iframe, where the domain is different from the top level frame.</span></span> <span data-ttu-id="d588d-115">A ITP pode afetar cenários complexos de autenticação, onde uma caixa de diálogo pop-up é usada para inserir credenciais e, em seguida, o acesso a cookies é necessário por um iframe de um complemento para concluir o fluxo de autenticação.</span><span class="sxs-lookup"><span data-stu-id="d588d-115">ITP could affect complex authentication scenarios, where a popup dialog is used to enter credentials and then the cookie access is needed by an add-in iframe to complete the authentication flow.</span></span> <span data-ttu-id="d588d-116">A ITP também pode afetar cenários de autenticação silenciosa, onde você já usou uma caixa de diálogo pop-up para autenticar, mas o uso subsequente do complemento tenta autenticar por meio de um iframe oculto.</span><span class="sxs-lookup"><span data-stu-id="d588d-116">ITP could also affect silent authentication scenarios, where you have previously used a popup dialog to authenticate, but subsequent use of the add-in tries to authenticate through a hidden iframe.</span></span>

<span data-ttu-id="d588d-117">Ao desenvolver Office no Mac, o acesso a cookies de terceiros é bloqueado pelo SDK do MacOS Big Sur.</span><span class="sxs-lookup"><span data-stu-id="d588d-117">When developing Office Add-ins on Mac, access to third-party cookies is blocked by the MacOS Big Sur SDK.</span></span> <span data-ttu-id="d588d-118">Isso porque a ITP WKWebView está habilitada por padrão no navegador Safari e o WKWebView bloqueia todos os cookies de terceiros.</span><span class="sxs-lookup"><span data-stu-id="d588d-118">This is because WKWebView ITP is enabled by default on the Safari browser, and WKWebView blocks all third-party cookies.</span></span> <span data-ttu-id="d588d-119">Office no Mac versão 16.44 ou posterior é integrado ao MacOS Big Sur SDK.</span><span class="sxs-lookup"><span data-stu-id="d588d-119">Office on Mac version 16.44 or later is integrated with the MacOS Big Sur SDK.</span></span>

<span data-ttu-id="d588d-120">No navegador Safari, os usuários finais podem alternar a caixa de seleção Impedir rastreamento entre **sites** em **Privacidade** de Preferência para desativar  >   a ITP.</span><span class="sxs-lookup"><span data-stu-id="d588d-120">In the Safari browser, end users can toggle the **Prevent cross-site tracking** checkbox under **Preference** > **Privacy** to turn off ITP.</span></span> <span data-ttu-id="d588d-121">No entanto, a ITP não pode ser desligada para o controle WKWebView incorporado.</span><span class="sxs-lookup"><span data-stu-id="d588d-121">However, ITP cannot be turned off for the embedded WKWebView control.</span></span>

## <a name="see-also"></a><span data-ttu-id="d588d-122">Confira também</span><span class="sxs-lookup"><span data-stu-id="d588d-122">See also</span></span>

- [<span data-ttu-id="d588d-123">Manipular a ITP no Safari e em outros navegadores onde cookies de terceiros são bloqueados</span><span class="sxs-lookup"><span data-stu-id="d588d-123">Handle ITP in Safari and other browsers where third-party cookies are blocked</span></span>](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [<span data-ttu-id="d588d-124">Prevenção de Rastreamento no WebKit</span><span class="sxs-lookup"><span data-stu-id="d588d-124">Tracking Prevention in WebKit</span></span>](https://webkit.org/tracking-prevention/)
- [<span data-ttu-id="d588d-125">"Área de Privacidade" do Chrome</span><span class="sxs-lookup"><span data-stu-id="d588d-125">Chrome’s “Privacy Sandbox”</span></span>](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [<span data-ttu-id="d588d-126">Apresentando a API Armazenamento Access</span><span class="sxs-lookup"><span data-stu-id="d588d-126">Introducing the Storage Access API</span></span>](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)