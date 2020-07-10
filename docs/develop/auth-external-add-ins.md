---
title: Autorizar serviços externos no seu suplemento do Office
description: Obter autorização para outras fontes de dados além da Microsoft como Google, Facebook, LinkedIn, SalesForce e GitHub, usando o OAuth 2.0, o código de autorização e os fluxos implícitos.
ms.date: 08/07/2019
localization_priority: Normal
ms.openlocfilehash: fd180e11106e7e1e2f20f539746535c4310ad81e
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093739"
---
# <a name="authorize-external-services-in-your-office-add-in"></a>Autorizar serviços externos no seu suplemento do Office

Popular online services, including Microsoft 365, Google, Facebook, LinkedIn, SalesForce, and GitHub, let developers give users access to their accounts in other applications. This gives you the ability to include these services in your Office Add-in.

> [!NOTE]
> O restante deste artigo é sobre o acesso a serviços que não são da Microsoft. Para obter informações sobre como acessar o Microsoft Graph (incluindo o Microsoft 365), confira [acessar o Microsoft Graph com SSO](overview-authn-authz.md#access-to-microsoft-graph-with-sso) e [acesso ao Microsoft Graph sem SSO](overview-authn-authz.md#access-to-microsoft-graph-without-sso).

The industry standard framework for enabling web application access to an online service is **OAuth 2.0**. In most situations, you don't need to know the details of how the framework works to use it in your add-in. Many libraries are available that simplify the details for you.

Uma ideia fundamental do OAuth é que um aplicativo pode ser uma [entidade de segurança](/windows/security/identity-protection/access-control/security-principals) por si só, assim como um usuário ou um grupo, com sua própria identidade e conjunto de permissões. Nos cenários mais comuns, quando o usuário realiza uma ação no Suplemento do Office que requer o serviço online, o suplemento envia ao serviço uma solicitação para um conjunto específico de permissões para a conta do usuário. Em seguida, o serviço solicita que o usuário conceda essas permissões ao suplemento. Após a concessão das permissões, o serviço envia ao suplemento um pequeno *token de acesso* codificado. O suplemento pode usar o serviço, incluindo o token, em todas as suas solicitações para as APIs do serviço. Porém, o suplemento só pode agir dentro das permissões concedidas a ele pelo usuário. O token também expira após um tempo especificado.

Vários padrões OAuth, chamados de *fluxos* ou *tipos de concessão*, foram projetados para diferentes cenários. Os dois padrões a seguir são os mais comumente implementados:

- **Fluxo Implícito**: a comunicação entre o suplemento e o serviço online é implementada com um JavaScript no lado do cliente. Esse fluxo costuma ser usado em aplicativos página única (SPAs).
- **Fluxo de Código de Autorização**: A comunicação é *de servidor para servidor* entre o aplicativo Web do seu suplemento e o serviço online. Portanto, a implementação é feita com código no lado do servidor.

A finalidade de um fluxo OAuth é garantir a identidade e autorização do aplicativo. No fluxo de Código de Autorização, você recebe um *segredo do cliente* que precisa permanecer oculto. Um aplicativo que não tem nenhum back-end do lado do servidor, como é o caso de um SPA, não tem como proteger o segredo; por isso recomendamos usar o fluxo Implícito em SPAs.

Você deve estar familiarizado com os prós e os contras do fluxo implícito e o fluxo do código de autorização. Para obter mais informações sobre esses dois fluxos, consulte [Código de Autorização](https://tools.ietf.org/html/rfc6749#section-1.3.1) e [Implícito](https://tools.ietf.org/html/rfc6749#section-1.3.2).

> [!NOTE]
> You also have the option of using a middleman service to perform authorization and pass the access token to your add-in. For details about this scenario, see the **Middleman services** section later in this article.

## <a name="using-the-implicit-flow-in-office-add-ins"></a>Usando o fluxo Implícito em suplementos do Office

A melhor maneira de descobrir se um serviço online suporta o fluxo implícito é consultar a documentação do serviço.

Para obter informações sobre outras bibliotecas que suportam o fluxo implícito, consulte a seção **Bibliotecas** mais adiante neste artigo.

## <a name="using-the-authorization-code-flow-in-office-add-ins"></a>Usando o fluxo de Código de Autorização em suplementos do Office

Many libraries are available for implementing the Authorization Code flow in various languages and frameworks. For more information about some of these libraries, see the **Libraries** section later in this article.

## <a name="libraries"></a>Bibliotecas

Libraries are available for many languages and platforms, for both the Implicit flow and the Authorization Code flow. Some libraries are general purpose, while others are for specific online services.

**Google**: Search [GitHub.com/Google](https://github.com/google) for "auth" or the name of your language. Most of the relevant repos are named `google-auth-library-[name of language]`.

**Facebook**: Pesquise "library" ou "sdk" no [Facebook para Desenvolvedores](https://developers.facebook.com).

**General OAuth 2.0**: A page of links to libraries for over a dozen languages is maintained by the IETF OAuth Working Group at: [OAuth Code](https://oauth.net/code/). Note that some of these libraries are for implementing an OAuth compliant service. The libraries of interest to you as a an add-in developer are called *client* libraries on this page because your web server is a client of the OAuth compliant service.

## <a name="middleman-services"></a>Serviços intermediários

Your add-in can use a middleman service such as [OAuth.io](https://oauth.io) or [Auth0](https://auth0.com) to perform authorization. A middleman service may either provide access tokens for popular online services or simplify the process of enabling social login for your add-in, or both. With very little code, your add-in can use either client-side script or server-side code to connect to the middleman service and it will send your add-in any required tokens for the online service. All of the authorization implementation code is in the middleman service. 

É recomendável que a interface do usuário de autenticação/autorização no suplemento usar nossos APIs de caixa de diálogo para abrir uma página de logon. Ver [usar as APIs REST de caixa de diálogo em um fluxo de autenticação](dialog-api-in-office-add-ins.md#use-the-dialog-apis-in-an-authentication-flow) para saber mais. Quando você abre uma caixa de diálogo do Office dessa forma, a caixa de diálogo tem uma instância totalmente nova e separada do navegador e mecanismo JavaScript da instância na página pai (por exemplo, painel de tarefas do suplemento ou FunctionFile). Um token e outras informações que podem ser convertidas em uma cadeia de caracteres é passado para o pai usando uma chamada de API `messageParent`. Página pai, em seguida, pode usar o token para fazer chamadas autorizadas ao recurso. Devido à arquitetura, tenha cuidado como usar as APIs REST fornecidas pelo serviço de intermediário. Muitas vezes o serviço fornecerá uma configuração API no qual o código cria algum tipo de objeto contexto que é um token e o utiliza para fazer chamadas subsequentes ao recurso. Muitas vezes o serviço fornece um método de API único que faz a chamada inicial *e* cria objeto contexto. Um objeto assim não pode ser stringificado completamente, para que não possam ser passado de caixa de diálogo Office para a página de pai. Normalmente, o serviço intermediário fornece um segundo conjunto de API um nível inferior de abstração, como uma API REST. Este segundo conjunto tem um API que recebe um token do serviço e outras APIs que passam o token para o serviço ao usa-lo para obter acesso autorizado ao recurso. Precisa trabalhar com uma API neste nível inferior de abstração para que você possa obter o token na caixa de diálogo do Office e, em seguida, usar `messageParent` para passar para a página de pai. 

## <a name="what-is-cors"></a>O que é CORS?

CORS stands for [Cross Origin Resource Sharing](https://developer.mozilla.org/docs/Web/HTTP/Access_control_CORS). For information about how to use CORS inside add-ins, see [Addressing same-origin policy limitations in Office Add-ins](addressing-same-origin-policy-limitations.md).
