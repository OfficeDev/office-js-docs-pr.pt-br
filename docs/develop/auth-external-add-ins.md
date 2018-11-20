---
title: Autorizar serviços externos no seu suplemento do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 6cdf07886ba883a7dfe935b59c918948c2b45afa
ms.sourcegitcommit: 86724e980f720ed05359c9525948cb60b6f10128
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/09/2018
ms.locfileid: "26237476"
---
# <a name="authorize-external-services-in-your-office-add-in"></a>Autorizar serviços externos no seu suplemento do Office

Serviços online populares, incluindo o Office 365, o Google, o Facebook, o LinkedIn, o SalesForce e o GitHub, permitem que os desenvolvedores forneçam acesso para os usuários a suas contas em outros aplicativos, o que possibilita que você inclua esses serviços no seu Suplemento do Office.

A estrutura padrão do setor para habilitar o acesso de aplicativos Web a um serviço online é **OAuth 2.0**. Na maioria das situações, você não precisa saber os detalhes de como a estrutura funciona para usá-la no seu suplemento. Estão disponíveis muitas bibliotecas que simplificam os detalhes para você.

Uma ideia fundamental do OAuth é que um aplicativo pode ser uma entidade de segurança por si só, assim como um usuário ou um grupo, com sua própria identidade e conjunto de permissões. Nos cenários mais comuns, quando o usuário realiza uma ação no Suplemento do Office que requer o serviço online, o suplemento envia ao serviço uma solicitação para um conjunto específico de permissões para a conta do usuário. Em seguida, o serviço solicita que o usuário conceda essas permissões ao suplemento. Após a concessão das permissões, o serviço envia ao suplemento um pequeno *token de acesso* codificado. O suplemento pode usar o serviço, incluindo o token, em todas as suas solicitações para as APIs do serviço. Porém, o suplemento só pode agir dentro das permissões concedidas a ele pelo usuário. O token também expira após um tempo especificado.

Vários padrões OAuth, chamados de *fluxos* ou *tipos de concessão*, foram projetados para diferentes cenários. Os dois padrões a seguir são os mais comumente implementados:

- **Fluxo Implícito**: a comunicação entre o suplemento e o serviço online é implementada com um JavaScript no lado do cliente.
- **Fluxo de Código de Autorização**: a comunicação é *de servidor para servidor* entre o aplicativo Web do seu suplemento e o serviço online. Portanto, a implementação é feita com código no lado do servidor.

A finalidade dos fluxos do OAuth é proteger a identidade e a autorização do aplicativo. No fluxo de Código de Autorização, você recebe um *segredo de cliente* que precisa permanecer oculto. Como um Aplicativo de Página Única (SPA) não tem como proteger o segredo, recomendamos que você use o fluxo Implícito em SPAs.

Você deve estar familiarizado com os prós e os contras do fluxo implícito e o fluxo do Código de Autorização. Para obter mais informações sobre esses dois fluxos, consulte [Código de Autorização](https://tools.ietf.org/html/rfc6749#section-1.3.1) e [Implícito](https://tools.ietf.org/html/rfc6749#section-1.3.2).

> [!NOTE]
> Você também tem a opção de usar um serviço intermediário para executar a autorização e passar o token de acesso ao seu suplemento. Para obter detalhes sobre esse cenário, consulte a seção **Serviços intermediários** mais adiante neste artigo.

## <a name="using-the-implicit-flow-in-office-add-ins"></a>Usando o fluxo Implícito em suplementos do Office
A melhor maneira de descobrir se um serviço online suporta o fluxo implícito é consultar a documentação do serviço. Para serviços que suportam o fluxo implícito, você pode usar a biblioteca de JavaScript **Office-js-helpers** para fazer todo o trabalho detalhado para você:

- [Office-js-helpers](https://github.com/OfficeDev/office-js-helpers)

Para obter informações sobre outras bibliotecas que suportam o fluxo implícito, consulte a seção **Bibliotecas** mais adiante neste artigo.

## <a name="using-the-authorization-code-flow-in-office-add-ins"></a>Usando o fluxo de Código de Autorização em suplementos do Office

Muitas bibliotecas estão disponíveis para implementar o fluxo de Código de Autorização em várias linguagens e estruturas. Para mais informações sobre algumas dessas bibliotecas, consulte a seção **Bibliotecas** mais adiante neste artigo.

As seguintes amostras fornecem exemplos de suplementos que implementam o Fluxo do Código de Autorização:

- [Office-Add-in-Nodejs-ServerAuth](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) (NodeJS)
- [PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) (ASP.NET MVC)

### <a name="relayproxy-functions"></a>Funções de retransmissão/Proxy

Você pode usar o fluxo do Código de Autorização mesmo com um aplicativo Web sem servidor armazenando os valores de **ID do cliente** e **segredo cliente** em uma função simples hospedada em um serviço como o [Azure Functions](https://azure.microsoft.com/services/functions) ou o [Amazon Lambda](https://aws.amazon.com/lambda). A função troca um código específico por um **token de acesso** e transmite-o de volta para o cliente. A segurança dessa abordagem depende do quanto protegido é o acesso à função.

Para usar essa técnica, o suplemento exibe uma interface do usuário/pop-up para mostrar a tela de logon do serviço online (Google, Facebook e assim por diante). Quando o usuário se conecta e concede ao suplemento a permissão para acessar seus recursos no serviço online, o suplemento recebe um código que então pode ser enviado para a função online. Os serviços descritos em **Serviços intermediários** neste artigo usam um fluxo semelhante a esse.

## <a name="libraries"></a>Bibliotecas

As bibliotecas estão disponíveis para vários idiomas e plataformas, tanto para o fluxo implícito quanto para o fluxo do Código de Autorização. Algumas bibliotecas são de propósito geral, enquanto outras são para serviços online específicos.

**Office 365 e outros serviços que usam o Azure Active Directory como provedor de autorização**: [Bibliotecas de autenticação do Active Directory do Azure](https://azure.microsoft.com/documentation/articles/active-directory-authentication-libraries/). Também está disponível uma visualização da [Biblioteca de Autenticação da Microsoft](https://www.nuget.org/packages/Microsoft.Identity.Client).

**Google**: Pesquise "auth" ou o nome da sua linguagem no [GitHub.com/Google](https://github.com/google). A maioria dos repositórios relevantes se chama `google-auth-library-[name of language]`.

**Facebook**: Pesquise "library" ou "sdk" no [Facebook para Desenvolvedores](https://developers.facebook.com).

**OAuth 2.0 Geral**: Uma página de links para bibliotecas de mais de uma dúzia de linguagens é mantida pelo IETF OAuth Working Group, em: [Código OAuth](http://oauth.net/code/). Observe que algumas dessas bibliotecas são para implementar um serviço compatível com o OAuth. As bibliotecas interessantes para você, como desenvolvedor do suplemento, são chamadas de bibliotecas de *cliente* nesta página, pois o servidor Web é um cliente do serviço compatível com o OAuth.

## <a name="middleman-services"></a>Serviços intermediários

Seu suplemento pode usar um serviço intermediário, como o OAuth.io ou o Auth0, para executar a autorização. O serviço intermediário fornece tokens de acesso para serviços online populares ou simplifica o processo de habilitar o logon social para esse suplemento. Com muito pouco código, o suplemento pode usar qualquer script no lado do cliente ou código no lado do servidor para se conectar ao serviço intermediário e enviar ao suplemento qualquer token necessário para o serviço online. Todo o código de implementação de autorização está no serviço intermediário.

Para obter exemplos de suplementos que usam um serviço intermediário para autorização, consulte as seguintes amostras:

- [Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0) usa o Auth0 para habilitar o login social com o Facebook, Google e contas da Microsoft.

- [Office-Add-in-OAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io) usa o OAuth.io para obter tokens de acesso a partir do Facebook e Google.

## <a name="what-is-cors"></a>O que é CORS?

CORS significa [Compartilhamento de Recursos Entre Origens](https://developer.mozilla.org/docs/Web/HTTP/Access_control_CORS). Para obter informações sobre como usar o CORS nos suplementos, confira [Como lidar com as limitações da política de mesma origem nos suplementos do Office](addressing-same-origin-policy-limitations.md).
