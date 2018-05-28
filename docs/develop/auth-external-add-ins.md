---
title: Autorizar servi?os externos no seu suplemento do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 34e8119d4ecf6432cde7f06552584d164b8c9b8e
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="authorize-external-services-in-your-office-add-in"></a>Autorizar servi?os externos no seu suplemento do Office

Servi?os online populares, incluindo o Office 365, o Google, o Facebook, o LinkedIn, o SalesForce e o GitHub, permitem que os desenvolvedores forne?am acesso para os usu?rios a suas contas em outros aplicativos, o que possibilita que voc? inclua esses servi?os no seu Suplemento do Office.

A estrutura padr?o do setor para habilitar o acesso de aplicativos Web a um servi?o online ? **OAuth 2.0**. Na maioria das situa??es, voc? n?o precisa saber os detalhes de como a estrutura funciona para us?-la no seu suplemento. Est?o dispon?veis muitas bibliotecas que simplificam os detalhes para voc?.

Uma ideia fundamental do OAuth ? que um aplicativo pode ser uma entidade de seguran?a por si s?, assim como um usu?rio ou um grupo, com sua pr?pria identidade e conjunto de permiss?es. Nos cen?rios mais comuns, quando o usu?rio realiza uma a??o no suplemento do Office que requer o servi?o online, o suplemento envia ao servi?o uma solicita??o para um conjunto espec?fico de permiss?es para a conta do usu?rio. Em seguida, o servi?o solicita que o usu?rio conceda essas permiss?es ao suplemento. Ap?s a concess?o das permiss?es, o servi?o envia ao suplemento um pequeno *token de acesso* codificado. O suplemento pode usar o servi?o, incluindo o token, em todas as suas solicita??es para as APIs do servi?o. Por?m, o suplemento s? pode agir dentro das permiss?es concedidas a ele pelo usu?rio. O token tamb?m expira ap?s um tempo especificado.

V?rios padr?es OAuth, chamados de *fluxos* ou *tipos de concess?o*, foram projetados para diferentes cen?rios. Os dois padr?es a seguir s?o os mais comumente implementados:

- **Fluxo Impl?cito**: a comunica??o entre o suplemento e o servi?o online ? implementada com um JavaScript no lado do cliente.
- **Fluxo de C?digo de Autoriza??o**: a comunica??o ? *de servidor para servidor* entre o aplicativo Web do seu suplemento e o servi?o online. Portanto, a implementa??o ? feita com c?digo no lado do servidor.

A finalidade dos fluxos do OAuth ? proteger a identidade e a autoriza??o do aplicativo. No fluxo de C?digo de Autoriza??o, voc? recebe um *segredo de cliente* que precisa permanecer oculto. Como um Aplicativo de P?gina ?nica (SPA) n?o tem como proteger o segredo, recomendamos que voc? use o fluxo Impl?cito em SPAs.

Voc? deve estar familiarizado com os pr?s e os contras do fluxo impl?cito e o fluxo do C?digo de Autoriza??o. Para obter mais informa??es sobre esses dois fluxos, consulte [C?digo de Autoriza??o](https://tools.ietf.org/html/rfc6749#section-1.3.1) e [Impl?cito](https://tools.ietf.org/html/rfc6749#section-1.3.2).

> [!NOTE]
> Voc? tamb?m tem a op??o de usar um servi?o intermedi?rio para executar a autoriza??o e passar o token de acesso ao seu suplemento. Confira detalhes sobre esse cen?rio na se??o **Servi?os intermedi?rios** mais adiante neste artigo.

## <a name="authorization-to-microsoft-graph"></a>Autoriza??o para o Microsoft Graph

Se o servi?o externo puder ser acessado por meio do Microsoft Graph, como o Office 365 ou o OneDrive, voc? poder? fornecer a melhor experi?ncia para os usu?rios e a experi?ncia de desenvolvimento mais f?cil para voc?, usando o sistema de logon ?nico descrito em [Autorizar para o Microsoft Graph no suplemento do Office](authorize-to-microsoft-graph.md) e seus artigos relacionados. As t?cnicas descritas neste artigo s?o ideais para servi?os externos que n?o podem ser acessados com o Microsoft Graph. No entanto, elas *podem* ser usadas para acessar o Microsoft Graph, e voc? pode preferir as vantagens do logon ?nico. Por exemplo, o sistema de logon ?nico requer c?digo do lado do servidor, portanto, ele n?o pode ser usado com um aplicativo de p?gina ?nica. Al?m disso, o sistema de logon ?nico ainda n?o ? compat?vel com todas as plataformas.

## <a name="using-the-implicit-flow-in-office-add-ins"></a>Usando o fluxo Impl?cito em suplementos do Office
A melhor maneira de descobrir se um servi?o online suporta o fluxo impl?cito ? consultar a documenta??o do servi?o. Para servi?os que suportam o fluxo impl?cito, voc? pode usar a biblioteca de JavaScript **Office-js-helpers** para fazer todo o trabalho detalhado para voc?:

- [Office-js-helpers](https://github.com/OfficeDev/office-js-helpers)

Para obter informa??es sobre outras bibliotecas que suportam o fluxo impl?cito, consulte a se??o **Bibliotecas** mais adiante neste artigo.

## <a name="using-the-authorization-code-flow-in-office-add-ins"></a>Usando o fluxo de C?digo de Autoriza??o em suplementos do Office

Muitas bibliotecas est?o dispon?veis para implementar o fluxo de C?digo de Autoriza??o em v?rias linguagens e estruturas. Para mais informa??es sobre algumas dessas bibliotecas, consulte a se??o **Bibliotecas** mais adiante neste artigo.

As seguintes amostras fornecem exemplos de suplementos que implementam o Fluxo do C?digo de Autoriza??o:

- [Office-Add-in-Nodejs-ServerAuth](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) (NodeJS)
- [PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) (ASP.NET MVC)

### <a name="relayproxy-functions"></a>Fun??es de retransmiss?o/Proxy

Voc? pode usar o fluxo do C?digo de Autoriza??o mesmo com um aplicativo Web sem servidor armazenando os valores de **ID do cliente** e **segredo cliente** em uma fun??o simples hospedada em um servi?o como o [Azure Functions](https://azure.microsoft.com/en-us/services/functions) ou o [Amazon Lambda](https://aws.amazon.com/lambda). A fun??o troca um c?digo espec?fico por um **token de acesso** e transmite-o de volta para o cliente. A seguran?a dessa abordagem depende do quanto protegido ? o acesso ? fun??o.

Para usar essa t?cnica, o suplemento exibe uma interface do usu?rio/pop-up para mostrar a tela de logon do servi?o online (Google, Facebook e assim por diante). Quando o usu?rio se conecta e concede ao suplemento a permiss?o para acessar seus recursos no servi?o online, o suplemento recebe um c?digo que ent?o pode ser enviado para a fun??o online. Os servi?os descritos em **Servi?os intermedi?rios** neste artigo usam um fluxo semelhante a esse.

## <a name="libraries"></a>Bibliotecas

As bibliotecas est?o dispon?veis para v?rios idiomas e plataformas, tanto para o fluxo impl?cito quanto para o fluxo do C?digo de Autoriza??o. Algumas bibliotecas s?o de prop?sito geral, enquanto outras s?o para servi?os online espec?ficos.

**Office 365 e outros servi?os que usam o Azure Active Directory como provedor de autoriza??o**: [Bibliotecas de autentica??o do Active Directory do Azure](https://azure.microsoft.com/en-us/documentation/articles/active-directory-authentication-libraries/). Tamb?m est? dispon?vel uma visualiza??o da [Biblioteca de Autentica??o da Microsoft](https://www.nuget.org/packages/Microsoft.Identity.Client).

**Google**: Pesquise "auth" ou o nome da sua linguagem no [GitHub.com/Google](https://github.com/google). A maioria dos reposit?rios relevantes se chama `google-auth-library-[name of language]`.

**Facebook**: Pesquise "library" ou "sdk" no [Facebook para Desenvolvedores](https://developers.facebook.com).

**OAuth 2.0 Geral**: Uma p?gina de links para bibliotecas de mais de uma d?zia de linguagens ? mantida pelo IETF OAuth Working Group, em: [C?digo OAuth](http://oauth.net/code/). Observe que algumas dessas bibliotecas s?o para implementar um servi?o compat?vel com o OAuth. As bibliotecas interessantes para voc?, como desenvolvedor do suplemento, s?o chamadas de bibliotecas de *cliente* nesta p?gina, pois o servidor Web ? um cliente do servi?o compat?vel com o OAuth.

## <a name="middleman-services"></a>Servi?os intermedi?rios

Seu suplemento pode usar um servi?o intermedi?rio, como o OAuth.io ou o Auth0, para executar a autoriza??o. O servi?o intermedi?rio fornece tokens de acesso para servi?os online populares ou simplifica o processo de habilitar o logon social para esse suplemento. Com muito pouco c?digo, o suplemento pode usar qualquer script no lado do cliente ou c?digo no lado do servidor para se conectar ao servi?o intermedi?rio e enviar ao suplemento qualquer token necess?rio para o servi?o online. Todo o c?digo de implementa??o de autoriza??o est? no servi?o intermedi?rio.

Para obter exemplos de suplementos que usam um servi?o intermedi?rio para autoriza??o, consulte as seguintes amostras:

- [Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0) usa o Auth0 para habilitar o login social com o Facebook, Google e contas da Microsoft.

- [Office-Add-in-OAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io) usa o OAuth.io para obter tokens de acesso a partir do Facebook e Google.

## <a name="what-is-cors"></a>O que ? CORS?

CORS significa [Compartilhamento de Recursos Entre Origens](https://developer.mozilla.org/en-US/docs/Web/HTTP/Access_control_CORS). Para obter informa??es sobre como usar o CORS nos suplementos, confira [Como lidar com as limita??es da pol?tica de mesma origem nos suplementos do Office](addressing-same-origin-policy-limitations.md).
