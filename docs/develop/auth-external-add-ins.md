---
title: Autorizar serviços externos no seu suplemento do Office
description: Obter autorização para outras fontes de dados além da Microsoft como Google, Facebook, LinkedIn, SalesForce e GitHub, usando o OAuth 2.0, o código de autorização e os fluxos implícitos.
ms.date: 08/07/2019
localization_priority: Normal
ms.openlocfilehash: 55f46a4cb381bc3f87434893d065f9ebf8147814
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608429"
---
# <a name="authorize-external-services-in-your-office-add-in"></a>Autorizar serviços externos no seu suplemento do Office

Serviços online populares, incluindo o Office 365, o Google, o Facebook, o LinkedIn, o SalesForce e o GitHub, permitem que os desenvolvedores forneçam acesso para os usuários a suas contas em outros aplicativos, o que possibilita que você inclua esses serviços no seu Suplemento do Office.

> [!NOTE]
> O restante deste artigo é sobre o acesso a serviços que não são da Microsoft. Para saber mais sobre como acessar o Microsoft Graph (incluindo o Office 365), confira[ acessar o Microsoft Graph com o SSO](overview-authn-authz.md#access-to-microsoft-graph-with-sso) e [acessar o Microsoft Graph sem o SSO](overview-authn-authz.md#access-to-microsoft-graph-without-sso).

A estrutura padrão do setor para habilitar o acesso de aplicativos Web a um serviço online é **OAuth 2.0**. Na maioria das situações, você não precisa saber os detalhes de como a estrutura funciona para usá-la no seu suplemento. Estão disponíveis muitas bibliotecas que simplificam os detalhes para você.

Uma ideia fundamental do OAuth é que um aplicativo pode ser uma [entidade de segurança](/windows/security/identity-protection/access-control/security-principals) por si só, assim como um usuário ou um grupo, com sua própria identidade e conjunto de permissões. Nos cenários mais comuns, quando o usuário realiza uma ação no Suplemento do Office que requer o serviço online, o suplemento envia ao serviço uma solicitação para um conjunto específico de permissões para a conta do usuário. Em seguida, o serviço solicita que o usuário conceda essas permissões ao suplemento. Após a concessão das permissões, o serviço envia ao suplemento um pequeno *token de acesso* codificado. O suplemento pode usar o serviço, incluindo o token, em todas as suas solicitações para as APIs do serviço. Porém, o suplemento só pode agir dentro das permissões concedidas a ele pelo usuário. O token também expira após um tempo especificado.

Vários padrões OAuth, chamados de *fluxos* ou *tipos de concessão*, foram projetados para diferentes cenários. Os dois padrões a seguir são os mais comumente implementados:

- **Fluxo Implícito**: a comunicação entre o suplemento e o serviço online é implementada com um JavaScript no lado do cliente. Esse fluxo costuma ser usado em aplicativos página única (SPAs).
- **Fluxo de Código de Autorização**: A comunicação é *de servidor para servidor* entre o aplicativo Web do seu suplemento e o serviço online. Portanto, a implementação é feita com código no lado do servidor.

A finalidade de um fluxo OAuth é garantir a identidade e autorização do aplicativo. No fluxo de Código de Autorização, você recebe um *segredo do cliente* que precisa permanecer oculto. Um aplicativo que não tem nenhum back-end do lado do servidor, como é o caso de um SPA, não tem como proteger o segredo; por isso recomendamos usar o fluxo Implícito em SPAs.

Você deve estar familiarizado com os prós e os contras do fluxo implícito e o fluxo do código de autorização. Para obter mais informações sobre esses dois fluxos, consulte [Código de Autorização](https://tools.ietf.org/html/rfc6749#section-1.3.1) e [Implícito](https://tools.ietf.org/html/rfc6749#section-1.3.2).

> [!NOTE]
> Você também tem a opção de usar um serviço intermediário para executar a autorização e passar o token de acesso ao seu suplemento. Para obter detalhes sobre esse cenário, consulte a seção **Serviços intermediários** mais adiante neste artigo.

## <a name="using-the-implicit-flow-in-office-add-ins"></a>Usando o fluxo Implícito em suplementos do Office

A melhor maneira de descobrir se um serviço online suporta o fluxo implícito é consultar a documentação do serviço.

Para obter informações sobre outras bibliotecas que suportam o fluxo implícito, consulte a seção **Bibliotecas** mais adiante neste artigo.

## <a name="using-the-authorization-code-flow-in-office-add-ins"></a>Usando o fluxo de Código de Autorização em suplementos do Office

Muitas bibliotecas estão disponíveis para implementar o fluxo de Código de Autorização em várias linguagens e estruturas. Para mais informações sobre algumas dessas bibliotecas, consulte a seção **Bibliotecas** mais adiante neste artigo.

## <a name="libraries"></a>Bibliotecas

As bibliotecas estão disponíveis para vários idiomas e plataformas, tanto para o fluxo implícito quanto para o fluxo do Código de Autorização. Algumas bibliotecas são de propósito geral, enquanto outras são para serviços online específicos.

**Google**: Pesquise "auth" ou o nome da sua linguagem no [GitHub.com/Google](https://github.com/google). A maioria dos repositórios relevantes se chama `google-auth-library-[name of language]`.

**Facebook**: Pesquise "library" ou "sdk" no [Facebook para Desenvolvedores](https://developers.facebook.com).

**OAuth 2.0 Geral**: Uma página de links para bibliotecas de mais de uma dúzia de linguagens é mantida pelo IETF OAuth Working Group, em: [Código OAuth](https://oauth.net/code/). Observe que algumas dessas bibliotecas são para implementar um serviço compatível com o OAuth. As bibliotecas interessantes para você, como desenvolvedor do suplemento, são chamadas de bibliotecas de *cliente* nesta página, pois o servidor Web é um cliente do serviço compatível com o OAuth.

## <a name="middleman-services"></a>Serviços intermediários

Seu suplemento pode usar um serviço intermediário, como o [OAuth.io](https://oauth.io) ou [Auth0](https://auth0.com), para executar a autorização. Um serviço intermediário fornece tokens de acesso para serviços online populares ou simplifica o processo de habilitar o logon social para esse suplemento. Com muito pouco código, o suplemento pode usar qualquer script no lado do cliente ou código no lado do servidor para se conectar ao serviço intermediário e enviar ao suplemento qualquer token necessário para o serviço online. Todo o código de implementação de autorização está no serviço intermediário. 

É recomendável que a interface do usuário de autenticação/autorização no suplemento usar nossos APIs de caixa de diálogo para abrir uma página de logon. Ver [usar as APIs REST de caixa de diálogo em um fluxo de autenticação](dialog-api-in-office-add-ins.md#use-the-dialog-apis-in-an-authentication-flow) para saber mais. Quando você abre uma caixa de diálogo do Office dessa forma, a caixa de diálogo tem uma instância totalmente nova e separada do navegador e mecanismo JavaScript da instância na página pai (por exemplo, painel de tarefas do suplemento ou FunctionFile). Um token e outras informações que podem ser convertidas em uma cadeia de caracteres é passado para o pai usando uma chamada de API `messageParent`. Página pai, em seguida, pode usar o token para fazer chamadas autorizadas ao recurso. Devido à arquitetura, tenha cuidado como usar as APIs REST fornecidas pelo serviço de intermediário. Muitas vezes o serviço fornecerá uma configuração API no qual o código cria algum tipo de objeto contexto que é um token e o utiliza para fazer chamadas subsequentes ao recurso. Muitas vezes o serviço fornece um método de API único que faz a chamada inicial *e* cria objeto contexto. Um objeto assim não pode ser stringificado completamente, para que não possam ser passado de caixa de diálogo Office para a página de pai. Normalmente, o serviço intermediário fornece um segundo conjunto de API um nível inferior de abstração, como uma API REST. Este segundo conjunto tem um API que recebe um token do serviço e outras APIs que passam o token para o serviço ao usa-lo para obter acesso autorizado ao recurso. Precisa trabalhar com uma API neste nível inferior de abstração para que você possa obter o token na caixa de diálogo do Office e, em seguida, usar `messageParent` para passar para a página de pai. 

## <a name="what-is-cors"></a>O que é CORS?

CORS significa [Compartilhamento de Recursos Entre Origens](https://developer.mozilla.org/docs/Web/HTTP/Access_control_CORS). Para obter informações sobre como usar o CORS nos suplementos, confira [Como lidar com as limitações da política de mesma origem nos suplementos do Office](addressing-same-origin-policy-limitations.md).
