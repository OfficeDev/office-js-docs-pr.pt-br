---
title: Visão geral da autenticação e autorização nos Suplementos do Office
description: ''
ms.date: 08/07/2019
localization_priority: Priority
ms.openlocfilehash: 2733f8af9f236347e52269c9e73b322b4310e2a9
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/13/2019
ms.locfileid: "36302925"
---
# <a name="overview-of-authentication-and-authorization-in-office-add-ins"></a>Visão geral da autenticação e autorização nos Suplementos do Office

Os aplicativos da Web e, portanto, os Suplementos do Office permitem acesso anônimo por padrão, mas é possível exigir que os usuários se autentiquem com um logon. Em particular, você pode exigir que os usuários se conectem com uma Conta da Microsoft, uma Conta Corporativa ou de Estudante (Office 365). Essa tarefa é chamada de autenticação do usuário, pois permite que o suplemento saiba quem é o usuário.

Seu suplemento também pode obter consentimento do usuário para acessar seus dados do Microsoft Graph (como seu perfil do Office 365, arquivos do OneDrive e dados do SharePoint) ou para dados de outras fontes externas, como Google, Facebook, LinkedIn, SalesForce e GitHub. Essa tarefa é chamada de autorização de suplemento (ou aplicativo), pois é o *suplemento* que está sendo autorizado, não o usuário.

Há duas maneiras de realizar essas autenticações.

- **Logon único do Office (SSO)**: Um sistema, * que está atualmente no modo de visualização.*, permite que o logon do usuário do Office funcione também como um logon no suplemento. Opcionalmente, o suplemento também pode usar as credenciais do usuário do Office para autorizar o suplemento ao Microsoft Graph. (As fontes que não são da Microsoft não podem ser acessadas por este sistema.)
- **Autenticação e Autorização de Aplicativos Web com o Azure Active Directory**: Não é algo novo ou especial. É apenas a maneira como o suplemento do Office (e outros aplicativos Web) autenticavam os usuários e aplicativos autorizados antes de haver um sistema de SSO do Office e ainda é usado em situações em que o SSO do Office não pode.

O fluxograma a seguir mostra as decisões que você precisa tomar como desenvolvedor de suplemento. Os detalhes estão incluídos mais adiante neste artigo.

![Uma imagem mostrando um fluxograma de decisão para habilitar a autenticação e a autorização nos suplementos do Office](../images/auth-decisions-flowchart.gif)

## <a name="user-authentication-without-sso"></a>Autenticação de usuário sem SSO

É possível autenticar um usuário em um Suplemento do Office com o Azure Active Directory (AAD) da mesma forma que em qualquer outro aplicativo Web com uma exceção: o AAD não permite que sua página de logon seja aberta em um iframe. Quando um suplemento do Office está sendo executado no *Office na Web*, o painel de tarefas é um iframe. Isso significa que você precisará abrir a tela de logon do AAD em uma caixa de diálogo aberta com a API de Diálogo do Office. Isso afeta o modo como você usa bibliotecas auxiliares de autenticação. Para saber mais, confira [Autenticação com a API de Diálogo do Office](auth-with-office-dialog-api.md).

Para obter informações sobre a autenticação de programação com o AAD, comece com [Visão geral da plataforma de Identidade da Microsoft (v 2.0)](/azure/active-directory/develop/v2-overview). Há muitos tutoriais e guias nesse conjunto de documentos, bem como links para exemplos e bibliotecas relevantes. Conforme explicado em [Autenticação com a API de Diálogo do Office](auth-with-office-dialog-api.md), talvez seja necessário ajustar o código nos exemplos para executar o Diálogo no Office.

## <a name="access-to-microsoft-graph-without-sso"></a>Acesso ao Microsoft Graph sem SSO

Você pode obter autorização para os dados do Microsoft Graph para seu suplemento obtendo um token de acesso ao Graph a partir do Azure Active Directory (AAD). Você pode fazer isso sem depender do SSO do Office. Para saber mais sobre como fazer isso, confira [Acesse o Microsoft Graph sem o SSO](authorize-to-microsoft-graph-without-sso.md) que tem mais detalhes e links para os exemplos.

## <a name="user-authentication-with-sso"></a>Autenticação do usuário com o SSO

Para usar o SSO para autenticar o usuário, seu código em um arquivo de função ou painel de tarefas chama o método [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-). Se o usuário não estiver conectado ao Office, o Office abrirá uma caixa de diálogo e o navegará para a página de logon do Azure Active Directory. Depois que o usuário estiver conectado ou se o usuário já tiver entrado, o método retorna um token de acesso. O token é um token de Bootstrap no fluxo **On Behalf Of**. (Confira [Acessar o Microsoft Graph com o SSO](#access-to-microsoft-graph-with-sso).) No entanto, ele também pode ser usado como um token de ID, pois inclui várias declarações exclusivas para o usuário atual, incluindo `preferred_username`, `name`, `sub` e `oid`. Para obter orientação sobre qual propriedade usar como a ID de usuário final, consulte [Tokensde acesso da plataforma de identidade da Microsoft](https://docs.microsoft.com/pt-BR/azure/active-directory/develop/access-tokens#payload-claims). Para obter um exemplo de um desses tokens, consulte o [Exemplo de token de acesso](sso-in-office-add-ins.md#example-access-token).

Após o seu código ter extraído a declaração desejada no token, ele usará esse valor para pesquisar o usuário em uma tabela de usuário ou banco de dados de usuário que você mantém. Use o banco de dados para armazenar informações relativas ao usuário, como as preferências do usuário ou o estado da conta do usuário. Uma vez que você está usando o SSO, os usuários não entram separadamente no seu suplemento, assim você não precisa armazenar uma senha para o usuário.

Antes de começar a implementar a autenticação do usuário com o SSO, certifique-se de que você está totalmente familiarizado com o artigo [Habilitar o logon único para Suplementos do Office](sso-in-office-add-ins.md). Observe também estes exemplos:

- [Suplemento do Office NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO), especialmente o arquivo [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts), que usa a biblioteca [jswebtoken](https://github.com/auth0/node-jsonwebtoken) para decodificar e analisar o token. (Esse exemplo, no entanto, não usa o token como um token de ID. Ele a usa para obter acesso ao Microsoft Graph com o fluxo **On Behalf Of**.)
- [O suplemento do Office ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO), especialmente o arquivo [ValuesController.ts](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Controllers/ValuesController.cs), que usa a biblioteca [System.Security.Claims.ClaimsPrincipal](https://docs.microsoft.com/dotnet/api/system.security.claims.claimsprincipal) para extrair declarações do token. (Esse exemplo, no entanto, não usa o token como um token de ID. Ele extrai uma `scope`declaração do token e a usa para obter acesso ao Microsoft Graph com o fluxo **On Behalf Of**).

## <a name="access-to-microsoft-graph-with-sso"></a>Acesso ao Microsoft Graph com oSSO

Para usar o SSO para acessar o Microsoft Graph, seu suplemento em um arquivo de função ou painel de tarefas chama o método [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-). Se o usuário não estiver conectado ao Office, o Office abrirá uma caixa de diálogo e o navegará para a página de logon do Azure Active Directory. Depois que o usuário estiver conectado ou se o usuário já tiver entrado, o método retorna um token de acesso. O token é um token de Bootstrap no fluxo **On Behalf Of**. Especificamente, ele tem uma `scope` Declaração com o valor `access_as_user`. Para mais instruções sobre as declarações no token, consulte [Tokens de acesso à plataforma de identidade da Microsoft](https://docs.microsoft.com/pt-BR/azure/active-directory/develop/access-tokens#payload-claims). Para obter um exemplo de um desses tokens, consulte o [Exemplo de token de acesso](sso-in-office-add-ins.md#example-access-token).

Após o código obter o token, ele o usará o fluxo **On Behalf Of** para obter um segundo token: um token de acesso ao Microsoft Graph.

Antes de começar a implementar o SSO do Office, certifique-se de que você está totalmente familiarizado com esses dois artigos:

- [Habilitar o logon único para Suplementos do Office](sso-in-office-add-ins.md)
- [Autorizar o Microsoft Graph com SSO](authorize-to-microsoft-graph.md)

Você também deve ler pelo menos um dos artigos descritos aqui. Mesmo que você não execute as etapas, elas contêm informações valiosas sobre a implementação do SSO do Office e o fluxo **On Behalf Of**. 

- [Criar um Suplemento do Office com ASP.NET que usa logon único](create-sso-office-add-ins-aspnet.md)
- [Crie um Suplemento do Office com Node.js que use logon único](create-sso-office-add-ins-nodejs.md)

Observe também estes exemplos:

- [SSO do NodeJS do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)
- [SSO do ASP.NET do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)

## <a name="access-to-non-microsoft-data-sources"></a>Acesso a fontes de dados que não são da Microsoft:

Serviços online populares, incluindo o Google, o Facebook, o LinkedIn, o SalesForce e o GitHub, permitem que os desenvolvedores forneçam acesso para os usuários a suas contas em outros aplicativos. Isso dá a você a capacidade de incluir esses serviços no seu Suplemento do Office. Para obter uma visão geral das maneiras como seu suplemento pode fazer isso, confira [Autorizar serviços externos em seu Suplemento do Office](auth-external-add-ins.md).

> [!IMPORTANT]
> Antes de começar a codificar, descubra se a fonte de dados permite que o logon na tela seja aberto em um iFrame. Quando um suplemento do Office está sendo executado no *Office na Web*, o painel de tarefas é um iFrame. Se a fonte de dados não permitir que sua tela de logon seja aberta em um iFrame, você precisará abrir a tela de logon em uma caixa de diálogo aberta com a API de Diálogo do Office. Para saber mais, confira [Autenticação com a API de Diálogo do Office](auth-with-office-dialog-api.md).
