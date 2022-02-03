---
title: Visão geral da autenticação e autorização nos Suplementos do Office
description: Saiba como a autenticação e a autorização funcionam nos Suplementos do Office.
ms.date: 01/25/2022
ms.localizationpriority: high
ms.openlocfilehash: 1dab5e7e4cd1d5a32115bdecca3fa742699a53b9
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/02/2022
ms.locfileid: "62320120"
---
# <a name="overview-of-authentication-and-authorization-in-office-add-ins"></a>Visão geral da autenticação e autorização nos Suplementos do Office

Os Suplementos do Office permitem o acesso anônimo por padrão, mas você pode exigir que os usuários se conectem para usar seu suplemento com um conta Microsoft, uma conta corporativa ou do Microsoft 365 Education ou outra conta comum. Essa tarefa é chamada de autenticação do usuário, pois permite que o suplemento saiba quem é o usuário.

Seu suplemento também pode obter consentimento do usuário para acessar seus dados do Microsoft Graph (como seu perfil do Microsoft 365, arquivos do OneDrive e dados do SharePoint) ou para dados de outras fontes externas, como Google, Facebook, LinkedIn, SalesForce e GitHub. Essa tarefa é chamada de autorização de suplemento (ou aplicativo), pois é o *suplemento* que está sendo autorizado, não o usuário.

## <a name="key-resources-for-authentication-and-authorization"></a>Principais recursos da autenticação e da autorização

Esta documentação explica como criar e configurar Suplementos do Office para implementar a autenticação e a autorização com êxito. No entanto, muitos conceitos e tecnologias de segurança mencionados estão fora do escopo desta documentação. Por exemplo, conceitos gerais de segurança como fluxos OAuth, cache de token ou gerenciamento de identidades não são explicados aqui. Esta documentação também não apresenta nada específico sobre o Microsoft Azure e a plataforma de identidade da Microsoft. Recomendamos que você confira os recursos a seguir se precisar de mais informações nessas áreas.

- [Plataforma de identidade da Microsoft](/azure/active-directory/develop)
- [Suporte à plataforma de identidade da Microsoft e opções de ajuda para desenvolvedores](/azure/active-directory/develop/developer-support-help-options)
- [Protocolos OAuth 2.0 e OpenID Connect na plataforma de identidade da Microsoft](/azure/active-directory/develop/active-directory-v2-protocols)

## <a name="sso-scenarios"></a>Cenários de SSO

Usar o SSO (logon único) é conveniente para o usuário porque ele só precisa entrar no Office uma vez. Eles não precisam se conectar separadamente no seu suplemento. O SSO não é compatível com todas as versões do Office, portanto, você ainda precisará implementar uma abordagem de entrada alternativa, [usando a plataforma de identidade da Microsoft](#authenticate-with-the-microsoft-identity-platform). Para obter mais informações sobre as versões compatíveis do Office, confira os [Conjuntos de requisitos da API de Identidade](../reference/requirement-sets/identity-api-requirement-sets.md)

### <a name="get-the-users-identity-through-sso"></a>Obter a identidade do usuário por meio do SSO

Geralmente, seu suplemento precisa apenas da identidade do usuário. Por exemplo, talvez você queira apenas personalizar seu suplemento e exibir o nome do usuário no painel de tarefas. Ou talvez você queira uma ID exclusiva para associar o usuário aos dados em seu banco de dados. Isso pode ser feito ao obter o token de acesso do usuário do Office.

Para obter a identidade do usuário por meio do SSO, chame o método [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_). O método retorna um token de acesso que também é um token de identidade que contém várias declarações exclusivas do usuário conectado no momento, incluindo `preferred_username`, `name`, `sub` e `oid`. Para obter mais informações sobre essas propriedades, confira os [Tokens de ID da plataforma de identidade da Microsoft](/azure/active-directory/develop/id-tokens). Para obter um exemplo do token retornado por [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_), confira o [Exemplo de token de acesso](sso-in-office-add-ins.md#example-access-token).

Se o usuário não estiver conectado, o Office abrirá uma caixa de diálogo e usará a plataforma de identidade da Microsoft para solicitar que o usuário entre. Em seguida, o método retornará um token de acesso ou gerará um erro se não for possível conectar o usuário.

Em um cenário em que você precisa armazenar dados do usuário, confira [Tokens de ID da plataforma de identidade da Microsoft](/azure/active-directory/develop/id-tokens) para saber como obter um valor do token para identificar o usuário de forma exclusiva. Use esse valor para pesquisar o usuário em uma tabela de usuário ou banco de dados de usuário que você mantém. Use o banco de dados para armazenar informações relativas ao usuário, como as preferências do usuário ou o estado da conta do usuário. Uma vez que você está usando o SSO, os usuários não entram separadamente no seu suplemento, assim você não precisa armazenar uma senha para o usuário.

Antes de começar a implementar a autenticação do usuário com o SSO, certifique-se de que você está totalmente familiarizado com o artigo [Habilitar o logon único para Suplementos do Office](sso-in-office-add-ins.md).

### <a name="access-your-web-apis-through-sso"></a>Acessar suas APIs Web por meio do SSO

Se seu suplemento tiver APIs no servidor que exigem um usuário autorizado, chame o método [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_) para obter um token de acesso. O token de acesso fornece acesso ao seu próprio servidor Web (configurado por meio de um [registro de aplicativo do Microsoft Azure](register-sso-add-in-aad-v2.md)). Ao chamar APIs no seu servidor Web, você também passa o token de acesso para autorizar o usuário.

O código a seguir mostra como construir uma solicitação HTTPS GET para a API do servidor Web do suplemento para obter alguns dados. O código é executado do lado do cliente, como em um painel de tarefas. Primeiro, ele obtém o token de acesso chamando `getAccessToken`. Em seguida, ele constrói uma chamada AJAX com a URL e o cabeçalho de autorização corretos para a API do servidor.

```javascript
function getOneDriveFileNames() {

    let accessToken = await Office.auth.getAccessToken();

    $.ajax({
        url: "/api/data",
        headers: { "Authorization": "Bearer " + accessToken },
        type: "GET"
    })
        .done(function (result) {
            //... work with data from the result...
        });
}
```

O código a seguir mostra um exemplo de manipulador /api/data para a chamada REST do exemplo de código anterior. O código é um ASP.NET executado em um servidor Web. O atributo `[Authorize]` exigirá que um token de acesso válido seja passado pelo cliente ou retornará um erro ao cliente.

```csharp
    [Authorize]
    // GET api/data
    public async Task<HttpResponseMessage> Get()
    {
        //... obtain and return data to the client-side code...
    }
```

### <a name="access-microsoft-graph-through-sso"></a>Acessar o Microsoft Graph por meio do SSO

Em alguns cenários, você não precisa somente da identidade do usuário, mas também precisa acessar os recursos do [Microsoft Graph](/graph) em nome do usuário. Por exemplo, talvez seja necessário enviar um email ou criar um chat no Teams em nome do usuário. Essas ações e muito mais podem ser realizadas por meio do Microsoft Graph. Você precisará seguir essas etapas: 

1. Obtenha o token de acesso para o usuário atual por meio do SSO chamando [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_). Se o usuário não estiver conectado, o Office abrirá uma caixa de diálogo e conectará o usuário com a plataforma de identidade da Microsoft. Após o usuário entrar ou se o usuário já tiver entrado, o método retorna um token de acesso.
1. Passe o token de acesso para o código do servidor.
1. No servidor, use o [Fluxo On-Behalf-Of do OAuth 2.0](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow) para trocar o token de acesso por um novo token de acesso que contém a identidade de usuário delegada necessária e as permissões para chamar o Microsoft Graph.

> [!NOTE]
> Para melhorar a segurança e evitar o vazamento do token de acesso, sempre execute o fluxo On-Behalf-Of no servidor. Chame as APIs do Microsoft Graph a partir do seu servidor, não do cliente. Não retorne o token de acesso para o código do lado do cliente.

Antes de começar a implementar o SSO para acessar o Microsoft Graph no seu suplemento, garanta que você está completamente familiarizado com os artigos a seguir.

- [Habilitar o logon único para Suplementos do Office](sso-in-office-add-ins.md)
- [Autorizar o Microsoft Graph com SSO](authorize-to-microsoft-graph.md)

Você também deve ler pelo menos um dos seguintes artigos que o orientarão na criação de um Suplemento do Office para usar o SSO e acessar o Microsoft Graph. Mesmo que você não execute as etapas, elas contêm informações valiosas sobre a implementação do SSO e o fluxo On-Behalf-Of.

- [Criar um Suplemento do Office com ASP.NET que usa logon único](create-sso-office-add-ins-aspnet.md) que orienta você pelo exemplo em [Suplemento do Office com ASP.NET e SSO](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO).
- [Criar um Suplemento do Office com Node.js que usa logon único](create-sso-office-add-ins-nodejs.md) que orienta você pelo exemplo em [Suplemento do Office com Node.js e SSO](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO).

## <a name="non-sso-scenarios"></a>Cenários sem SSO

Em alguns cenários, talvez você não queira usar o SSO. Por exemplo, talvez seja necessário autenticar usando um provedor de identidade diferente da plataforma de identidade da Microsoft. Além disso, o SSO não é compatível com todos os cenários. Por exemplo, as versões mais antigas do Office não são compatíveis com o SSO. Nesse caso, você precisaria recorrer a um sistema de autenticação alternativo para seu suplemento.

### <a name="authenticate-with-the-microsoft-identity-platform"></a>Autenticar com a plataforma de identidade da Microsoft

Seu suplemento pode conectar usuários usando a [ plataforma de identidade da Microsoft](/azure/active-directory/develop) como provedor de autenticação. Depois que o usuário estiver conectado, você poderá usar a plataforma de identidade da Microsoft para autorizar o suplemento no [Microsoft Graph](/graph) ou em outros serviços gerenciados pela Microsoft. Use essa abordagem como um método de entrada alternativo quando o SSO por meio do Office não estiver disponível. Além disso, existem cenários nos quais você deseja que seus usuários façam logon em seu suplemento separadamente, mesmo quando o SSO estiver disponível. Por exemplo, se você quiser que os eles tenham a opção de fazer o logon no suplemento com uma ID diferente daquela com o qual eles estão atualmente conectados no Office.

É importante observar que a plataforma de identidade da Microsoft não permite que sua página de entrada seja aberta em um iframe. Quando um Suplemento do Office está sendo executado no *Office na Web*, o painel de tarefas é um iframe. Isso significa que será necessário abrir a página de entrada usando uma caixa de diálogo aberta com a API de diálogo do Office. Isso afeta o modo como você usa bibliotecas auxiliares de autenticação. Para saber mais, confira [Autenticação com a API de diálogo do Office](auth-with-office-dialog-api.md).

Para obter informações sobre como implementar a autenticação com a plataforma de identidade da Microsoft, confira a [Visão geral da plataforma de Identidade da Microsoft (v 2.0)](/azure/active-directory/develop/v2-overview). A documentação contém muitos tutoriais e guias, bem como links para exemplos e bibliotecas relevantes. Conforme explicado em [Autenticação com a API de diálogo do Office](auth-with-office-dialog-api.md), talvez seja necessário ajustar o código nos exemplos para executar na caixa de diálogo do Office.

### <a name="access-to-microsoft-graph-without-sso"></a>Acesso ao Microsoft Graph sem SSO

Você pode obter autorização para os dados do Microsoft Graph para seu suplemento obtendo um token de acesso ao Microsoft Graph a partir da plataforma de identidade da Microsoft. Você pode fazer isso sem depender do SSO por meio do Office (ou se o SSO falhou ou não é compatível). Para obter mais informações, confira [Acesse o Microsoft Graph sem o SSO](authorize-to-microsoft-graph-without-sso.md) que tem mais detalhes e links para os exemplos.

### <a name="access-to-non-microsoft-data-sources"></a>Acesso a fontes de dados que não são da Microsoft:

Serviços online populares, incluindo o Google, o Facebook, o LinkedIn, o SalesForce e o GitHub, permitem que os desenvolvedores forneçam acesso para os usuários a suas contas em outros aplicativos. Isso dá a você a capacidade de incluir esses serviços no seu Suplemento do Office. Para obter uma visão geral das maneiras como seu suplemento pode fazer isso, confira [Autorizar serviços externos em seu Suplemento do Office](auth-external-add-ins.md).

> [!IMPORTANT]
> Antes de começar a codificar, descubra se a fonte de dados permite que a página de entrada seja aberta em um iframe. Quando um Suplemento do Office está sendo executado no *Office na Web*, o painel de tarefas é um iframe. Se a fonte de dados não permitir que a página de entrada seja aberta em um iframe, você precisará abrir a página de entrada em uma caixa de diálogo aberta com a API de diálogo do Office. Para saber mais, confira [Autenticação com a API de diálogo do Office](auth-with-office-dialog-api.md).

## <a name="see-also"></a>Confira também

- [Documentação da plataforma de identidade da Microsoft](/azure/active-directory/develop/)
- [Tokens de acesso da plataforma de identidade da Microsoft](/azure/active-directory/develop/access-tokens)
- [Protocolos OAuth 2.0 e OpenID Connect na plataforma de identidade da Microsoft](/azure/active-directory/develop/active-directory-v2-protocols)
- [Plataforma de identidade da Microsoft e Fluxo On-Behalf-Of do OAuth 2.0](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
- [Token Web JSON (JWT)](https://en.wikipedia.org/wiki/JSON_Web_Token)
- [Visualizador de Token Web JSON](https://jwt.ms/)
