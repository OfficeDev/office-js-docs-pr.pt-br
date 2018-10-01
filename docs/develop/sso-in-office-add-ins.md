---
title: Habilitar o logon único para Suplementos do Office
description: ''
ms.date: 09/26/2018
ms.openlocfilehash: 05b5088a61df3f77a09b60dbdc3129074d5f8530
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348167"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a>Habilitar o logon único para Suplementos do Office (versão prévia)

Os usuários entram no Office (plataformas online, móveis e desktop) usando sua conta pessoal da Microsoft, sua conta corporativa ou de estudante (Office 365). Você pode aproveitar isso e usar o logon único (SSO) para autorizar o usuário ao seu suplemento sem exigir que o usuário faça login uma segunda vez.

![Imagem mostrando o processo de inicio de sessão para um suplemento](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a>Status de versão prévia

A API de Logon Único no momento é suportada somente em versão prévia. Está disponível para desenvolvedores para testes; mas não deve ser usado em um suplemento de produção. Além disso, os suplementos que usam SSO não são aceitos no [AppSource](https://appsource.microsoft.com).

Nem todos os aplicativos do Office oferecem suporte a visualização SSO. Está disponível no Word, Excel, Outlook e PowerPoint. Confira mais informações sobre o suporte da API de Logon Único em [Conjuntos de requisitos da IdentityAPI](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js).

### <a name="requirements-and-best-practices"></a>Requisitos e práticas recomendadas

Para usar SSO, você precisa carregar a versão beta da Biblioteca JavaScript para Office de `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` na página HTML de inicialização do suplemento.

Se você estiver trabalhando com um suplemento do **Outlook** , certifique-se de habilitar a Autenticação Moderna para a locação do Office 365. Confira mais informações sobre como fazer isso em [Exchange Online: Como habilitar seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Você *não* deve depender do SSO como único método de autenticação de suplementos. Você deve implementar um sistema de autenticação alternativo ao qual seu suplemento retorne em determinadas situações de erro. Você pode usar um sistema de tabelas de usuários e autenticação, ou pode utilizar um provedor de logon social. Para obter mais informações sobre como fazer isso com um suplemento do Office, consulte [Autorizar serviços externos no seu Suplemento do Office](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins). Para o *Outlook*, existe um sistema de contingência recomendado. Para mais informações, confira [Cenário: Implementar o logon único no seu serviço em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).

### <a name="how-sso-works-at-runtime"></a>Funcionamento do SSO em tempo de execução

O diagrama a seguir mostra como funciona o processo de SSO.

![Diagrama que mostra o processo de SSO](../images/sso-overview-diagram.png)

1. No suplemento, o JavaScript chama uma nova API Office.js [getAccessTokenAsync](#sso-api-reference). Isso informa ao aplicativo host do Office para obter um token de acesso para o suplemento. Veja [Exemplo de token de acesso](#example-access-token).
2. Se o usuário não estiver conectado, o aplicativo host do Office abrirá uma janela pop-up para o usuário entrar.
3. Se essa é a primeira vez que o usuário atual usa o suplemento, será solicitado que informe seu consentimento.
4. O aplicativo host do Office solicita o **token do suplemento** do ponto de extremidade v 2.0 do Azure AD para o usuário atual.
5. O Azure AD envia o token do suplemento para o aplicativo host do Office.
6. O aplicativo host do Office envia o **token do suplemento** ao suplemento como parte do objeto de resultado que retornou pela chamada `getAccessTokenAsync`.
7. O JavaScript no suplemento pode analisar o token e extrair as informações necessárias, como o endereço de email do usuário. 
8. Opcionalmente, o suplemento pode enviar uma solicitação HTTP para o lado servidor para obter mais dados sobre o usuário, como as preferências do usuário. Ou então, o próprio token de acesso pode ser enviado para o servidor para análise e validação. 

## <a name="develop-an-sso-add-in"></a>Desenvolver um suplemento com SSO

Esta seção descreve as tarefas envolvidas na criação de um Suplemento do Office que usa SSO. Essas tarefas são descritas aqui de maneira agnóstica em termos de linguagem e estrutura. Para exemplos de orientações detalhadas, consulte:

* [Criar um Suplemento do Office com Node.js que usa logon único](create-sso-office-add-ins-nodejs.md)
* [Criar um Suplemento do Office com ASP.NET que usa logon único](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>Criar o aplicativo de serviço

Registrar o suplemento no portal de registro para o ponto de extremidade v 2.0 Azure: https://apps.dev.microsoft.com. Esse processo leva de 5 a 10 minutos e inclui as seguintes tarefas:

* Obter uma ID do cliente e o segredo para o suplemento.
* Especificar as permissões que o suplemento precisa para o AAD v. Ponto de extremidade 2.0 (e, opcionalmente, para o Microsoft Graph). A permissão de "perfil" é sempre necessária.
* Conceder a relação de confiança do aplicativo host do Office para o suplemento.
* Autorizar previamente o aplicativo host do Office para o suplemento com a permissão padrão *access_as_user*.

Para mais detalhes sobre este processo, veja [Registrar um suplemento do Office que usa SSO com o ponto de extremidade v2.0 do Azure AD](register-sso-add-in-aad-v2.md).

### <a name="configure-the-add-in"></a>Configurar o suplemento

Adicione novas marcações ao manifesto do suplemento:

* **WebApplicationInfo** – o pai dos seguintes elementos.
* **Id** - A ID do cliente do suplemento. Esta é uma ID do aplicativo que você obtém como parte do registro do suplemento. Veja [Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do D do Azure v2.0](register-sso-add-in-aad-v2.md).
* **Recurso** – A URL do suplemento.
* **Escopos** – O pai de uma ou mais elementos **Escopo**.
* **Escopo** – Especifica uma permissão que o suplemento precisa para o AAD. A permissão `profile` é sempre necessária e pode ser a única permissão necessária, se seu suplemento não acessar o Microsoft Graph. Se isso acontecer, você também precisará dos elementos do **Escopo** para as permissões necessárias do Microsoft Graph; por exemplo, `User.Read`, `Mail.Read`. Bibliotecas que você usa no seu código para acessar o Microsoft Graph podem precisar de permissões adicionais. Por exemplo, a Microsoft Authentication Library (MSAL) para .NET requer a permissão `offline_access`. Para mais informações, veja [Autorizar para o Microsoft Graph de um suplemento do Office](authorize-to-microsoft-graph.md).

Para hosts do Office diferentes do Outlook, adicione a marcação no final da seção `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Para o Outlook, adicione a marcação no final da seção `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.

Veja a seguir um exemplo da marcação:

```xml
<WebApplicationInfo>
    <Id>5661fed9-f33d-4e95-b6cf-624a34a2f51d</Id>
    <Resource>api://addin.contoso.com/5661fed9-f33d-4e95-b6cf-624a34a2f51d</Resource>
    <Scopes>
        <Scope>user.read</Scope>
        <Scope>files.read</Scope>
        <Scope>profile</Scope>
    </Scopes>
</WebApplicationInfo>
```

### <a name="add-client-side-code"></a>Adicionar código do lado do cliente

Adicione o JavaScript ao suplemento para:

* Chame [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).

* Analisar o token de acesso ou passá-lo para o código do servidor do suplemento. 

Aqui está um exemplo simples de uma chamada para `getAccessTokenAsync`. 

> [!NOTE]
> Este exemplo manipula apenas um tipo de erro explicitamente. Para exemplos de manipulação de erro mais elaboradoa, veja [Home.js no Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) e [program.js em Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js). E veja [Solucionar problemas de mensagens de erro para logon único (SSO)](troubleshoot-sso-in-office-add-ins.md).
 

```js
Office.context.auth.getAccessTokenAsync(function (result) {
    if (result.status === "succeeded") {
        // Use this token to call Web API
        var ssoToken = result.value;
        ...
    } else {
        if (result.error.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // work or school (Office 365) or Microsoft Account IDs.
        } else {
            // Handle error
        }
    }
});
```

Aqui está um exemplo simples de passagem do token do suplemento para o servidor. O token é incluído como um cabeçalho `Authorization` ao enviar uma solicitação de volta para o servidor. Este exemplo prevê o envio de dados JSON e, portanto, ele usa o método `POST`, mas `GET` é suficiente para enviar o token de acesso quando você não estiver gravando no servidor.

```js
$.ajax({
    type: "POST",
    url: "/api/DoSomething",
    headers: {
        "Authorization": "Bearer " + ssoToken
    },
    data: { /* some JSON payload */ },
    contentType: "application/json; charset=utf-8"
}).done(function (data) {
    // Handle success
}).fail(function (error) {
    // Handle error
}).always(function () {
    // Cleanup
});
```

#### <a name="when-to-call-the-method"></a>Quando chamar o método

Se o seu suplemento não puder ser usado quando nenhum usuário está conectado no Office, deverá chamar `getAccessTokenAsync` *quando o suplemento for iniciado*.

Se o suplemento tiver alguma funcionalidade que não exija um usuário conectado, você poderá chamar `getAccessTokenAsync` *quando o usuário realizar uma ação que exija um usuário conectado*. Não há uma degradação do desempenho significativa com chamadas redundantes de `getAccessTokenAsync` porque o Office armazena em cache o token de acesso e o reutiliza até que ele expire, sem fazer outra chamada para o AAD v. O ponto de extremidade 2.0 sempre `getAccessTokenAsync` é chamado. Portanto, você pode adicionar chamadas de `getAccessTokenAsync` para todas as funções e manipuladores que iniciam uma ação onde o token é necessário.

### <a name="add-server-side-code"></a>Adicionar código no lado do servidor

Na maioria dos cenários, não haverá muitas razões para obter o token de acesso se o suplemento não passar para uso no servidor. Algumas tarefas do servidor que seu suplemento pode fazer:

* Crie um ou mais métodos da API Web que usem informações sobre o usuário extraídas do token; por exemplo, um método que procura as preferências do usuário na base de dados hospedada. (Veja **Usar o token SSO como uma identidade** abaixo.) Dependendo do seu idioma e estrutura, as bibliotecas podem estar disponíveis para simplificar o código que você precisa escrever.
* Obter dados do Microsoft Graph. Seu código do lado do servidor deve fazer o seguinte:

    * Validar o token de acesso (veja **Validar o token de acesso** abaixo).
    * Iniciar o fluxo "em nome de" com uma chamada para o ponto de extremidade do AD do Azure  v2.0 que inclui o token de acesso, alguns metadados sobre o usuário e as credenciais do suplemento (sua ID e segredo). Nesse contexto, o token de acesso é chamado de token de inicialização.
    * Armazenar em cache o novo token de acesso que é retornado do fluxo em nome de.
    * Obter os dados do Microsoft Graph usando o novo token.

 Para mais detalhes sobre como obter acesso autorizado aos dados do Microsoft Graph do usuário, veja [Autorizar o Microsoft Graph no Suplemento do Office](authorize-to-microsoft-graph.md).

#### <a name="validate-the-access-token"></a>Validar o token de acesso

Após a API Web receber o token de acesso, ela deve validá-lo antes de usá-lo O token é um JSON Web Token (JWT), o que significa que a validação funciona como a validação de token na maioria dos fluxos padrão OAuth. Há diversas bibliotecas disponíveis que podem manipular a validação de JWT, mas as noções básicas incluem:

- Verificar se o token foi bem formado
- Verificando se o token foi emitido pela autoridade desejada
- Verificar se o token está direcionado para a API Web

Ao validar o token, lembre-se das seguintes diretrizes:

- Os tokens SSO válidos serão emitidos pela autoridade do Azure, `https://login.microsoftonline.com`. A declaração `iss` no token deve começar com esse valor.
- O parâmetro `aud` do token será configurado para a ID de aplicativo do registro do suplemento.
- O parâmetro `scp` do token será definido como `access_as_user`.

#### <a name="using-the-sso-token-as-an-identity"></a>Usar o token SSO como uma identidade

Se o suplemento necessita verificar a identidade do usuário, o token SSO contém informações que podem ser usadas para estabelecer a identidade. As seguintes declarações no token estão relacionadas à identidade.

- `name` - Nome de exibição do usuário.
- `preferred_username` - O endereço de email do usuário.
- `oid` – Uma GUID que representa a ID do usuário no Active Directory do Azure.
- `tid` – Uma GUID que representa a ID da organização do usuário no Active Directory do Azure.

Como os valores `name` e `preferred_username` podem mudar, recomendamos que os valores `oid` e `tid` sejam usados ​​para correlacionar a identidade com o serviço de autorização do seu back-end.

Por exemplo, o serviço poderia formatar esses valores juntos como `{oid-value}@{tid-value}` e armazená-los como um valor no registro do usuário no banco de dados interno de usuário. Em seguida, nas solicitações subsequentes, o usuário pode ser recuperado usando o mesmo valor e o acesso a recursos específicos pode ser determinado com base nos mecanismos de controle de acesso existentes.

### <a name="example-access-token"></a>Exemplo de token de acesso

A seguir, um conteúdo decodificado típico de um token de acesso. Para mais informações sobre as propriedades, veja [Referência de tokens do Active Directory do Azure v2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).


```js
{
    aud: "2c3caa80-93f9-425e-8b85-0745f50c0d24",         
    iss: "https://login.microsoftonline.com/fec4f964-8bc9-4fac-b972-1c1da35adbcd/v2.0",         
    iat: 1521143967,         
    nbf: 1521143967,         
    exp: 1521147867,         
    aio: "ATQAy/8GAAAA0agfnU4DTJUlEqGLisMtBk5q6z+6DB+sgiRjB/Ni73q83y0B86yBHU/WFJnlMQJ8",         
    azp: "e4590ed6-62b3-5102-beff-bad2292ab01c",         
    azpacr: "0",         
    e_exp: 262800,         
    name: "Mila Nikolova",         
    oid: "6467882c-fdfd-4354-a1ed-4e13f064be25",         
    preferred_username: "milan@contoso.com",         
    scp: "access_as_user",         
    sub: "XkjgWjdmaZ-_xDmhgN1BMP2vL2YOfeVxfPT_o8GRWaw",         
    tid: "fec4f964-8bc9-4fac-b972-1c1da35adbcd",         
    uti: "MICAQyhrH02ov54bCtIDAA",         
    ver: "2.0"
}
```

## <a name="using-sso-with-an-outlook-add-in"></a>Como usar o SSO com um suplemento do Outlook

Existem algumas diferenças pequenas, mas importantes, no uso do SSO em um suplemento do Outlook para usá-lo como suplemento do Excel, PowerPoint ou Word. Certifique-se de ler [Autenticar um usuário com um token de logon único em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) e [Cenário: Implementar o logon único em seu serviço em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).

## <a name="sso-api-reference"></a>Referência da API de SSO

### <a name="getaccesstokenasync"></a>getAccessTokenAsync

O namespace Office Auth, `Office.context.auth`, fornece um método `getAccessTokenAsync` que permite que o host do Office obtenha um token de acesso para o aplicativo da Web do suplemento. Indiretamente, isso também permite que o suplemento acesse os dados do Microsoft Graph do usuário conectado sem exigir que o usuário entre uma segunda vez.

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

Chama o ponto de extremidade do Active Directory do Azure V 2.0 para obter um token de acesso para o aplicativo web do seu suplemento. Isso permite que os suplementos identifiquem usuários. O código do lado do servidor pode usar esse token para acessar o Microsoft Graph do aplicativo da web do suplemento usando o [fluxo OAuth "on behalf of"](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

> [!NOTE]
> No Outlook, essa API não é suportada se o suplemento for carregado em uma caixa de correio do Outlook.com ou do Gmail.

<table><tr><td>Hosts</td><td>Excel, OneNote, Outlook, PowerPoint, Word</td></tr>

 <tr><td>Conjuntos de requisitos</td><td>[IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

#### <a name="parameters"></a>Parâmetros

`options` - Opcional. Aceita um objeto `AuthOptions` (veja abaixo) para definir os comportamentos de logon.

`callback` - Opcional. Aceita um método de retorno de chamada que pode analisar o token para a ID do usuário ou usar o token no fluxo de "em nome de" para obter acesso ao Microsoft Graph. Se [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` é "sucedido", então `AsyncResult.value` é o AAD v não processado. Token de acesso formatado para 2.0.

A `AuthOptions` interface oferece opções para a experiência do usuário quando o Office obtém um token de acesso para o suplemento do AAD v. 2.0 com o método `getAccessTokenAsync`.

```typescript
interface AuthOptions {
    /**
        * Causes Office to display the add-in consent experience. Useful if the add-in's Azure permissions have changed or if the user's consent has 
        * been revoked.
        */
    forceConsent?: boolean,
    /**
        * Prompts the user to add their Office account (or to switch to it, if it is already added).
        */
    forceAddAccount?: boolean,
    /**
        * Causes Office to prompt the user to provide the additional factor when the tenancy being targeted by Microsoft Graph requires multifactor 
        * authentication. The string value identifies the type of additional factor that is required. In most cases, you won't know at development 
        * time whether the user's tenant requires an additional factor or what the string should be. So this option would be used in a "second try" 
        * call of getAccessTokenAsync after Microsoft Graph has sent an error requesting the additional factor and containing the string that should 
        * be used with the authChallenge option.
        */
    authChallenge?: string
    /**
        * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
        */
    asyncContext?: any
}
```



