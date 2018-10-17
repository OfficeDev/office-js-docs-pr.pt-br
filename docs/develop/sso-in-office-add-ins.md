---
title: Habilitar o logon único para Suplementos do Office
description: ''
ms.date: 09/26/2018
ms.openlocfilehash: 1a75f7d619d2375a2f7fcb07f6afb7e0d6261ead
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579902"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a>Habilitar o logon único para Suplementos do Office (versão prévia)

Os usuários entram no Office (plataformas online, móveis e desktop) usando uma conta pessoal da Microsoft ou contas do trabalho ou da escola (Office 365). Você pode aproveitar isso e usar o logon único (SSO) para autorizar que o usuário use o seu suplemento sem exigir que ele entre uma segunda vez.

![Imagem mostrando o processo de logon de um suplemento](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a>Status da versão prévia

A API de logon único é suportada somente no modo de visualização neste momento. Ela está disponível para experimentação dos desenvolvedores; mas não deve ser usada em um suplemento de produção. Além disso, os suplementos que usam SSO não são aceitos no [AppSource](https://appsource.microsoft.com).

Nem todos os aplicativos do Office oferecem suporte para versão prévia do SSO. Ele está disponível no Word, Excel, Outlook e PowerPoint. Para obter mais informações sobre onde a API de logon único é suportada no momento, confira [Conjuntos de requisitos IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).

### <a name="requirements-and-best-practices"></a>Requisitos e melhores práticas

Para usar o SSO, você deve carregar a versão beta da Biblioteca JavaScript do Office em `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` na página HTML de inicialização do suplemento.

Se estiver usando um suplemento do **Outlook** , você deve habilitar a autenticação moderna para os locatários do Office 365. Para obter mais informações sobre isso, confira  [Exchange Online: como habilitar o seu locatário para autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Você *não* deve depender do SSO como único método de autenticação do seu suplemento. Você deve implementar um sistema alternativo de autenticação ao qual seu suplemento possa recorrer em determinadas situações de erro. Você pode usar um sistema de autenticação e de tabelas de usuário, ou você pode aproveitar um dos provedores de logon social. Para mais informações sobre como fazer isso com um suplemento do Office, confira [Autorizar serviços externos no seu suplemento do Office](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins). Para o *Outlook*, existe um sistema alternativo recomendado. Para mais informações, confira [Cenário: implementar o logon único para seu serviço em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).

### <a name="how-sso-works-at-runtime"></a>Funcionamento do SSO em tempo de execução

O diagrama a seguir mostra como funciona o processo de SSO.

![Diagrama que mostra o processo de SSO](../images/sso-overview-diagram.png)

1. No suplemento, o JavaScript chama uma nova API Office.js[getAccessTokenAsync](#sso-api-reference). Isso informa ao aplicativo host do Office para obter um token de acesso para o suplemento. Consulte [Exemplo de token de acesso](#example-access-token).
2. Se o usuário não estiver conectado, o aplicativo host do Office abrirá uma janela pop-up para o usuário entrar.
3. Se essa é a primeira vez que o usuário atual usa o suplemento, será solicitado que informe seu consentimento.
4. O aplicativo host do Office solicita o **token do suplemento** do ponto de extremidade v 2.0 do Azure AD para o usuário atual.
5. O Azure AD envia o token do suplemento para o aplicativo host do Office.
6. O aplicativo host do Office envia o **token do suplemento** ao suplemento como parte do objeto de resultado que retornou pela chamada `getAccessTokenAsync`.
7. O JavaScript no suplemento pode analisar o token e extrair as informações necessárias, como o endereço de email do usuário. 
8. Opcionalmente, o suplemento pode enviar a solicitação HTTP para o seu servidor visando coletar mais dados sobre o usuário, como as preferências do usuário, por exemplo. Ou o próprio token de acesso pode ser enviado para o servidor visando a análise e a validação. 

## <a name="develop-an-sso-add-in"></a>Desenvolver um suplemento com SSO

Esta seção descreve as tarefas envolvidas na criação de um suplemento do Office que usa SSO. Essas tarefas são descritas aqui de forma independente de idioma e estrutura. Para ver exemplos de passo a passo detalhado, confira:

* [Criar um suplemento do Office com Node.js que usa logon único](create-sso-office-add-ins-nodejs.md)
* [Criar um Suplemento do Office com ASP.NET que usa logon único](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>Criar o aplicativo de serviço

Registre o suplemento no portal de registro para o ponto de extremidade v2.0 do Azure: https://apps.dev.microsoft.com. Esse é um processo que leva de 5 a 10 minutos e inclui as seguintes tarefas:

* Obter uma ID do cliente e o segredo para o suplemento.
* Especifique as permissões que seu suplemento precisa para o ponto de extremidade AAD v. 2.0 (e, opcionalmente, para o Microsoft Graph). A permissão "perfil" sempre será necessária.
* Conceda a relação de confiança do aplicativo host do Office para o suplemento.
* Autorizar previamente o aplicativo host do Office para o suplemento com a permissão padrão *access_as_user*.

Para mais detalhes sobre este processo, veja [Registrar um suplemento do Office que usa SSO com o ponto de extremidade v2.0 do Azure AD](register-sso-add-in-aad-v2.md).

### <a name="configure-the-add-in"></a>Configurar o suplemento

Adicione novas marcações ao manifesto do suplemento:

* **WebApplicationInfo** – O pai dos seguintes elementos.
* **Id** - ID do cliente do suplemento. É uma ID de aplicativo que você obtém como parte do processo de registro do suplemento. Confira [Registrar um suplemento do Office que usa o SSO com o ponto de extremidade do Azure AD v2.0](register-sso-add-in-aad-v2.md).
* **Recurso** – A URL do suplemento.
* **Escopos** – O pai de um ou mais elementos **Escopo**.
* **Escopo** - Especifica uma permissão que o suplemento precisa para AAD. A permissão `profile` sempre é necessária e pode ser a única permissão necessária se seu suplemento não acessar o Microsoft Graph. Se ele tiver acesso, elementos **Escopo** também são necessários para as permissões do Microsoft Graph; por exemplo, `User.Read`, `Mail.Read`. As bibliotecas que você usar em seu código para acessar o Microsoft Graph podem precisar de permissões adicionais. Por exemplo, a biblioteca de autenticação da Microsoft (MSAL) para .NET requer a permissão `offline_access`. Para mais informações, confira [Autorizar para o Microsoft Graph a partir de um suplemento do Office](authorize-to-microsoft-graph.md).

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
> Este exemplo trata apenas de um tipo de erro explicitamente. Para obter exemplos de manipulação de erro mais elaborada, confira [Home.js no Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) e [program.js no Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js). Confira [Solucionar mensagens de erro para logon único (SSO)](troubleshoot-sso-in-office-add-ins.md).
 

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

Veja um exemplo simples de como passar o token do suplemento para o servidor. O token é incluído como um cabeçalho `Authorization` ao enviar uma solicitação de volta para o servidor. Este exemplo visualiza o envio de dados JSON. Portanto, ele usa o método `POST`, mas `GET` é suficiente para enviar o token de acesso quando você não estiver gravando para o servidor.

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

Se o seu suplemento não puder ser usado quando nenhum usuário estiver conectado ao Office, você deverá chamar `getAccessTokenAsync` *quando o suplemento for iniciado*.

Se o suplemento tiver algumas funcionalidades que não exijam um usuário conectado, você deverá chamar `getAccessTokenAsync` *quando o usuário executar uma ação onde seja necessário que o usuário esteja conectado*. Não há nenhuma degradação de desempenho significativa com chamadas redundantes de `getAccessTokenAsync` porque o Office armazena em cache o token de acesso e irá reutilizá-lo até que ele expire, sem fazer outra chamada para o ponto de extremidade AAD v. 2.0 sempre que `getAccessTokenAsync` for chamado. Portanto, você pode adicionar chamadas de `getAccessTokenAsync` para todas as funções e manipuladores que iniciam uma ação onde o token é necessário.

### <a name="add-server-side-code"></a>Adicionar código do servidor

Na maioria dos cenários, não há razão para obter o token de acesso, caso o seu suplemento não o passe para um servidor e use-o. Veja algumas tarefas do servidor que seu suplemento pode fazer:

* Criar um ou mais métodos da API da Web que usam as informações sobre o usuário extraídas do token; por exemplo, um método que procura as preferências do usuário em sua base de dados hospedada. (Confira **Usar o token de SSO como uma identidade** abaixo). Dependendo do seu idioma e da estrutura, as bibliotecas podem estar disponíveis, simplificando o código que você precisa escrever.
* Obtenha os dados do Microsoft Graph. Seu código do servidor deve fazer o seguinte:

    * Validar o token de acesso (confira **Validar o token de acesso** abaixo).
    * Inicie o fluxo "em nome de" com uma chamada para o ponto de extremidade do Azure AD v2.0 que inclui o token de acesso, alguns metadados sobre o usuário e as credenciais do suplemento (seu ID e segredo). Nesse contexto, o token de acesso é chamado token de inicialização.
    * Armazenar em cache o novo token de acesso que é retornado do fluxo em nome de.
    * Obter os dados do Microsoft Graph usando o novo token.

 Para mais detalhes sobre como obter acesso autorizado aos dados do Microsoft Graph do usuário, veja [Autorizar o Microsoft Graph no Suplemento do Office](authorize-to-microsoft-graph.md).

#### <a name="validate-the-access-token"></a>Validar o token de acesso

Depois que a API da Web receber o token de acesso, ela deverá validá-lo para utilizá-lo. O token é um JSON Web Token (JWT) e isso significa que a validação funciona como validação do token nos fluxos padrão de OAuth. Há um número de bibliotecas disponíveis que pode manipular a validação de JWT, mas os fundamentos incluem:

- Verificar se o token foi bem formado
- Verificando se o token foi emitido pela autoridade desejada
- Verificar se o token está direcionado para a API Web

Ao validar o token, lembre-se das seguintes diretrizes:

- Tokens válidos de SSO serão emitidos pela autoridade do Azure, `https://login.microsoftonline.com`. A declaração `iss` no token deve começar com esse valor.
- O parâmetro `aud` do token será configurado para a ID de aplicativo do registro do suplemento.
- O parâmetro `scp` do token será definido como `access_as_user`.

#### <a name="using-the-sso-token-as-an-identity"></a>Usar o token SSO como uma identidade

Se seu suplemento precisa verificar a identidade do usuário, o token SSO contém informações que podem ser usadas para estabelecer a identidade. As seguintes declarações no token relacionam-se com a identidade.

- `name` – O nome para exibição do usuário.
- `preferred_username` - O endereço de email do usuário.
- `oid` – Uma GUID que representa o ID do usuário no Active Directory do Azure.
- `tid` – Uma GUID que representa o ID da organização do usuário no Active Directory do Azure.

Como os valores `name` e `preferred_username` podem mudar, recomendamos que os valores `oid` e `tid` sejam usados ​​para correlacionar a identidade com o serviço de autorização do seu back-end.

Por exemplo, o seu serviço pode formatar esses valores em conjunto como `{oid-value}@{tid-value}` e armazená-los como um valor no registro do usuário em seu banco de dados de usuário interno. Nas solicitações subsequentes, o usuário pode ser recuperado usando o mesmo valor, enquanto o acesso a recursos específicos pode ser determinado com base nos seus mecanismos existentes de controle de acesso.

### <a name="example-access-token"></a>Exemplo de token de acesso

A seguir está uma carga decodificada típica de um token de acesso. Para obter informações sobre as propriedades, confira [Referência de tokens do Active Directory do Azure v2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).


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

Existem algumas diferenças pequenas, mas importantes, entre usar o SSO em um suplemento do Outlook em lugar de usá-lo em um suplemento do Excel, PowerPoint ou Word. Leia [Autenticar um usuário com um token de logon único em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) e [Cenário: implementar único logon único para seu serviço em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).

## <a name="sso-api-reference"></a>Referência da API de SSO

### <a name="getaccesstokenasync"></a>getAccessTokenAsync

O namespace do Office Auth, `Office.context.auth`, fornece um método, `getAccessTokenAsync` que permite ao host do Office obter um token de acesso para o aplicativo da web do suplemento. Indiretamente, isso também permite que o suplemento acesse dados do Microsoft Graph do usuário conectado sem exigir que o usuário entre uma segunda vez.

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

O método chama o ponto de extremidade do Active Directory do Azure V 2.0 para obter um token de acesso para o aplicativo da web do seu suplemento. Isso permite que os suplementos identifiquem usuários. O código do servidor pode usar este token para acessar o Microsoft Graph para o aplicativo da web do suplemento usando o [fluxo de OAuth "em nome de"](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

> [!NOTE]
> No Outlook, essa API não é suportada se o suplemento for carregado em uma caixa de correio do Outlook.com ou do Gmail.

<table><tr><td>Hosts</td><td>Excel, OneNote, Outlook, PowerPoint, Word</td></tr>

 <tr><td>[Conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td><td>[IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)</td></tr></table>

#### <a name="parameters"></a>Parâmetros

`options` Opcional. Aceita um objeto `AuthOptions` (veja abaixo) para definir os comportamentos de logon.

`callback` - Opcional. Aceita um método de retorno de chamada que pode analisar o token para o ID do usuário ou usar o token no fluxo de "em nome de" para obter acesso ao Microsoft Graph. Se [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` tiver "êxito", `AsyncResult.value` será o token de acesso formatado do AAD v. 2.0 bruto.

A interface `AuthOptions` oferece opções para a experiência do usuário quando o Office obtém um token de acesso para o suplemento do AAD v. 2.0 com o método `getAccessTokenAsync`.

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



