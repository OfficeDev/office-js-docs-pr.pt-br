---
title: Habilitar o logon único para Suplementos do Office
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: f7430bdec99fc52998a43bca98e0256dd23ce400
ms.sourcegitcommit: 28fc652bded31205e393df9dec3a9dedb4169d78
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/23/2018
ms.locfileid: "22927437"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a>Habilitar o logon único para Suplementos do Office (visualização)

Os usuários entram no Office (online, em dispositivos móveis e plataformas desktop) usando tanto a conta pessoal deles da Microsoft, como a conta corporativa ou de estudante (Office 365). Você pode aproveitar isso e usar o logon único (SSO) para autorizar o usuário ao seu suplemento sem exigir que o usuário faça login uma segunda vez.


![Imagem mostrando o processo de entrada de um suplemento](../images/office-host-title-bar-sign-in.png)

> [!NOTE]
> A API de Logon único é suportada atualmente em versão prévia para Word, Excel, Outlook e PowerPoint. Para obter mais informações sobre onde a API de Logon único é suportada no momento, consulte [conjuntos de requisitos da IdentityAPI](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets). Para usar SSO, você deverá carregar a versão beta do Office JavaScript Library na https://appsforoffice.microsoft.com/lib/beta/hosted/office.js na página HTML de inicialização do suplemento. Se você estiver trabalhando com um suplemento do Outlook, não esqueça de habilitar a Autenticação Moderna para a locação do Office 365. Para obter informações sobre como fazer isso, consulte [Exchange Online: Como habilitar o seu locatário para a autenticação moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Para os usuários, isso torna a experiência de execução do suplemento mais fácil com um único logon. Para os desenvolvedores, isso significa que o suplemento não precisa manter suas próprias tabelas de usuário com senhas criptografadas.

### <a name="how-it-works-at-runtime"></a>Como ele funciona em tempo de execução

O diagrama a seguir mostra como funciona o processo de SSO.

![Diagrama que mostra o processo de SSO](../images/sso-overview-diagram.png)

1. No suplemento, o JavaScript chama uma nova API Office.js `getAccessTokenAsync`. Isso informa ao aplicativo host do Office para obter um token de acesso para o suplemento. Veja [Exemplo de token de acesso](#example-access-token).
2. Se o usuário não estiver conectado, o aplicativo host do Office abrirá uma janela pop-up para o usuário entrar.
3. Se essa é a primeira vez que o usuário atual usa seu suplemento, será solicitado que ele dê o consentimento.
4. O aplicativo host do Office solicita o **token do suplemento** do ponto de extremidade v 2.0 do Azure AD para o usuário atual.
5. O Azure AD envia o token do suplemento ao aplicativo host do Office.
6. O aplicativo host do Office envia o **token do suplemento** ao suplemento como parte do objeto de resultado que retornou pela chamada de `getAccessTokenAsync`.
7. O JavaScript no suplemento pode analisar o token e extrair as informações necessárias, como o endereço de email do usuário. 
8. Opcionalmente, o suplemento pode enviar uma solicitação HTTP para o servidor para obter mais dados sobre o usuário, tais como as preferências do usuário. Ou então, o próprio token de acesso pode ser enviado para o servidor para análise e validação. 

## <a name="develop-an-sso-add-in"></a>Desenvolver um suplemento com SSO

Esta seção descreve as tarefas envolvidas na criação de um suplemento do Office que usa SSO. Essas tarefas descritas aqui apresentam uma linguagem e uma estrutura de forma agnóstica. Confira exemplos de explicações detalhadas em:

* [Criar um Suplemento do Office com Node.js que usa logon único](create-sso-office-add-ins-nodejs.md)
* [Criar um Suplemento do Office com ASP.NET que usa logon único](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>Criar o aplicativo de serviço

Registrar o suplemento no portal de registro para o ponto de extremidade v 2.0 Azure: https://apps.dev.microsoft.com. Esse processo leva de 5 a 10 minutos e inclui as seguintes tarefas:

* Obter uma ID de cliente e um segredo para o suplemento.
* Especificar as permissões que o seu suplemento precisa para o AAD v. Ponto de extremidade 2.0 (e, opcionalmente, para o Microsoft Graph). A permissão "perfil" é sempre necessária.
* Conceder a relação de confiança do aplicativo host do Office para o suplemento.
* Autorizar previamente o aplicativo host do Office para o suplemento com a permissão padrão *access_as_user*.

Para mais detalhes sobre este processo, veja [Registrar um suplemento do Office que usa SSO com o ponto de extremidade v2.0 do Azure AD](register-sso-add-in-aad-v2.md).

### <a name="configure-the-add-in"></a>Configurar o suplemento

Adicione novas marcações ao manifesto do suplemento:

* **WebApplicationInfo** – O pai dos seguintes elementos.
* **Id** - A ID do cliente do suplemento. Esta é uma ID do aplicativo que você obtém como parte do registro do suplemento. Veja [Registrar um Suplemento do Office que usa SSO com o ponto de extremidade v2.0 do Azure AD](register-sso-add-in-aad-v2.md)
* **Resource** – A URL do suplemento.
* **Scopes** – O pai de um ou mais elementos **Scope**.
* **Scope** – Especifica uma permissão que o suplemento precisa para o AAD. A permissão `profile` é sempre necessária e pode ser a única permissão necessária, se seu suplemento não acessar o Microsoft Graph. Se isso acontecer, você também precisará dos elementos do **Escopo** para as permissões necessárias do Microsoft Graph; por exemplo, `User.Read`, `Mail.Read`. Bibliotecas que você usa no seu código para acessar o Microsoft Graph podem precisar de permissões adicionais. Por exemplo, a Microsoft Authentication Library (MSAL) para .NET requer a permissão `offline_access`. Para mais informações, veja [Autorizar para o Microsoft Graph de um suplemento do Office](authorize-to-microsoft-graph.md).

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

* Chamar [Office.context.auth.getAccessTokenAsync](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync).
* Analisar o token de acesso ou passá-lo para o código do servidor do suplemento. 

Aqui está um exemplo simples de uma chamada para `getAccessTokenAsync`. 

> [!Note]
> Este exemplo manipula apenas um tipo de erro explicitamente. Para exemplos de manipulação de erro mais elaborados, veja [Home.js no Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) e [program.js em Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js). E veja [Solucionar problemas de mensagens de erro no logon único (SSO)](troubleshoot-sso-in-office-add-ins.md).
 

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

Se o seu suplemento não puder ser usado quando nenhum usuário estiver conectado no Office, você deverá chamar `getAccessTokenAsync` *quando o suplemento for iniciado*.

Se o suplemento tiver alguma funcionalidade que não exija um usuário conectado, você poderá chamar `getAccessTokenAsync` *quando o usuário realizar uma ação que exija um usuário conectado*. Não há uma degradação do desempenho significativa com chamadas redundantes de `getAccessTokenAsync` porque o Office armazena em cache o token de acesso e o reutiliza até que ele expire, sem fazer outra chamada para o AAD v. sempre que o `getAccessTokenAsync` for chamado. Portanto, você pode adicionar chamadas de `getAccessTokenAsync` para todas as funções e manipuladores que iniciam uma ação onde o token é necessário.

### <a name="add-server-side-code"></a>Adicionar código no lado do servidor

Na maioria dos cenários, não haverá muitas razões para obter o token de acesso, se o suplemento não o passar no lado do servidor e o utilizar lá. Algumas tarefas do servidor que seu suplemento pode fazer:

* Criar um ou mais métodos da API da Web que usem informações sobre o usuário extraído do token; por exemplo, um método que procura as preferências do usuário em sua base de dados hospedada. (Veja **Usar o token SSO como uma identidade** abaixo.) Dependendo do seu idioma e estrutura, as bibliotecas podem estar disponíveis para simplificar o código que você precisa escrever.
* Obter dados do Microsoft Graph. O código do lado do servidor precisa fazer o seguinte:

    * Validar o token de acesso (veja **Validar o token de acesso** abaixo).
    * Iniciar o fluxo "em nome de" com uma chamada para o ponto de extremidade v2.0 do Azure AD que inclui o token de acesso, alguns metadados sobre o usuário e as credenciais do suplemento (sua ID e segredo). Nesse contexto, o token de acesso é chamado de token de inicialização.
    * Armazenar em cache o novo token de acesso que é retornado do fluxo em nome de.
    * Obter os dados do Microsoft Graph usando o novo token.

 Para mais detalhes sobre como obter acesso autorizado aos dados do Microsoft Graph do usuário, veja [Autorizar para o Microsoft Graph no seu Suplemento do Office](authorize-to-microsoft-graph.md).

#### <a name="validate-the-access-token"></a>Validar o token de acesso

Após a API Web receber o token de acesso, ela deve validá-lo antes que ele possa ser usado. O token é um Token Web JSON (JWT) e isso significa que validação funciona como uma validação de token na maioria dos fluxos padrão do OAuth. Há diversas bibliotecas disponíveis que podem lidar com a validação de JWT. No entanto, as noções básicas incluem:

- Verificar se o token foi bem formado
- Verificar se o token foi emitido pela autoridade desejada
- Verificar se o token está direcionado para a API Web

Ao validar o token, lembre-se das seguintes diretrizes:

- Os tokens SSO válidos serão emitidos pela autoridade do Azure, `https://login.microsoftonline.com`. A declaração `iss` no token deve começar com esse valor.
- O parâmetro `aud` do token será configurado como a ID de aplicativo do registro do suplemento.
- O parâmetro `scp` do token será definido como `access_as_user`.

#### <a name="using-the-sso-token-as-an-identity"></a>Usar o token SSO como uma identidade

Se o suplemento precisar verificar a identidade do usuário, o token SSO contém informações que podem ser usadas para estabelecer a identidade. As seguintes declarações no token estão relacionadas à identidade.

- `name` – O nome de exibição do usuário.
- `preferred_username` O endereço de email do usuário.
- `oid` – Um GUID que representa a ID do usuário no Azure Active Directory.
- `tid` – Um GUID que representa a ID da organização do usuário no Azure Active Directory.

Como os valores `name` e `preferred_username` podem mudar, recomendamos que os valores `oid` e `tid` sejam usados ​​para correlacionar a identidade com o serviço de autorização do back-end.

Por exemplo, o serviço poderia formatar os valores em conjunto como `{oid-value}@{tid-value}` e armazená-los como um valor no registro do usuário no banco de dados do usuário interno. Em seguida, nas solicitações subsequentes, o usuário poderia ser recuperado usando o mesmo valor e o acesso a recursos específicos poderia ser determinado com base em seus mecanismos de controle de acesso existentes.

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

## <a name="using-sso-with-and-outlook-add-in"></a>Usar o SSO com o suplemento do Outlook

Existem algumas diferenças pequenas, mas importantes, no uso do SSO com o suplemento do Outlook para usá-lo como suplemento do Excel, PowerPoint ou Word. Certifique-se de ler [Autenticar um usuário com um token de logon único em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) e [Cenário: implementar o logon único no serviço em um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).