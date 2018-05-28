---
title: Habilitar o logon ?nico para Suplementos do Office
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 45bd63150ffa8e46bf9c0fa54711ac907b8490ce
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a>Habilitar o logon ?nico para Suplementos do Office (visualiza??o)

Os usu?rios entram no Office (online, em dispositivos m?veis e plataformas desktop) usando tanto a conta pessoal deles da Microsoft, como a conta corporativa ou de estudante (Office 365). Voc? pode aproveitar isso e usar o logon ?nico (SSO) para autorizar o usu?rio ao seu suplemento sem exigir que o usu?rio fa?a login uma segunda vez.


![Imagem mostrando o processo de logon de um suplemento](../images/office-host-title-bar-sign-in.png)

> [!NOTE]
> Atualmente a API de logon ?nico tem suporte para Word, Excel e PowerPoint. Confira mais informa??es sobre os programas para os quais a API de logon ?nico tem suporte no momento em [Conjuntos de requisitos da IdentityAPI](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets).
> Se voc? estiver trabalhando com um suplemento do Outlook, certifique-se de habilitar a Autentica??o Moderna para a loca??o do Office 365. Confira mais informa??es sobre como fazer isso em [Exchange Online: como habilitar seu locat?rio para autentica??o moderna](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

Para os usu?rios, isso torna a experi?ncia de execu??o do suplemento mais f?cil com apenas um ?nico logon. Para os desenvolvedores, isso significa que o suplemento n?o precisa manter suas pr?prias tabelas de usu?rio com senhas criptografadas.

### <a name="how-it-works-at-runtime"></a>Como ele funciona em tempo de execu??o

O diagrama a seguir mostra como funciona o processo de SSO.

![Diagrama que mostra o processo de SSO](../images/sso-overview-diagram.png)

1. No suplemento, o JavaScript chama uma nova API Office.js `getAccessTokenAsync`. Isso notifica o aplicativo host do Office para que obtenha um token de acesso para o suplemento. Veja [Exemplo de token de acesso](#example-access-token).
2. Se o usu?rio n?o estiver conectado, o aplicativo host do Office abrir? uma janela pop-up para o usu?rio entrar.
3. Se essa ? a primeira vez que o usu?rio atual usa seu suplemento, ser? solicitado que ele d? o consentimento.
4. O aplicativo host do Office solicita o **token do suplemento** do ponto de extremidade v 2.0 do Azure AD para o usu?rio atual.
5. O Azure AD envia o token do suplemento ao aplicativo host do Office.
6. O aplicativo host do Office envia o **token do suplemento** ao suplemento como parte do objeto de resultado que retornou pela chamada de `getAccessTokenAsync`.
7. O JavaScript no suplemento pode analisar o token e extrair as informa??es necess?rias, como o endere?o de email do usu?rio. 
8. Opcionalmente, o suplemento pode enviar uma solicita??o HTTP para o servidor para obter mais dados sobre o usu?rio, tais como as prefer?ncias do usu?rio. Ou ent?o, o pr?prio token de acesso pode ser enviado para o servidor para an?lise e valida??o. 

## <a name="develop-an-sso-add-in"></a>Desenvolver um suplemento com SSO

Esta se??o descreve as tarefas envolvidas na cria??o de um suplemento do Office que usa SSO. Essas tarefas descritas aqui apresentam uma linguagem e uma estrutura de forma agn?stica. Confira exemplos de explica??es detalhadas em:

* [Criar um Suplemento do Office com Node.js que usa logon ?nico](create-sso-office-add-ins-nodejs.md)
* [Criar um Suplemento do Office com ASP.NET que usa logon ?nico](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>Criar o aplicativo de servi?o

Registre o suplemento no portal de registro para o ponto de extremidade v2.0 do Azure: https://apps.dev.microsoft.com. Esse ? um processo que leva de 5 a 10 minutos e inclui as seguintes tarefas:

* Obter uma ID do cliente e o segredo para o suplemento.
* Especificar as permiss?es que seu suplemento precisa para o AAD v. Ponto de extremidade 2.0 (e, opcionalmente, para o Microsoft Graph). A permiss?o "perfil" ? sempre necess?ria.
* Conceder a rela??o de confian?a do aplicativo host do Office para o suplemento.
* Autorizar previamente o aplicativo host do Office para o suplemento com a permiss?o padr?o *access_as_user*.

Para mais detalhes sobre este processo, veja [Registrar um suplemento do Office que usa SSO com o ponto de extremidade v2.0 do Azure AD](register-sso-add-in-aad-v2.md).

### <a name="configure-the-add-in"></a>Configurar o suplemento

Adicione novas marca??es ao manifesto do suplemento:

* **WebApplicationInfo** ? O respons?vel dos seguintes elementos.
* **Id** - A ID do cliente do suplemento. Esta ? uma ID do aplicativo que voc? obt?m como parte do registro do suplemento. Veja [Registrar um Suplemento do Office que usa SSO com o ponto de extremidade v2.0 do Azure AD](register-sso-add-in-aad-v2.md).
* **Recurso** ? O URL do suplemento.
* **Escopos** ? O respons?vel de um ou mais elementos de **Escopo**.
* **Escopo** ? Especifica uma permiss?o que o suplemento precisa para o AAD. A permiss?o `profile` ? sempre necess?ria e pode ser a ?nica permiss?o necess?ria, se seu suplemento n?o acessar o Microsoft Graph. Se isso acontecer, voc? tamb?m precisar? dos elementos do **Escopo** para as permiss?es necess?rias do Microsoft Graph; por exemplo, `User.Read`, `Mail.Read`. Bibliotecas que voc? usa no seu c?digo para acessar o Microsoft Graph podem precisar de permiss?es adicionais. Por exemplo, a Microsoft Authentication Library (MSAL) para .NET requer a permiss?o `offline_access`. Para mais informa??es, veja [Autorizar para o Microsoft Graph de um suplemento do Office](authorize-to-microsoft-graph.md).

Para hosts do Office diferentes do Outlook, adicione a marca??o no final da se??o `<VersionOverrides ... xsi:type="VersionOverridesV1_0">`. Para o Outlook, adicione a marca??o no final da se??o `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.

Veja a seguir um exemplo da marca??o:

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

### <a name="add-client-side-code"></a>Adicionar c?digo do cliente

Adicionar o JavaScript ao suplemento para:

* Chamar [Office.context.auth.getAccessTokenAsync](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync).
* Analisar o token de acesso ou pass?-lo para o c?digo do servidor do suplemento. 

Aqui est? um exemplo simples de uma chamada para `getAccessTokenAsync`. 

> [!Note]
> Este exemplo manipula apenas um tipo de erro explicitamente. Para exemplos de manipula??o de erro mais elaborados, veja [Home.js no Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) e [program.js em Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js). E veja [Solucionar problemas de mensagens de erro no logon ?nico (SSO)](troubleshoot-sso-in-office-add-ins.md).
 

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

Aqui est? um exemplo simples de passagem do token do suplemento para o servidor. O token ? inclu?do como um cabe?alho `Authorization` ao enviar uma solicita??o de volta para o servidor. Este exemplo prev? o envio de dados JSON e, portanto, ele usa o m?todo `POST`, mas `GET` ? suficiente para enviar o token de acesso quando voc? n?o estiver gravando no servidor.

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

#### <a name="when-to-call-the-method"></a>Quando chamar o m?todo

Se o seu suplemento n?o puder ser usado quando nenhum usu?rio estiver conectado no Office, voc? dever? chamar `getAccessTokenAsync` *quando o suplemento for iniciado*.

Se o suplemento tiver alguma funcionalidade que n?o exija um usu?rio conectado, voc? poder? chamar `getAccessTokenAsync` *quando o usu?rio realizar uma a??o que exija um usu?rio conectado*. N?o h? uma degrada??o do desempenho significativa com chamadas redundantes de `getAccessTokenAsync` porque o Office armazena em cache o token de acesso e o reutiliza at? que ele expire, sem fazer outra chamada para o ponto de extremidade v2.0 do AAD sempre que o `getAccessTokenAsync` for chamado. Portanto, voc? pode adicionar chamadas de `getAccessTokenAsync` para todas as fun??es e manipuladores que iniciam uma a??o onde o token ? necess?rio.

### <a name="add-server-side-code"></a>Adicionar c?digo do servidor

Na maioria dos cen?rios, n?o haver? muitas raz?es para obter o token de acesso, se o suplemento n?o o passar no lado do servidor e o utilizar l?. Algumas tarefas do servidor que seu suplemento pode fazer:

* Criar um ou mais m?todos da API da Web que usem informa??es sobre o usu?rio extra?do do token; por exemplo, um m?todo que procura as prefer?ncias do usu?rio em sua base de dados hospedada. (Veja **Usar o token SSO como uma identidade** abaixo.) Dependendo do seu idioma e estrutura, as bibliotecas podem estar dispon?veis para simplificar o c?digo que voc? precisa escrever.
* Obter dados do Microsoft Graph. O c?digo do servidor precisa fazer o seguinte:

    * Validar o token de acesso (veja **Validar o token de acesso** abaixo).
    * Iniciar o fluxo "em nome de" com uma chamada para o ponto de extremidade v2.0 do Azure AD que inclui o token de acesso, alguns metadados sobre o usu?rio e as credenciais do suplemento (sua ID e segredo). Nesse contexto, o token de acesso ? chamado de token de inicializa??o.
    * Armazenar em cache o novo token de acesso que ? retornado do fluxo em nome de.
    * Obter os dados do Microsoft Graph usando o novo token.

 Para mais detalhes sobre como obter acesso autorizado aos dados do Microsoft Graph do usu?rio, veja [Autorizar para o Microsoft Graph no seu Suplemento do Office](authorize-to-microsoft-graph.md).

#### <a name="validate-the-access-token"></a>Validar o token de acesso

Ap?s a API Web receber o token de acesso, ela deve valid?-lo antes que ele possa ser usado. O token ? um Token Web JSON (JWT) e isso significa que valida??o funciona como uma valida??o de token na maioria dos fluxos padr?o do OAuth. H? diversas bibliotecas dispon?veis que podem lidar com a valida??o de JWT. No entanto, as no??es b?sicas incluem:

- Verificar se o token foi bem formado
- Verificar se o token foi emitido pela autoridade desejada
- Verificar se o token est? direcionado para a API Web

Ao validar o token, lembre-se das seguintes diretrizes:

- Os tokens SSO v?lidos ser?o emitidos pela autoridade do Azure, `https://login.microsoftonline.com`. A declara??o `iss` no token deve come?ar com esse valor.
- O par?metro `aud` do token ser? configurado como a ID de aplicativo do registro do suplemento.
- O par?metro `scp` do token ser? definido como `access_as_user`.

#### <a name="using-the-sso-token-as-an-identity"></a>Usar o token SSO como uma identidade

Se o suplemento precisar verificar a identidade do usu?rio, o token SSO cont?m informa??es que podem ser usadas para estabelecer a identidade. As seguintes declara??es no token est?o relacionadas ? identidade.

- `name` ? O nome para exibi??o do usu?rio.
- `preferred_username` O endere?o de email do usu?rio.
- `oid` ? Um GUID que representa a ID do usu?rio no Active Directory do Azure.
- `tid` ? Um GUID que representa a ID da organiza??o do usu?rio no Active Directory do Azure.

Como os valores `name` e `preferred_username` podem mudar, recomendamos que os valores `oid` e `tid` sejam usados ??para correlacionar a identidade com o servi?o de autoriza??o do back-end.

Por exemplo, o servi?o poderia formatar os valores em conjunto como `{oid-value}@{tid-value}` e armazen?-los como um valor no registro do usu?rio no banco de dados do usu?rio interno. Em seguida, nas solicita??es subsequentes, o usu?rio poderia ser recuperado usando o mesmo valor e o acesso a recursos espec?ficos poderia ser determinado com base em seus mecanismos de controle de acesso existentes.

### <a name="example-access-token"></a>Exemplo de token de acesso

A seguir, um conte?do decodificado t?pico de um token de acesso. Para mais informa??es sobre as propriedades, veja [Refer?ncia de tokens do Active Directory do Azure v2.0](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-v2-tokens).


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

Existem algumas diferen?as pequenas, mas importantes, no uso do SSO com o suplemento do Outlook para us?-lo como suplemento do Excel, PowerPoint ou Word. Certifique-se de ler [Autenticar um usu?rio com um token de logon ?nico em um suplemento do Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/authenticate-a-user-with-an-sso-token) e [Cen?rio: implementar o logon ?nico no servi?o em um suplemento do Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/implement-sso-in-outlook-add-in).