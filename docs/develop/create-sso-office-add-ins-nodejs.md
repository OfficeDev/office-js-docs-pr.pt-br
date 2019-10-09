---
title: Crie um Suplemento do Office com Node.js que use logon único
description: ''
ms.date: 08/21/2019
localization_priority: Priority
ms.openlocfilehash: 65efb7b4423a2764bcc07e3105dfb87292895297
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695795"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a>Crie um Suplemento do Office com Node.js que use logon único (prévia)

Os usuários podem entrar no Office, e o Suplemento Web do Office pode aproveitar esse processo de entrada para autorizá-los a acessar seu suplemento e o Microsoft Graph sem exigir que os eles entrem uma segunda vez. Para obter uma visão geral, confira o artigo [Habilitar o SSO em um Suplemento do Office](sso-in-office-add-ins.md).

Este artigo apresenta o processo passo a passo de habilitação do logon único (SSO) em um suplemento que foi criado com Node.js e Express.

> [!NOTE]
> Para ler um artigo semelhante sobre um suplemento baseado em ASP.NET, confira [Criar um Suplemento do Office com ASP.NET que usa o logon único](create-sso-office-add-ins-aspnet.md).

## <a name="prerequisites"></a>Pré-requisitos

* [Node e npm](https://nodejs.org/en/), versão 6.9.4 ou posterior

* [Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)

* TypeScript, versão 2.2.2 ou posterior

* Office 365 (a versão de assinatura do Office). Build e versão mensal mais recentes do canal de Participante do programa Office Insider. É necessário ingressar no programa Office Insider para obter essa versão. Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1). Observe que, quando um build é promovido ao Canal Semestral de produção, o suporte para recursos de visualização, como o SSO, é desativado para esse build.

## <a name="set-up-the-starter-project"></a>Configure o projeto inicial

1. Clone ou baixe o repositório em [SSO com Suplemento NodeJS do Office](https://github.com/officedev/office-add-in-nodejs-sso).

    > [!NOTE]
    > Há três versões do exemplo:  
    > * A pasta **Before** (antes) traz um projeto inicial. A interface do usuário e outros aspectos do suplemento que não estão diretamente ligados ao SSO ou à autorização já estão prontos. As próximas seções deste artigo apresentam uma orientação passo a passo para concluir o projeto.
    > * A versão **Completed** (concluído) do exemplo apresenta como seria o suplemento quando concluídos os procedimentos apresentados neste artigo, com exceção de que o projeto concluído traz comentários de códigos que seriam redundantes neste artigo. Para usar a versão concluída, apenas siga as instruções apresentadas neste artigo, substituindo "Before" por "Completed" e pulando as seções **Codificar o lado do cliente** e **Codificar o lado do servidor**.
    > * A versão **Multilocatário completa** é um exemplo completo que ofereça suporte para multilocação. Explore este exemplo, se você pretende oferecer suporte para contas da Microsoft de domínios diferentes com SSO.

    > [!IMPORTANT]
    > Independentemente de qual versão você usa, será necessário confiar em um certificado para um host local. Siga [essas instruções para instalar certificados autoassinados](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md), exceto que as pastas `certs` de cada uma das versões neste repositório estão na pasta `/src`, não na pasta raiz.

1. Abra um console Git bash na pasta **Before**.

1. Insira `npm install` no console para instalar todas as dependências discriminadas no arquivo package.json.

1. Insira `npm run build` no console para compilar o projeto.

    > [!NOTE]
    > Talvez você veja alguns erros de build informando que algumas variáveis estão declaradas mas não são usadas. Ignore esses erros. Eles são um efeito colateral, pois na versão "Before" do exemplo estão faltando alguns códigos que serão adicionados posteriormente.

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Registre o suplemento com o ponto de extremidade v2.0 do Azure AD

As instruções a seguir são escritas de forma geral, elas podem ser usadas em vários locais. Para este artigo faça o seguinte:

- Substitua o espaço reservado **$ADD-IN-NAME$** por `Office-Add-in-NodeJS-SSO`.
- Substitua o espaço reservado **$FQDN-WITHOUT-PROTOCOL$** por `localhost:3000`.
- Quando você especificar permissões na caixa de diálogo **Selecionar permissões**, marque as caixas das seguintes permissões. Somente a primeira permissão é realmente necessária pelo suplemento em si, mas a permissão `profile` é necessária para que o host do Office obtenha um token no aplicativo Web do seu suplemento.
  * Files.Read.All
  * profile

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]


## <a name="grant-administrator-consent-to-the-add-in"></a>Conceder consentimento do administrador ao suplemento

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a>Configurar o suplemento

1. Em seu editor de códigos, abra o arquivo src\server.ts. Perto da parte superior, há uma chamada para um construtor de uma classe `AuthModule`. Há alguns parâmetros de cadeia de caracteres no construtor aos quais você precisa atribuir valores.

1. Na propriedade `client_id`, substitua o espaço reservado `{client GUID}` pela ID do aplicativo que você salvou ao registrar o suplemento. Quando terminar, deverá haver apenas um GUID entre aspas simples. Não deverá haver nenhum caractere "{}"

1. Na propriedade `client_secret`, substitua o espaço reservado `{client secret}` pelo segredo do aplicativo que você salvou ao registrar o suplemento.

1. Na propriedade `audience`, substitua o espaço reservado `{audience GUID}` pela ID do aplicativo que você salvou ao registrar o suplemento. (Exatamente o mesmo valor que você atribuiu à propriedade `client_id`.)
  
1. Na cadeia de caracteres atribuída à propriedade `issuer`, você verá o espaço reservado *{O365 tenant GUID}*. Substitua pela ID de locatário do Office 365. Se você não copiou a ID de locatário quando você registrou o suplemento com AAD, use um dos métodos em [Encontrar sua ID de locatário do Office 365](/onedrive/find-your-office-365-tenant-id) para obtê-la. Quando terminar, o valor da propriedade `issuer` deve ser algo parecido com isto:

    `https://login.microsoftonline.com/12345678-1234-1234-1234-123456789012/v2.0`

1. Não altere os demais parâmetros no construtor `AuthModule`. Salve e feche o arquivo.

1. Na raiz do projeto, abra o arquivo do manifesto do suplemento "Office-Add-in-NodeJS-SSO.xml".

1. Role até o final do arquivo.

1. Logo acima da marca de fim `</VersionOverrides>`, você encontrará a marcação a seguir:

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:3000/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. Substitua o espaço reservado "{application_GUID aqui}" *nos dois lugares*, na marcação, pela ID do Aplicativo que você copiou ao registrar seu suplemento. Os "{}" não fazem parte da ID, portanto, não os inclua. Essa é a mesma ID usada para ClientID e Audience no web.config.

    > [!NOTE]
    > * O valor de **Resource** é o **URI da ID do Aplicativo** que você definiu quando adicionou a plataforma API Web no registro do suplemento.
    > * A seção **Scopes** só será usada para gerar uma caixa de diálogo de consentimento se o suplemento for vendido no AppSource.

1. Salve e feche o arquivo.

## <a name="code-the-client-side"></a>Codificar o lado do cliente

1. Abra o arquivo program.js da pasta **public**. Ele já apresenta alguns códigos:

    * Uma atribuição ao método `Office.initialize` que, por sua vez, atribui um manipulador ao evento clicar do botão `getGraphAccessTokenButton`.
    * Um método `showResult` que exibirá os dados retornados do Microsoft Graph (ou uma mensagem de erro) na parte inferior do painel de tarefas.
    * Um método `logErrors` que registrará erros de console que não são destinados ao usuário final.

1. Abaixo da atribuição a `Office.initialize`, adicione o código a seguir. Observe o seguinte sobre este código:

    * O processamento de erros no suplemento às vezes tentará novamente obter um token de acesso automaticamente, usando um conjunto diferente de opções. A variável de contador `timesGetOneDriveFilesHasRun` e as variáveis sinalizador `triedWithoutForceConsent` e `timesMSGraphErrorReceived` são usadas para garantir que o usuário não seja trocado repetidas vezes em tentativas falhas de obter um token.
    * Você criará um método `getDataWithToken` na próxima etapa, mas observe que ele define uma opção chamada `forceConsent` como `false`. Trataremos mais disso na etapa seguinte.

    ```js
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;
    var timesMSGraphErrorReceived = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }
    ```

1. Abaixo do método `getOneDriveFiles`, adicione o código a seguir. Observe o seguinte sobre este código:

    * O [getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) é a nova API no Office.js que permite que um suplemento solicite ao aplicativo host do Office (Excel, PowerPoint, Word etc.) um token de acesso ao suplemento (para o usuário conectado ao Office). O aplicativo host do Office, por sua vez, solicita o token ao ponto de extremidade 2.0 do Azure AD. Uma vez que você previamente autorizou o host do Office para o seu suplemento ao registrá-lo, o Azure AD enviará o token.
    * Se nenhum usuário estiver conectado ao Office, o host do Office solicitará que o usuário se conecte.
    * O parâmetro de opções configura o `forceConsent` como `false`. Dessa forma, não será solicitado que o usuário consinta o acesso ao host do Office ao seu suplemento sempre que ele o usar. Na primeira vez que o usuário tiver o suplemento, a chamada de `getAccessTokenAsync` falhará, mas lógica de processamento de erros que você adicionará em uma etapa posterior será automaticamente chamada com a opção `forceConsent` definida como `true` e o usuário será solicitado a consentir, mas somente essa primeira vez.
    * Você criará o método `handleClientSideErrors` em uma etapa posterior.

    ```js
    function getDataWithToken(options) {
    Office.context.auth.getAccessTokenAsync(options,
        function (result) {
            if (result.status === "succeeded") {
                TODO1: Use the access token to get Microsoft Graph data.
            }
            else {
                handleClientSideErrors(result);
            }
        });
    }
    ```

1. Substitua TODO1 pelas linhas a seguir. Você criará o método `getData` e a rota "/api/values" do lado do servidor nas etapas posteriores. Uma URL relativa é usada para o ponto de extremidade porque ela deve ser hospedada no mesmo domínio que seu suplemento.

    ```js
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. Abaixo do método `getOneDriveFiles`, adicione o seguinte. Observe isto sobre este código:

    * Este método utilitário chama um ponto de extremidade da API Web especificado e transmite a ela o mesmo token de acesso que aplicativo host do Office usou para obter acesso ao seu suplemento. No lado do servidor, esse token de acesso será usado no fluxo "on behalf of" (em nome de) para obter um token de acesso para o Microsoft Graph.
    * Você criará o método `handleServerSideErrors` em uma etapa posterior.

    ```js
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            handleServerSideErrors(result);
        });
    }
    ```

### <a name="create-the-error-handling-methods"></a>Crie os métodos de processamento de erros

1. Abaixo do método `getData`, adicione o método a seguir. Esse método processará os erros no cliente do suplemento quando o host do Office não puder obter um token de acesso para o serviço Web do suplemento. Esses erros são relatados com um código de erro, portanto, o método usa uma instrução `switch` para distingui-los.

    ```js
    function handleClientSideErrors(result) {

        switch (result.error.code) {

            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor.

            // TODO3: Handle the case where the user's sign-in or consent was aborted.

            // TODO4: Handle the case where the user is logged in with an account that is neither work or school,
            //        nor Microsoft Account.

            // TODO5: Handle the case where the Office host has not been authorized to the add-in's web service or
            //        the user has not granted the service permission to their `profile`.

            // TODO6: Handle an unspecified error from the Office host.

            // TODO7: Handle the case where the Office host cannot get an access token to the add-ins
            //        web service/application.

            // TODO8: Handle the case where the user triggered an operation that calls `getAccessTokenAsync`
            //        before a previous call of it completed.

            // TODO9: Handle the case where the add-in does not support forcing consent.

            // TODO10: Log all other client errors.
        }
    }
    ```

1. Substitua `TODO2` pelo código a seguir. O erro 13001 ocorre quando o usuário não está conectado ou quando ele cancela, sem responder, uma solicitação para fornecer um segundo fator de autenticação. Em ambos os casos, o código executará novamente o método `getDataWithToken` e definirá uma opção para forçar uma solicitação de entrada.

    ```js
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. Substitua `TODO3` pelo código a seguir. O erro 13002 ocorre quando a entrada ou o consentimento do usuário é anulado. Peça que o usuário tente novamente, mas não mais de uma vez.

    ```js
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }
        break;
    ```

1. Substitua `TODO4` pelo código a seguir. O erro 13003 ocorre quando o usuário está conectado com uma conta que não é corporativa, de estudante, nem da Microsoft. Peça que o usuário saia e entre novamente com um tipo de conta suportado.

    ```js
    case 13003:
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;
    ```

    > [!NOTE]
    > O erro 13004 não é processado neste método, pois eles ocorre apenas em desenvolvimento. Não é possível corrigi-lo pelo código de tempo de execução e não seria útil reportá-lo a um usuário final.

1. Substitua `TODO5` pelo código a seguir. O erro 13005 ocorre quando o Office não tem autorização para o serviço Web do suplemento ou o usuário não concedeu permissão ao serviço para o respectivo `profile`.

    ```js
    case 13005:
        getDataWithToken({ forceConsent: true });
        break;
    ```

1. Substitua `TODO6` pelo seguinte código. O Erro 13006 ocorre quando houve um erro não especificado no host do Office, que pode indicar a instabilidade do host. Peça ao usuário para reiniciar o Office.

    ```js
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;
    ```

1. Substitua `TODO7` pelo código a seguir. O erro 13007 ocorre quando algo deu errado com a interação do host do Office com o AAD de forma que o host não pode obter um token de acesso para o serviço Web/aplicativo dos suplementos. É possível que esse seja um problema de rede temporário. Peça que o usuário tente novamente mais tarde.

    ```js
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;
    ```

1. Substitua `TODO8` pelo código a seguir. O erro 13008 ocorre quando o usuário aciona uma operação que chama o `getAccessTokenAsync` antes que uma chamada anterior dele seja concluída.

    ```js
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```

1. Substitua `TODO9` pelo código a seguir. O erro 13009 ocorre quando o suplemento não permite forçar consentimento, mas `getAccessTokenAsync` foi chamado com a opção `forceConsent` definida como `true`. Normalmente, quando isso acontece, o código deve ser reexecutar `getAccessTokenAsync` automaticamente com a opção de consentimento definida como `false`. No entanto, em alguns casos, chamar o método com `forceConsent` definido como `true` é uma resposta automática para um erro em uma chamada para o método com a opção definida como `false`. Nesse caso, o código não deve tentar novamente, mas, em vez disso, ele deve solicitar que o usuário saia e entre novamente.

    ```js
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;


1. Replace `TODO10` with the following code.

    ```js
    default:
        logError(result);
        break;
    ```  

1. Abaixo do método `handleClientSideErrors`, adicione o seguinte método. Esse método processará os erros no serviço Web do suplemento quando algo der errado na execução do fluxo on-behalf-of ou ao obter dados do Microsoft Graph.

    ```js
    function handleServerSideErrors(result) {

        // TODO11: Handle the case where AAD asks for an additional form of authentication.

        // TODO12: Handle the case where consent has not been granted, or has been revoked.

        // TODO13: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow

        // TODO14: Handle the case where the token that the add-in's client-side sends to its
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO15: Handle the case where the token sent to Microsoft Graph in the request for
        //         data is expired or invalid.

        // TODO16: Log all other server errors.
    }
    ```

1. Substitua `TODO11` pelo código a seguir. Observação sobre este código:

    * Existem configurações do Azure Active Directory nas quais o usuário precisa fornecer fator(es) de autenticação adicional(ais) para acessar alguns objetivos do Microsoft Graph (por exemplo, o OneDrive), mesmo que o usuário possa fazer login no Office apenas com uma senha. Nesse caso, o AAD enviará uma resposta com o erro 50076, que tem uma propriedade `Claims`.
    * O host do Office deve obter um novo token com o valor **Claims** como a opção `authChallenge`. Isso instrui o AAD a solicitar ao usuário todas as formas de autenticação requeridas.

    ```js
    if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 50076){
        getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
    }
    ```

1. Substitua `TODO12` pelo seguinte código *logo abaixo da última chave de fechamento do código adicionado na etapa anterior*. Observação sobre esse código:

    * O erro 65001 significa que o consentimento para acessar o Microsoft Graph não foi concedido (ou foi revogado) para uma ou mais permissões.
    * O suplemento deverá obter um novo token com a opção `forceConsent` definida como `true`.

    ```js
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 65001){
        getDataWithToken({ forceConsent: true });
    }
    ```

1. Substitua `TODO13` pelo seguinte código *logo abaixo da última chave de fechamento do código adicionado na etapa anterior*. Observação sobre esse código:

    * O erro 70011 significa que um escopo inválido (permissão) foi solicitado. O suplemento deverá relatar o erro.
    * O código registra qualquer outro erro com um número de erro do AAD.

    ```js
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 70011){
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. Substitua `TODO14` pelo seguinte código *logo abaixo da última chave de fechamento do código adicionado na etapa anterior*. Observação sobre esse código:

    * Código de servidor criado em uma etapa posterior enviará a mensagem terminada em `... expected access_as_user` se a o escopo `access_as_user` (permissão) não for o token de acesso que o cliente do suplemento enviar para o ADD para ser usado no fluxo on-behalf-of.
    * O suplemento deverá relatar o erro.

    ```js
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('expected access_as_user') !== -1){
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. Substitua `TODO15` pelo seguinte código *logo abaixo da última chave de fechamento do código adicionado na etapa anterior*. Observação sobre esse código:

    * É improvável que um token expirado ou inválido seja enviado para o Microsoft Graph, mas, se isso acontecer, o código de servidor que você criará em uma etapa posterior terminará com a cadeia de caracteres `Microsoft Graph error`.
    * Nesse caso, o suplemento deverá iniciar o processo de autenticação completo ao redefinir o contador `timesGetOneDriveFilesHasRun` e as variáveis de sinalizador `timesGetOneDriveFilesHasRun` e, em seguida, chamando novamente o método de identificador de botão. No entanto, isso deve ser feito apenas uma vez. Se isso acontecer novamente, o erro deve ser apenas registrado.
    * O código registra o erro se isso acontecer duas vezes em sequência.

    ```js
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('Microsoft Graph error') !== -1) {
        if (!timesMSGraphErrorReceived) {
            timesMSGraphErrorReceived = true;
            timesGetOneDriveFilesHasRun = 0;
            triedWithoutForceConsent = false;
            getOneDriveFiles();
        } else {
            logError(result);
        }
    }
    ```

1. Substitua `TODO16` pelo seguinte código *logo abaixo da última chave de fechamento do código adicionado na etapa anterior*.

    ```js
    else {
        logError(result);
    }
    ```

## <a name="code-the-server-side"></a>Codifique o lado do servidor

Há dois arquivos do lado do servidor que precisam ser modificados.

- O src\auth.js fornece funções auxiliares de autorização. Ele já tem membros genéricos que são usados em uma variedade de fluxos de autorização. É preciso adicionar funções a esse arquivo para implementar o fluxo "on behalf of".
- O arquivo de src\server.js tem os membros básicos necessários para executar um servidor e o middleware do express. É necessário adicionar funções a ele que ajudam a API Web e a página inicial a obterem os dados do Microsoft Graph.

### <a name="create-a-method-to-exchange-tokens"></a>Criar um método para troca de tokens

1. Abra o arquivo \src\auth.ts. Adicione o método abaixo à classe `AuthModule`. Observe o seguinte sobre este código:

    * O parâmetro `jwt` é o token de acesso ao aplicativo. No fluxo de "on behalf of" (em nome de), ele é trocado com AAD por um token de acesso ao recurso.
    * O parâmetro scopes (escopos) tem um valor padrão, mas neste exemplo será substituído pelo código de chamada.
    * O parâmetro de recurso é opcional. Ele não deverá ser usado quando o [STS (Secure Token Service)](/previous-versions/windows-identity-foundation/ee748490(v=msdn.10)) for o ponto de extremidade do AAD V 2.0. O ponto de extremidade V 2.0 infere o recurso dos escopos e retorna um erro se um recurso é enviado na Solicitação HTTP.
    * Gerar uma exceção no bloco `catch` *não* causará o envio imediato do "500 Erro Interno do Servidor" para o cliente. Chamar o código no arquivo server.js acionará essa exceção e a transformará em uma mensagem de erro que será enviada para o cliente.

        ```typescript
        private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
            try {
                // TODO3: Construct the parameters that will be sent in the body of the
                //        HTTP Request to the STS that starts the "on behalf of" flow.
                // TODO4: Send the request to the STS.
                // TODO5: Catch errors from the STS and relay them to the client.
                // TODO6: Process the response and persist the access token to resource.
            }
            catch (exception) {
                throw new UnauthorizedError('Unable to obtain an access token to the resource'
                                            + JSON.stringify(exception),
                                            exception);
            }
        }
        ```

1. Substitua `TODO3` pelo código a seguir. Sobre este código, observe:
    * Um STS com suporte para o fluxo "on behalf of" espera determinados pares de valor/propriedade no corpo da solicitação HTTP. Esse código constrói um objeto que se tornará o corpo da solicitação.
    * Uma propriedade de recurso é adicionada ao corpo se, e somente se, um recurso é transmitido para o método.

        ```typescript
        const v2Params = {
                client_id: this.clientId,
                client_secret: this.clientSecret,
                grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
                assertion: jwt,
                requested_token_use: 'on_behalf_of',
                scope: scopes.join(' ')
            };
            let finalParams = {};
            if (resource) {
                // In JavaScript we could just add the resource property to the v2Params
                // object, but that won't compile in TypeScript.
                let v1Params  = { resource: resource };  
                for(var key in v2Params) { v1Params[key] = v2Params[key]; }
                finalParams = v1Params;
            } else {
                finalParams = v2Params;
            }
        ```

1. Substitua `TODO4` pelo código a seguir que envia a solicitação HTTP para o ponto de extremidade do token do STS.

    ```typescript
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    });
    ```

1. Substitua `TODO5` pelo código a seguir. Observe que gerar uma exceção *não* causará o envio imediato do "500 Erro Interno do Servidor" para o cliente. Chamar o código no arquivo server.js acionará essa exceção e a transformará em uma mensagem de erro que será enviada para o cliente.

    ```typescript
     if (res.status !== 200) {
        const exception = await res.json();
        throw exception;
    }
    ```

1. Substitua `TODO6` pelo código a seguir. Observe que o código persiste no token de acesso ao recurso, e é a hora de expiração, além de retorná-lo. O código de chamada pode evitar chamadas desnecessárias ao STS reutilizando um token de acesso não expirado ao recurso. Você verá como fazer isso na próxima seção.

    ```typescript  
    const json = await res.json();
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken;
    ```

1. Salve o arquivo, mas não o feche.

### <a name="create-a-method-to-get-access-to-the-resource-using-the-on-behalf-of-flow"></a>Criar um método para obter acesso ao recurso usando o fluxo "on behalf of"

1. Ainda no arquivo src/auth.ts, adicione o método abaixo à classe `AuthModule`. Observe o seguinte sobre este código:

    * Os comentários acima sobre os parâmetros para o método `exchangeForToken` aplicam-se aos parâmetros deste método também.
    * O método primeiro verifica o armazenamento persistente para um token de acesso ao recurso que não expirou e não vai expirar no próximo minuto. Ele chama o método `exchangeForToken` que você criou na última seção somente se necessário.

    ```typescript
    async acquireTokenOnBehalfOf(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        const resourceTokenExpirationTime = ServerStorage.retrieve('ResourceTokenExpiresAt');
        if (moment().add(1, 'minute').diff(await resourceTokenExpirationTime) < 1 ) {
            return ServerStorage.retrieve('ResourceToken');
        } else if (resource) {
            return this.exchangeForToken(jwt, scopes, resource);
        } else {
            return this.exchangeForToken(jwt, scopes);
        }
    }
    ```

1. Salve e feche o arquivo.

### <a name="create-the-endpoints-that-will-serve-the-add-ins-home-page-and-data"></a>Criar os pontos de extremidade que servirão aos dados e à página inicial do suplemento

1. Abra o arquivo src\server.ts.

1. Adicione o método a seguir na parte inferior do arquivo. Esse método servirá à página inicial do suplemento. O manifesto do suplemento especifica a URL da página inicial.

    ```typescript
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    }));
    ```

1. Adicione o método a seguir na parte inferior do arquivo. Este método lidará com todas as solicitações para a API `values`.

    ```typescript
    app.get('/api/values', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Relay any errors from Microsoft Graph to the client.
        // TODO11: Send to the client only the data that it actually needs.
    }));
    ```

1. Substitua `TODO7` pelo seguinte código que valida o token de acesso recebido do aplicativo host do Office. O método `verifyJWT` é definido no arquivo src\auth.ts. Ele sempre valida a audiência e o emissor. Usamos o parâmetro opcional para especificar que também desejamos que ele verifique se o escopo no token de acesso é `access_as_user`. Esta é a única permissão ao suplemento que o usuário e o host do Office precisam para obter um token de acesso para o Microsoft Graph por meio do fluxo "on behalf of".

    ```typescript
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' });
    ```

    > [!NOTE]
    > Você deve usar apenas o escopo `access_as_user` para autorizar a API que lida com o fluxo Em Nome De para os suplementos do Office. Outras APIs em seu serviço devem ter seus próprios requisitos de escopo. Isso limita o que pode ser acessado com os tokens que o Office adquire.

1. Substitua `TODO8` pelo código a seguir. Observe o seguinte sobre este código:

    * A chamada para `acquireTokenOnBehalfOf` não inclui um parâmetro de recurso porque construímos o objeto `AuthModule` (`auth`) com o ponto de extremidade V2.0 do AAD que não oferece suporte à propriedade de recurso.
    * O segundo parâmetro da chamada especifica as permissões que o suplemento precisará para obter uma lista dos arquivos e das pastas do usuário no OneDrive. (A permissão `profile` não é solicitada, porque só é necessária quando o host do Office obtém o token de acesso ao seu suplemento, e não quando você está negociando nesse token para um token de acesso para o Microsoft Graph.)

    ```typescript
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    ```

1. Substitua `TODO9` pela linha a seguir. Observe o seguinte sobre este código:

    * A classe MSGraphHelper é definida no src\msgraph-helper.ts.
    * Podemos minimizar os dados que devem ser retornados especificando que só queremos a propriedade de nome e somente os três primeiros itens.

    ```typescript
    const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");
    ```

1. Substitua `TODO10` pelo código a seguir. Observe que esse código processa erros "401 Não Autorizado" do Microsoft Graph que indicariam um token expirado ou inválido. É muito improvável que isso aconteça, pois a lógica persistente do token impede essa situação. (Confira a seção **Criar um método para obter acesso ao recurso usando o fluxo "on behalf of"** acima.) Se isso acontecer, o código transmitirá o erro para o cliente com "Erro do Microsoft Graph" no nome do erro. (Confira o método `handleClientSideErrors` que você criou no arquivo program.js em uma etapa anterior.) O código adicionado ao arquivo ODataHelper.js em uma etapa posterior ajuda a processar erros do Microsoft Graph.

    ```typescript
    if (graphData.code) {
        if (graphData.code === 401) {
            throw new UnauthorizedError('Microsoft Graph error', graphData);
        }
    }
    ```


1. Substitua `TODO11` pelo código a seguir. Observe que o Microsoft Graph retorna alguns metadados OData e uma propriedade **eTag** para cada item, mesmo se `name` é a única propriedade solicitada. O código envia somente os nomes de item para o cliente.

    ```typescript
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

1. Salve e feche o arquivo.

### <a name="add-response-handling-to-the-odatahelper"></a>Adicione processamento de respostas ao ODataHelper

1. Abra o arquivo src\odata-helper.ts. O arquivo está quase pronto. O que está ausente é o corpo do retorno de chamada para o identificador do evento “end” da solicitação. Substitua o `TODO` pelo código a seguir. Sobre este código, observe:

    * A resposta do ponto de extremidade OData pode ser um erro, por exemplo, 401, se o ponto de extremidade exigir um token de acesso e ele for inválido ou estiver expirado. Uma mensagem de erro é ainda um *mensagem*, não um erro, nas chamadas de `https.get`, portanto, a linha `on('error', reject)` no final do `https.get` não é acionada. Portanto, o código distingue mensagens de sucesso (200) de mensagens de erro e envia um objeto JSON para o chamador com o OData solicitado ou informações de erro.

    ```typescript
    var error;
    if (response.statusCode === 200) {
        // TODO1: Return the data to the caller and resolve the Promise.
    } else {
       // TODO2: Return an error object to the caller and resolve the Promise.
    }
    ```

1. Substitua `TODO1` pelo código a seguir. Observe que o código pressupõe que os dados retornados são JSON.

    ```typescript
    let parsedBody = JSON.parse(body);
    resolve(parsedBody);
    ```

1. Substitua `TODO2` pelo código a seguir. Observação sobre este código:

    * Uma resposta de erro de uma fonte de OData sempre terá um statusCode e, normalmente, um statusMessage. Algumas fontes de OData também adicionam uma propriedade de erro ao corpo da mensagem com mais informações, como uma solicitação interna ou, mais especificamente, um código e uma mensagem.
    * O objeto Promise é resolvido, não rejeitado. O `https.get` é executado quando um serviço Web chama um ponto de extremidade OData de servidor para servidor. No entanto, essa chamada chega no contexto de uma chamada de um cliente para uma Web API do serviço Web. A solicitação "externa" do cliente para o serviço Web nunca é concluída se essa solicitação "interna" for rejeitada. Além disso, a solicitação com o objeto `Error` personalizado é necessária se o chamador de `http.get` precisar transmitir erros do ponto de extremidade OData para o cliente.

    ```typescript
    error = new Error();
    error.code = response.statusCode;
    error.message = response.statusMessage;

    // The error body sometimes includes an empty space
    // before the first character, remove it or it causes an error.
    body = body.trim();
    error.bodyCode = JSON.parse(body).error.code;
    error.bodyMessage = JSON.parse(body).error.message;
    resolve(error);
    ```

1. Salve e feche o arquivo.

## <a name="deploy-the-add-in"></a>Implantar o suplemento

Agora é preciso informar ao Office onde encontrar o suplemento.

1. Crie um compartilhamento de rede ou [compartilhe uma pasta na rede](/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc770880(v=ws.11)).

1. Coloque uma cópia do arquivo de manifesto Office-Add-in-NodeJS-SSO.xml, da raiz do projeto, dentro da pasta compartilhada.

1. Inicie o PowerPoint e abra um documento.

1. Escolha a guia **Arquivo** e, então, **Opções**.

1. Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.

1. Escolha **Catálogos de Suplementos Confiáveis**.

1. No campo **URL do Catálogo**, insira o caminho de rede para o compartilhamento de pasta que contém o arquivo Office-Add-in-NodeJS-SSO.xml e escolha **Adicionar Catálogo**.

1. Selecione a caixa de seleção **Mostrar no Menu** e, em seguida, escolha **OK**.

1. Uma mensagem será exibida para informá-lo de que suas configurações serão aplicadas na próxima vez que você iniciar o Microsoft Office. Feche o PowerPoint.

## <a name="build-and-run-the-project"></a>Criar e executar o projeto

Há duas maneiras de criar e executar o projeto dependendo se você estiver ou não usando o Visual Studio Code. Em ambas as maneiras, o projeto cria e recria automaticamente e entra novamente em execução quando você faz alterações no código.

1. Se não estiver usando o Visual Studio Code:
   1. Abra um nó terminal e vá até a pasta raiz do projeto.
   1. No terminal, insira **npm run build**.
   1. Abra um segundo nó terminal e vá até a pasta raiz do projeto.
   1. No terminal, insira **npm run start**.

1. Se estiver usando o VS Code:
   1. Abra o projeto no VS Code.
   1. Pressione Ctrl+Shift+B para compilar o projeto.
   1. Pressione **F5** para executar o projeto em uma sessão de depuração.


## <a name="add-the-add-in-to-an-office-document"></a>Adicionar o suplemento em um documento do Office

1. Reinicie o PowerPoint, abra ou crie uma apresentação.

1. Se a guia **Desenvolvedor** não estiver visível na faixa de opções, habilite-a através das seguintes etapas:
   1. Navegue até **Arquivo** | **Opções** | **Personalizar faixa de opções**.
   1. Clique na caixa de seleção para habilitar o **Desenvolvedor** na árvore de nomes de controle do lado direito da página **Personalizar faixa de opções**.
   1. Pressione **OK**.

1. Na guia **Desenvolvedor** no PowerPoint, escolha **Meus Suplementos**.

1. Selecione a guia **PASTA COMPARTILHADA**.

1. Escolha **Exemplo de SSO NodeJS**e selecione **OK**.

1. Na faixa de opções **Página Inicial**, há um novo grupo chamado **SSO NodeJS** com um botão com o rótulo **Mostrar Suplemento** e um ícone.

## <a name="test-the-add-in"></a>Testar o suplemento

1. Certifique-se de ter alguns arquivos no seu OneDrive para que você possa verificar os resultados.

1. Clique no botão **Exibir Suplemento** para abrir o suplemento.

1. O suplemento é aberto na página inicial. Clique no botão **Obter Meus Arquivos do OneDrive**.

1. Se você estiver conectado ao Office, será exibida uma lista de seus arquivos e suas pastas no OneDrive, abaixo do botão. Isso poderá demorar mais de 15 segundos na primeira vez.

1. Se você não tiver entrado no Office, um pop-up será aberto e pedirá que você entre. Depois de concluir a entrada, a lista de arquivos e pastas aparecerá após alguns segundos. *Você não deve pressionar o botão uma segunda vez.*

> [!NOTE]
> Se você entrou no Office com uma ID diferente e se alguns aplicativos do Office que estavam abertos no momento continuam abertos, o Office pode não alterar de forma confiável sua ID, mesmo que pareça ter feito isso no PowerPoint. Se isso acontecer, a chamada para o Microsoft Graph pode falhar ou os dados da ID anterior podem ser retornados. Para evitar isso, certifique-se de *fechar todos os outros aplicativos do Office* antes de pressionar **Obter meus arquivos do OneDrive**.
