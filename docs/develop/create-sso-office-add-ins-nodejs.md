---
title: Criar um Suplemento do Office com Node.js que usa logon ?nico
description: 23/01/2018
ms.openlocfilehash: 4086471bec2ded671e1b3eafebc4fe69e9818344
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a>Crie um Suplemento do Office com Node.js que use logon ?nico (pr?via)

Os usu?rios podem entrar no Office, e o Suplemento Web do Office pode aproveitar esse processo de entrada para autoriz?-los a acessar seu suplemento e o Microsoft Graph sem exigir que os eles entrem uma segunda vez. Para obter uma vis?o geral, confira o artigo [Habilitar o SSO em um Suplemento do Office](sso-in-office-add-ins.md).

Este artigo apresenta o processo passo a passo de habilita??o do logon ?nico (SSO) em um suplemento que foi criado com Node.js e Express. 

> [!NOTE]
> Para ler um artigo semelhante sobre um suplemento baseado em ASP.NET, confira [Criar um Suplemento do Office com ASP.NET que usa o logon ?nico](create-sso-office-add-ins-aspnet.md).

## <a name="prerequisites"></a>Pr?-requisitos

* [Node e npm](https://nodejs.org/en/), vers?o 6.9.4 ou posterior

* [Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)

* TypeScript, vers?o 2.2.2 ou posterior

* Office 2016, vers?o 1708, build 8424.nnnn ou posterior (a vers?o de assinatura do Office 365, ?s vezes chamada de "Clique para Executar")

  Talvez seja necess?rio ser um Office Insider para obter essa vers?o. Para saber mais, confira [Seja um Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1).

## <a name="set-up-the-starter-project"></a>Configure o projeto inicial

1. Clone ou baixe o reposit?rio em [SSO com Suplemento NodeJS do Office](https://github.com/officedev/office-add-in-nodejs-sso). 

    > [!NOTE]
    > H? duas vers?es do exemplo:  
    > * A pasta **Before** (antes) traz um projeto inicial. A interface do usu?rio e outros aspectos do suplemento que n?o est?o diretamente ligados ao SSO ou ? autoriza??o j? est?o prontos. As pr?ximas se??es deste artigo apresentam uma orienta??o passo a passo para concluir o projeto. 
    > * A vers?o **Completed** (conclu?do) do exemplo apresenta como seria o suplemento quando conclu?dos os procedimentos apresentados neste artigo, com exce??o de que o projeto conclu?do traz coment?rios de c?digos que seriam redundantes neste artigo. Para usar a vers?o conclu?da, apenas siga as instru??es apresentadas neste artigo, substituindo "Before" por "Completed" e pulando as se??es **Codificar o lado do cliente** e **Codificar o lado do servidor**.

2. Abra um console Git bash na pasta **Before**.

3. Insira `npm install` no console para instalar todas as depend?ncias discriminadas no arquivo package.json.

4. Insira `npm run build ` no console para compilar o projeto. 

    > [!NOTE]
    > Talvez voc? veja alguns erros de build informando que algumas vari?veis est?o declaradas mas n?o s?o usadas. Ignore esses erros. Eles s?o um efeito colateral, pois na vers?o "Before" do exemplo est?o faltando alguns c?digos que ser?o adicionados posteriormente.

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Registre o suplemento com o ponto de extremidade v2.0 do Azure AD

As instru??es a seguir foram escritas de modo gen?rico para que possam ser usadas em diversos lugares. Para este artigo, fa?a o seguinte:
- Substitua o espa?o reservado **$ADD-IN-NAME$** por `?Office-Add-in-NodeJS-SSO`.
- Substitua o espa?o reservado **$FQDN-WITHOUT-PROTOCOL$** por `localhost:3000`.
- Quando voc? especificar permiss?es na caixa de di?logo **Selecionar Permiss?es**, marque as caixas para as permiss?es a seguir. Apenas a primeira ? realmente necess?ria pelo seu suplemento; mas a `profile` permiss?o ? necess?ria para que o host do Office obtenha um token para seu suplemento de aplicativo da Web.
    * Files.Read.All
    * perfil

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]


## <a name="grant-administrator-consent-to-the-add-in"></a>Conceder autoriza??o do administrador ao suplemento

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a>Configure o suplemento

1. Em seu editor de c?digos, abra o arquivo src\server.ts. Perto da parte superior, h? uma chamada para um construtor de uma classe `AuthModule`. H? alguns par?metros de cadeia de caracteres no construtor aos quais voc? precisa atribuir valores.

2. Na propriedade `client_id`, substitua o espa?o reservado `{client GUID}` pelo ID do aplicativo que voc? salvou ao registrar o suplemento. Ao terminar, deve haver apenas um GUID entre aspas simples. N?o deve haver nenhum caractere "{}".

3. Na propriedade `client_secret`, substitua o espa?o reservado `{client secret}` pelo segredo do aplicativo que voc? salvou ao registrar o suplemento.

4. Na propriedade `audience`, substitua o espa?o reservado `{audience GUID}` pela ID do aplicativo que voc? salvou ao registrar o suplemento. (Exatamente o mesmo valor que voc? atribuiu ? propriedade `client_id`.)
  
3. Na sequ?ncia atribu?da ? propriedade `issuer`, voc? ver? o espa?o reservado *{O365 tenant GUID}*. Substitua-o pelo ID de loca??o do Office 365. Use um dos m?todos em [Encontre seu ID de locat?rio do Office 365](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) para obt?-lo. Quando voc? terminar, o `issuer` valor da propriedade deve ser algo como isto:

    `https://login.microsoftonline.com/12345678-1234-1234-1234-123456789012/v2.0`

1. N?o altere os demais par?metros no construtor `AuthModule`. Salve e feche o arquivo.

1. Na raiz do projeto, abra o arquivo do manifesto do suplemento "Office-Add-in-NodeJS-SSO.xml".

1. Role at? o final do arquivo.

1. Logo acima da marca de fim `</VersionOverrides>`, voc? encontrar? a marca??o a seguir:

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

1. Substitua o espa?o reservado "{application_GUID here}" *nos dois lugares* na marca??o pelo ID do Aplicativo que voc? copiou ao registrar seu suplemento. (Os "{}"n?o fazem parte do ID, portanto, n?o inclua-os.) Esse ? o mesmo ID usado para o ClientID e Audience no web.config.

    > [!NOTE]
    > * O valor de **Resource** ? o **URI da ID do Aplicativo** que voc? definiu quando adicionou a plataforma API Web no registro do suplemento.
    > * A se??o **Scopes** s? ser? usada para gerar uma caixa de di?logo de consentimento se o suplemento for vendido no AppSource.

1. Salve e feche o arquivo.

## <a name="code-the-client-side"></a>Codificar o lado do cliente

1. Abra o arquivo program.js da pasta **public**. Ele j? apresenta alguns c?digos:

    * Uma atribui??o ao m?todo `Office.initialize` que, por sua vez, atribui um manipulador ao evento clicar do bot?o `getGraphAccessTokenButton`.
    * Um m?todo `showResult` que exibir? os dados retornados do Microsoft Graph (ou uma mensagem de erro) na parte inferior do painel de tarefas.
    * Um m?todo `logErrors` que registrar? erros de console que n?o s?o destinados ao usu?rio final.

11. Abaixo da atribui??o a `Office.initialize`, adicione o c?digo a seguir. Observe o seguinte sobre este c?digo:

    * O processamento de erros no suplemento ?s vezes tentar? novamente obter um token de acesso automaticamente, usando um conjunto diferente de op??es. A vari?vel de contador `timesGetOneDriveFilesHasRun` e as vari?veis sinalizador `triedWithoutForceConsent` e `timesMSGraphErrorReceived` s?o usadas para garantir que o usu?rio n?o seja trocado repetidas vezes em tentativas falhas de obter um token. 
    * Voc? criar? um m?todo `getDataWithToken` na pr?xima etapa, mas observe que ele define uma op??o chamada `forceConsent` como `false`. Trataremos mais disso na etapa seguinte.

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;
    var timesMSGraphErrorReceived = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }   
    ```

1. Abaixo do m?todo `getOneDriveFiles`, adicione o c?digo a seguir. Observe o seguinte sobre este c?digo:

    * O `getAccessTokenAsync` ? a nova API no Office.js que permite que um suplemento solicite ao aplicativo host do Office (Excel, PowerPoint, Word, etc.) um token de acesso para o suplemento (para o usu?rio conectado ao Office). O aplicativo host do Office, por sua vez, solicita o token ao ponto de extremidade 2.0 do Azure AD. Uma vez que voc? previamente autorizou o host do Office para o seu suplemento ao registr?-lo, o Azure AD enviar? o token.
    * Se nenhum usu?rio estiver conectado ao Office, o host do Office solicitar? que o usu?rio se conecte.
    * O par?metro de op??es configura o `forceConsent` como `false`. Dessa forma, n?o ser? solicitado que o usu?rio consinta o acesso ao host do Office ao seu suplemento sempre que ele o usar. Na primeira vez que o usu?rio tiver o suplemento, a chamada de `getAccessTokenAsync` falhar?, mas l?gica de processamento de erros que voc? adicionar? em uma etapa posterior ser? automaticamente chamada com a op??o `forceConsent` definida como `true` e o usu?rio ser? solicitado a consentir, mas somente essa primeira vez.
    * Voc? criar? o m?todo `handleClientSideErrors` em uma etapa posterior.

    ```javascript
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

1. Substitua TODO1 pelas linhas a seguir. Voc? criar? o m?todo `getData` e a rota "/api/values" do lado do servidor nas etapas posteriores. Uma URL relativa ? usada para o ponto de extremidade porque ela deve ser hospedada no mesmo dom?nio que seu suplemento.

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. Abaixo do m?todo `getOneDriveFiles`, adicione o seguinte. Observe isto sobre este c?digo:

    * Este m?todo utilit?rio chama um ponto de extremidade da API Web especificado e transmite a ela o mesmo token de acesso que aplicativo host do Office usou para obter acesso ao seu suplemento. No lado do servidor, esse token de acesso ser? usado no fluxo "on behalf of" (em nome de) para obter um token de acesso para o Microsoft Graph.
    * Voc? criar? o m?todo `handleServerSideErrors` em uma etapa posterior.

    ```javascript
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

### <a name="create-the-error-handling-methods"></a>Crie os m?todos de processamento de erros

1. Abaixo do m?todo `getData`, adicione o m?todo a seguir. Esse m?todo processar? os erros no cliente do suplemento quando o host do Office n?o puder obter um token de acesso para o servi?o Web do suplemento. Esses erros s?o relatados com um c?digo de erro, portanto, o m?todo usa uma instru??o `switch` para distingui-los.

    ```javascript
    function handleClientSideErrors(result) {

        switch (result.error.code) {
    
            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor. 
    
            // TODO3: Handle the case where the user's sign-in or consent was aborted.
    
            // TODO4: Handle the case where the user is logged in with an account that is neither work or school, 
            //        nor Micrososoft Account.
    
            // TODO5: Handle an unspecified error from the Office host.
    
            // TODO6: Handle the case where the Office host cannot get an access token to the add-ins 
            //        web service/application.
    
            // TODO7: Handle the case where the user tiggered an operation that calls `getAccessTokenAsync` 
            //        before a previous call of it completed.
    
            // TODO8: Handle the case where the add-in does not support forcing consent.
    
            // TODO9: Log all other client errors.
        }
    }
    ```

1. Substitua `TODO2` pelo c?digo a seguir. O erro 13001 ocorre quando o usu?rio n?o est? conectado ou quando ele cancela, sem responder, uma solicita??o para fornecer um segundo fator de autentica??o. Em ambos os casos, o c?digo executar? novamente o m?todo `getDataWithToken` e definir? uma op??o para for?ar uma solicita??o de entrada.

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. Substitua `TODO3` pelo c?digo a seguir. O erro 13002 ocorre quando a entrada ou o consentimento do usu?rio ? anulado. Pe?a que o usu?rio tente novamente, mas n?o mais de uma vez.

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }          
        break; 
    ```

1. Substitua `TODO4` pelo c?digo a seguir. O erro 13003 ocorre quando o usu?rio est? conectado com uma conta que n?o ? corporativa, de estudante nem da Microsoft. Pe?a que o usu?rio saia e entre novamente com um tipo de conta suportado.

    ```javascript
    case 13003: 
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;   
    ```

    > [!NOTE]
    > Os erros 13004 e 13005 n?o s?o processados neste m?todo, pois eles s? ocorrem em desenvolvimento. Eles n?o podem ser corrigidos pelo c?digo de tempo de execu??o e n?o seria ?til report?-lo a um usu?rio final.

1. Substitua `TODO5` pelo seguinte c?digo. O Erro 13006 ocorre quando houve um erro n?o especificado no host do Office, que pode indicar a instabilidade do host. Pe?a ao usu?rio para reiniciar o Office.

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;        
    ```

1. Substitua `TODO6` pelo c?digo a seguir. O erro 13007 ocorre quando algo deu errado com a intera??o do host do Office com o AAD de forma que o host n?o pode obter um token de acesso para o servi?o Web/aplicativo dos suplementos. ? poss?vel que esse seja um problema de rede tempor?rio. Pe?a que o usu?rio tente novamente mais tarde.

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;      
    ```

1. Substitua `TODO7` pelo c?digo a seguir. O Erro 13008 ocorre quando o usu?rio aciona uma opera??o que chama `getAccessTokenAsync` antes que uma chamada anterior dele seja conclu?da.

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```      

1. Substitua `TODO8` pelo c?digo a seguir. O erro 13009 ocorre quando o suplemento n?o permite for?ar consentimento, mas `getAccessTokenAsync` foi chamado com a op??o `forceConsent` definida como `true`. Normalmente, quando isso acontece, o c?digo deve ser reexecutar `getAccessTokenAsync` automaticamente com a op??o de consentimento definida como `false`. No entanto, em alguns casos, chamar o m?todo com `forceConsent` definido como `true` ? uma resposta autom?tica para um erro em uma chamada para o m?todo com a op??o definida como `false`. Nesse caso, o c?digo n?o deve tentar novamente, mas, em vez disso, ele deve solicitar que o usu?rio saia e entre novamente.

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```      
    
1. Substitua `TODO9` pelo c?digo a seguir.

    ```javascript
    default:
        logError(result);
        break;
    ```  

1. Abaixo do m?todo `handleClientSideErrors`, adicione o seguinte m?todo. Esse m?todo processar? os erros no servi?o Web do suplemento quando algo der errado na execu??o do fluxo on-behalf-of ou ao obter dados do Microsoft Graph.

    ```javascript
    function handleServerSideErrors(result) {
    
        // TODO10: Handle the case where AAD asks for an additional form of authentication.

        // TODO11: Handle the case where consent has not been granted, or has been revoked.

        // TODO12: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow

        // TODO13: Handle the case where the token that the add-in's client-side sends to it's 
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO14: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        // TODO15: Log all other server errors.
    }
    ```

1. Substitua `TODO10` pelo c?digo a seguir. Observa??o sobre este c?digo:

    * Existem configura??es do Azure Active Directory nas quais o usu?rio precisa fornecer fator(es) de autentica??o adicional(ais) para acessar alguns objetivos do Microsoft Graph (por exemplo, o OneDrive), mesmo que o usu?rio possa fazer login no Office apenas com uma senha. Nesse caso, o AAD enviar? uma resposta com o erro 50076, que tem uma propriedade `Claims`. 
    * O host do Office deve obter um novo token com o valor **Claims** como a op??o `authChallenge`. Isso instrui o AAD a solicitar ao usu?rio todas as formas de autentica??o requeridas. 

    ```javascript
    if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 50076){
        getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
    }
    ```

1. Substitua `TODO11` pelo seguinte c?digo *logo abaixo da ?ltima chave de fechamento do c?digo adicionado na etapa anterior*. Observa??o sobre esse c?digo:

    * O erro 65001 significa que o consentimento para acessar o Microsoft Graph n?o foi concedido (ou foi revogado) para uma ou mais permiss?es. 
    * O suplemento dever? obter um novo token com a op??o `forceConsent` definida como `true`.

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 65001){
        showResult(['Please grant consent to this add-in to access your Microsoft Graph data.']);        
        /*
            THE FORCE CONSENT OPTION IS NOT AVAILABLE IN DURING PREVIEW. WHEN SSO FOR
            OFFICE ADD-INS IS RELEASED, REMOVE THE showResult LINE ABOVE AND UNCOMMENT
            THE FOLLOWING LINE.
        */
        // getDataWithToken({ forceConsent: true });
    }
    ```

1. Substitua `TODO12` pelo seguinte c?digo *logo abaixo da ?ltima chave de fechamento do c?digo adicionado na etapa anterior*. Observa??o sobre esse c?digo:

    * O erro 70011 significa que um escopo inv?lido (permiss?o) foi solicitado. O suplemento dever? relatar o erro.
    * O c?digo registra qualquer outro erro com um n?mero de erro do AAD.

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 70011){
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. Substitua `TODO13` pelo seguinte c?digo *logo abaixo da ?ltima chave de fechamento do c?digo adicionado na etapa anterior*. Observa??o sobre esse c?digo:

    * C?digo de servidor criado em uma etapa posterior enviar? a mensagem terminada em `... expected access_as_user` se a o escopo `access_as_user` (permiss?o) n?o for o token de acesso que o cliente do suplemento enviar para o ADD para ser usado no fluxo on-behalf-of.
    * O suplemento dever? relatar o erro.

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('expected access_as_user') !== -1){
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. Substitua `TODO14` pelo seguinte c?digo *logo abaixo da ?ltima chave de fechamento do c?digo adicionado na etapa anterior*. Observa??o sobre esse c?digo:

    * ? improv?vel que um token expirado ou inv?lido seja enviado para o Microsoft Graph, mas, se isso acontecer, o c?digo de servidor que voc? criar? em uma etapa posterior terminar? com a cadeia de caracteres `Microsoft Graph error`.
    * Nesse caso, o suplemento dever? iniciar o processo de autentica??o completo ao redefinir o contador `timesGetOneDriveFilesHasRun` e as vari?veis de sinalizador `timesGetOneDriveFilesHasRun` e, em seguida, chamando novamente o m?todo de identificador de bot?o. No entanto, isso deve ser feito apenas uma vez. Se isso acontecer novamente, o erro deve ser apenas registrado.
    * O c?digo registra o erro se isso acontecer duas vezes em sequ?ncia.

    ```javascript
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

1. Substitua `TODO15` pelo seguinte c?digo *logo abaixo da ?ltima chave de fechamento do c?digo adicionado na etapa anterior*.

    ```javascript
    else {
        logError(result);
    }
    ```

## <a name="code-the-server-side"></a>Codifique o lado do servidor

H? dois arquivos do lado do servidor que precisam ser modificados. 
- O src\auth.js fornece fun??es auxiliares de autoriza??o. Ele j? tem membros gen?ricos que s?o usados em uma variedade de fluxos de autoriza??o. ? preciso adicionar fun??es a esse arquivo para implementar o fluxo "on behalf of".
- O arquivo de src\server.js tem os membros b?sicos necess?rios para executar um servidor e o middleware do express. ? necess?rio adicionar fun??es a ele que ajudam a API Web e a p?gina inicial a obterem os dados do Microsoft Graph.

### <a name="create-a-method-to-exchange-tokens"></a>Criar um m?todo para troca de tokens

1. Abra o arquivo \src\auth.ts. Adicione o m?todo abaixo ? classe `AuthModule`. Observe o seguinte sobre este c?digo:

    * O par?metro `jwt` ? o token de acesso ao aplicativo. No fluxo de "on behalf of" (em nome de), ele ? trocado com AAD por um token de acesso ao recurso.
    * O par?metro scopes (escopos) tem um valor padr?o, mas neste exemplo ser? substitu?do pelo c?digo de chamada.
    * O par?metro de recurso ? opcional. N?o deve ser usado quando o STS ? o ponto de extremidade V 2.0 do AAD. ele infere o recurso dos escopos e retorna um erro se um recurso ? enviado na Solicita??o HTTP. 
    * Gerar uma exce??o no bloco `catch` *n?o* causar? o envio imediato do "500 Erro Interno do Servidor" para o cliente. Chamar o c?digo no arquivo server.js acionar? essa exce??o e a transformar? em uma mensagem de erro que ser? enviada para o cliente.

        ```javascript
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

2. Substitua `TODO3` pelo c?digo a seguir. Sobre este c?digo, observe:
    * Um STS com suporte para o fluxo "on behalf of" espera determinados pares de valor/propriedade no corpo da solicita??o HTTP. Esse c?digo constr?i um objeto que se tornar? o corpo da solicita??o. 
    * Uma propriedade de recurso ? adicionada ao corpo se, e somente se, um recurso ? transmitido para o m?todo.

        ```javascript
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

3. Substitua `TODO4` pelo c?digo a seguir que envia a solicita??o HTTP para o ponto de extremidade do token do STS.

    ```javascript
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    }); 
    ```

4. Substitua `TODO5` pelo c?digo a seguir. Observe que gerar uma exce??o *n?o* causar? o envio imediato do "500 Erro Interno do Servidor" para o cliente. Chamar o c?digo no arquivo server.js acionar? essa exce??o e a transformar? em uma mensagem de erro que ser? enviada para o cliente.

    ```javascript
     if (res.status !== 200) {
        const exception = await res.json();
        throw exception;                
    } 
    ```

5. Substitua `TODO6` pelo c?digo a seguir. Observe que o c?digo persiste no token de acesso ao recurso, e ? a hora de expira??o, al?m de retorn?-lo. O c?digo de chamada pode evitar chamadas desnecess?rias ao STS reutilizando um token de acesso n?o expirado ao recurso. Voc? ver? como fazer isso na pr?xima se??o.

    ```javascript  
    const json = await res.json();
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken; 
    ```

6. Salve o arquivo, mas n?o o feche.

### <a name="create-a-method-to-get-access-to-the-resource-using-the-on-behalf-of-flow"></a>Criar um m?todo para obter acesso ao recurso usando o fluxo "on behalf of"

1. Ainda no arquivo src/auth.ts, adicione o m?todo abaixo ? classe `AuthModule`. Observe o seguinte sobre este c?digo:

    * Os coment?rios acima sobre os par?metros para o m?todo `exchangeForToken` aplicam-se aos par?metros deste m?todo tamb?m.
    * O m?todo primeiro verifica o armazenamento persistente para um token de acesso ao recurso que n?o expirou e n?o vai expirar no pr?ximo minuto. Ele chama o m?todo `exchangeForToken` que voc? criou na ?ltima se??o somente se necess?rio.

    ```javascript
    async acquireTokenOnBehalfOf(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        const resourceTokenExpirationTime = ServerStorage.retrieve('ResourceTokenExpiresAt');
        if (moment().add(1, 'minute').diff(resourceTokenExpirationTime) < 1 ) {
            return ServerStorage.retrieve('ResourceToken');
        } else if (resource) {
            return this.exchangeForToken(jwt, scopes, resource);
        } else {
            return this.exchangeForToken(jwt, scopes);
        }
    } 
    ```

2. Salve e feche o arquivo.

### <a name="create-the-endpoints-that-will-serve-the-add-ins-home-page-and-data"></a>Criar os pontos de extremidade que servir?o aos dados e ? p?gina inicial do suplemento

1. Abra o arquivo src\server.ts. 

2. Adicione o m?todo a seguir na parte inferior do arquivo. Esse m?todo servir? ? p?gina inicial do suplemento. O manifesto do suplemento especifica a URL da p?gina inicial.

    ```javascript
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    })); 
    ```

3. Adicione o m?todo a seguir na parte inferior do arquivo. Este m?todo lidar? com todas as solicita??es para a API `onedriveitems`.
    ```javascript
    app.get('/api/onedriveitems', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token 
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage 
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Relay any errors from Microsoft Graph to the client.
        // TODO11: Send to the client only the data that it actually needs.
    })); 
    ```

4. Substitua `TODO7` pelo seguinte c?digo que valida o token de acesso recebido do aplicativo host do Office. O m?todo `verifyJWT` ? definido no arquivo src\auth.ts. Ele sempre valida a audi?ncia e o emissor. Usamos o par?metro opcional para especificar que tamb?m desejamos que ele verifique se o escopo no token de acesso ? `access_as_user`. Esta ? a ?nica permiss?o ao suplemento que o usu?rio e o host do Office precisam para obter um token de acesso para o Microsoft Graph por meio do fluxo "on behalf of". 

    ```javascript
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    ```

    > [!NOTE]
    > Voc? deve usar apenas o escopo `access_as_user` para autorizar a API que lida com o fluxo Em Nome De para os suplementos do Office. Outras APIs em seu servi?o devem ter seus pr?prios requisitos de escopo. Isso limita o que pode ser acessado com os tokens que o Office adquire.

5. Substitua `TODO8` pelo c?digo a seguir. Observe o seguinte sobre este c?digo:

    * A chamada para `acquireTokenOnBehalfOf` n?o inclui um par?metro de recurso porque constru?mos o objeto `AuthModule` (`auth`) com o ponto de extremidade V2.0 do AAD que n?o oferece suporte ? propriedade de recurso.
    * O segundo par?metro da chamada especifica as permiss?es que o suplemento precisar? para obter uma lista dos arquivos e das pastas do usu?rio no OneDrive. (A permiss?o `profile` n?o ? solicitada, porque s? ? necess?ria quando o host do Office obt?m o token de acesso ao seu suplemento, e n?o quando voc? est? negociando nesse token para um token de acesso para o Microsoft Graph.)

    ```javascript
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    ```

6. Substitua `TODO9` pela linha a seguir. Observe o seguinte sobre este c?digo:

    * A classe MSGraphHelper ? definida no src\msgraph-helper.ts. 
    * Podemos minimizar os dados que devem ser retornados especificando que s? queremos a propriedade de nome e somente os tr?s primeiros itens.

    `const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");`

7. Substitua `TODO10` pelo c?digo a seguir. Observe que esse c?digo processa erros "401 N?o Autorizado" do Microsoft Graph que indicariam um token expirado ou inv?lido. ? muito improv?vel que isso aconte?a, pois a l?gica persistente do token impede essa situa??o. (Confira a se??o **Criar um m?todo para obter acesso ao recurso usando o fluxo "on behalf of"** acima.) Se isso acontecer, o c?digo transmitir? o erro para o cliente com "Erro do Microsoft Graph" no nome do erro. (Confira o m?todo `handleClientSideErrors` que voc? criou no arquivo program.js em uma etapa anterior.) O c?digo adicionado ao arquivo ODataHelper.js em uma etapa posterior ajuda a processar erros do Microsoft Graph.

    ```javascript
    if (graphData.code) {
        if (graphData.code === 401) {
            throw new UnauthorizedError('Microsoft Graph error', graphData);
        }
    }
    ```


1. Substitua `TODO11` pelo c?digo a seguir. Observe que o Microsoft Graph retorna alguns metadados OData e uma propriedade **eTag** para cada item, mesmo se `name` ? a ?nica propriedade solicitada. O c?digo envia somente os nomes de item para o cliente.

    ```javascript
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

8. Salve e feche o arquivo.

### <a name="add-response-handling-to-the-odatahelper"></a>Adicione processamento de respostas ao ODataHelper

1. Abra o arquivo src\odata-helper.ts. O arquivo est? quase pronto. O que est? ausente ? o corpo do retorno de chamada para o identificador do evento ?end? da solicita??o. Substitua o `TODO` pelo c?digo a seguir. Sobre este c?digo, observe:

    * A resposta do ponto de extremidade OData pode ser um erro, por exemplo, 401, se o ponto de extremidade exigir um token de acesso e ele for inv?lido ou estiver expirado. Uma mensagem de erro ? ainda um *mensagem*, n?o um erro, nas chamadas de `https.get`, portanto, a linha `on('error', reject)` no final do `https.get` n?o ? acionada. Portanto, o c?digo distingue mensagens de sucesso (200) de mensagens de erro e envia um objeto JSON para o chamador com o OData solicitado ou informa??es de erro.

    ```javascript
    var error;
    if (response.statusCode === 200) {
        // TODO1: Return the data to the caller and resolve the Promise.
    } else {
       // TODO2: Return an error object to the caller and resolve the Promise.
    }
    ```

1.  Substitua `TODO1` pelo c?digo a seguir. Observe que o c?digo pressup?e que os dados retornados s?o JSON.

    ```javascript
    let parsedBody = JSON.parse(body);
    resolve(parsedBody);
    ```

1.  Substitua `TODO2` pelo c?digo a seguir. Observa??o sobre este c?digo:

    * Uma resposta de erro de uma fonte de OData sempre ter? um statusCode e, normalmente, um statusMessage. Algumas fontes de OData tamb?m adicionam uma propriedade de erro ao corpo da mensagem com mais informa??es, como uma solicita??o interna ou, mais especificamente, um c?digo e uma mensagem.
    * O objeto Promise ? resolvido, n?o rejeitado. O `https.get` ? executado quando um servi?o Web chama um ponto de extremidade OData de servidor para servidor. No entanto, essa chamada chega no contexto de uma chamada de um cliente para uma Web API do servi?o Web. A solicita??o "externa" do cliente para o servi?o Web nunca ? conclu?da se essa solicita??o "interna" for rejeitada. Al?m disso, a solicita??o com o objeto `Error` personalizado ? necess?ria se o chamador de `http.get` precisar transmitir erros do ponto de extremidade OData para o cliente.

    ```javascript
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

Agora ? preciso informar ao Office onde encontrar o suplemento.

1. Crie um compartilhamento de rede ou [compartilhe uma pasta na rede](https://technet.microsoft.com/en-us/library/cc770880.aspx).

2. Coloque uma c?pia do arquivo de manifesto Office-Add-in-NodeJS-SSO.xml, da raiz do projeto, dentro da pasta compartilhada.

3. Inicie o PowerPoint e abra um documento.

4. Escolha a guia **Arquivo** e, ent?o, **Op??es**.

5. Escolha **Central de Confiabilidade**, e escolha o bot?o **Configura??es da Central de Confiabilidade**.

6. Escolha **Cat?logos de Suplementos Confi?veis**.

7. No campo **URL do Cat?logo**, insira o caminho de rede para o compartilhamento de pasta que cont?m o arquivo Office-Add-in-NodeJS-SSO.xml e escolha **Adicionar Cat?logo**.

8. Selecione a caixa de sele??o **Mostrar no Menu** e, em seguida, escolha **OK**.

9. Uma mensagem ser? exibida para inform?-lo de que suas configura??es ser?o aplicadas na pr?xima vez que voc? iniciar o Microsoft Office. Feche o PowerPoint.

## <a name="build-and-run-the-project"></a>Criar e executar o projeto

H? duas maneiras de criar e executar o projeto dependendo se voc? estiver ou n?o usando o Visual Studio Code. Em ambas as maneiras, o projeto cria e recria automaticamente e entra novamente em execu??o quando voc? faz altera??es no c?digo.

1. Se n?o estiver usando o Visual Studio Code: 
 1. Abra um n? terminal e v? at? a pasta raiz do projeto.
 2. No terminal, insira **npm run build**. 
 3. Abra um segundo n? terminal e v? at? a pasta raiz do projeto.
 4. No terminal, insira **npm run start**.

2. Se estiver usando o VS Code:
 1. Abra o projeto no VS Code.
 2. Pressione Ctrl+Shift+B para compilar o projeto.
 3. Pressione F5 para executar o projeto em uma sess?o de depura??o.


## <a name="add-the-add-in-to-an-office-document"></a>Adicionar o suplemento em um documento do Office

1. Reinicie o PowerPoint, abra ou crie uma apresenta??o. 

2. Na guia **Desenvolvedor** no PowerPoint, escolha **Meus Suplementos**.

3. Selecione a guia **PASTA COMPARTILHADA**.

4. Escolha **Exemplo de SSO NodeJS**e selecione **OK**.

5. Na faixa de op??es **P?gina Inicial**, h? um novo grupo chamado **SSO NodeJS** com um bot?o com o r?tulo **Mostrar Suplemento** e um ?cone. 

## <a name="test-the-add-in"></a>Testar o suplemento

1. Certifique-se de ter alguns arquivos no seu OneDrive para que voc? possa verificar os resultados.

2. Clique no bot?o **Exibir Suplemento** para abrir o suplemento.

2. O suplemento ? aberto na p?gina inicial. Clique no bot?o **Obter Meus Arquivos do OneDrive**.

2. Se voc? estiver conectado ao Office, ser? exibida uma lista de seus arquivos e suas pastas no OneDrive, abaixo do bot?o. Isso poder? demorar mais de 15 segundos na primeira vez.

3. Se voc? n?o tiver entrado no Office, um pop-up ser? aberto e pedir? que voc? entre. Depois de concluir a entrada, a lista de arquivos e pastas aparecer? ap?s alguns segundos. *N?o pressione o bot?o uma segunda vez.*

> [!NOTE]
> Se voc? entrou no Office com uma ID diferente e se alguns aplicativos do Office que estavam abertos no momento continuam abertos, o Office pode n?o alterar de forma confi?vel sua ID, mesmo que pare?a ter feito isso no PowerPoint. Se isso acontecer, a chamada para o Microsoft Graph pode falhar ou os dados da ID anterior podem ser retornados. Para evitar isso, certifique-se de *fechar todos os outros aplicativos do Office* antes de pressionar **Obter meus arquivos do OneDrive**.
