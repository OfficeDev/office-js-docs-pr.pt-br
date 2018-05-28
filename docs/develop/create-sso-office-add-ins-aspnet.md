---
title: Criar um Suplemento do Office com ASP.NET que usa logon ?nico
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 6a1c8ea7a8634d701a43e08fd8bb9c5f9c1863cd
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a>Criar um Suplemento do Office com ASP.NET que use logon ?nico (visualiza??o)

Quando os usu?rios est?o conectados ao Office, o seu suplemento pode usar as mesmas credenciais para permitir que os usu?rios acessem v?rios aplicativos sem exigir que eles entrem uma segunda vez. Para obter uma vis?o geral, consulte [Habilitar o SSO em um Suplemento do Office](sso-in-office-add-ins.md).

Este artigo apresenta o processo passo a passo de habilita??o do logon ?nico (SSO) em um suplemento que foi criado com ASP.NET, OWIN e com a Biblioteca de Autentica??o da Microsoft (MSAL) para .NET.

> [!NOTE]
> Para ler um artigo semelhante sobre um suplemento baseado em Node.js, confira [Criar um Suplemento do Office com Node.js que use logon ?nico](create-sso-office-add-ins-nodejs.md).

## <a name="prerequisites"></a>Pr?-requisitos

* A vers?o mais recente dispon?vel do Visual Studio 2017 Preview.

* Office 2016, vers?o 1708, build 8424.nnnn ou posterior (a vers?o de assinatura do Office 365, ?s vezes chamada de "Clique para Executar"). Voc? talvez precise ser um participante do programa Office Insider para obter essa vers?o. Para obter mais informa??es, confira a p?gina [Seja um Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1).

## <a name="set-up-the-starter-project"></a>Configure o projeto inicial

1. Clone ou baixe o reposit?rio em [SSO com Suplemento ASPNET do Office](https://github.com/officedev/office-add-in-aspnet-sso).

1. Abra a pasta **Before** (antes) e abra o arquivo .sln no Visual Studio. Esse ? um projeto inicial. A interface do usu?rio e outros aspectos do suplemento que n?o est?o diretamente ligados ao SSO ou ? autoriza??o j? est?o prontos.

    > [!NOTE]
    > H? tamb?m uma vers?o conclu?da do exemplo no mesmo reposit?rio. Essa vers?o apresenta como seria o suplemento quando conclu?dos os procedimentos apresentados neste artigo, com exce??o de que o projeto conclu?do traz coment?rios de c?digos que seriam redundantes neste artigo. Para usar a vers?o conclu?da, apenas abra o arquivo `sln` e siga as instru??es apresentadas neste artigo, mas pule as se??es **Codificar o lado do cliente** e **Codificar o lado do servidor**.

1. Depois que o projeto for aberto, compile-o no Visual Studio, que instalar? os pacotes listados no arquivo packages.config. Esse procedimento poder? levar entre alguns segundos e alguns minutos dependendo de quantos pacotes estiverem no cache de pacote local do computador.

    > [!NOTE]
    > Voc? receber? um erro sobre o namespace Identity. Este ? um efeito colateral de um problema de configura??o que voc? corrigir? no pr?ximo passo. O importante ? que os pacotes estejam instalados.

1. Atualmente, a vers?o da biblioteca MSAL (Microsoft.Identity.Client) necess?ria para SSO (vers?o `1.1.1-alpha0393`) n?o faz parte do cat?logo padr?o de nuget, portanto, n?o est? listada no package.config e deve ser instalada separadamente. 

   > 1. No menu **Ferramentas**, navegue at? **Nuget Package Manager** > **Console do Gerenciador de Pacotes**. 

   > 2. No console, execute o seguinte comando: Pode levar um minuto ou mais para concluir, mesmo com uma conex?o r?pida ? Internet. Quando terminar, voc? deve ver **Microsoft.Identity.Client 1.1.1-alpha0393' instalado com sucesso...** perto do final da sa?da no console.

   >    `Install-Package Microsoft.Identity.Client -Version 1.1.1-alpha0393 -Source https://www.myget.org/F/aad-clients-nightly/api/v3/index.json`

   > 3. No **Explorador de solu??es**, clique com o bot?o direito em **Refer?ncias**. Confirme que o **Microsoft.Identity.Client** est? listado. Se n?o estiver, ou se houver um ?cone de aviso na entrada dele, exclua a entrada e use o assistente do Visual Studio Add Reference para adicionar uma refer?ncia ? montagem em **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.1-alpha0393\lib\net45\Microsoft.Identity.Client.dll**

1. Crie o projeto pela segunda vez.

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Registre o suplemento com o ponto de extremidade do Azure AD v2.0

As instru??es a seguir foram escritas de modo gen?rico para que possam ser usadas em diversos lugares. Para este artigo, fa?a o seguinte:
- Substitua o espa?o reservado **$ADD-IN-NAME$** por `Office-Add-in-ASPNET-SSO`.
- Substitua o espa?o reservado **$FQDN-WITHOUT-PROTOCOL$** por `localhost:44355`.
- Quando voc? especifica permiss?es no di?logo **Selecionar Permiss?es**, marque as caixas para as permiss?es a seguir. Somente a primeira ? realmente exigida pelo suplemento propriamente dito, mas a biblioteca MSAL usada pelo c?digo de servidor exige `offline_access` e `openid`. A permiss?o `profile` ? necess?ria para que o host do Office obtenha um token no aplicativo Web do seu suplemento.
    * Files.Read.All
    * offline_access
    * openid
    * perfil


[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="grant-administrator-consent-to-the-add-in"></a>Conceder autoriza??o do administrador ao suplemento

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a>Configurar o suplemento

1. Na cadeia de caracteres a seguir, substitua o espa?o reservado "{tenant_ID}" pelo ID de locat?rio do Office 365. Use um dos m?todos em [Encontre seu ID de locat?rio do Office 365](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) para obt?-lo.

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

2. No Visual Studio, abra o Web.config. Existem algumas chaves na se??o **appSettings** ?s quais voc? precisa atribuir valores.

3. Use a cadeia de caracteres constru?da na etapa 1 como o valor para a chave denominada "ida:Issuer". N?o deixe espa?os em branco no valor.

4. Atribua os seguintes valores para as chaves correspondentes:

    |Chave|Valor|
    |:-----|:-----|
    |ida:ClientID|A ID do aplicativo obtida ao registrar o suplemento.|
    |ida:Audience|A ID do aplicativo obtida ao registrar o suplemento.|
    |ida:Password|A senha obtida ao registrar o suplemento.|

   Veja a seguir um exemplo de como as quatro chaves que voc? alterou devem se parecer. *Observe que as chaves ClientID e Audience s?o iguais*. Voc? tamb?m pode usar uma ?nica chave para ambos os fins, mas sua marca??o web.config ? mais reutiliz?vel se for mantida separada, pois ela n?o ? sempre a mesma. Al?m disso, ter chaves separadas refor?a a ideia de que seu suplemento ? tanto um recurso de OAuth, em rela??o a um host do Office, e um cliente OAuth, em rela??o ao Microsoft Graph.

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />
    
    ```

   > [!NOTE]
   > N?o altere as demais configura??es na se??o **appSettings**.

1. Salve e feche o arquivo.

1. Na raiz do projeto, abra o arquivo do manifesto do suplemento "Office-Add-in-ASPNET-SSO.xml".

1. Role at? o final do arquivo.

1. Logo acima da marca de fim `</VersionOverrides>`, voc? encontrar? a marca??o a seguir:

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:44355/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. Substitua o espa?o reservado "{application_GUID here}" *nos dois lugares* na marca??o pela ID do Aplicativo que voc? copiou ao registrar seu suplemento. Os "{}" n?o fazem parte do ID, portanto n?o os inclua. Essa ? a mesma ID usada para a ClientID e a Audience no web.config.

    > [!NOTE]
    > * O valor de **Resource** ? o **URI da ID do Aplicativo** que voc? definiu quando adicionou a plataforma API Web no registro do suplemento.
    > * A se??o **Scopes** s? ser? usada para gerar uma caixa de di?logo de consentimento se o suplemento for vendido no AppSource.

1. Abra a guia **Avisos** da **Lista de Erros** no Visual Studio. Se houver um aviso informando que `<WebApplicationInfo>` n?o ? um filho v?lido de `<VersionOverrides>`, sua vers?o do Visual Studio 2017 Preview n?o reconhecer? a marca??o SSO. Para solucionar esse problema, fa?a o seguinte para um suplemento do Word, Excel ou PowerPoint. Se voc? estiver trabalhando com um suplemento do Outlook, confira a solu??o abaixo.

   - **Solu??o alternativa para Word, Excel e PowerPoint**

        1. Comente a se??o `<WebApplicationInfo>` do manifesto logo acima do final de `</VersionOverrides>`.

        2. Pressione F5 para iniciar uma sess?o de depura??o. Isso criar? uma c?pia do manifesto na seguinte pasta (que pode ser acessada mais facilmente pelo **Gerenciador de Arquivos** do que pelo Visual Studio): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`

        3. Na c?pia do manifesto, remova a sintaxe do coment?rio em torno da se??o `<WebApplicationInfo>`.

        4. Salve a c?pia do manifesto.

        5. Agora, ? preciso evitar que o Visual Studio substitua a c?pia do manifesto quando voc? terminar na pr?xima vez que pressionar F5. Clique com bot?o direito do mouse no n? da solu??o na parte superior do **Gerenciador de Solu??es** (n?o nos n?s do projeto).

        6. Escolha **Propriedades** no menu de contexto e uma caixa de di?logo **P?ginas de Propriedades da Solu??o** ser? aberta.

        7. Expanda **Propriedades da Configura??o** e escolha **Configura??o**.

        8. Desmarque **Criar** e **Implantar** na linha do projeto **Office-Add-in-ASPNET-SSO** (*n?o* o projeto **Office-Add-in-ASPNET-SSO-WebAPI**).

        9. Pressione **OK** para fechar a caixa de di?logo.

   - **Solu??o alternativa para Outlook**

        1. Em sua m?quina de desenvolvimento, localize o `MailAppVersionOverridesV1_1.xsd` existente. Ele deve estar localizado no diret?rio de instala??o do Visual Studio em `./Xml/Schemas/{lcid}`. Por exemplo, em uma instala??o t?pica do VS 2017 de 32 bits em um sistema em ingl?s (EUA), o caminho completo seria `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.

        2. Renomeie o arquivo existente para `MailAppVersionOverridesV1_1.old`.

        3. Copie essa vers?o modificada do arquivo para a pasta: [Esquema MailAppVersionOverrides modificado](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)

1. Salve e feche o arquivo de manifesto principal no Visual Studio.

## <a name="code-the-client-side"></a>Codificar o lado do cliente

1. Abra o arquivo Home.js da pasta **Scripts**. Ele j? apresenta alguns c?digos:
    * Uma atribui??o ao m?todo `Office.initialize` que, por sua vez, atribui um manipulador ao evento clicar do bot?o `getGraphAccessTokenButton`.
    * Um m?todo `showResult` que exibir? os dados retornados do Microsoft Graph (ou uma mensagem de erro) na parte inferior do painel de tarefas.
    * Um m?todo `logErrors` que registrar? erros de console que n?o s?o destinados ao usu?rio final.

1. Abaixo da atribui??o a `Office.initialize`, adicione o c?digo a seguir. Observe o seguinte sobre este c?digo:

    * O processamento de erros no suplemento ?s vezes tentar? novamente obter um token de acesso automaticamente, usando um conjunto diferente de op??es. A vari?vel de contador `timesGetOneDriveFilesHasRun` e a vari?veis de sinalizador `triedWithoutForceConsent` s?o usadas para garantir que o usu?rio n?o seja trocado repetidas vezes em tentativas falhas de obter um token. 
    * Voc? criar? um m?todo `getDataWithToken` na pr?xima etapa, mas observe que ele define uma op??o chamada `forceConsent` como `false`. Trataremos mais disso na etapa seguinte.

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;

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
    
        // TODO10: Parse the JSON response.

        // TODO11: Handle the case where AAD asks for an additional form of authentication.

        // TODO12: Handle the case where consent has not been granted, or has been revoked.

        // TODO13: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow.

        // TODO14: Handle the case where the token that the add-in's client-side sends to it's 
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO15: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        // TODO16: Log all other server errors.
    }
    ```

1. Substitua `TODO10` pelo c?digo a seguir. Observe que, para a maioria dos erros `4xx` que o servi?o Web do suplemento passar? para o suplemento do lado do cliente, haver? uma propriedade **ExceptionMessage** em resposta com o n?mero de erro AADSTS (Azure Active Directory Secure Token Service) al?m de outros dados. No entanto, quando AAD envia uma mensagem para o servi?o Web do suplemento solicitando um fator de autentica??o adicional, a mensagem cont?m uma propriedade **Claims** especial que especifica (com um n?mero de c?digo) qual fator adicional ? necess?rio. As APIs ASP.NET que criam e enviam respostas HTTP para clientes n?o conhecem a propriedade **Claims**, portanto, elas n?o a incluem no objeto Response. O c?digo de servidor que ser? criado em uma etapa posterior lidar? com isso adicionando manualmente o valor **Claims** no objeto Response. Esse valor ser? salvo na propriedade **Message**, portanto, o c?digo tamb?m precisar? analisar essa propriedade.

    ```javascript
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    var message = JSON.parse(result.responseText).Message;
    ```

1. Substitua `TODO11` pelo c?digo a seguir. Observa??o sobre este c?digo:

    * O erro 50076 ocorre quando o Microsoft Graph requer uma forma adicional de autentica??o.
    * O host do Office deve obter um novo token com o valor **Claims** como a op??o `authChallenge`. Isso instrui o AAD a solicitar ao usu?rio todas as formas de autentica??o requeridas. 

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
        }
    }    
    ```

1. Substitua `TODO12` pelo c?digo a seguir. Observa??o sobre este c?digo:

    * O erro 65001 significa que o consentimento para acessar o Microsoft Graph n?o foi concedido (ou foi revogado) para uma ou mais permiss?es. 
    * O suplemento dever? obter um novo token com a op??o `forceConsent` definida como `true`.

    ```javascript
    if (exceptionMessage.indexOf('AADSTS65001') !== -1) {
        showResult(['Please grant consent to this add-in to access your Microsoft Graph data.']);        
        /*
            THE FORCE CONSENT OPTION IS NOT AVAILABLE IN DURING PREVIEW. WHEN SSO FOR
            OFFICE ADD-INS IS RELEASED, REMOVE THE showResult LINE ABOVE AND UNCOMMENT
            THE FOLLOWING LINE.
        */
       // getDataWithToken({ forceConsent: true });
    }    
    ```

1. Substitua `TODO13` pelo c?digo a seguir. Observa??o sobre este c?digo:

    * O Erro 70011 tem muitos significados. O que importa para este suplemento ? quando ele significa que um escopo inv?lido (permiss?o) foi solicitado, ent?o o c?digo verifica a descri??o completa do erro, n?o apenas o n?mero.
    * O suplemento dever? relatar o erro.

    ```javascript
     else if (exceptionMessage.indexOf("AADSTS70011: The provided value for the input parameter 'scope' is not valid.") !== -1) {
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }    
    ```

1. Substitua `TODO14` pelo c?digo a seguir. Observa??o sobre este c?digo:

    * C?digo de servidor criado em uma etapa posterior enviar? a mensagem `Missing access_as_user` se o escopo `access_as_user` (permiss?o) n?o for o token de acesso que o cliente do suplemento enviar para o ADD para ser usado no fluxo on-behalf-of.
    * O suplemento dever? relatar o erro.

    ```javascript
    else if (exceptionMessage.indexOf('Missing access_as_user.') !== -1) {
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }    
    ```

1. Substitua `TODO15` pelo c?digo a seguir. Observa??o sobre este c?digo:

    * A biblioteca de identidade que voc? usar? no c?digo do lado do servidor (Biblioteca de Autentica??o da Microsoft - MSAL) deve garantir que nenhum token inv?lido ou expirado seja enviado para o Microsoft Graph. Contudo, se isso ocorrer, o erro retornado para servi?o Web do suplemento do Microsoft Graph ter? o c?digo `InvalidAuthenticationToken`. O c?digo do lado do servidor que voc? criar? em uma etapa futura transmitir? essa mensagem ao cliente do suplemento.
    * Nesse caso, o suplemento dever? iniciar o processo de autentica??o completo ao redefinir o contador e as vari?veis de sinalizador e, em seguida, chamando novamente o m?todo de identificador de bot?o.

    ```javascript
    // If the token sent to MS Graph is expired or invalid, start the whole process over.
    else if (result.code === 'InvalidAuthenticationToken') {
        timesGetOneDriveFilesHasRun = 0;
        triedWithoutForceConsent = false;
        getOneDriveFiles();
    }    
    ```

1. Substitua `TODO16` pelo c?digo a seguir.

    ```javascript
    else {
        logError(result);
    }    
    ```

1. Salve e feche o arquivo.

## <a name="code-the-server-side"></a>Codifique o lado do servidor

### <a name="configure-the-owin-middleware"></a>Configurar o middleware OWIN

1. Abra o arquivo Startup.cs na raiz do projeto.

1. Adicione a palavra-chave `partial` para a declara??o da classe Startup, se ainda n?o estiver l?. A linha dever? ser assim:

    `public partial class Startup`

1. Adicione a linha a seguir ao corpo do m?todo `Configuration`. Voc? criar? o m?todo `ConfigureAuth` em uma etapa posterior.

    `ConfigureAuth(app);`

1. Salve e feche o arquivo.

1. Clique com bot?o direito do mouse na pasta **App_Start** e selecione **Adicionar > Classe**.

1. Na caixa de di?logo **Adicionar novo item** nomeie o arquivo **Startup.Auth.cs** e, em seguida, clique em **Adicionar**.

1. Encurte o nome do namespace no novo arquivo para `Office_Add_in_ASPNET_SSO_WebAPI`.

1. Verifique se todas as seguintes instru??es `using` est?o na parte superior do arquivo.

    ```csharp
    using Owin;
    using System.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. Adicione a palavra-chave `partial` ? declara??o da classe `Startup`, se ainda n?o estiver l?. A linha dever? ser assim:

    `public partial class Startup`

1. Adicione o m?todo a seguir ? classe `Startup`. Este m?todo especifica como o middleware OWIN validar? os tokens de acesso que s?o transmitidos a ele do m?todo `getData` no arquivo Home.js do lado do cliente. O processo de autoriza??o ? disparado sempre que um ponto de extremidade da API Web decorado com o atributo `[Authorize]` ? chamado.

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. Substitua TODO3 pelo seguinte c?digo. Observa??o sobre o c?digo:

    * O c?digo instrui o OWIN a garantir que o emissor de token e audi?ncia especificado no token de acesso que vem do host do Office (e ? transmitido pela chamada de `getData` do lado do cliente) deve coincidir com os valores especificados no Web.config.
    * Definir `SaveSigninToken` como `true` faz com que o OWIN salve o token bruto do host do Office. O suplemento precisa dele para obter um token de acesso para o Microsoft Graph com o fluxo "on behalf of".
    * Os escopos n?o s?o validados pelo middleware OWIN. Os escopos do token de acesso, que devem conter `access_as_user`, s?o validados no controlador.

    ```csharp
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. Substitua TODO4 pelo seguinte. Observa??o sobre este c?digo:

    * O m?todo `UseOAuthBearerAuthentication` ? chamado em vez do `UseWindowsAzureActiveDirectoryBearerAuthentication` que ? mais comum, porque este ?ltimo n?o ? compat?vel com o ponto de extremidade V2 do Azure AD.
    * A URL de descoberta transmitida ao m?todo ? onde o middleware OWIN obt?m instru??es para conseguir a chave que precisa para verificar a assinatura no token de acesso recebido do host do Office.

    ```csharp
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
        {
            AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
        });
    ```

1. Salve e feche o arquivo.

### <a name="create-the-apivalues-controller"></a>Criar o controlador /api/values

1. Abra o arquivo **Controllers\ValueController.cs**.

2. Verifique se as seguintes instru??es `using` est?o na parte superior do arquivo.

    ```csharp
    using Microsoft.Identity.Client;
    using System.IdentityModel.Tokens;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    using Office_Add_in_ASPNET_SSO_WebAPI.Models;
    ```

3. Logo acima da linha que declara o `ValuesController`, adicione o atributo `[Authorize]`. Isso garante que seu suplemento executar? o processo de autoriza??o configurado no ?ltimo procedimento sempre que um m?todo controlador for chamado. Apenas os chamadores com um token de acesso v?lido para o seu suplemento podem invocar os m?todos do controlador.

    > [!NOTE]
    > Um servi?o da ASP.NET MVC Web API de produ??o deve ter l?gica personalizada para o fluxo on-behalf-of em uma ou mais classes [FilterAttribute](https://msdn.microsoft.com/en-us/library/system.web.http.filters(v=vs.108).aspx) personalizadas. Este exemplo educacional coloca a l?gica no controlador de principal para que o fluxo de autoriza??o e dados busca l?gica inteiro possa ser acompanhado facilmente. Isso tamb?m faz com que o exemplo fique consistente com os exemplos de padr?o de autoriza??o nos [Exemplos do Azure](https://github.com/Azure-Samples/).    

4. Adicione o m?todo a seguir ao `ValuesController`. Observe que ? o valor de retorno ? `Task<HttpResponseMessage>` em vez de `Task<IEnumerable<string>>`, como seria mais comum para um m?todo `GET api/values`. Este ? um efeito colateral do fato de que nossa l?gica de autoriza??o personalizada estar? no controlador: algumas condi??es de erro nessa l?gica exigem que um objeto de resposta HTTP seja enviado para o cliente do suplemento. 

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO1: Validate the scopes of the access token.
    }
    ```

5. Substitua `TODO1` pelo seguinte c?digo para validar que os escopos especificados no token incluam `access_as_user`.

    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (addinScopes.Contains("access_as_user"))
    {
        // TODO2: Assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.
        // TODO3: Get the access token for Microsoft Graph.
        // TODO4: Get the names of files and folders in OneDrive by using the Microsoft Graph API.
        // TODO5: Remove excess information from the data and send the data to the client.
    }
    return SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    ```

    > [!NOTE]
    > Voc? deve usar apenas o escopo `access_as_user` para autorizar a API que lida com o fluxo Em Nome De para os suplementos do Office. Outras APIs em seu servi?o devem ter seus pr?prios requisitos de escopo. Isso limita o que pode ser acessado com os tokens que o Office adquire.

6. Substitua `TODO2` pelo c?digo a seguir. Observa??o sobre este c?digo:
    * Ele transforma o token de acesso bruto recebido do host do Office em um objeto de `UserAssertion` que ser? transmitido para outro m?todo.
    * Seu suplemento n?o est? mais desempenhando o papel de um recurso (ou p?blico) para o qual o host do Office e o usu?rio precisam de acesso. Agora, ele mesmo ? um cliente que precisa de acesso ao Microsoft Graph. `ConfidentialClientApplication` ? o objeto "client context" da MSAL.
    * O terceiro par?metro para o construtor `ConfidentialClientApplication` ? uma URL de redirecionamento que n?o ? realmente usada no fluxo "on behalf of", mas usar a URL correta ? uma boa pr?tica. O quarto e o quinto par?metros podem ser usados para definir um armazenamento persistente que permitiria a reutiliza??o de tokens n?o expirados em diferentes sess?es com o suplemento. Este exemplo n?o implementa nenhum armazenamento persistente.
    * A MSAL exige os escopos `openid` e `offline_access` para funcionar, mas ela lan?a um erro se o c?digo solicit?-los de forma redundante. Ela tamb?m lan?ar? um erro se o seu c?digo solicitar o `profile`, que realmente ? usado apenas quando o aplicativo host do Office recebe o token para o aplicativo Web do seu suplemento. Ent?o, apenas `Files.Read.All` ? explicitamente solicitado.

    ```csharp
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

7. Substitua `TODO3` pelo c?digo a seguir. Observa??o sobre este c?digo:

    * O m?todo `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` procurar? primeiro no cache da MSAL, que est? na mem?ria, para fazer a correspond?ncia com o token de acesso. Somente se n?o houver um, ele iniciar? o fluxo "on behalf of" com o ponto de extremidade V2 do Azure AD.
    * Se a autentica??o multi-fator for requerida pelo recurso MS Graph e o usu?rio ainda n?o a tiver fornecido, o AAD lan?ar? uma exce??o contendo uma propriedade de Declara??es.
    * O valor da propriedade de Declara??es deve ser passado para o cliente, que o passar? para o host do Office, que, em seguida, o incluir? em um pedido para um novo token. O AAD solicitar? ao usu?rio todas as formas de autentica??o necess?rias.
    * Quaisquer exce??es que n?o forem do tipo `MsalServiceException` s?o intencionalmente n?o detectadas, e, portanto, se propagar?o para o cliente como mensagens `500 Server Error`.

    ```csharp
    AuthenticationResult result = null;
    try
    {
        result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
    }
    catch (MsalServiceException e)
    {        
        // TODO3a: Handle request for multi-factor authentication.
        // TODO3b: Handle lack of consent.
        // TODO3c: Handle invalid scope (permission).
        // TODO3d: Handle all other MsalServiceExceptions.
    }
    ```

8. Substitua `TODO3a` pelo c?digo a seguir. Observa??o sobre este c?digo:

    * Se a autentica??o multifator for exigida pelo recurso MS Graph e o usu?rio ainda n?o a tiver fornecido, o AAD retornar? "400 Bad Request" com o erro AADSTS50076 e uma propriedade **Declara??es**. O MSAL lan?ar? uma **MsalUiRequiredException** (que herda de **MsalServiceException**) com essas informa??es. 
    * O valor da propriedade **Declara??es** deve ser passado para o cliente, que deve pass?-lo para o host do Office, que, por sua vez, o incluir? em um pedido para um novo token. O AAD solicitar? ao usu?rio todas as formas de autentica??o necess?rias.
    * As APIs que criam respostas HTTP a partir de exce??es n?o conhecem a propriedade **Claims**, portanto, elas n?o a incluem no objeto de resposta. ? necess?rio criar manualmente uma mensagem que inclua esse recurso. Uma propriedade **Message** personalizada, no entanto, impede a cria??o de uma propriedade **ExceptionMessage**, portanto, a ?nica maneira de obter a ID de erro `AADSTS50076` para o cliente ? adicion?-la ? **Message** personalizada. O JavaScript no cliente precisar? descobrir se uma resposta tem uma **Message** ou **ExceptionMessage** para saber qual ler.
    * A mensagem personalizada ? formatada como JSON para que o JavaScript do cliente possa analis?-la com m?todos de objeto `JSON` conhecidos.
    * Voc? criar? o m?todo `SendErrorToClient` em uma etapa posterior. ? segundo par?metro ? um objeto **Exception**. Nesse caso, o c?digo passa `null` porque incluir o objeto **Exception** bloqueia a inclus?o da propriedade **Message** na resposta HTTP que ? gerada.

    ```csharp
    if (e.Message.StartsWith("AADSTS50076")) {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

9. Substitua `TODO3b` e `TODO3c` pelo c?digo a seguir. Observa??o sobre este c?digo:

    * Se a chamada para o AAD contiver pelo menos um escopo (permiss?o) que n?o tenha sido consentido pelo usu?rio ou por um administrador de locat?rios (ou se o consentimento foi revogado), o AAD retornar? "400 Solicita??o Incorreta" com o erro `AADSTS65001`. O MSAL exibe um **MsalUiRequiredException** com essas informa??es. O cliente deve chamar `getAccessTokenAsync` novamente com a op??o `{ forceConsent: true }`.
    *  Se a chamada para o AAD contiver pelo menos um escopo que AAD n?o reconhece, o AAD retornar? "400 Solicita??o Incorreta" com o erro `AADSTS70011`. O MSAL exibe um **MsalUiRequiredException** com essas informa??es. O cliente deve informar o usu?rio.
    *  A descri??o completa ? inclu?da porque 70011 ? retornado em outras condi??es e ele deve ser processado nesse suplemento somente quando significar que h? um escopo inv?lido. 
    *  O objeto **MsalUiRequiredException** ? passado para `SendErrorToClient`. Isso garante que uma propriedade **ExceptionMessage** contendo as informa??es de erro seja inclu?da na resposta HTTP.
    *  N?o h? uma mensagem personalizada, portanto, `null` ? passado para o terceiro par?metro.

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001"))
    || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

10. Substitua `TODO3d` pelo c?digo a seguir. Observe que o c?digo exibe a exce??o em vez de transmiti-la em uma resposta HTTP personalizada com **HttpStatusCode.Forbidden** (401). O efeito disso ? que o ASP.NET enviar? sua pr?pria resposta HTTP com o status "500 Erro de Servidor".

    ```csharp
    else
    {
        throw e;
    }  
    ```

11. Substitua `TODO4` pelo seguinte. Observa??o sobre este c?digo:

    * As classes `GraphApiHelper` e `ODataHelper` s?o definidas nos arquivos da pasta **Helpers**. A classe `OneDriveItem` ? definida em um arquivo da pasta **Models**. A discuss?o detalhada dessas classes n?o ? relevante para a autoriza??o ou o SSO, portanto, est? fora do escopo deste artigo.
    * O desempenho ? aprimorado ao se solicitar ao Microsoft Graph apenas os dados que s?o realmente necess?rios. Desse modo, o c?digo usa um par?metro de consulta ` $select` para especificar que desejamos somente a propriedade de nome, e usa um par?metro `$top` para especificar que desejamos somente os tr?s primeiros nomes de pasta ou de arquivo.
    * Se o token enviado para o Microsoft Graph for inv?lido, o Microsoft Graph enviar? um erro "401 N?o Autorizado" com o c?digo "InvalidAuthenticationToken". Em seguida, o ASP.NET exibe um **RuntimeBinderException**. Isso tamb?m ocorre quando o token expira, embora o MSAL deva impedir que isso aconte?a. 

    ```csharp
    var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
    IEnumerable<OneDriveItem> filesResult;
    try
    {
        filesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
    }
    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
    {
        return SendErrorToClient(HttpStatusCode.Unauthorized, e, null);                    
    }
    ```

12. Substitua `TODO5` pelo seguinte. Observa??o sobre este c?digo: 

    * Embora o c?digo acima solicite somente a propriedade *name* dos itens do OneDrive, o Microsoft Graph sempre inclui a propriedade *eTag* para os itens do OneDrive. Para reduzir a carga enviada para o cliente, o c?digo a seguir reconstr?i os resultados apenas com os nomes dos itens.
    * A lista de tr?s pastas e arquivos do OneDrive ? enviada para o cliente como uma resposta HTTP "200 OK".

    ```csharp
    List<string> itemNames = new List<string>();
    foreach (OneDriveItem item in filesResult)
    {
        itemNames.Add(item.Name);
    }

    var requestMessage = new HttpRequestMessage();
    requestMessage.SetConfiguration(new HttpConfiguration());
    var response = requestMessage.CreateResponse<List<string>>(HttpStatusCode.OK, itemNames); 
    return response;
    ```

13. Abaixo do m?todo Get, adicione o m?todo a seguir. Sobre este c?digo, observe:  

    * O m?todo transmite ao cliente informa??es sobre uma exce??o do servidor. 
    * Se a exce??o original for passada para o m?todo, o construtor HttpError incluir? informa??es do objeto de exce??o em uma propriedade **ExceptionMessage**.  
    * Se `null` for passado para a exce??o, o construtor HttpError incluir? o par?metro de mensagem em uma propriedade **Message** e n?o haver? uma propriedade **ExceptionMessage**.

    ```csharp
    private HttpResponseMessage SendErrorToClient(HttpStatusCode statusCode, Exception e, string message)
    {
        HttpError error;
        if (e != null)
        {
            error = new HttpError(e, true);
        }
        else
        {
            error = new HttpError(message);
        }
        var requestMessage = new HttpRequestMessage();
        var errorMessage = requestMessage.CreateErrorResponse(statusCode, error);
        return errorMessage;
    }        
    ```

## <a name="run-the-add-in"></a>Execute o suplemento

1. Certifique-se de ter alguns arquivos no seu OneDrive para que voc? possa verificar os resultados.

1. No Visual Studio, pressione F5. O PowerPoint ser? aberto e haver? um grupo **SSO ASP.NET** na faixa de op??es **P?gina Inicial**.

1. Pressione o bot?o **Mostrar Suplemento** nesse grupo para ver a interface do usu?rio do suplemento no painel de tarefas.

1. Pressione o bot?o **Obter meus arquivos do OneDrive**. Se voc? n?o estiver conectado ao Office, voc? ser? solicitado a entrar.
    
    > [!NOTE]
    > Se voc? entrou no Office com uma ID diferente e se alguns aplicativos do Office que estavam abertos no momento continuam abertos, o Office pode n?o alterar de forma confi?vel sua ID, mesmo que pare?a ter feito isso no PowerPoint. Se isso acontecer, a chamada para o Microsoft Graph pode falhar ou os dados da ID anterior podem ser retornados. Para evitar isso, certifique-se de *fechar todos os outros aplicativos do Office* antes de pressionar **Obter meus arquivos do OneDrive**.

1. Depois de entrar, ser? exibida uma lista de seus arquivos e suas pastas no OneDrive, abaixo do bot?o. Esse procedimento pode levar mais de 15 segundos, principalmente na primeira vez.
