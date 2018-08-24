---
title: Criar um Suplemento do Office com ASP.NET que usa logon único
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 70662a01d86d3fa111b39deb4c16702a4f8530f5
ms.sourcegitcommit: e1c92ba882e6eb03a165867c6021a6aa742aa310
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/20/2018
ms.locfileid: "22925637"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a>Criar um Suplemento do Office com ASP.NET que use logon único (visualização)

Quando os usuários estão conectados ao Office, o seu suplemento pode usar as mesmas credenciais para permitir que os usuários acessem vários aplicativos sem exigir que eles entrem uma segunda vez. Para obter uma visão geral, consulte [Habilitar o SSO em um Suplemento do Office](sso-in-office-add-ins.md).

Este artigo apresenta o processo passo a passo de habilitação do logon único (SSO) em um suplemento que foi criado com ASP.NET, OWIN e com a Biblioteca de Autenticação da Microsoft (MSAL) para .NET.

> [!NOTE]
> Para ler um artigo semelhante sobre um suplemento baseado em Node.js, confira [Criar um Suplemento do Office com Node.js que use logon único](create-sso-office-add-ins-nodejs.md).

## <a name="prerequisites"></a>Pré-requisitos

* A versão mais recente disponível do Visual Studio 2017 Preview.

* Office 2016, versão 1708, build 8424.nnnn ou posterior (a versão de assinatura do Office 365, às vezes chamada de "Clique para Executar"). Você talvez precise ser um participante do programa Office Insider para obter essa versão. Para obter mais informações, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1).

## <a name="set-up-the-starter-project"></a>Configure o projeto inicial

1. Clone ou baixe o repositório em [SSO com Suplemento ASPNET do Office](https://github.com/officedev/office-add-in-aspnet-sso).

1. Abra a pasta **Before** (antes) e abra o arquivo .sln no Visual Studio. Esse é um projeto inicial. A interface do usuário e outros aspectos do suplemento que não estão diretamente ligados ao SSO ou à autorização já estão prontos.

    > [!NOTE]
    > Há também uma versão concluída do exemplo no mesmo repositório. Essa versão apresenta como seria o suplemento quando concluídos os procedimentos apresentados neste artigo, com exceção de que o projeto concluído traz comentários de códigos que seriam redundantes neste artigo. Para usar a versão concluída, apenas abra o arquivo `sln` e siga as instruções apresentadas neste artigo, mas pule as seções **Codificar o lado do cliente** e **Codificar o lado do servidor**.

1. Depois que o projeto for aberto, compile-o no Visual Studio, que instalará os pacotes listados no arquivo packages.config. Esse procedimento poderá levar entre alguns segundos e alguns minutos dependendo de quantos pacotes estiverem no cache de pacote local do computador.

    > [!NOTE]
    > Você receberá um erro sobre o namespace Identity. Este é um efeito colateral de um problema de configuração que você corrigirá no próximo passo. O importante é que os pacotes estejam instalados.

1. Atualmente, a versão da biblioteca MSAL (Microsoft.Identity.Client) necessária para SSO (versão `1.1.4-preview0002`) não faz parte do catálogo padrão de nuget, portanto, não está listada no package.config e deve ser instalada separadamente. 

   > 1. No menu **Ferramentas**, navegue até **Nuget Package Manager** > **Console do Gerenciador de Pacotes**. 

   > 2. No console, execute o seguinte comando: Pode levar um minuto ou mais para concluir, mesmo com uma conexão de Internet rápida. Quando terminar, você deve ver **'Microsoft.Identity.Client 1.1.4-preview0002' instalado com sucesso...** perto do final da saída no console.

   >    `Install-Package Microsoft.Identity.Client -Version 1.1.4-preview0002`

   > 3. No **Gerenciador de Soluções**, clique com o botão direito do mouse em **Referências**. Verifique se **Microsoft.Identity.Client** está na lista. Se não estiver ou se houver um ícone de aviso na entrada, exclua a entrada e use o assistente de Adicionar Referência do Visual Studio para adicionar uma referência à montagem em **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.4-preview0002\lib\net45\Microsoft.Identity.Client.dll**

1. Crie o projeto pela segunda vez.

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Registre o suplemento com o ponto de extremidade v2.0 do Azure AD

As instruções a seguir foram escritas de modo genérico para que possam ser usadas em diversos lugares. Para este artigo, faça o seguinte:
- Substitua o espaço reservado **$ADD-IN-NAME$** por `Office-Add-in-ASPNET-SSO`.
- Substitua o espaço reservado **$FQDN-WITHOUT-PROTOCOL$** por `localhost:44355`.
- Quando você especifica permissões no diálogo **Selecionar Permissões**, marque as caixas para as permissões a seguir. Somente a primeira é realmente exigida pelo suplemento propriamente dito, mas a biblioteca MSAL usada pelo código do servidor exige `offline_access` e `openid`. A permissão `profile` é necessária para que o host do Office obtenha um token no aplicativo Web do seu suplemento.
    * Files.Read.All
    * offline_access
    * openid
    * perfil


[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="grant-administrator-consent-to-the-add-in"></a>Conceder autorização do administrador ao suplemento

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a>Configurar o suplemento

1. Na cadeia de caracteres a seguir, substitua o espaço reservado "{tenant_ID}" pelo ID de locatário do Office 365. Use um dos métodos em [Encontre seu ID de locatário do Office 365](https://docs.microsoft.com/onedrive/find-your-office-365-tenant-id) para obtê-lo.

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

2. No Visual Studio, abra o web.config. Existem algumas chaves na seção **appSettings** às quais você precisa atribuir valores.

3. Use a cadeia de caracteres construída na etapa 1 como o valor para a chave denominada "ida:Issuer". Não deixe espaços em branco no valor.

4. Atribua os seguintes valores para as chaves correspondentes:

    |Chave|Valor|
    |:-----|:-----|
    |ida:ClientID|A ID do aplicativo obtida ao registrar o suplemento.|
    |ida:Audience|A ID do aplicativo obtida ao registrar o suplemento.|
    |ida:Password|A senha obtida ao registrar o suplemento.|

   Veja a seguir um exemplo de como as quatro chaves que você alterou devem se parecer. *Observe que as chaves ClientID e Audience são iguais*. Você também pode usar uma única chave para ambos os fins, mas sua marcação web.config é mais reutilizável se for mantida separada, pois ela não é sempre a mesma. Além disso, ter chaves separadas reforça a ideia de que seu suplemento é tanto um recurso de OAuth, em relação a um host do Office, e um cliente OAuth, em relação ao Microsoft Graph.

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />
    
    ```

   > [!NOTE]
   > Não altere as demais configurações na seção **appSettings**.

1. Salve e feche o arquivo.

1. Na raiz do projeto, abra o arquivo do manifesto do suplemento "Office-Add-in-ASPNET-SSO.xml".

1. Role até o final do arquivo.

1. Logo acima da marca de fim `</VersionOverrides>`, você encontrará a marcação a seguir:

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

1. Substitua o espaço reservado "{application_GUID here}" *nos dois lugares* na marcação pela ID do Aplicativo que você copiou ao registrar seu suplemento. O símbolo "{}" não faz parte da ID, portanto não o inclua. Essa é a mesma ID usada para ClientID e Audience no web.config.

    > [!NOTE]
    > * O valor de **Resource** é o **URI da ID do Aplicativo** que você definiu quando adicionou a plataforma API Web no registro do suplemento.
    > * A seção **Scopes** só será usada para gerar uma caixa de diálogo de consentimento se o suplemento for vendido no AppSource.

1. Abra a guia **Avisos** da **Lista de Erros** no Visual Studio. Se houver um aviso informando que `<WebApplicationInfo>` não é um filho válido de `<VersionOverrides>`, sua versão prévia do Visual Studio 2017 não reconhece a marcação SSO. Para solucionar esse problema, faça o seguinte para um suplemento do Word, Excel ou PowerPoint. Se você estiver trabalhando com um suplemento do Outlook, confira a solução abaixo.

   - **Solução alternativa para Word, Excel e PowerPoint**

        1. Comente a seção `<WebApplicationInfo>` do manifesto logo acima do final de `</VersionOverrides>`.

        2. Pressione F5 para iniciar uma sessão de depuração. Isso criará uma cópia do manifesto na seguinte pasta (que pode ser acessada mais facilmente pelo **Gerenciador de Arquivos** do que pelo Visual Studio): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`

        3. Na cópia do manifesto, remova a sintaxe do comentário em torno da seção `<WebApplicationInfo>`.

        4. Salve a cópia do manifesto.

        5. Agora, é preciso evitar que o Visual Studio substitua a cópia do manifesto quando você terminar na próxima vez que pressionar F5. Clique com botão direito do mouse no nó da solução na parte superior do **Gerenciador de Soluções** (não nos nós do projeto).

        6. Escolha **Propriedades** no menu de contexto e uma caixa de diálogo **Páginas de Propriedades da Solução** será aberta.

        7. Expanda **Propriedades da Configuração** e escolha **Configuração**.

        8. Desmarque **Criar** e **Implantar** na linha do projeto **Office-Add-in-ASPNET-SSO** (*não* o projeto **Office-Add-in-ASPNET-SSO-WebAPI**).

        9. Pressione **OK** para fechar a caixa de diálogo.

   - **Solução alternativa para Outlook**

        1. Em sua máquina de desenvolvimento, localize o `MailAppVersionOverridesV1_1.xsd` existente. Ele deve estar localizado no diretório de instalação do Visual Studio em `./Xml/Schemas/{lcid}`. Por exemplo, em uma instalação típica do VS 2017 de 32 bits em um sistema em inglês (EUA), o caminho completo seria `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.

        2. Renomeie o arquivo existente para `MailAppVersionOverridesV1_1.old`.

        3. Copie essa versão modificada do arquivo para a pasta: [Esquema MailAppVersionOverrides modificado](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)

1. Salve e feche o arquivo de manifesto principal no Visual Studio.

## <a name="code-the-client-side"></a>Codificar o lado do cliente

1. Abra o arquivo Home.js da pasta **Scripts**. Ele já apresenta alguns códigos:
    * Uma atribuição ao método `Office.initialize` que, por sua vez, atribui um manipulador ao evento clicar do botão `getGraphAccessTokenButton`.
    * Um método `showResult` que exibirá os dados retornados do Microsoft Graph (ou uma mensagem de erro) na parte inferior do painel de tarefas.
    * Um método `logErrors` que registrará erros de console que não são destinados ao usuário final.

1. Abaixo da atribuição a `Office.initialize`, adicione o código a seguir. Observe o seguinte sobre este código:

    * O processamento de erros no suplemento às vezes tentará novamente obter um token de acesso automaticamente, usando um conjunto diferente de opções. A variável de contador `timesGetOneDriveFilesHasRun` e a variáveis de sinalizador `triedWithoutForceConsent` são usadas para garantir que o usuário não seja trocado repetidas vezes em tentativas falhas de obter um token. 
    * Você criará um método `getDataWithToken` na próxima etapa, mas observe que ele define uma opção chamada `forceConsent` como `false`. Trataremos mais disso na etapa seguinte.

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }   
    ```

1. Abaixo do método `getOneDriveFiles`, adicione o código a seguir. Observe o seguinte sobre este código:

    * O `getAccessTokenAsync` é a nova API no Office.js que permite que um suplemento solicite ao aplicativo host do Office (Excel, PowerPoint, Word, etc.) um token de acesso para o suplemento (para o usuário conectado ao Office). O aplicativo host do Office, por sua vez, solicita o token ao ponto de extremidade 2.0 do Azure AD. Uma vez que você previamente autorizou o host do Office para o seu suplemento ao registrá-lo, o Azure AD enviará o token.
    * Se nenhum usuário estiver conectado ao Office, o host do Office solicitará que o usuário se conecte.
    * O parâmetro de opções configura o `forceConsent` como `false`. Dessa forma, não será solicitado que o usuário consinta o acesso ao host do Office ao seu suplemento sempre que ele o usar. Na primeira vez que o usuário tiver o suplemento, a chamada de `getAccessTokenAsync` falhará, mas lógica de processamento de erros que você adicionará em uma etapa posterior será automaticamente chamada com a opção `forceConsent` definida como `true` e o usuário será solicitado a consentir, mas somente essa primeira vez.
    * Você criará o método `handleClientSideErrors` em uma etapa posterior.

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

1. Substitua TODO1 pelas linhas a seguir. Você criará o método `getData` e a rota "/api/values" do lado do servidor nas etapas posteriores. Uma URL relativa é usada para o ponto de extremidade porque ela deve ser hospedada no mesmo domínio que seu suplemento.

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. Abaixo do método `getOneDriveFiles`, adicione o seguinte. Observe isto sobre este código:

    * Este método utilitário chama um ponto de extremidade da API Web especificado e transmite a ela o mesmo token de acesso que aplicativo host do Office usou para obter acesso ao seu suplemento. No lado do servidor, esse token de acesso será usado no fluxo "on behalf of" (em nome de) para obter um token de acesso para o Microsoft Graph.
    * Você criará o método `handleServerSideErrors` em uma etapa posterior.

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

### <a name="create-the-error-handling-methods"></a>Crie os métodos de processamento de erros

1. Abaixo do método `getData`, adicione o método a seguir. Esse método processará os erros no cliente do suplemento quando o host do Office não puder obter um token de acesso para o serviço Web do suplemento. Esses erros são relatados com um código de erro, portanto, o método usa uma instrução `switch` para distingui-los.

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

1. Substitua `TODO2` pelo código a seguir. O erro 13001 ocorre quando o usuário não está conectado ou quando ele cancela, sem responder, uma solicitação para fornecer um segundo fator de autenticação. Em ambos os casos, o código executará novamente o método `getDataWithToken` e definirá uma opção para forçar uma solicitação de entrada.

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. Substitua `TODO3` pelo código a seguir. O erro 13002 ocorre quando a entrada ou o consentimento do usuário é anulado. Peça que o usuário tente novamente, mas não mais de uma vez.

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }          
        break; 
    ```

1. Substitua `TODO4` pelo código a seguir. O erro 13003 ocorre quando o usuário está conectado com uma conta que não é corporativa, de estudante nem da Microsoft. Peça que o usuário saia e entre novamente com um tipo de conta suportado.

    ```javascript
    case 13003: 
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;   
    ```

    > [!NOTE]
    > Os erros 13004 e 13005 não são processados neste método, pois eles só ocorrem em desenvolvimento. Eles não podem ser corrigidos pelo código de tempo de execução e não seria útil reportá-lo a um usuário final.

1. Substitua `TODO5` pelo seguinte código. O Erro 13006 ocorre quando houve um erro não especificado no host do Office, que pode indicar a instabilidade do host. Peça ao usuário para reiniciar o Office.

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;        
    ```

1. Substitua `TODO6` pelo código a seguir. O erro 13007 ocorre quando algo deu errado com a interação do host do Office com o AAD de forma que o host não pode obter um token de acesso para o serviço Web/aplicativo dos suplementos. É possível que esse seja um problema de rede temporário. Peça que o usuário tente novamente mais tarde.

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;      
    ```

1. Substitua `TODO7` pelo código a seguir. O Erro 13008 ocorre quando o usuário aciona uma operação que chama `getAccessTokenAsync` antes que uma chamada anterior dele seja concluída.

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```      

1. Substitua `TODO8` pelo código a seguir. O erro 13009 ocorre quando o suplemento não permite forçar consentimento, mas `getAccessTokenAsync` foi chamado com a opção `forceConsent` definida como `true`. Normalmente, quando isso acontece, o código deve ser reexecutar `getAccessTokenAsync` automaticamente com a opção de consentimento definida como `false`. No entanto, em alguns casos, chamar o método com `forceConsent` definido como `true` é uma resposta automática para um erro em uma chamada para o método com a opção definida como `false`. Nesse caso, o código não deve tentar novamente, mas, em vez disso, ele deve solicitar que o usuário saia e entre novamente.

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```      
    
1. Substitua `TODO9` pelo código a seguir.

    ```javascript
    default:
        logError(result);
        break;
    ```  


1. Abaixo do método `handleClientSideErrors`, adicione o seguinte método. Esse método processará os erros no serviço Web do suplemento quando algo der errado na execução do fluxo on-behalf-of ou ao obter dados do Microsoft Graph.

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

1. Substitua `TODO10` pelo código a seguir. Observe que, para a maioria dos erros `4xx` que o serviço Web do suplemento passará para o suplemento do lado do cliente, haverá uma propriedade **ExceptionMessage** em resposta com o número de erro AADSTS (Azure Active Directory Secure Token Service) além de outros dados. No entanto, quando AAD envia uma mensagem para o serviço Web do suplemento solicitando um fator de autenticação adicional, a mensagem contém uma propriedade **Claims** especial que especifica (com um número de código) qual fator adicional é necessário. As APIs ASP.NET que criam e enviam respostas HTTP para clientes não conhecem a propriedade **Claims**, portanto, elas não a incluem no objeto Response. O código de servidor que será criado em uma etapa posterior lidará com isso adicionando manualmente o valor **Claims** no objeto Response. Esse valor será salvo na propriedade **Message**, portanto, o código também precisará analisar essa propriedade.

    ```javascript
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    var message = JSON.parse(result.responseText).Message;
    ```

1. Substitua `TODO11` pelo código a seguir. Observação sobre este código:

    * O erro 50076 ocorre quando o Microsoft Graph requer uma forma adicional de autenticação.
    * O host do Office deve obter um novo token com o valor **Claims** como a opção `authChallenge`. Isso instrui o AAD a solicitar ao usuário todas as formas de autenticação requeridas. 

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
        }
    }    
    ```

1. Substitua `TODO12` pelo código a seguir. Observação sobre este código:

    * O erro 65001 significa que o consentimento para acessar o Microsoft Graph não foi concedido (ou foi revogado) para uma ou mais permissões. 
    * O suplemento deverá obter um novo token com a opção `forceConsent` definida como `true`.

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

1. Substitua `TODO13` pelo código a seguir. Observação sobre este código:

    * O Erro 70011 tem muitos significados. O que importa para este suplemento é quando ele significa que um escopo inválido (permissão) foi solicitado, então o código verifica a descrição completa do erro, não apenas o número.
    * O suplemento deverá relatar o erro.

    ```javascript
     else if (exceptionMessage.indexOf("AADSTS70011: The provided value for the input parameter 'scope' is not valid.") !== -1) {
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }    
    ```

1. Substitua `TODO14` pelo código a seguir. Observação sobre este código:

    * Código de servidor criado em uma etapa posterior enviará a mensagem `Missing access_as_user` se o escopo `access_as_user` (permissão) não for o token de acesso que o cliente do suplemento enviar para o ADD para ser usado no fluxo on-behalf-of.
    * O suplemento deverá relatar o erro.

    ```javascript
    else if (exceptionMessage.indexOf('Missing access_as_user.') !== -1) {
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }    
    ```

1. Substitua `TODO15` pelo código a seguir. Observação sobre este código:

    * A biblioteca de identidade que você usará no código do lado do servidor (Biblioteca de Autenticação da Microsoft - MSAL) deve garantir que nenhum token inválido ou expirado seja enviado para o Microsoft Graph. Contudo, se isso ocorrer, o erro retornado para serviço Web do suplemento do Microsoft Graph terá o código `InvalidAuthenticationToken`. O código do lado do servidor que você criará em uma etapa futura transmitirá essa mensagem ao cliente do suplemento.
    * Nesse caso, o suplemento deverá iniciar o processo de autenticação completo ao redefinir o contador e as variáveis de sinalizador e, em seguida, chamando novamente o método de identificador de botão.

    ```javascript
    // If the token sent to MS Graph is expired or invalid, start the whole process over.
    else if (result.code === 'InvalidAuthenticationToken') {
        timesGetOneDriveFilesHasRun = 0;
        triedWithoutForceConsent = false;
        getOneDriveFiles();
    }    
    ```

1. Substitua `TODO16` pelo código a seguir.

    ```javascript
    else {
        logError(result);
    }    
    ```

1. Salve e feche o arquivo.

## <a name="code-the-server-side"></a>Codifique o lado do servidor

### <a name="configure-the-owin-middleware"></a>Configurar o middleware OWIN

1. Abra o arquivo Startup.cs na raiz do projeto.

1. Adicione a palavra-chave `partial` para a declaração da classe Startup, se ainda não estiver lá. A linha deverá ser assim:

    `public partial class Startup`

1. Adicione a linha a seguir ao corpo do método `Configuration`. Você criará o método `ConfigureAuth` em uma etapa posterior.

    `ConfigureAuth(app);`

1. Salve e feche o arquivo.

1. Clique com botão direito do mouse na pasta **App_Start** e selecione **Adicionar > Classe**.

1. Na caixa de diálogo **Adicionar novo item** nomeie o arquivo **Startup.Auth.cs** e, em seguida, clique em **Adicionar**.

1. Encurte o nome do namespace no novo arquivo para `Office_Add_in_ASPNET_SSO_WebAPI`.

1. Verifique se todas as seguintes instruções `using` estão na parte superior do arquivo.

    ```csharp
    using Owin;
    using System.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. Adicione a palavra-chave `partial` à declaração da classe `Startup`, se ainda não estiver lá. A linha deverá ser assim:

    `public partial class Startup`

1. Adicione o método a seguir à classe `Startup`. Este método especifica como o middleware OWIN validará os tokens de acesso que são transmitidos a ele do método `getData` no arquivo Home.js do lado do cliente. O processo de autorização é disparado sempre que um ponto de extremidade da API Web decorado com o atributo `[Authorize]` é chamado.

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. Substitua TODO3 pelo seguinte código. Observação sobre o código:

    * O código instrui o OWIN a garantir que o emissor de token e audiência especificado no token de acesso que vem do host do Office (e é transmitido pela chamada de `getData` do lado do cliente) deve coincidir com os valores especificados no Web.config.
    * Definir `SaveSigninToken` como `true` faz com que o OWIN salve o token bruto do host do Office. O suplemento precisa dele para obter um token de acesso para o Microsoft Graph com o fluxo "on behalf of".
    * Os escopos não são validados pelo middleware OWIN. Os escopos do token de acesso, que devem conter `access_as_user`, são validados no controlador.

    ```csharp
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. Substitua TODO4 pelo seguinte. Observação sobre este código:

    * O método `UseOAuthBearerAuthentication` é chamado em vez do `UseWindowsAzureActiveDirectoryBearerAuthentication` que é mais comum, porque este último não é compatível com o ponto de extremidade V2 do Azure AD.
    * A URL de descoberta transmitida ao método é onde o middleware OWIN obtém instruções para conseguir a chave que precisa para verificar a assinatura no token de acesso recebido do host do Office.

    ```csharp
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
        {
            AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
        });
    ```

1. Salve e feche o arquivo.

### <a name="create-the-apivalues-controller"></a>Criar o controlador /api/values

1. Abra o arquivo **Controllers\ValueController.cs**.

2. Verifique se as seguintes instruções `using` estão na parte superior do arquivo.

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

3. Logo acima da linha que declara o `ValuesController`, adicione o atributo `[Authorize]`. Isso garante que seu suplemento executará o processo de autorização configurado no último procedimento sempre que um método controlador for chamado. Apenas os chamadores com um token de acesso válido para o seu suplemento podem invocar os métodos do controlador.

    > [!NOTE]
    > Um serviço da ASP.NET MVC Web API de produção deve ter lógica personalizada para o fluxo on-behalf-of em uma ou mais classes [FilterAttribute](https://docs.microsoft.com/previous-versions/aspnet/web-frameworks/hh834645(v=vs.108)) personalizadas. Este exemplo educacional coloca a lógica no controlador de principal para que o fluxo de autorização e dados busca lógica inteiro possa ser acompanhado facilmente. Isso também faz com que o exemplo fique consistente com os exemplos de padrão de autorização nos [Exemplos do Azure](https://github.com/Azure-Samples/).    

4. Adicione o método a seguir ao `ValuesController`. Observe que é o valor de retorno é `Task<HttpResponseMessage>` em vez de `Task<IEnumerable<string>>`, como seria mais comum para um método `GET api/values`. Este é um efeito colateral do fato de que nossa lógica de autorização personalizada estará no controlador: algumas condições de erro nessa lógica exigem que um objeto de resposta HTTP seja enviado para o cliente do suplemento. 

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO1: Validate the scopes of the access token.
    }
    ```

5. Substitua `TODO1` pelo seguinte código para validar que os escopos especificados no token incluam `access_as_user`.

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
    > Você deve usar apenas o escopo `access_as_user` para autorizar a API que lida com o fluxo Em Nome De para os suplementos do Office. Outras APIs em seu serviço devem ter seus próprios requisitos de escopo. Isso limita o que pode ser acessado com os tokens que o Office adquire.

6. Substitua `TODO2` pelo código a seguir. Observação sobre este código:
    * Ele transforma o token de acesso bruto recebido do host do Office em um objeto de `UserAssertion` que será transmitido para outro método.
    * Seu suplemento não está mais desempenhando o papel de um recurso (ou público) para o qual o host do Office e o usuário precisam de acesso. Agora, ele mesmo é um cliente que precisa de acesso ao Microsoft Graph. `ConfidentialClientApplication` é o objeto "client context" da MSAL.
    * O terceiro parâmetro para o construtor `ConfidentialClientApplication` é uma URL de redirecionamento que não é realmente usada no fluxo "on behalf of", mas usar a URL correta é uma boa prática. O quarto e o quinto parâmetros podem ser usados para definir um armazenamento persistente que permitiria a reutilização de tokens não expirados em diferentes sessões com o suplemento. Este exemplo não implementa nenhum armazenamento persistente.
    * A MSAL exige os escopos `openid` e `offline_access` para funcionar, mas ela lança um erro se o código solicitá-los de forma redundante. Ela também lançará um erro se o seu código solicitar o `profile`, que realmente é usado apenas quando o aplicativo host do Office recebe o token para o aplicativo Web do seu suplemento. Então, apenas `Files.Read.All` é explicitamente solicitado.

    ```csharp
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

7. Substitua `TODO3` pelo código a seguir. Observação sobre este código:

    * O método `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` procurará primeiro no cache da MSAL, que está na memória, para fazer a correspondência com o token de acesso. Somente se não houver um, ele iniciará o fluxo "on behalf of" com o ponto de extremidade V2 do Azure AD.
    * Se a autenticação multi-fator for requerida pelo recurso MS Graph e o usuário ainda não a tiver fornecido, o AAD lançará uma exceção contendo uma propriedade de Declarações.
    * O valor da propriedade de Declarações deve ser passado para o cliente, que o passará para o host do Office, que, em seguida, o incluirá em um pedido para um novo token. O AAD solicitará ao usuário todas as formas de autenticação necessárias.
    * Quaisquer exceções que não forem do tipo `MsalServiceException` são intencionalmente não detectadas, e, portanto, se propagarão para o cliente como mensagens `500 Server Error`.

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

8. Substitua `TODO3a` pelo código a seguir. Observação sobre este código:

    * Se a autenticação multifator for exigida pelo recurso MS Graph e o usuário ainda não a tiver fornecido, o AAD retornará "400 Bad Request" com o erro AADSTS50076 e uma propriedade **Declarações**. O MSAL lançará uma **MsalUiRequiredException** (que herda de **MsalServiceException**) com essas informações. 
    * O valor da propriedade **Declarações** deve ser passado para o cliente, que deve passá-lo para o host do Office, que, por sua vez, o incluirá em um pedido para um novo token. O AAD solicitará ao usuário todas as formas de autenticação necessárias.
    * As APIs que criam respostas HTTP a partir de exceções não conhecem a propriedade **Claims**, portanto, elas não a incluem no objeto de resposta. É necessário criar manualmente uma mensagem que inclua esse recurso. Uma propriedade **Message** personalizada, no entanto, impede a criação de uma propriedade **ExceptionMessage**, portanto, a única maneira de obter a ID de erro `AADSTS50076` para o cliente é adicioná-la à **Message** personalizada. O JavaScript no cliente precisará descobrir se uma resposta tem uma **Message** ou **ExceptionMessage** para saber qual ler.
    * A mensagem personalizada é formatada como JSON para que o JavaScript do cliente possa analisá-la com métodos de objeto `JSON` conhecidos.
    * Você criará o método `SendErrorToClient` em uma etapa posterior. É segundo parâmetro é um objeto **Exception**. Nesse caso, o código passa `null` porque incluir o objeto **Exception** bloqueia a inclusão da propriedade **Message** na resposta HTTP que é gerada.

    ```csharp
    if (e.Message.StartsWith("AADSTS50076")) {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

9. Substitua `TODO3b` e `TODO3c` pelo código a seguir. Observação sobre este código:

    * Se a chamada para o AAD contiver pelo menos um escopo (permissão) que não tenha sido consentido pelo usuário ou por um administrador de locatários (ou se o consentimento foi revogado), o AAD retornará "400 Solicitação Incorreta" com o erro `AADSTS65001`. O MSAL exibe um **MsalUiRequiredException** com essas informações. O cliente deve chamar `getAccessTokenAsync` novamente com a opção `{ forceConsent: true }`.
    *  Se a chamada para o AAD contiver pelo menos um escopo que AAD não reconhece, o AAD retornará "400 Solicitação Incorreta" com o erro `AADSTS70011`. O MSAL exibe um **MsalUiRequiredException** com essas informações. O cliente deve informar o usuário.
    *  A descrição completa é incluída porque 70011 é retornado em outras condições e ele deve ser processado nesse suplemento somente quando significar que há um escopo inválido. 
    *  O objeto **MsalUiRequiredException** é passado para `SendErrorToClient`. Isso garante que uma propriedade **ExceptionMessage** contendo as informações de erro seja incluída na resposta HTTP.
    *  Não há uma mensagem personalizada, portanto, `null` é passado para o terceiro parâmetro.

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001"))
    || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

10. Substitua `TODO3d` pelo código a seguir. Observe que o código exibe a exceção em vez de transmiti-la em uma resposta HTTP personalizada com **HttpStatusCode.Forbidden** (401). O efeito disso é que o ASP.NET enviará sua própria resposta HTTP com o status "500 Erro de Servidor".

    ```csharp
    else
    {
        throw e;
    }  
    ```

11. Substitua `TODO4` pelo seguinte. Observação sobre este código:

    * As classes `GraphApiHelper` e `ODataHelper` são definidas nos arquivos da pasta **Helpers**. A classe `OneDriveItem` é definida em um arquivo da pasta **Models**. A discussão detalhada dessas classes não é relevante para a autorização ou o SSO, portanto, está fora do escopo deste artigo.
    * O desempenho é aprimorado ao se solicitar ao Microsoft Graph apenas os dados que são realmente necessários. Desse modo, o código usa um parâmetro de consulta ` $select` para especificar que desejamos somente a propriedade de nome, e usa um parâmetro `$top` para especificar que desejamos somente os três primeiros nomes de pasta ou de arquivo.
    * Se o token enviado para o Microsoft Graph for inválido, o Microsoft Graph enviará um erro "401 Não Autorizado" com o código "InvalidAuthenticationToken". Em seguida, o ASP.NET exibe um **RuntimeBinderException**. Isso também ocorre quando o token expira, embora o MSAL deva impedir que isso aconteça. 

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

12. Substitua `TODO5` pelo seguinte. Observação sobre este código: 

    * Embora o código acima solicite somente a propriedade *name* dos itens do OneDrive, o Microsoft Graph sempre inclui a propriedade *eTag* para os itens do OneDrive. Para reduzir a carga enviada para o cliente, o código a seguir reconstrói os resultados apenas com os nomes dos itens.
    * A lista de três pastas e arquivos do OneDrive é enviada para o cliente como uma resposta HTTP "200 OK".

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

13. Abaixo do método Get, adicione o método a seguir. Sobre este código, observe:  

    * O método transmite ao cliente informações sobre uma exceção do servidor. 
    * Se a exceção original for passada para o método, o construtor HttpError incluirá informações do objeto de exceção em uma propriedade **ExceptionMessage**.  
    * Se `null` for passado para a exceção, o construtor HttpError incluirá o parâmetro de mensagem em uma propriedade **Message** e não haverá uma propriedade **ExceptionMessage**.

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

1. Certifique-se de ter alguns arquivos no seu OneDrive para que você possa verificar os resultados.

1. No Visual Studio, pressione F5. O PowerPoint será aberto e haverá um grupo **SSO ASP.NET** na faixa de opções **Página Inicial**.

1. Pressione o botão **Mostrar Suplemento** nesse grupo para ver a interface do usuário do suplemento no painel de tarefas.

1. Pressione o botão **Obter meus arquivos do OneDrive**. Se você não estiver conectado ao Office, você será solicitado a entrar.
    
    > [!NOTE]
    > Se você entrou no Office com uma ID diferente e se alguns aplicativos do Office que estavam abertos no momento continuam abertos, o Office pode não alterar de forma confiável sua ID, mesmo que pareça ter feito isso no PowerPoint. Se isso acontecer, a chamada para o Microsoft Graph pode falhar ou os dados da ID anterior podem ser retornados. Para evitar isso, certifique-se de *fechar todos os outros aplicativos do Office* antes de pressionar **Obter meus arquivos do OneDrive**.

1. Depois de entrar, será exibida uma lista de seus arquivos e suas pastas no OneDrive, abaixo do botão. Esse procedimento pode levar mais de 15 segundos, principalmente na primeira vez.
