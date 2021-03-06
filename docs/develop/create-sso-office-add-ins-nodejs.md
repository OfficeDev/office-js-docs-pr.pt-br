---
title: Crie um Suplemento do Office com Node.js que use logon único
description: Aprenda a criar um suplemento baseado em node.js que usa o logon único do Office
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: a587f6a1ad596f4006734e0d850e009c3fc145b0
ms.sourcegitcommit: 6e6de48f746416ec68b2cf4c298253986486fbfd
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/14/2020
ms.locfileid: "47651934"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a>Crie um Suplemento do Office com Node.js que use logon único

Os usuários podem entrar no Office, e o Suplemento Web do Office pode aproveitar esse processo de entrada para autorizá-los a acessar seu suplemento e o Microsoft Graph sem exigir que os eles entrem uma segunda vez. Para obter uma visão geral, confira o artigo [Habilitar o SSO em um Suplemento do Office](sso-in-office-add-ins.md).

Este artigo apresenta o processo passo a passo de habilitação do logon único (SSO) em um suplemento que foi criado com Node.js e Express. Para ler um artigo semelhante sobre um suplemento baseado em ASP.NET, confira [Criar um Suplemento do Office com ASP.NET que usa o logon único](create-sso-office-add-ins-aspnet.md).

> [!NOTE]
> Como alternativa para concluir as etapas descritas neste artigo, você pode usar o gerador Yeoman para criar um Suplemento do Office com Node.js habilitado para SSO. O gerador Yeoman simplifica o processo de criação de um suplemento habilitado para SSO, automatizando as etapas necessárias para configurar o SSO no Azure e gerando o código necessário para um suplemento usar o SSO. Para obter mais informações, confira [Início rápido de logon único (SSO)](../quickstarts/sso-quickstart.md).

## <a name="prerequisites"></a>Pré-requisitos

* [Node.js](https://nodejs.org/) (a versão mais recente de [LTS](https://nodejs.org/about/releases))

* [Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)

* TypeScript, versão 3.6.2 ou posterior.

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* Um editor de códigos. Recomendamos o código do Visual Studio.

* Pelo menos alguns arquivos e pastas armazenados no OneDrive for Business em sua assinatura do Microsoft 365.

* Uma assinatura do Microsoft Azure. Este suplemento requer o Azure Active Directory (AD). O Active AD fornece serviços de identidade que os aplicativos usam para autenticação e autorização. Você pode adquirir uma assinatura de avaliação no [Microsoft Azure](https://account.windowsazure.com/SignUp).

## <a name="set-up-the-starter-project"></a>Configure o projeto inicial

1. Clone ou baixe o repositório em [SSO com Suplemento NodeJS do Office](https://github.com/officedev/office-add-in-nodejs-sso).

    > [!NOTE]
    > Há três versões do exemplo:  
    > * A pasta **inicial** é um projeto inicial. A interface de usuário e outros aspectos do suplemento que não estejam diretamente conectados ao SSO ou à autorização já foram feitos. Seções posteriores deste artigo orientam você durante o processo de conclusão.
    > * A versão **Complete** (concluído) do exemplo apresenta como seria o suplemento quando concluídos os procedimentos apresentados neste artigo, com exceção de que o projeto concluído traz comentários de códigos que seriam redundantes neste artigo. Para usar a versão completa, basta seguir as instruções deste artigo, mas substituir "Begin" por "concluído" e ignorar as seções **codificadas pelo cliente** e **codificar o** lado do servidor.
    > * A versão **SSOAutoSetup** é um exemplo concluído que automatiza a maioria das etapas para registrar o suplemento com o Azure AD e configurá-lo. Use esta versão se desejar ver um suplemento de trabalho com SSO rapidamente. Basta seguir as etapas no README da pasta. É recomendável que, em algum momento, você siga as etapas de configuração e registro manuais deste artigo para entender melhor a relação entre o Azure AD e um suplemento. 

1. Abra um prompt de comando na pasta **Iniciar** .

1. Insira `npm install` no console para instalar todas as dependências discriminadas no arquivo package.json.

1. Execute o comando `npm run install-dev-certs`. Selecione **Sim** à solicitação para instalar o certificado.

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Registre o suplemento com o ponto de extremidade v2.0 do Azure AD

1. Acesse a página [Portal do Azure - Registros de aplicativo](https://go.microsoft.com/fwlink/?linkid=2083908) para registrar o seu aplicativo.

1. Entre com as credenciais de ***administrador*** em seu Microsoft 365 locação. Por exemplo, MeuNome@contoso.onmicrosoft.com.

1. Selecione **Novo registro**. Na página **Registrar um aplicativo**, defina os valores da seguinte forma.

    * Defina **Nome** para `Office-Add-in-NodeJS-SSO`.
    * Defina **Tipos de conta com suporte** para **Contas em qualquer diretório organizacional e contas pessoais da Microsoft (por exemplo, Skype, Xbox, Outlook.com)**.
    * Defina o tipo de aplicativo como **Web** e, em seguida, defina **URI de redirecionamento** como ` https://localhost:44355/dialog.html` .
    * Escolha **Registrar**.

1. Na página **Office-Add-in-NodeJS-SSO**, copie e salve os valores para a **ID do aplicativo (cliente)** e a **ID do diretório (locatário)**. Use ambos os valores nos procedimentos posteriores.

    > [!NOTE]
    > Este **aplicativo (cliente) ID** é o valor de "audiência" quando outros aplicativos, como o aplicativo cliente do Office (por exemplo, PowerPoint, Word, Excel), procuram acesso autorizado ao aplicativo. Também é a "ID do cliente" do aplicativo quando ela, por sua vez, busca acesso autorizado ao Microsoft Graph.

1. Selecione **Autenticação** em **Gerenciar**. Na seção **concessão implícita** , habilite as caixas de seleção para token de **acesso** e **token de ID**. O exemplo tem um sistema de autorização de fallback que é chamado quando o SSO não está disponível. Esse sistema usa o fluxo implícito.

1. Na parte superior da página, selecione **Salvar**.

1. Selecione **Certificados e segredos** sob **Gerenciar**. Selecione o botão **Novo segredo do cliente**. Insira um valor para **Descrição** e, em seguida, selecione uma opção adequada para **Expira** e escolha **Adicionar**. *Copiar o valor de segredo do cliente imediatamente e salvá-lo com a ID de aplicativo* antes de prosseguir, pois ele será necessário em um procedimento posterior.

1. Selecionar **Expor uma API** em **Gerenciar**. Selecione o link **definir** . Isso gerará o URI da ID do aplicativo no formato "api://$App ID GUID $", onde $App GUID de ID $ é a **ID do aplicativo (cliente)**.

1. Na ID gerada, insira `localhost:44355/` (Observe a barra "/" anexada ao final) entre as barras duplas e o GUID. Quando você terminar, a ID inteira deverá ter a forma `api://localhost:44355/$App ID GUID$` ; por exemplo `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7` .

1. Selecione o botão **Adicionar um escopo**. No painel que se abre, insira `access_as_user` como o **Nome de escopo**.

1. Definir **Quem pode consentir?** aos **Administradores e usuários**.

1. Preencha os campos para configurar os prompts de consentimento de usuário e administrador com os valores que são apropriados para o `access_as_user` escopo, o que permite que o aplicativo cliente do Office Use as APIs Web do seu suplemento com os mesmos direitos que o usuário atual. Sugestões:

    - **Nome para exibição do consentimento do administrador**: o Office pode atuar como o usuário.
    - **Descrição de autorização de administrador:** Permite ao Office chamar os APIs de suplemento da web com os mesmos direitos que o usuário atual.
    - **Nome para exibição do consentimento do usuário**: o Office pode agir como você.
    - **Descrição do consentimento do usuário**: habilitar o Office para chamar as APIs Web do suplemento com os mesmos direitos que você tem.

1. Verifique se o **Estado** está definido como **Habilitado**.

1. Selecione **Adicionar escopo**.

    > [!NOTE]
    > A parte de domínio do nome de **Escopo** exibidos logo abaixo do campo de texto deve corresponder automaticamente ao URI de ID do aplicativo definidos na etapa anterior com `/access_as_user` acrescentado ao final; por exemplo, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. Na seção **Aplicativos clientes autorizados**, você identifica os aplicativos que deseja autorizar para o aplicativo da Web do seu suplemento. Cada uma das seguintes IDs precisa ser pré-autorizada.

    - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
    - `57fb890c-0dab-4253-a5e0-7188c88b2bb4`(Office na Web)
    - `08e18876-6177-487e-b8b5-cf950c1e598c`(Office na Web)
    - `bc59ab01-8403-45c6-8796-ac3ef710b3e3`(Outlook na Web)

    Para cada ID, siga estas etapas:

    a. Selecione o botão **Adicionar um aplicativo cliente** e, no painel que se abre, defina o ID do cliente para o respectivo GUID e marque a caixa `api://localhost:44355/$App ID GUID$/access_as_user`.

    b. Selecione **Adicionar aplicativo**.

1. Selecione **Permissões para API** em **Gerenciar** e selecione **Adicionar uma permissão**. No painel que se abre, escolha **Microsoft Graph** e, em seguida, escolha **Permissões delegadas**.

1. Use a caixa de pesquisa **Selecionar permissões** para procurar as permissões que o seu suplemento precisa. Selecione estas opções. Somente o primeiro é realmente necessário para o seu suplemento; Mas a `profile` permissão é necessária para que o aplicativo do Office obtenha um token para o aplicativo Web do seu suplemento.

    * Files.Read.All
    * perfil

    > [!NOTE]
    > A permissão `User.Read` pode já estar listada por padrão. É uma boa prática não pedir permissões desnecessárias, por isso recomendamos desmarcar a caixa para essa permissão se o suplemento não precisar dela.

1. Marque a caixa de seleção para cada permissão conforme elas forem exibidas. Depois de selecionar as permissões que o suplemento precisa, selecione o botão **Adicionar permissões** na parte inferior do painel.

1. Na mesma página, escolha o botão **conceder permissão de administrador para [nome do locatário]** e, em seguida, selecione **Sim** para a confirmação exibida.

## <a name="configure-the-add-in"></a>Configurar o suplemento

1. Abra a pasta `\Begin` no projeto clonado no editor de códigos.

1. Abra o arquivo `.ENV` e use os valores que você copiou anteriormente. Defina o **CLIENT_ID** para a identificação do seu **ID de aplicativo (cliente)** e defina **CLIENT_SECRET** para o seu segredo de cliente. Os valores **não** devem estar entre aspas. Quando terminar, o arquivo deverá ser semelhante ao seguinte: 

    ```javascript
    CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
    CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
    NODE_ENV=development
    ```

1. Abra o arquivo `\public\javascripts\fallbackAuthDialog.js`. Na declaração `msalConfig` substitua o espaço reservado "{application_GUID here}", pela ID do Aplicativo que você copiou ao registrar seu suplemento. O valor deve estar entre aspas.

1. Abra o arquivo de manifesto de suplemento "manifest\ manifest_local.xml" e role até a parte inferior do arquivo. Logo acima da marca de fim `</VersionOverrides>`, você encontrará a marcação a seguir:

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. Substitua o espaço reservado "{$application_GUID here$}" *nos dois lugares* na marcação pela ID do Aplicativo que você copiou ao registrar seu suplemento. O símbolo "$" não faz parte da ID, portanto não o inclua. Esta é a mesma ID usada para o CLIENT_ID e audiência no. ENV arquivo.

    > [!NOTE]
    > O valor **Recurso** é o**URI da ID de aplicativo** que você definiu quando registrou o suplemento. A seção **Scopes** só será usada para gerar uma caixa de diálogo de consentimento se o suplemento for vendido no AppSource.

## <a name="code-the-client-side"></a>Codificar o lado do cliente

### <a name="create-the-sso-logic"></a>Criar a lógica de SSO

1. No editor de códigos, abra o arquivo `public\javascripts\ssoAuthES6.js`. Ele já tem um código que garante que o Promises seja suportado, mesmo no Internet Explorer 11, e uma chamada`Office.onReady` para atribuir um manipulador para o botão somente suplemento.

    > [!NOTE]
    > Como o nome sugere, o ssoAuthES6.js usa a sintaxe JavaScript ES6, pois usar `async` e `await` mostra melhor a simplicidade fundamental da API de SSO. Quando o servidor localhost for iniciado, esse arquivo será transformado em uma sintaxe ES5 para que o exemplo seja executado no Internet Explorer 11. 

1. Adicione o seguinte código abaixo do método Office. onReady:

    ```javascript
    async function getGraphData() {
        try {
            
            // TODO 1: Tell Office to get a bootstrap token from Azure AD.
            
            // TODO 2: Attempt to exchange the bootstrap token for an 
            //         access token to Microsoft Graph.

            // TODO 3: Handle case where Microsoft Graph requires an 
            //         additional form of authentication.

            // TODO 4: Use the access token in a call to Microsoft Graph 
            //         or handle any error from the attempted token exchange.

        }
        catch(exception) {

            // TODO 5: Respond to exceptions thrown by the
            //         OfficeRuntime.auth.getAccessToken call.

        }
    }
    ```

1. Substitua `TODO 1` pelo código a seguir. Sobre este código, observe:

    - `OfficeRuntime.auth.getAccessToken` instrui o Office a obter um token de bootstrap do Azure AD. Um token de bootstrap é semelhante a um token de ID, mas tem uma propriedade `scp` (Scope) com o valor `access-as-user`. Esse tipo de token pode ser trocado por um aplicativo Web para um token de acesso ao Microsoft Graph.
    - Definir a `allowSignInPrompt` opção como true significa que, se nenhum usuário estiver conectado ao Office no momento, o Office abrirá um prompt de entrada pop-up.
    - A definição da `allowConsentPrompt` opção como true significa que, se o usuário não tiver sido remetido para permitir que o suplemento acesse o perfil AAD do usuário, o Office abrirá um prompt de consentimento. (O prompt permite que o usuário consentisse no perfil AAD do usuário, e não nos escopos do Microsoft Graph).
    - Definir a `forMSGraphAccess` opção como true indica ao Office que o suplemento pretende usar o token de inicialização para obter um token de acesso ao Microsoft Graph, em vez de apenas usá-lo como um token de ID. Se o administrador locatário não tiver concedido consentimento para o acesso do suplemento ao Microsoft Graph, `OfficeRuntime.auth.getAccessToken` retornará o erro **13012**. O suplemento pode responder voltando para um sistema alternativo de autorização. Isso é necessário porque o Office pode solicitar apenas consentimento para o perfil do Azure AD do usuário, não para escopos do Microsoft Graph. O sistema de autorização de fallback exige que o usuário entre novamente e o usuário *pode* ser solicitado a se concordar com escopos do Microsoft Graph. Portanto, a opção `forMSGraphAccess` garante que o suplemento não fará uma troca de tokens que falhará devido à falta de consentimento. Uma vez que você concedeu consentimento de administrador em uma etapa anterior, esse cenário não acontecerá para esse suplemento. No entanto, a opção é incluída aqui para ilustrar uma prática recomendada.

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true }); 
    ```

1. Substitua `TODO 2` pelo código a seguir. Você criará o método `getGraphToken` em uma etapa posterior.

    ```javascript
    let exchangeResponse = await getGraphToken(bootstrapToken);
    ```

1. Substitua `TODO 3` pelo seguinte. Sobre este código, observe: 

    - Se o Microsoft 365 locatário tiver sido configurado para exigir a autenticação multifator, o `exchangeResponse` incluirá uma `claims` propriedade com informações sobre os outros fatores necessários. Nesse caso, `OfficeRuntime.auth.getAccessToken` deve ser chamado novamente com a opção `authChallenge` definida como o valor da propriedade de declarações. Isso instrui o AAD a solicitar ao usuário todas as formas de autenticação requeridas.

    ```javascript
    if (exchangeResponse.claims) {
        let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
        exchangeResponse = await getGraphToken(mfaBootstrapToken);
    }
    ```

1. Substitua `TODO 4` pelo seguinte. Sobre este código, observe: 

    - Você criará o método `handleAADErrors` em uma etapa posterior. Os erros do Azure AD são retornados para o cliente como respostas HTTP # 200. Eles não geram erros, portanto, não disparam o bloco `catch` do método `getGraphData`.
    - Você criará o método `makeGraphApiCall` em uma etapa posterior. Ele faz uma chamada AJAX para o ponto de extremidade do MS Graph. Os erros são detectados na callback`.fail` da chamada, não no bloco `catch` do método `getGraphData`.

    ```javascript
    if (exchangeResponse.error) {
        handleAADErrors(exchangeResponse);
    } 
    else {
        makeGraphApiCall(exchangeResponse.access_token);
    }
    ```

1. Substitua `TODO 5` pelo seguinte

    - Os erros da chamada `getAccessToken` terão uma propriedade `code` com um número de erro, normalmente no intervalo 13xxx. Você criará o método `handleClientSideErrors` em uma etapa posterior.
    - O método `showMessage` exibe o texto no painel de tarefas.

    ```javascript
    if (exception.code) { 
        handleClientSideErrors(exception);
    }
    else {
        showMessage("EXCEPTION: " + JSON.stringify(exception));
    }
    ```

1. Abaixo do método `getGraphData`, adicione a função a seguir. Observe que `/auth` é uma rota expressa do servidor que troca o token de inicialização com o Azure ad para obter um token de acesso para o Microsoft Graph.

    ```javascript
    async function getGraphToken(bootstrapToken) {
        let response = await $.ajax({type: "GET", 
            url: "/auth",
            headers: {"Authorization": "Bearer " + bootstrapToken }, 
            cache: false
        });
        return response;
    }
    ```

1. Abaixo do método `getGraphToken`, adicione a função a seguir. Observe que `error.code` é um número, normalmente no intervalo 13xxx.

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 6: Handle errors where the add-in should NOT invoke 
            //         the alternative system of authorization.

            // TODO 7: Handle errors where the add-in should invoke 
            //         the alternative system of authorization.

        }
    }
    ```

1. Substitua `TODO 6` pelo código a seguir. Para saber mais sobre esses erros, confira [Solucionar problemas de SSO em suplementos do Office em](troubleshoot-sso-in-office-add-ins.md). 

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one 
        // is logged into Office, then the first call of getAccessToken should pass the 
        // `allowSignInPrompt: true` option. Since this add-in does that, you should not see
        // this error. 
        showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");  
        break;
    case 13002:
        // OfficeRuntime.auth.getAccessToken was called with the allowConsentPrompt 
        // option set to true. But, the user aborted the consent prompt. 
        showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."); 
        break;
    case 13006:
        // Only seen in Office on the web.
        showMessage("Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."); 
        break;
    case 13008:
        // The OfficeRuntime.auth.getAccessToken method has already been called and 
        // that call has not completed yet. Only seen in Office on the web.
        showMessage("Office is still working on the last operation. When it completes, try this operation again."); 
        break;
    case 13010:
        // Only seen in Office on the web.
        showMessage("Follow the instructions to change your browser's zone configuration.");
        break;
    ```

1. Substitua `TODO 7` pelo código a seguir. Para saber mais sobre esses erros, confira [Solucionar problemas de SSO em suplementos do Office](troubleshoot-sso-in-office-add-ins.md). A função `dialogFallback` invoca o sistema de autorização alternativo. Neste suplemento, o sistema de fallback abre uma caixa de diálogo que exige que o usuário entre, mesmo que o usuário já esteja, e use o msal.js e Implicit Flow para obter um token de acesso ao Microsoft Graph.

    ```javascript
    default:
    // For all other errors, including 13000, 13003, 13005, 13007, 13012, 
    // and 50001, fall back to non-SSO sign-in.
    dialogFallback();
    break;
    ```

1. Abaixo da função `handleClientSideErrors`, adicione a função a seguir. 

    ```javascript
    function handleAADErrors(exchangeResponse) {

    // TODO 8: Handle case where the bootstrap token is expired.

    // TODO 9: Handle all other Azure AD errors.
    
    }
    ```

1. Em raras ocasiões, o token de bootstrap no cache do Office fica não vencido quando o Office o valida, mas vence no momento em que ele atinge o Azure AD para o Exchange. O Azure AD responderá com o erro **AADSTS500133**. Nesse caso, o suplemento deve simplesmente ligar recursivamente o `getGraphData`. Como o token de inicialização em cache já expirou, o Office receberá um novo token do Azure AD. Portanto, substitua `TODO 8` pelo seguinte. 

    ```javascript
    if (exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
    {
        getGraphData();
    }
    ```

1. Para garantir que o suplemento não insira um loop infinito de chamadas para `getGraphData`, o suplemento deve controlar quantas vezes `getGraphData` foi chamado e ter a certeza de que o não é chamado recursivomente chamado mais de uma vez. Portanto, crie uma variável de contador em um escopo global para as funções `handleAADErrors` e `getGraphData`. Um bom lugar para as variáveis globais está logo abaixo da chamada de método `Office.onReady`.

    ```javascript
    let retryGetAccessToken = 0;
    ```

1. Altere a estrutura `if` no método `handleAADErrors` para que ela:

    - Incremente o contador antes de chamar `getGraphData`.
    - Teste para garantir que `getGraphData` ainda não tenha sido chamado pela segunda vez. 

    Portanto, a versão final da estrutura `if` deve ter a seguinte aparência:

    ```javascript
    if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
        &&
        (retryGetAccessToken <= 0)) 
    {
        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. Substitua `TODO 9` pelo seguinte. 

    ```javascript
    else {
        dialogFallback();
    }
    ```

1. Salve e feche o arquivo.

### <a name="get-the-data-and-add-it-to-the-office-document"></a>Obtenha os dados e adicione-os ao documento do Office

1. Na pasta `public\javascripts`, crie um novo arquivo chamado `data.js`.

1. Adicione a seguinte função ao arquivo. Esta é a função que é chamada pela função `getGraphData` quando tiver adquirido um token de acesso ao Microsoft Graph. 

    ```javascript
    function makeGraphApiCall(accessToken) {
        $.ajax(

            // TODO 10: Call an Express route on the add-in's server-side 
            //          code and pass the access token to Microsoft Graph.

        )
        .done(function (response) {

            // TODO 11: Write the data received from Microsoft Graph to 
            //          the Office document.

        })
        .fail(function (errorResult) {
            showMessage("Error from Microsoft Graph: " + JSON.stringify(errorResult));
        });
    }
    ```

1. Substitua `TODO 10` pelo seguinte. Sobre este código, observe: 

    - Esse objeto é o parâmetro para o método `$.ajax`.
    - O `/getuserdata` é uma rota expressa no servidor do suplemento criado em uma etapa posterior. Ele chamará um ponto de extremidade do Microsoft Graph e incluiremos o token de acesso em sua chamada. 

    ```javascript
    {
        type: "GET",
        url: "/getuserdata",
        headers: {"access_token": accessToken },
        cache: false
    }
    ```

1. Substitua `TODO11` pelo seguinte. Sobre este código, observe:

    - `writeFileNamesToOfficeDocument` inserirá os dados do gráfico no documento do Office. Ela é definida no arquivo `public\javascripts\document.js`. 
    - Se `writeFileNamesToOfficeDocument` retornar um erro, ele começará com "não é possível adicionar nomes de arquivo ao documento".

    ```javascript
    writeFileNamesToOfficeDocument(response)
    .then(function () {
        showMessage("Your data has been added to the document.");
    })
    .catch(function (error) {
        showMessage(error);
    });
    ```

1. Salve e feche o arquivo.

## <a name="code-the-server-side"></a>Codifique o lado do servidor

### <a name="create-the-auth-router-and-the-token-exchange-logic"></a>Crie o roteador de autenticação e a lógica de troca de tokens

1. Abra o arquivo `routes\authRoute.js` e adicione a seguinte função de rota logo abaixo das instruções `require` e acima da instrução `module.exports`. Observe que o parâmetro de URL de `router.get` é '/'. Como esta rota está sendo definida em um roteador que tratará todas as solicitações HTTP para a URL "/auth", esta rota manipula todas as solicitações de "/auth". A função `getGraphToken` do lado do cliente que você criou anteriormente chama essa rota.  

    ```javascript
    router.get('/', async function(req, res, next) {

        // TODO 12: Test for the presence of the Authorization header.

        // TODO 13: Create the hidden form that will be sent to Azure AD 
        //          to request the access token in exchange for the 
        //          bootstrap token.

        // TODO 14: Send the POST request to Azure AD and relay the 
        //          access token (or an error) to the client.

    });
    ```

1. Substitua `TODO 12` pelo código a seguir.

    ```javascript
    const authorization = req.get('Authorization');
    if (authorization == null) {
        let error = new Error('No Authorization header was found.');
        next(error);
    } 
    ```

1. Substitua `TODO 13` pelo código a seguir. Sobre este código, observe: 

    - Este é o início de um bloco `else` longo, mas o `}` de fechamento não está no final, já que você adicionará mais código a ele. 
    - A cadeia de caracteres `authorization` é um "transportador" seguido pelo token bootstrap, portanto, a primeira linha do bloco `else` está atribuindo o token para `jwt`. ("JWT" significa "JSON Web Token".)
    - Os dois valores `process.env.*` são as constantes que você atribuiu ao configurar o suplemento. 
    - O parâmetro de formulário `requested_token_use` está definido como ' on_behalf_of '. Isso informa ao Azure AD que o suplemento está solicitando um token de acesso ao Microsoft Graph usando o fluxo On-Behalf-Of. O Azure responderá validando que o token de bootstrap, que é atribuído ao parâmetro de formulário `assertion`, tem uma propriedade `scp` que está definida como `access-as-user`.
    - O parâmetro de formulário `scope` está definido como "Files.Read.All', que é o único escopo do Microsoft Graph necessário para o suplemento.

    ```javascript
     else {
        const [schema, jwt] = authorization.split(' ');
        const formParams = {
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: jwt,
        requested_token_use: 'on_behalf_of',
        scope: ['Files.Read.All'].join(' ')
        };
    ```

1. Substitua `TODO 14` pelo código a seguir, que completa o bloco `else`. Sobre este código, observe:

    - A constante `tenant` é definida como "comum" porque você configurou o suplemento como multilocatário ao registrá-lo no Azure AD, especificamente quando você define **Tipos de conta com suporte** para **Contas em qualquer diretório corporativo e contas pessoais da Microsoft (por exemplo, Skype, Xbox, Outlook.com)**. Se, em vez disso, você optou por suportar apenas contas no mesmo locatário do Microsoft 365 em que o suplemento está registrado, o código `tenant` seria definido como o GUID do locatário. 
    - Se a solicitação POST não for recebida, a resposta do Azure AD será convertida para JSON e enviada para o cliente. Esse objeto JSON tem uma propriedade `access_token` à qual o Azure AD atribuiu o token de acesso ao Microsoft Graph.

    ```javascript
        const stsDomain = 'https://login.microsoftonline.com';
        const tenant = 'common';
        const tokenURLSegment = 'oauth2/v2.0/token';

        try {
            const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
                method: 'POST',
                body: form(formParams),
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            });
            const json = await tokenResponse.json();

            res.send(json);
        }
        catch(error) {
            res.status(500).send(error);
        }
    }
    ```

1. Salve e feche o arquivo.

### <a name="create-the-route-that-will-fetch-the-data-from-microsoft-graph"></a>Criar o roteiro que obterá os dados do Microsoft Graph

1. Abra o arquivo `app.js` na raiz do projeto. Logo abaixo da rota para "/dialog.html", adicione a seguinte rota. Esse roteiro é chamado pela função `makeGraphApiCall` que você criou em uma etapa anterior.

    ```javascript
    app.get('/getuserdata', async function(req, res, next) {
        
        // TODO 15: Send a request to the Microsoft Graph REST endpoint.

        // TODO 16: Trim excess information from the returned data and relay it
        //          to the client.
        
    });
    ```

1. Substitua `TODO 15` pelo seguinte. Sobre este código, observe:

    - O chamador dessa rota, `makeGraphApiCall`, adicionou o token de acesso ao Microsoft Graph à solicitação HTTP como um cabeçalho chamado "access_token".
    - A função `getGraphData` é definida no arquivo`msgraph-helper.js`. (Essa não é a mesma função que a função do lado do cliente`getGraphData` definida no arquivo de `ssoAuthES6.js`.)
    - O último parâmetro, por `queryParamsSegment`, é codificado. Se você reutilizar o código em um suplemento de produção e provenientes de qualquer parte do `queryParamsSegment` de entrada do usuário, certifique-se de que estão limpos para que não possam ser usados em um ataque de inserção de cabeçalho de resposta.
    - O código minimiza os dados que devem ser provenientes do Microsoft Graph especificando apenas a propriedade de que precisamos ("nome") e somente os 10 primeiros nomes de pasta ou arquivo.

    ```javascript
    const graphToken = req.get('access_token');
    const graphData = await getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=10");
    ```

1. Substitua `TODO 16` pelo seguinte. Sobre este código, observe:

    - Se o Microsoft Graph retornar um erro, como um token inválido ou expirado, haverá uma propriedade de código no conjunto de objetos retornados para um status HTTP (por exemplo, 401). O código retransmite o erro para o cliente. Ele será pego na callback `.fail` do `makeGraphApiCall`.
    - Os dados do Microsoft Graph incluem metadados OData e eTags que o suplemento não precisa, portanto, o código cria uma nova matriz contendo somente os nomes de arquivos a serem enviados para o cliente.

    ```javascript
    if (graphData.code) {
        next(createError(graphData.code, "Microsoft Graph error: " + JSON.stringify(graphData)));
    }
    else {
        const itemNames = [];
        const oneDriveItems = graphData['value'];
        for (let item of oneDriveItems) {
            itemNames.push(item['name']);
        }

        res.send(itemNames)
    }
    ```

1. Salve e feche o arquivo.

## <a name="run-the-project"></a>Executar o projeto

1. Certifique-se de ter alguns arquivos no seu OneDrive para que você possa verificar os resultados.

1. Abra um aviso de comando na raiz da pasta `\Begin`. 

1. Execute o comando `npm start`. 

1. Você deve fazer o sideload do suplemento em um aplicativo do Office (Excel, Word ou PowerPoint) para testá-lo. As instruções dependem da plataforma. Há links para instruções no [Fazer sideload de suplemento para teste](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

1. No aplicativo do Office, na faixa de opções **Home**, selecione o botão **Mostrar suplemento** no grupo**SSO Node.js** para abrir o suplemento do painel de tarefas.

1. Clique no botão **Definir Nome de Arquivos do One Drive**. Se você estiver conectado ao Office com uma conta de educação ou de trabalho do Microsoft 365, ou uma conta da Microsoft, e o SSO estiver funcionando conforme o esperado, os 10 primeiros nomes de arquivos e pastas no OneDrive for Business serão inseridos no documento. Isso pode levar até 15 segundos pela primeira vez. Se você não estiver conectado ou se você estiver em um cenário que não tem suporte para SSO, ou se o SSO não estiver funcionando por nenhum motivo, você será solicitado a fazer logon. Depois de entrar, os nomes de arquivos e pastas serão exibidos.

> [!NOTE]
> Se você entrou no Office com uma ID diferente e se alguns aplicativos do Office que estavam abertos no momento continuam abertos, o Office pode não alterar de forma confiável sua ID, mesmo que pareça ter feito isso. Se isso acontecer, a chamada para o Microsoft Graph pode falhar ou os dados da ID anterior podem ser retornados. Para evitar isso, certifique-se de *fechar todos os outros aplicativos do Office* antes de pressionar **Obter nomes de arquivos do OneDrive**.
