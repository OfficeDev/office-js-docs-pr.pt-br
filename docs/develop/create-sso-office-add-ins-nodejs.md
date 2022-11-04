---
title: Crie um Suplemento do Office com Node.js que use logon único
description: Saiba como criar um suplemento baseado em Node.js que usa o Logon Único do Office.
ms.date: 10/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: 35128da43b3f27a58df5e188a5001bfa8aba4a4c
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/28/2022
ms.locfileid: "68841663"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a>Crie um Suplemento do Office com Node.js que use logon único

Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).

Este artigo orienta você sobre o processo de habilitação do SSO (logon único) em um suplemento. O suplemento de exemplo que você cria tem duas partes; um painel de tarefas que carrega no Microsoft Excel e um servidor de camada média que manipula chamadas para o Microsoft Graph para o painel de tarefas. O servidor de camada intermediária é criado com Node.js e Express e expõe uma única API REST, `/getuserfilenames`, que retorna uma lista dos primeiros 10 nomes de arquivo na pasta OneDrive do usuário. O painel de tarefas usa o `getAccessToken()` método para obter um token de acesso para o usuário conectado ao servidor de camada intermediária. O servidor de camada intermediária usa o OBO (fluxo em nome de nome) para trocar o token de acesso por um novo com acesso ao Microsoft Graph. Você pode estender esse padrão para acessar todos os dados do Microsoft Graph. O painel de tarefas sempre chama uma API REST de camada intermediária (passando o token de acesso) quando precisa de serviços do Microsoft Graph. A camada intermediária usa o token obtido via OBO para chamar os serviços do Microsoft Graph e retornar os resultados para o painel de tarefas.

Este artigo funciona com um suplemento que usa Node.js e Express. Para ler um artigo semelhante sobre um suplemento baseado em ASP.NET, confira [Criar um Suplemento do Office com ASP.NET que usa o logon único](create-sso-office-add-ins-aspnet.md).

## <a name="prerequisites"></a>Pré-requisitos

- [Node.js](https://nodejs.org/) (a versão mais recente de [LTS](https://nodejs.org/about/releases))

- [Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)

- Um editor de código - recomendamos Visual Studio Code

- Pelo menos alguns arquivos e pastas armazenados em OneDrive for Business na assinatura do Microsoft 365

- Um build do Microsoft 365 que aceita o [conjunto de requisitos do IdentityAPI 1.3](/javascript/api/requirement-sets/common/identity-api-requirement-sets). Você pode obter uma [área restrita gratuita para desenvolvedores](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) que fornece uma assinatura de desenvolvedor de Microsoft 365 E5 de 90 dias renovável. A área restrita do desenvolvedor inclui uma assinatura do Microsoft Azure que você pode usar para registros de aplicativos em etapas posteriores neste artigo. Se preferir, você pode usar uma assinatura separada do Microsoft Azure para registros de aplicativo. Obtenha uma assinatura de avaliação no [Microsoft Azure](https://account.windowsazure.com/SignUp).

## <a name="set-up-the-starter-project"></a>Configure o projeto inicial

1. Clone ou baixe o repositório em [SSO com Suplemento NodeJS do Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO).

   > [!NOTE]
   > Há duas versões do exemplo:
   >
   > - A pasta **Begin** é um projeto inicial. A interface do usuário e outros aspectos do suplemento que não estão diretamente ligados ao SSO ou à autorização já estão prontos. As próximas seções deste artigo apresentam uma orientação passo a passo para concluir o projeto.
   > - A pasta **Concluir** contém o mesmo exemplo com todas as etapas de codificação deste artigo concluídas. Para usar a versão concluída, basta seguir as instruções neste artigo, mas substituir "Iniciar" por "Concluir" e ignorar as seções **Codificar o lado do cliente** e **Codificar o lado do servidor de camada intermediária** .

1. Abra um prompt de comando na pasta **Iniciar** .

1. Insira `npm install` no console para instalar todas as dependências discriminadas no arquivo package.json.

1. Execute o comando `npm run install-dev-certs`. Selecione **Sim** à solicitação para instalar o certificado.

Use os valores a seguir para espaços reservados para as etapas de registro de aplicativo subsequentes.

| Espaço reservado           | Valor                                 |
|-----------------------|---------------------------------------|
| `<add-in-name>`       | **Office-Add-in-NodeJS-SSO**          |
| `<redirect-platform>` | **Aplicativo de página única (SPA)**     |
| `<redirect-uri>`      | `https://localhost:44355/dialog.html` |

[!INCLUDE [register-sso-add-in-aad-v2-include](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="configure-the-add-in"></a>Configurar o suplemento

1. Abra a pasta `\Begin` no projeto clonado no editor de códigos.

1. Abra o `.ENV` arquivo e use os valores copiados anteriormente do registro do aplicativo **Office-Add-in-NodeJS-SSO** . Defina os valores da seguinte maneira:

   | Nome              | Valor                                                            |
   | ----------------- | ---------------------------------------------------------------- |
   | **CLIENT_ID**     | **ID do aplicativo (cliente)** da página de visão geral do registro do aplicativo. |
   | **CLIENT_SECRET** | **Segredo do cliente** salvo da página **Certificados & Segredos** .       |
   | **DIRECTORY_ID**  | **ID do diretório (locatário)** da página de visão geral do registro do aplicativo.   |

   Os valores **não** devem estar entre aspas. Quando terminar, o arquivo deverá ser semelhante ao seguinte:

   ```javascript
   CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
   CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
   DIRECTORY_ID=478aa78e-20ba-4c0d-9ffe-c4f62e5de3d5
   NODE_ENV=development
   SERVER_SOURCE=https://localhost:44355

1. Open the add-in manifest file "manifest\manifest_local.xml" and then scroll to the bottom of the file. Just above the `</VersionOverrides>` end tag, you'll find the following markup.

   ```xml
   <WebApplicationInfo>
     <Id>$app-id-guid$</Id>
     <Resource>api://localhost:44355/$app-id-guid$</Resource>
     <Scopes>
         <Scope>Files.Read</Scope>
         <Scope>profile</Scope>
         <Scope>openid</Scope>
     </Scopes>
   </WebApplicationInfo>
   ```

1. Substitua o espaço reservado "$app-id-guid$" _em ambos os lugares_ na marcação pela **ID do aplicativo** copiada quando você criou o registro do aplicativo **Office-Add-in-NodeJS-SSO** . Os símbolos "$" não fazem parte da ID, portanto, não os inclua. Essa é a mesma ID que você usou para o CLIENT_ID no . Arquivo ENV.

   > [!NOTE]
   > O **\<Resource\>** valor é o **URI da ID do Aplicativo** que você define quando registrou o suplemento. A **\<Scopes\>** seção é usada apenas para gerar uma caixa de diálogo de consentimento se o suplemento for vendido por meio do AppSource.

1. Abra o arquivo `\public\javascripts\fallback-msal\authConfig.js`. Substitua o espaço reservado "$app-id-guid$" pela ID do aplicativo que você salvou do registro de aplicativo **Office-Add-in-NodeJS-SSO** que você criou anteriormente.

1. Salve as alterações no arquivo.

## <a name="code-the-client-side"></a>Codificar o lado do cliente

### <a name="create-client-request-and-response-handler"></a>Criar manipulador de solicitação e resposta do cliente

1. No editor de códigos, abra o arquivo `public\javascripts\ssoAuthES6.js`. Ele já tem um código que garante que o Promises seja suportado, mesmo no Internet Explorer 11, e uma chamada`Office.onReady` para atribuir um manipulador para o botão somente suplemento.

   > [!NOTE]
   > Como o nome sugere, o ssoAuthES6.js usa a sintaxe JavaScript ES6, pois usar `async` e `await` mostra melhor a simplicidade fundamental da API de SSO. Quando o servidor localhost é iniciado, esse arquivo é transpilado para a sintaxe ES5 para que o exemplo dê suporte ao Internet Explorer 11.

    Uma parte fundamental do código de exemplo é a solicitação do cliente. A solicitação do cliente é um objeto que rastreia informações sobre a solicitação para chamar APIs REST no servidor de camada intermediária. É necessário porque o estado da solicitação do cliente precisa ser rastreado ou atualizado no seguinte cenário:

    - O SSO falha e a autenticação de fallback é necessária. O token de acesso é adquirido por meio do MSAL em uma caixa de diálogo pop-up. O objetivo é não falhar nesse cenário e voltar à abordagem de autenticação alternativa.

    O objeto de solicitação do cliente rastreia os seguintes dados:

    - `authSSO` – true se estiver usando o SSO, caso contrário, false.
    - `verb` - Verbo de API REST, como GET e POST.
    - `accessToken`- O token de acesso ao servidor ASP.NET Core.
    - `url`- A URL da API REST a ser chamada no servidor ASP.NET Core.
    - `callbackRESTApiHandler` - A função para passar os resultados da chamada da API REST.
    - `callbackFunction` – a função para a qual passar a solicitação do cliente quando estiver pronto.

1. Para inicializar o objeto de solicitação do cliente, na `createRequest` função, substitua `TODO 1` pelo código a seguir.

    ```javascript
    const clientRequest = {
      authSSO: authSSO,
      verb: verb,
      accessToken: null,
      url: url,
      callbackRESTApiHandler: restApiCallback,
        callbackFunction: callbackFunction,
    };
    ```

1. Substitua `TODO 2` pelo código a seguir. Sobre este código, observe:

    - Ele verifica se o SSO está sendo usado. O método para adquirir o token de acesso é diferente para SSO do que para a auth fallback.
    - Se o SSO retornar o token de acesso, ele chamará a `callbackfunction` função. Para autenticação de fallback, ele chama `dialogFallback`, que eventualmente chamará a função de retorno de chamada depois que o usuário entrar por meio do MSAL.

    ```javascript
    // Get access token.

    if (authSSO) {
    try {
      // Get access token from Office SSO.
      clientRequest.accessToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true,
      });
      callbackFunction(clientRequest);
    } catch (error) {
      // handle the SSO error which will inform us if we need to switch to fallback auth.
      let fallbackRequired = handleSSOErrors(error);
      if (fallbackRequired) switchToFallbackAuth(clientRequest);
    }
   } else {
     // Use fallback auth to get access token.
     dialogFallback(clientRequest);
   }
    ```

1. Na função `getFileNameList`, substitua `TODO 3` pelo seguinte código. Sobre este código, observe:

    - A função `getFileNameList` é chamada quando o usuário escolhe o botão **Obter Nomes de Arquivo do OneDrive** no painel de tarefas.
    - Ele cria uma solicitação de cliente para acompanhar informações sobre a chamada, como a URL da API REST.
    - Quando a API REST retorna um resultado, ela é passada para a `handleGetFileNameResponse` função. Esse retorno de chamada é passado como um parâmetro para `createRequest` e é rastreado em `clientRequest.callbackRESTApiHandler`.
    - O código chama `callWebServer` com a solicitação do cliente para executar as próximas etapas e chamar a API REST.

    ```javascript
    createRequest(
      "GET",
      "/getuserfilenames",
      handleGetFileNameResponse,
      async (clientRequest) => {
        await callWebServer(clientRequest);
      }
    );
    ```

1. Na função `handleGetFileNameResponse`, substitua `TODO 4` pelo seguinte código. Sobre este código, observe:

    - O código passa a resposta (que contém uma lista de nomes de arquivo) para `writeFileNamesToOfficeDocument` gravar os nomes de arquivo no documento.
    - O código verifica se há erros. Ele mostra uma mensagem de sucesso se os nomes de arquivo forem gravados, caso contrário, ele mostra um erro.

    ```javascript
    if (response !== null) {
      try {
        await writeFileNamesToOfficeDocument(response);
        showMessage("Your OneDrive filenames are added to the document.");
      } catch (error) {
        // The error from writeFileNamesToOfficeDocument will begin
        // "Unable to add filenames to document."
        showMessage(error);
      }
    } else
    showMessage("A null response was returned to handleGetFileNameResponse.");
    ```

1. Na função `handleSSOErrors`, substitua `TODO 5` pelo seguinte código. Para saber mais sobre esses erros, confira [Solucionar problemas de SSO em suplementos do Office em](troubleshoot-sso-in-office-add-ins.md).

    ```javascript
    let fallbackRequired = false;

   switch (err.code) {
     case 13001:
       // No one is signed into Office. If the add-in cannot be effectively used when no one
       // is logged into Office, then the first call of getAccessToken should pass the
       // `allowSignInPrompt: true` option. Since this sample does that, you should not see
       // this error.
       showMessage(
         "No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."
       );
       break;
     case 13002:
       // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
       // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
       showMessage(
         "You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."
       );
       break;
     case 13006:
       // Only seen in Office on the web.
       showMessage(
         "Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."
       );
       break;
     case 13008:
       // Only seen in Office on the web.
       showMessage(
        "Office is still working on the last operation. When it completes, try this operation again."
       );
       break;
     case 13010:
       // Only seen in Office on the web.
       showMessage(
         "Follow the instructions to change your browser's zone configuration."
       );
       break;
    ```

1. Substitua `TODO 6` pelo código a seguir. Para obter mais informações sobre esses erros, consulte [Solucionar problemas de SSO em Suplementos do Office](troubleshoot-sso-in-office-add-ins.md). Para quaisquer erros que não podem ser tratados, `true` é retornado ao chamador. Isso indica que o chamador deve mudar para o uso do MSAL como auth fallback.

    ```javascript
     default:
      // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
      // to non-SSO sign-in.
      fallbackRequired = true;
      break;
    }
    return fallbackRequired;
    ```

### <a name="call-the-rest-api-on-the-middle-tier-server"></a>Chamar a API REST no servidor de camada intermediária

1. Na função `callWebServer`, substitua `TODO 7` pelo seguinte código. Sobre este código, observe:

    - A chamada real do AJAX será feita pela `ajaxCallToRESTApi` função.
    - Essa função tentará obter um novo token de acesso se o servidor de camada média retornar um erro indicando que o token atual expirou.
    - Se a chamada AJAX não puder ser concluída com êxito, `switchToFallbackAuth` será chamada para usar a autenticação MSAL em vez do SSO do Office.

    ```javascript
    try {
    const data = await $.ajax({
      type: clientRequest.verb,
      url: clientRequest.url,
      headers: { Authorization: "Bearer " + clientRequest.accessToken },
      cache: false,
    });
    clientRequest.callbackRESTApiHandler(data);

    } catch (error) {
     // TODO 8: Check for expired SSO token and refresh if needed.

    // TODO 9: Check for Microsoft Graph and other errors.

    }
    ```

1. Substitua `TODO 8` pelo código a seguir. Sobre este código, observe:

    - Quando o servidor identifica um token expirado, ele retorna um erro com o tipo "TokenExpiredError".
    - A tentativa... Catch chamará Office.auth.getAccessToken para obter um token atualizado com uma nova expiração.
    - O código tentará chamar a API do servidor novamente.

    ```javascript
    // Check for expired SSO token. Refresh and retry the call if it expired.
    if (
      error.responseJSON &&
      authSSO === true &&
      error.responseJSON.type === "TokenExpiredError"
    ) {
      try {
        const accessToken = await Office.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true,
        });
        const data = await $.ajax({
          type: clientRequest.verb,
          url: clientRequest.url,
          headers: { Authorization: "Bearer " + accessToken },
          cache: false,
        });
        clientRequest.callbackRESTApiHandler(data);
      } catch (error) {
        showMessage(error.responseText);
        switchToFallbackAuth(clientRequest);
        return;
      }
    }
    ```

1. Substitua `TODO 9` pelo código a seguir. Sobre este código, observe:

    - Para erros do **Microsoft Graph** , mostre a mensagem no painel de tarefas.
    - Para todas as outras mensagens, mostre a mensagem no painel de tarefas.

    ```javascript
    // Check for a Microsoft Graph API call error. which is returned as bad request (403)
    if (error.status === 403) {
      if (error.responseJSON && error.responseJSON.type === "Microsoft Graph") {
        showMessage(error.responseJSON.errorDetails);
      } else {
        showMessage(error);
      }
      return;
    }

    // For all other error scenarios, display the message and use fallback auth.
    showMessage("Unknown error from web server: " + JSON.stringify(error));
    if (clientRequest.authSSO) switchToFallbackAuth(clientRequest);
    ```

A autenticação fallback usa a biblioteca MSAL para entrar no usuário. O suplemento em si é um SPA e usa um registro de aplicativo SPA para acessar o servidor de camada intermediária.

1. Na função `switchToFallbackAuth`, substitua `TODO 10` pelo seguinte código. Sobre este código, observe:

    - Ele define o global `authSSO` como false e cria uma nova solicitação de cliente que usa o MSAL para auth. A nova solicitação tem um token de acesso MSAL para o servidor de camada intermediária.
    - Depois que a solicitação é criada, ela chama `callWebServer` para continuar tentando chamar o servidor de camada intermediária com êxito.

    ```javascript
    // Guard against accidental call to this function when fallback is already in use.

    if (authSSO === false) return;

    showMessage("Switching from SSO to fallback auth.");
    authSSO = false;
    // Create a new request for fallback auth.
    createRequest(
      clientRequest.verb,
      clientRequest.url,
      clientRequest.callbackRESTApiHandler,
      async (fallbackRequest) => {
        // Hand off to call using fallback auth.
        await callWebServer(fallbackRequest);
      }
    );
    ```

## <a name="code-the-middle-tier-server"></a>Codificar o servidor de camada intermediária

O servidor de camada intermediária fornece APIs REST para o cliente chamar. Por exemplo, a API `/getuserfilenames` REST obtém uma lista de nomes de arquivo da pasta OneDrive do usuário. Cada chamada de API REST requer um token de acesso do cliente para garantir que o cliente correto esteja acessando seus dados. O token de acesso é trocado por um token do Microsoft Graph por meio do OBO (fluxo em nome de nome). O novo token do Microsoft Graph é armazenado em cache pela biblioteca MSAL para chamadas de API subsequentes. Ele nunca é enviado para fora do servidor de camada intermediária. Para obter mais informações, confira [Solicitação de token de acesso de camada média](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#middle-tier-access-token-request)

### <a name="create-the-route-and-implement-on-behalf-of-flow"></a>Criar a rota e implementar o fluxo On-Behalf-Of

1. Abra o arquivo `routes\getFilesRoute.js` e substitua `TODO 11` pelo código a seguir. Sobre este código, observe:

    - Ele chama `authHelper.validateJwt`. Isso garante que o token de acesso seja válido e não tenha sido adulterado.
    - Para obter mais informações, confira [Validando tokens](/azure/active-directory/develop/access-tokens#validating-tokens).

    ```javascript
    router.get(
     "/getuserfilenames",
     authHelper.validateJwt,
     async function (req, res) {
       // TODO 12: Exchange the access token for a Microsoft Graph token
       //          by using the OBO flow.
     }
    );
    ```

1. Substitua `TODO 12` pelo código a seguir. Sobre este código, observe:

    - Ele só solicita os escopos mínimos necessários, como `files.read`.
    - Ele usa o MSAL `authHelper` para executar o fluxo OBO na chamada para `acquireTokenOnBehalfOf`.

    ```javascript
    try {
      const authHeader = req.headers.authorization;
      let oboRequest = {
        oboAssertion: authHeader.split(" ")[1],
        scopes: ["files.read"],
      };

      // The Scope claim tells you what permissions the client application has in the service.
      // In this case we look for a scope value of access_as_user, or full access to the service as the user.
      const tokenScopes = jwt.decode(oboRequest.oboAssertion).scp.split(" ");
      const accessAsUserScope = tokenScopes.find(
        (scope) => scope === "access_as_user"
      );
      if (!accessAsUserScope) {
        res.status(401).send({ type: "Missing access_as_user" });
        return;
      }
      const cca = authHelper.getConfidentialClientApplication();
      const response = await cca.acquireTokenOnBehalfOf(oboRequest);
      // TODO 13: Call Microsoft Graph to get list of filenames.
    } catch (err) {
      // TODO 14: Handle any errors.
    }
    ```

1. Substitua `TODO 13` pelo código a seguir. Sobre este código, observe:

    - Ele constrói a URL para a chamada microsoft API do Graph e, em seguida, faz a chamada por meio da `getGraphData` função.
    - Ele retorna erros enviando uma resposta HTTP 500 junto com detalhes.
    - Com êxito, ele retorna o JSON com a lista de nome de arquivo para o cliente.

    ```javascript
    // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
    // and only the top 10 folder or file names.
    const rootUrl = "/me/drive/root/children";

    // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
    // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
    // sanitized so that it cannot be used in a Response header injection attack.
    const params = "?$select=name&$top=10";

    const graphData = await getGraphData(
      response.accessToken,
      rootUrl,
      params
    );

    // If Microsoft Graph returns an error, such as invalid or expired token,
    // there will be a code property in the returned object set to a HTTP status (e.g. 401).
    // Return it to the client. On client side it will get handled in the fail callback of `makeWebServerApiCall`.
    if (graphData.code) {
      res
        .status(403)
        .send({
          type: "Microsoft Graph",
          errorDetails:
            "An error occurred while calling the Microsoft Graph API.\n" +
            graphData,
        });
    } else {
      // MS Graph data includes OData metadata and eTags that we don't need.
      // Send only what is actually needed to the client: the item names.
      const itemNames = [];
      const oneDriveItems = graphData["value"];
      for (let item of oneDriveItems) {
        itemNames.push(item["name"]);
      }

      res.status(200).send(itemNames);
    }
    ```

1. Substitua `TODO 14` pelo código a seguir. Esse código verifica especificamente se o token expirou porque o cliente pode solicitar um novo token e chamar novamente.

   ```javascript
   // On rare occasions the SSO access token is unexpired when Office validates it,
   // but expires by the time it is used in the OBO flow. Microsoft identity platform will respond
   // with "The provided value for the 'assertion' is not valid. The assertion has expired."
   // Construct an error message to return to the client so it can refresh the SSO token.
   if (err.errorMessage.indexOf("AADSTS500133") !== -1) {
     res.status(401).send({ type: "TokenExpiredError", errorDetails: err });
   } else {
     res.status(403).send({ type: "Unknown", errorDetails: err });
   }
   ```

O exemplo deve lidar com a autenticação de fallback por meio da autenticação MSAL e SSO por meio do Office. O exemplo tentará primeiro o SSO e o `authSSO` booliano na parte superior do arquivo rastreará se o exemplo estiver usando o SSO ou tiver mudado para auth fallback.

## <a name="run-the-project"></a>Executar o projeto

1. Certifique-se de ter alguns arquivos no seu OneDrive para que você possa verificar os resultados.

1. Abra um aviso de comando na raiz da pasta `\Begin`.

1. Execute o comando `npm install` para instalar todas as dependências do pacote.

1. Execute o comando `npm start` para iniciar o servidor de camada intermediária.

1. Você deve fazer o sideload do suplemento em um aplicativo do Office (Excel, Word ou PowerPoint) para testá-lo. As instruções dependem da plataforma. Há links para instruções no [Fazer sideload de suplemento para teste](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

1. No aplicativo do Office, na faixa de opções **Home**, selecione o botão **Mostrar suplemento** no grupo **SSO Node.js** para abrir o suplemento do painel de tarefas.

1. Clique no botão **Definir Nome de Arquivos do One Drive**. Se você estiver conectado ao Office com uma conta de trabalho ou Microsoft 365 Education ou ou uma conta microsoft, e o SSO estiver funcionando conforme o esperado, os primeiros 10 nomes de arquivo e pasta em seu OneDrive for Business serão inseridos no documento. (Pode levar até 15 segundos na primeira vez.) Se você não estiver conectado ou estiver em um cenário que não dê suporte ao SSO ou o SSO não estiver funcionando por nenhum motivo, você será solicitado a entrar. Depois de entrar, os nomes do arquivo e da pasta aparecem.

> [!NOTE]
> Se você entrou no Office com uma ID diferente e se alguns aplicativos do Office que estavam abertos no momento continuam abertos, o Office pode não alterar de forma confiável sua ID, mesmo que pareça ter feito isso. Se isso acontecer, a chamada para o Microsoft Graph pode falhar ou os dados da ID anterior podem ser retornados. Para evitar isso, certifique-se de _fechar todos os outros aplicativos do Office_ antes de pressionar **Obter nomes de arquivos do OneDrive**.

## <a name="security-notes"></a>Notas de segurança

- A `/getuserfilenames` rota em `getFilesroute.js` usa uma cadeia de caracteres literal para compor a chamada para o Microsoft Graph. Se você alterar a chamada para que qualquer parte da cadeia de caracteres venha da entrada do usuário, higienize a entrada para que ela não possa ser usada em um ataque de injeção de cabeçalho de resposta.

- Na `app.js` política de segurança de conteúdo a seguir está em vigor para scripts. Talvez você queira especificar restrições adicionais dependendo de suas necessidades de segurança de suplemento.

    `"Content-Security-Policy": "script-src https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com https://alcdn.msauth.net " +  process.env.SERVER_SOURCE,`

Siga sempre as práticas recomendadas de segurança na [documentação plataforma de identidade da Microsoft](/azure/active-directory/develop/).
