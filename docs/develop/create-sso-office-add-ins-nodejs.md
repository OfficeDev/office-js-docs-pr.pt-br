---
title: Crie um Suplemento do Office com Node.js que use logon único
description: Saiba como criar um Node.js baseado em Node.js que usa o Logon Único do Office.
ms.date: 07/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6f71630f2694db9c53ba6d2e3e6d07f54ab91cb8
ms.sourcegitcommit: c62d087c27422db51f99ed7b14216c1acfda7fba
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/08/2022
ms.locfileid: "66689401"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a>Crie um Suplemento do Office com Node.js que use logon único

Os usuários podem entrar no Office, e o Suplemento Web do Office pode aproveitar esse processo de entrada para autorizá-los a acessar seu suplemento e o Microsoft Graph sem exigir que os eles entrem uma segunda vez. Para obter uma visão geral, confira o artigo [Habilitar o SSO em um Suplemento do Office](sso-in-office-add-ins.md).

Este artigo orienta você pelo processo de habilitar o SSO (logon único) em um suplemento. O suplemento de exemplo que você cria tem duas partes; um painel de tarefas que é carregado no Microsoft Excel e um servidor de camada intermediária que manipula chamadas para o Microsoft Graph para o painel de tarefas. O servidor de camada intermediária é criado com Node.js e Express e expõe uma única API REST, `/getuserfilenames`que retorna uma lista dos primeiros 10 nomes de arquivo na pasta do OneDrive do usuário. O painel de tarefas usa o `getAccessToken()` método para obter um token de acesso para o usuário conectado ao servidor de camada intermediária. O servidor de camada intermediária usa o fluxo On-Behalf-Of (OBO) para trocar o token de acesso por um novo com acesso ao Microsoft Graph. Você pode estender esse padrão para acessar todos os dados do Microsoft Graph. O painel de tarefas sempre chama uma API REST de camada intermediária (passando o token de acesso) quando precisa de serviços do Microsoft Graph. A camada intermediária usa o token obtido por meio do OBO para chamar os serviços do Microsoft Graph e retornar os resultados para o painel de tarefas.

Este artigo funciona com um suplemento que usa Node.js e Express. Para ler um artigo semelhante sobre um suplemento baseado em ASP.NET, confira [Criar um Suplemento do Office com ASP.NET que usa o logon único](create-sso-office-add-ins-aspnet.md).

## <a name="prerequisites"></a>Pré-requisitos

- [Node.js](https://nodejs.org/) (a versão mais recente de [LTS](https://nodejs.org/about/releases))

- [Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)

- Um editor de código – recomendamos Visual Studio Code

- Pelo menos alguns arquivos e pastas armazenados OneDrive for Business em sua assinatura do Microsoft 365

- Um build do Microsoft 365 que aceita o [conjunto de requisitos do IdentityAPI 1.3](/javascript/api/requirement-sets/common/identity-api-requirement-sets). Você pode obter uma [área restrita de desenvolvedor gratuita](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) que fornece uma assinatura renovável de 90 dias Microsoft 365 E5 desenvolvedor. A área restrita do desenvolvedor inclui uma assinatura do Microsoft Azure que você pode usar para registros de aplicativo em etapas posteriores neste artigo. Se preferir, você pode usar uma assinatura separada do Microsoft Azure para registros de aplicativo. Obtenha uma assinatura de avaliação [no Microsoft Azure](https://account.windowsazure.com/SignUp).

## <a name="set-up-the-starter-project"></a>Configure o projeto inicial

1. Clone ou baixe o repositório em [SSO com Suplemento NodeJS do Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO).

   > [!NOTE]
   > Há duas versões do exemplo:
   >
   > - A **pasta Begin** é um projeto inicial. A interface do usuário e outros aspectos do suplemento que não estão diretamente ligados ao SSO ou à autorização já estão prontos. As próximas seções deste artigo apresentam uma orientação passo a passo para concluir o projeto.
   > - A **pasta** Concluir contém o mesmo exemplo com todas as etapas de codificação deste artigo concluídas. Para usar a versão concluída, basta seguir as instruções neste artigo, mas substituir "Begin" por "Complete" e ignorar as seções Codificar o lado do cliente e codificar o lado do servidor de **camada intermediária.** 

1. Abra um prompt de comando na **pasta** Begin.

1. Insira `npm install` no console para instalar todas as dependências discriminadas no arquivo package.json.

1. Execute o comando `npm run install-dev-certs`. Selecione **Sim** à solicitação para instalar o certificado.

## <a name="register-the-add-in-with-microsoft-identity-platform"></a>Registrar o suplemento com o plataforma de identidade da Microsoft

Você precisa criar um registro de aplicativo no Azure que representa o servidor de camada intermediária. Isso habilita o suporte à autenticação para que os tokens de acesso adequados possam ser emitidos para o código do cliente em JavaScript. Esse registro dá suporte ao SSO no cliente e à autenticação de fallback usando a MSAL (Biblioteca de Autenticação da Microsoft).

1. Para registrar seu aplicativo, navegue até a [página portal do Azure - Registros de aplicativo](https://go.microsoft.com/fwlink/?linkid=2083908) para registrar seu aplicativo.

1. Entre com as **_credenciais de_** administrador no locatário do Microsoft 365. Por exemplo, MeuNome@contoso.onmicrosoft.com.

1. Selecione **Novo registro**. Na página **Registrar um aplicativo**, defina os valores da seguinte forma.

   - Defina **Nome** para `Office-Add-in-NodeJS-SSO`.
   - **Defina os tipos de** conta com suporte como Contas em qualquer diretório organizacional **(qualquer diretório Azure AD - Multilocatário) e contas pessoais da Microsoft (por exemplo, Skype, Xbox).**.
   - Na seção **URI de redirecionamento** , defina a plataforma como **SPA (** aplicativo de página única) com um valor de URI de redirecionamento de `https://localhost:44355/dialog.html`.
   - Escolha **Registrar**.

   > [!NOTE]
   > O tipo de aplicativo SPA só é usado quando o cliente usa a MSAL para autenticação de fallback.

1. Na página **Office-Add-in-NodeJS-SSO**, copie e salve os valores para a **ID do aplicativo (cliente)** e a **ID do diretório (locatário)**. Use ambos os valores nos procedimentos posteriores.

   > [!NOTE]
   > Essa **ID** de Aplicativo (cliente) é o valor de "público", quando outros aplicativos, como o aplicativo cliente do Office (por exemplo, PowerPoint, Word, Excel), buscam acesso autorizado ao aplicativo. Também é a "ID do cliente" do aplicativo quando ele busca acesso autorizado ao Microsoft Graph.

1. Na barra lateral mais à esquerda, selecione **Autenticação** em **Gerenciar**. Na seção **Concessão implícita e fluxos** híbridos, marque as caixas de seleção para **tokens de acesso** e **tokens de ID**. O exemplo usa a MSAL (Biblioteca de Autenticação da Microsoft) para autenticação de fallback quando o SSO não está disponível.

1. Escolha **Salvar**.

1. Em **Gerenciar**, selecione **Certificados & segredos e** selecione **Novo segredo do cliente**. Insira um valor para **Descrição** e, em seguida, selecione uma opção adequada para **Expira** e escolha **Adicionar**.

   O aplicativo Web usa o valor do segredo **do** cliente para provar sua identidade quando solicita tokens. _Registre esse valor para uso em uma etapa posterior – ele é mostrado apenas uma vez._

1. Na barra lateral mais à esquerda, selecione **Expor uma API** em **Gerenciar**. Selecione **o link** Definir. Isso gerará o URI da ID do Aplicativo no formato "api://$App ID GUID$", em que $App ID GUID$ é a ID do aplicativo **(cliente**).

1. Na ID gerada, insira `localhost:44355/` (observe a barra "/" acrescentada ao final) entre as barras duplas e o GUID. Quando terminar, a ID inteira deverá ter o formulário `api://localhost:44355/$App ID GUID$`; por exemplo `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`. Em seguida, **Salvar**.

1. Selecione o botão **Adicionar um escopo**. No painel que se abre, insira `access_as_user` como o **Nome de escopo**.

1. Definir **Quem pode consentir?** aos **Administradores e usuários**.

1. Preencha os campos para configurar os prompts de consentimento do administrador e do usuário com valores apropriados `access_as_user` para o escopo que permite que o aplicativo cliente do Office use as APIs Web do suplemento com os mesmos direitos que o usuário atual. Sugestões:

   - **Administração nome de exibição de consentimento**: o Office pode atuar como o usuário.
   - **Descrição de autorização de administrador:** Permite ao Office chamar os APIs de suplemento da web com os mesmos direitos que o usuário atual.
   - **Nome de exibição de consentimento do** usuário: o Office pode agir como você.
   - **Descrição de** consentimento do usuário: habilite o Office para chamar as APIs Web do suplemento com os mesmos direitos que você tem.

1. Verifique se o **Estado** está definido como **Habilitado**.

1. Selecione **Adicionar escopo**.

   > [!NOTE]
   > A parte de domínio do nome de **Escopo** exibidos logo abaixo do campo de texto deve corresponder automaticamente ao URI de ID do aplicativo definidos na etapa anterior com `/access_as_user` acrescentado ao final; por exemplo, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. Na seção **Aplicativos** cliente autorizados, selecione Adicionar um botão de aplicativo cliente e, em seguida, no painel que é aberto, defina **a** ID `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`do Cliente como e, em seguida, marque a caixa de seleção **Escopos autorizados** para `api://localhost:44355/$app-id-guid$/access_as_user`.

1. Selecione **Adicionar aplicativo**.

   > [!NOTE]
   > A `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` ID autoriza previamente todos os pontos de extremidade de aplicativo do Microsoft Office. Também será necessário se você quiser dar suporte a MSA (contas da Microsoft) no Office no Windows e mac. Como alternativa, você pode inserir um subconjunto adequado das IDs a seguir se, por algum motivo, quiser negar a autorização ao Office em algumas plataformas. Basta deixar de fora as IDs das plataformas das quais você deseja reprisar a autorização. Os usuários do suplemento nessas plataformas não poderão chamar suas APIs Web, mas outras funcionalidades no suplemento ainda funcionarão.
   >
   > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
   > - `93d53678-613d-4013-afc1-62e9e444a0a5`(Office na Web)
   > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3`(Outlook na Web)

1. Na barra lateral mais à esquerda, selecione **permissões de API** em **Gerenciar** e selecione **Adicionar uma permissão**. No painel que se abre, escolha **Microsoft Graph** e, em seguida, escolha **Permissões delegadas**.

1. Use a caixa de pesquisa **Selecionar permissões** para procurar as permissões que o seu suplemento precisa. Selecione estas opções. Somente o primeiro é realmente exigido pelo seu suplemento em si; mas o `profile` e as `openid` permissões são necessários para que o aplicativo do Office obtenha um token de acesso com a identidade do usuário para acessar o servidor de camada intermediária.

   - **Files.Read**
   - **perfil**
   - **openid**

   > [!NOTE]
   > A permissão `User.Read` pode já estar listada por padrão. É uma boa prática não solicitar permissões que não são necessárias, portanto, recomendamos que você desmarque a caixa dessa permissão se o suplemento não precisar dela.

1. Marque a caixa de seleção para cada permissão conforme elas forem exibidas. Depois de selecionar as permissões que o suplemento precisa, selecione o botão **Adicionar permissões** na parte inferior do painel.

1. Na mesma página, escolha o botão **conceder permissão de administrador para [nome do locatário]** e, em seguida, selecione **Sim** para a confirmação exibida.

## <a name="configure-the-add-in"></a>Configurar o suplemento

1. Abra a pasta `\Begin` no projeto clonado no editor de códigos.

1. Abra o `.ENV` arquivo e use os valores copiados anteriormente do registro do aplicativo **Office-Add-in-NodeJS-SSO** . Defina os valores da seguinte maneira:

   | Nome              | Valor                                                            |
   | ----------------- | ---------------------------------------------------------------- |
   | **CLIENT_ID**     | **ID do aplicativo (cliente) na** página de visão geral do registro do aplicativo. |
   | **CLIENT_SECRET** | **Segredo do cliente** salvo da **página Certificados & Segredos** .       |
   | **DIRECTORY_ID**  | **ID do diretório (locatário)** na página de visão geral do registro do aplicativo.   |

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

1. Substitua o espaço reservado "$app-id-guid _$" em_ ambos os locais na marcação pela **ID** do Aplicativo que você copiou quando criou o registro de aplicativo **Office-Add-in-NodeJS-SSO** . Os símbolos "$" não fazem parte da ID, portanto, não os inclua. Esta é a mesma ID que você usou para o CLIENT_ID no . Arquivo ENV.

   > [!NOTE]
   > O **\<Resource\>** valor é o **URI da ID** do Aplicativo que você definiu quando registrou o suplemento. A **\<Scopes\>** seção é usada apenas para gerar uma caixa de diálogo de consentimento se o suplemento for vendido por meio do AppSource.

1. Abra o arquivo `\public\javascripts\fallback-msal\authConfig.js`. Substitua o espaço reservado "$app-id-guid$" pela ID do aplicativo que você salvou do registro de aplicativo **Office-Add-in-NodeJS-SSO** criado anteriormente.

1. Salve as alterações no arquivo.

## <a name="code-the-client-side"></a>Codificar o lado do cliente

### <a name="create-client-request-and-response-handler"></a>Criar manipulador de solicitação e resposta do cliente

1. No editor de códigos, abra o arquivo `public\javascripts\ssoAuthES6.js`. Ele já tem um código que garante que o Promises seja suportado, mesmo no Internet Explorer 11, e uma chamada`Office.onReady` para atribuir um manipulador para o botão somente suplemento.

   > [!NOTE]
   > Como o nome sugere, o ssoAuthES6.js usa a sintaxe JavaScript ES6, pois usar `async` e `await` mostra melhor a simplicidade fundamental da API de SSO. Quando o servidor localhost é iniciado, esse arquivo é transcompilado para a sintaxe ES5 para que o exemplo seja compatível com o Internet Explorer 11.

    Uma parte importante do código de exemplo é a solicitação do cliente. A solicitação do cliente é um objeto que rastreia informações sobre a solicitação para chamar APIs REST no servidor de camada intermediária. Isso é necessário porque o estado da solicitação do cliente precisa ser rastreado ou atualizado por meio dos seguintes cenários:

    - O SSO tenta novamente quando a chamada à API REST falha porque precisa de consentimento adicional. O código de exemplo chama `getAccessToken` com opções de autenticação atualizadas, obtém o consentimento do usuário necessário e chama a API REST novamente. A meta é não falhar em cenários em que uma API REST precisa de consentimento adicional.
    - O SSO falha e a autenticação de fallback é necessária. O token de acesso é adquirido por meio da MSAL em uma caixa de diálogo pop-up. A meta é não falhar nesse cenário e retornar normalmente para a abordagem de autenticação alternativa.

    O objeto de solicitação do cliente rastreia os seguintes dados:

    - `authOptions` - [Parâmetros de configuração de autenticação](/javascript/api/office/office.authoptions) para SSO.
    - `authSSO` – true se estiver usando SSO; caso contrário, false.
    - `accessToken` – O token de acesso para o servidor de camada intermediária. O método para obter esse token é diferente para SSO do que a autenticação de fallback.
    - `url` - A URL da API REST a ser chamada no servidor de camada intermediária.
    - `callbackHandler` - A função para passar os resultados da chamada à API REST.
    - `callbackFunction` - A função para a qual passar a solicitação do cliente quando estiver pronto.

1. Para inicializar o objeto de solicitação do cliente, na função `createRequest` , substitua `TODO 1` pelo código a seguir.

   ```javascript
   const clientRequest = {
     authOptions: {
       allowSignInPrompt: true,
       allowConsentPrompt: true,
       forMSGraphAccess: true,
     },
     authSSO: authSSO,
     accessToken: null,
     url: url,
     callbackRESTApiHandler: restApiCallback,
     callbackFunction: callbackFunction,
   };
   ```

1. Substitua `TODO 2` pelo código a seguir. Sobre este código, observe:

   - Ele verifica se o SSO está sendo usado. O método para adquirir o token de acesso é diferente para SSO do que para autenticação de fallback.
   - Se o SSO retornar o token de acesso, ele chamará a `callbackfunction` função. Para autenticação de fallback, `dialogFallback`ele chama, que eventualmente chamará a função de retorno de chamada depois que o usuário entrar por meio da MSAL.

   ```javascript
   // Get access token.

   if (authSSO) {
     try {
       // Get access token from Office SSO.
       clientRequest.accessToken = await getAccessTokenFromSSO(
         clientRequest.authOptions
       );
       callbackFunction(clientRequest);
     } catch {
       // Use fallback authentication if SSO failed to get access token.
       switchToFallbackAuth(clientRequest);
     }
   } else {
     // Use fallback authentication to get access token.
     dialogFallback(clientRequest);
   }
   ```

1. Na função `getFileNameList`, substitua `TODO 3` pelo seguinte código. Sobre este código, observe:

   - A função `getFileNameList` é chamada quando o usuário escolhe o **botão Obter Nomes de Arquivo do OneDrive** no painel de tarefas.
   - Ele cria uma solicitação de cliente para controlar informações sobre a chamada, como a URL da API REST.
   - Quando a API REST retorna um resultado, ele é passado para a `handleGetFileNameResponse` função. Esse retorno de chamada é passado como um parâmetro para `createRequest` e é rastreado em `clientRequest.callbackRESTApiHandler`.
   - O código chama com `callWebServer` a solicitação do cliente para executar as próximas etapas e chamar a API REST.

   ```javascript
   createRequest(
     "/getuserfilenames",
     handleGetFileNameResponse,
     async (clientRequest) => {
       await callWebServer(clientRequest);
     }
   );
   ```

1. Na função `handleGetFileNameResponse`, substitua `TODO 4` pelo seguinte código. Sobre este código, observe:

   - O código passa a resposta (que contém uma lista de nomes de arquivo) `writeFileNamesToOfficeDocument` para gravar os nomes de arquivo no documento.
   - O código verifica se há erros. Ele mostrará uma mensagem de êxito se os nomes de arquivo forem gravados; caso contrário, ele mostrará um erro.

   ```javascript
   if (response != null) {
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

### <a name="get-the-sso-access-token"></a>Obter o token de acesso de SSO

1. Na função `getAccessTokenFromSSO`, substitua `TODO 5` pelo seguinte código. Sobre este código, observe:

   - Ele chama `Office.auth.getAccessToken` para obter o token de acesso do Office.
   - Se ocorrer um erro, ele chamará a `handleSSOErrors` função. Se o erro não puder ser tratado, ele gerará um erro para o chamador. Essa é a indicação para o chamador alternar para autenticação de fallback.

   ```javascript
   try {
     // The access token returned from getAccessToken only has permissions to your middle-tier server APIs,
     // and it contains the identity claims of the signed-in user.

     const accessToken = await Office.auth.getAccessToken(authOptions);
     return accessToken;
   } catch (error) {
     let fallbackRequired = handleSSOErrors(error);
     if (fallbackRequired) throw error; // Rethrow the error and caller will switch to fallback auth.
     return null; // Returning a null token indicates no need for fallback (an explanation about the error condition was shown by handleSSOErrors).
   }
   ```

1. Na função `handleSSOErrors`, substitua `TODO 6` pelo seguinte código. Para saber mais sobre esses erros, confira [Solucionar problemas de SSO em suplementos do Office em](troubleshoot-sso-in-office-add-ins.md).

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

1. Substitua `TODO 7` pelo código a seguir. Para obter mais informações sobre esses erros, consulte [Solucionar problemas de SSO em Suplementos do Office](troubleshoot-sso-in-office-add-ins.md). Para quaisquer erros que não possam ser tratados, `true` é retornado ao chamador. Isso indica que o chamador deve alternar para usar a MSAL como autenticação de fallback.

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

1. Na função `callWebServer`, substitua `TODO 8` pelo seguinte código. Sobre este código, observe:

   - A chamada AJAX real será feita pela `ajaxCallToRESTApi` função.
   - Essa função tentará obter um novo token de acesso se o servidor de camada intermediária retornar um erro indicando que o token atual expirou.
   - Se a chamada do AJAX não puder ser concluída com êxito, `switchToFallbackAuth` será chamada para usar a autenticação MSAL em vez do SSO do Office.

   ```javascript
   try {
     await ajaxCallToRESTApi(clientRequest);
   } catch (error) {
     if (error.statusText === "Internal Server Error") {
       const retryCall = handleWebServerErrors(error, clientRequest);
       if (retryCall && clientRequest.authSSO) {
         try {
           clientRequest.accessToken = await getAccessTokenFromSSO(
             clientRequest.authOptions
           );
           await ajaxCallToRESTApi(clientRequest);
         } catch {
           // If still an error go to fallback.
           switchToFallbackAuth(clientRequest);
           return;
         }
       }
     } else {
       console.log(JSON.stringify(error)); // Log any errors.
       showMessage(error.responseText);
     }
   }
   ```

1. Na função `ajaxCallToRESTApi`, substitua `TODO 9` pelo seguinte código. Sobre este código, observe:

   - A função relança explicitamente quaisquer erros para o chamador manipular.

   ```javascript
   try {
     await $.ajax({
       type: "GET",
       url: clientRequest.url,
       headers: { Authorization: "Bearer " + clientRequest.accessToken },
       cache: false,
       success: function (data) {
         result = data;
         // Send result to the callback handler.
         clientRequest.callbackRESTApiHandler(result);
       },
     });
   } catch (error) {
     // This function explicitly requires the caller to handle any errors
     throw error;
   }
   ```

1. Na função `handleWebServerErrors`, substitua `TODO 10` pelo seguinte código. Sobre este código, observe:

   - O erro é retornado pelo servidor de camada intermediária, que indica o tipo de erro e facilita o tratamento aqui.
   - Para **erros do Microsoft Graph** , mostre a mensagem no painel de tarefas.
   - Para o **erro AADSTS500133** , retorne true para que o chamador saiba que o token expirou e deve obter um novo.
   - Para todas as outras mensagens, mostre a mensagem no painel de tarefas.

   ```javascript
   let retryCall = false;
   // Our middle-tier server returns a type to help handle the known cases.
   switch (err.responseJSON.type) {
     case "Microsoft Graph":
       // An error occurred when the middle-tier server called Microsoft Graph.
       showMessage(
         "Error from Microsoft Graph: " +
           JSON.stringify(err.responseJSON.errorDetails)
       );
       retryCall = false;
       break;
     case "Missing access_as_user":
       // The access_as_user scope was missing.
       showMessage("Error: Access token is missing the access_as_user scope.");
       retryCall = false;
       break;
     case "AADSTS500133": // expired token
       // On rare occasions the access token could expire after it was sent to the middle-tier server.
       // Microsoft identity platform will respond with
       // "The provided value for the 'assertion' is not valid. The assertion has expired."
       // Return true to indicate to caller they should refresh the token.
       retryCall = true;
       break;
     default:
       showMessage(
         "Unknown error from web server: " +
           JSON.stringify(err.responseJSON.errorDetails)
       );
       retryCall = false;
       if (clientRequest.authSSO) switchToFallbackAuth(clientRequest);
   }
   return retryCall;
   ```

A autenticação de fallback usará a biblioteca MSAL para conectar o usuário. O suplemento em si é um SPA e usa um registro de aplicativo SPA para acessar o servidor de camada intermediária.

1. Na função `switchToFallbackAuth`, substitua `TODO 11` pelo seguinte código. Sobre este código, observe:

   - Ele define o global `authSSO` como false e cria uma nova solicitação de cliente que usa a MSAL para autenticação. A nova solicitação tem um token de acesso MSAL para o servidor de camada intermediária.
   - Depois que a solicitação é criada, ela chama `callWebServer` para continuar tentando chamar o servidor de camada intermediária com êxito.

   ```javascript
   showMessage("Switching from SSO to fallback auth.");
   authSSO = false;
   // Create a new request for fallback auth.
   createRequest(
     clientRequest.url,
     clientRequest.callbackRESTApiHandler,
     async (fallbackRequest) => {
       // Hand off to call using fallback auth.
       await callWebServer(fallbackRequest);
     }
   );
   ```

## <a name="code-the-middle-tier-server"></a>Codificar o servidor de camada intermediária

O servidor de camada intermediária fornece APIs REST para o cliente chamar. Por exemplo, a API `/getuserfilenames` REST obtém uma lista de nomes de arquivo da pasta do OneDrive do usuário. Cada chamada à API REST requer um token de acesso pelo cliente para garantir que o cliente correto esteja acessando seus dados. O token de acesso é trocado por um token do Microsoft Graph por meio do fluxo On-Behalf-Of (OBO). O novo token do Microsoft Graph é armazenado em cache pela biblioteca MSAL para chamadas à API subsequentes. Ele nunca é enviado fora do servidor de camada intermediária. Para obter mais informações, consulte [a solicitação de token de acesso de camada intermediária](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#middle-tier-access-token-request)

### <a name="create-the-route-and-implement-on-behalf-of-flow"></a>Criar a rota e implementar o fluxo On-Behalf-Of

1. Abra o arquivo `routes\getFilesRoute.js` e substitua `TODO 12` pelo código a seguir. Sobre este código, observe:

   - Ele chama `authHelper.validateJwt`. Isso garante que o token de acesso seja válido e não tenha sido adulterado.
   - Para obter mais informações, consulte [Validando tokens](/azure/active-directory/develop/access-tokens#validating-tokens).

   ```javascript
   router.get(
     "/getuserfilenames",
     authHelper.validateJwt,
     async function (req, res) {
       // TODO 13: Exchange the access token for a Microsoft Graph token
       //          by using the OBO flow.
     }
   );
   ```

1. Substitua `TODO 13` pelo código a seguir. Sobre este código, observe:

   - Ele solicita apenas os escopos mínimos necessários, como `files.read`.
   - Ele usa a MSAL `authHelper` para executar o fluxo OBO na chamada para `acquireTokenOnBehalfOf`.

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
     // TODO 14: Call Microsoft Graph to get list of filenames.
   } catch (err) {
     // TODO 15: Handle any errors.
   }
   ```

1. Substitua `TODO 14` pelo código a seguir. Sobre este código, observe:

   - Ele constrói a URL para a chamada API do Graph Microsoft e, em seguida, faz a chamada por meio da `getGraphData` função.
   - Ele retorna erros enviando uma resposta HTTP 500 juntamente com detalhes.
   - Em caso de êxito, ele retorna o JSON com a lista de nomes de arquivo para o cliente.

   ```javascript
   // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
   // and only the top 10 folder or file names.
   const rootUrl = "/me/drive/root/children";

   // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
   // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
   // sanitized so that it cannot be used in a Response header injection attack.
   const params = "?$select=name&$top=10";

   const graphData = await getGraphData(response.accessToken, rootUrl, params);

   // If Microsoft Graph returns an error, such as invalid or expired token,
   // there will be a code property in the returned object set to a HTTP status (e.g. 401).
   // Return it to the client. On client side it will get handled in the fail callback of `makeWebServerApiCall`.
   if (graphData.code) {
     res.status(500).send({ type: "Microsoft Graph", errorDetails: graphData });
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

1. Substitua `TODO 15` pelo código a seguir. Esse código verifica especificamente se o token expirou porque o cliente pode solicitar um novo token e chamar novamente.

   ```javascript
   // On rare occasions the SSO access token is unexpired when Office validates it,
   // but expires by the time it is used in the OBO flow. Microsoft identity platform will respond
   // with "The provided value for the 'assertion' is not valid. The assertion has expired."
   // Construct an error message to return to the client so it can refresh the SSO token.
   if (err.errorMessage.indexOf("AADSTS500133") !== -1) {
     res.status(500).send({ type: "AADSTS500133", errorDetails: err });
   } else {
     res.status(500).send({ type: "Unknown", errorDetails: err });
   }
   ```

O exemplo deve lidar com a autenticação de fallback por meio da autenticação MSAL e SSO por meio do Office. O exemplo tentará primeiro o SSO e `authSSO` o booliano na parte superior do arquivo rastreia se o exemplo estiver usando SSO ou tiver mudado para autenticação de fallback.

## <a name="run-the-project"></a>Executar o projeto

1. Certifique-se de ter alguns arquivos no seu OneDrive para que você possa verificar os resultados.

1. Abra um aviso de comando na raiz da pasta `\Begin`.

1. Execute o comando para `npm install` instalar todas as dependências do pacote.

1. Execute o comando para `npm start` iniciar o servidor de camada intermediária.

1. Você deve fazer o sideload do suplemento em um aplicativo do Office (Excel, Word ou PowerPoint) para testá-lo. As instruções dependem da plataforma. Há links para instruções no [Fazer sideload de suplemento para teste](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).

1. No aplicativo do Office, na faixa de opções **Home**, selecione o botão **Mostrar suplemento** no grupo **SSO Node.js** para abrir o suplemento do painel de tarefas.

1. Clique no botão **Definir Nome de Arquivos do One Drive**. Se você estiver conectado ao Office com uma conta corporativa ou Microsoft 365 Education ou uma conta da Microsoft, e o SSO estiver funcionando conforme o esperado, os 10 primeiros nomes de arquivo e pasta no OneDrive for Business serão inseridos no documento. (Pode levar até 15 segundos na primeira vez.) Se você não estiver conectado ou estiver em um cenário que não dá suporte ao SSO ou que o SSO não esteja funcionando por nenhum motivo, você será solicitado a entrar. Depois de entrar, os nomes de arquivo e pasta são exibidos.

> [!NOTE]
> Se você entrou no Office com uma ID diferente e se alguns aplicativos do Office que estavam abertos no momento continuam abertos, o Office pode não alterar de forma confiável sua ID, mesmo que pareça ter feito isso. Se isso acontecer, a chamada para o Microsoft Graph pode falhar ou os dados da ID anterior podem ser retornados. Para evitar isso, certifique-se de _fechar todos os outros aplicativos do Office_ antes de pressionar **Obter nomes de arquivos do OneDrive**.

## <a name="security-notes"></a>Notas de segurança

* A `/getuserfilenames` rota em `getFilesroute.js` usa uma cadeia de caracteres literal para compor a chamada para o Microsoft Graph. Se você alterar a chamada para que qualquer parte da cadeia de caracteres venha da entrada do usuário, desaitize a entrada para que ela não possa ser usada em um ataque de injeção de cabeçalho de resposta.

* Na `app.js` política de segurança de conteúdo a seguir está em vigor para scripts. Talvez você queira especificar restrições adicionais dependendo das suas necessidades de segurança do suplemento.

    `"Content-Security-Policy": "script-src https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com https://alcdn.msauth.net " +  process.env.SERVER_SOURCE,`

Sempre siga as práticas recomendadas de segurança [na documentação plataforma de identidade da Microsoft segurança](/azure/active-directory/develop/).
