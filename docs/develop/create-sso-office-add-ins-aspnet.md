---
title: Criar um Suplemento do Office com ASP.NET que use logon único
description: ''
ms.date: 12/04/2019
localization_priority: Priority
ms.openlocfilehash: a5a05c8299f22bbc1a7593eebc9bde60f52b2b84
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/17/2020
ms.locfileid: "41216999"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a>Criar um Suplemento do Office com ASP.NET que use logon único (visualização)

Quando os usuários estão conectados ao Office, o seu suplemento pode usar as mesmas credenciais para permitir que os usuários acessem vários aplicativos sem exigir que eles entrem uma segunda vez. Confira uma visão geral no artigo [Habilitar o SSO em um Suplemento do Office](sso-in-office-add-ins.md).
Este artigo apresenta o processo passo a passo de habilitação do logon único (SSO) em um suplemento que foi criado com Node.js e Express.

> [!NOTE]
> Para ler um artigo semelhante sobre um suplemento baseado em ASP.NET, confira [Criar um Suplemento do Office com Node.js que usa o logon único](create-sso-office-add-ins-nodejs.md).

## <a name="prerequisites"></a>Pré-requisitos

* Visual Studio 2019 ou posterior.

* [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* Em pelo menos algumas pastas e arquivos armazenados no OneDrive for Business na assinatura do Office 365.

* Uma assinatura do Microsoft Azure. Este suplemento requer o Azure Active Directory (AD). O Active AD fornece serviços de identidade que os aplicativos usam para autenticação e autorização. Você pode adquirir uma assinatura de avaliação no [Microsoft Azure](https://account.windowsazure.com/SignUp).

## <a name="set-up-the-starter-project"></a>Configure o projeto inicial

Clone ou baixe o repositório em [SSO com Suplemento ASPNET do Office](https://github.com/officedev/office-add-in-aspnet-sso).

> [!NOTE]
> Há duas versões do exemplo:
>
> * A pasta **Before** (antes) traz um projeto inicial. A interface do usuário e outros aspectos do suplemento que não estão diretamente ligados ao SSO ou à autorização já estão prontos. As próximas seções deste artigo apresentam uma orientação passo a passo para concluir o projeto.
> * A versão **Complete** (concluído) do exemplo apresenta como seria o suplemento quando concluídos os procedimentos apresentados neste artigo, com exceção de que o projeto concluído traz comentários de códigos que seriam redundantes neste artigo. Para usar a versão concluída, apenas siga as instruções apresentadas neste artigo, substituindo "Before" por "Complete" e pulando as seções **Codificar o lado do cliente** e **Codificar o lado do servidor**.


## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Registre o suplemento com o ponto de extremidade v2.0 do Azure AD

1. Acesse a página [Portal do Azure - Registros de aplicativo](https://go.microsoft.com/fwlink/?linkid=2083908) para registrar o seu aplicativo.

1. Entre com as credenciais de ***administrador*** em sua locação do Office 365. Por exemplo, MeuNome@contoso.onmicrosoft.com.

1. Selecione **Novo registro**. Na página **Registrar um aplicativo**, defina os valores da seguinte forma.

    * Defina **Nome** para `Office-Add-in-ASPNET-SSO`.
    * Defina **Tipos de conta com suporte** para **Contas em qualquer diretório organizacional (Qualquer diretório do Azure AD – Multilocatário) e contas pessoais da Microsoft (por exemplo, Skype, Xbox)**. (Se você quiser que o suplemento possa ser usado somente por usuários no locatário em que você está os registrando, escolha **Contas somente neste diretório organizacional...**, mas execute algumas etapas adicionais. Confira **Configuração para locatário único** abaixo.)
    * Na seção **URI de redirecionamento**, verifique se **Web** está selecionado no menu suspenso e defina o URI como ` https://localhost:44355/AzureADAuth/Authorize`.
    * Escolha **Registrar**.

1. Na página **Office-Add-in-NodeJS-SSO**, copie e salve os valores para a **ID do aplicativo (cliente)** e a **ID do diretório (locatário)**. Use ambos os valores nos procedimentos posteriores.

    > [!NOTE]
    > Essa ID é o valor "audience" (público) quando outros aplicativos, como o aplicativo host do Office (por exemplo, PowerPoint, Word, Excel), buscam o acesso autorizado ao aplicativo. Também é a "ID do cliente" do aplicativo quando ela, por sua vez, busca acesso autorizado ao Microsoft Graph.

1. Em **Gerenciar**, selecione **Certificados e segredos**. Selecione o botão **Novo segredo do cliente**. Insira um valor para **Descrição** e, em seguida, selecione uma opção adequada para **Expira** e escolha **Adicionar**. *Copiar o valor de segredo do cliente imediatamente e salvá-lo com a ID de aplicativo* antes de prosseguir, pois ele será necessário em um procedimento posterior.

1. Em **Gerenciar**, selecione **Expor uma API**. Selecione o link **Definir** para gerar o URI da ID de Aplicativo no formato "api: / / $App ID GUID$", em que $App ID GUID$ é **ID do aplicativo (cliente)**. Insira `localhost:44355/` (Observe a barra "/" anexada ao fim) após o `//` e antes do GUID. A ID inteira deve ter o formulário `api://localhost:44355/$App ID GUID$`; por exemplo `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

1. Marque **Salvar** na caixa de diálogo.

1. Selecione o botão **Adicionar um escopo**. No painel que se abre, insira `access_as_user` como o **Nome de escopo**.

1. Definir **Quem pode consentir?** aos **Administradores e usuários**.

1. Preencha os campos para configurar a solicitação de consentimento de administrador e usuário com valores apropriados ao `access_as_user` escopo que permite que o aplicativo de host do Office use os seus APIs de suplemento da web com os mesmos direitos que o usuário atual. Sugestões:

    - **Título de autorização de administrador:** Office pode funcionar como o usuário.
    - **Descrição de autorização de administrador:** Permite ao Office chamar os APIs de suplemento da web com os mesmos direitos que o usuário atual.
    - **Título de autorização de usuário:** O Office pode funcionar como se fosse você.
    - **Descrição de autorização de administrador:** Permite ao Office chamar os APIs de suplemento da web com os mesmos direitos que você possui.

1. Verifique se o **Estado** está definido como **Habilitado**.

1. Selecione **Adicionar escopo**.

    > [!NOTE]
    > A parte de domínio do nome de **Escopo** exibidos logo abaixo do campo de texto deve corresponder automaticamente ao URI de ID do aplicativo definidos na etapa anterior com `/access_as_user` acrescentado ao final; por exemplo, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. Na seção **Aplicativos clientes autorizados**, você identifica os aplicativos que deseja autorizar para o aplicativo da Web do seu suplemento. Cada uma das seguintes IDs precisa ser pré-autorizada.

    - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
    - `57fb890c-0dab-4253-a5e0-7188c88b2bb4`(Office na Web)
    - `bc59ab01-8403-45c6-8796-ac3ef710b3e3`(Office na Web)

    Para cada ID, siga estas etapas:

    a. Selecione o botão **Adicionar um aplicativo cliente** e, no painel que se abre, defina o ID do cliente para o respectivo GUID e marque a caixa `api://localhost:44355/$App ID GUID$/access_as_user`.

    b. Selecione **Adicionar aplicativo**.

1. Em **Gerenciar**, selecione **Permissões para API** e selecione **Adicionar uma permissão**. No painel que se abre, escolha **Microsoft Graph** e, em seguida, escolha **Permissões delegadas**.

1. Use a caixa de pesquisa **Selecionar permissões** para procurar as permissões que o seu suplemento precisa. Selecione estas opções. Somente a primeira permissão é realmente necessária pelo suplemento em si, mas a permissão `profile` é necessária para que o host do Office obtenha um token no aplicativo Web do seu suplemento. (Somente Files.Read.All e o perfil são, de fato, necessários para o suplemento. Solicite os outros dois porque a biblioteca MSAL.NET exige.)

    * Files.Read.All
    * offline_access
    * openid
    * perfil

    > [!NOTE]
    > A permissão `User.Read` pode já estar listada por padrão. É uma boa prática não pedir permissões desnecessárias, por isso recomendamos desmarcar a caixa para essa permissão se o suplemento não precisar dela.

1. Marque a caixa de seleção para cada permissão conforme elas forem exibidas. Depois de selecionar as permissões que o suplemento precisa, selecione o botão **Adicionar permissões** na parte inferior do painel.

1. Na mesma página, escolha o botão **conceder permissão de administrador para [nome do locatário]** e, em seguida, selecione **Aceitar** para a confirmação exibida.

    > [!NOTE]
    > Depois de escolher **Conceder consentimento de administrador para [nome do locatário]**, você verá uma mensagem solicitando que você tente novamente alguns minutos depois, para que a solicitação de consentimento possa ser construída. Em caso afirmativo, você pode começar a trabalhar na próxima seção, ***mas não se esqueça de voltar para o portal e pressionar este botão***!

## <a name="configure-the-solution"></a>Configurar a solução

1. Na raiz da pasta **Before** (antes), abra o arquivo de solução (.sln) no **Visual Studio**. Clique com o botão direito do mouse no nó superior no **Gerenciador de Soluções** (no nó Solução, não em qualquer um dos nós do projeto) e selecione **Configurar projetos de inicialização**.

1. Em **Propriedades Comuns**, selecione **Projeto de Inicialização** e **Vários projetos de inicialização**. Verifique se a **Ação** para ambos os projetos está definida como **Iniciar** e se o projeto terminado em "...WebAPI" está listado primeiro. Feche a caixa de diálogo.

1. No **Gerenciador de Soluções**, selecione (não clique com o botão direito) o projeto **Office-Add-in-Microsoft-Graph-ASPNETWebAPI**. O painel **Propriedades** é exibido. Verifique se **SSL Habilitado** é **Verdadeiro**. Verifique se a **URL do SSL** é `http://localhost:44355/`.

1. Em "Web.config", use os valores copiados anteriormente. Defina **ida:ClientID** e **ida:Audience** para sua **ID do aplicativo (cliente)** e defina **ida:Password** para a senha de cliente.

    > [!NOTE]
    > A **ID do aplicativo (cliente)** é o valor "audience" (público) quando outros aplicativos, como o aplicativo host do Office (por exemplo, PowerPoint, Word, Excel), buscam o acesso autorizado ao aplicativo. Também é a "ID do cliente" do aplicativo quando ela, por sua vez, busca acesso autorizado ao Microsoft Graph.

1. Se você não tiver escolhido "Somente contas neste diretório organizacional" para **TIPOS DE CONTA COM SUPORTE** ao registrar o suplemento, salve e feche o Web.config. Caso contrário, salve, mas deixe-o aberto.

1. Ainda no **Gerenciador de Soluções**, escolha o projeto **Office-Add-in-Microsoft-Graph-ASPNET** e abra o arquivo de manifesto do suplemento "Office-Add-in-ASPNET-SSO.xml" e role até a parte inferior do arquivo. Logo acima da marca de fim `</VersionOverrides>`, você encontrará a marcação a seguir:

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. Substitua o espaço reservado "{$application_GUID here$}" *nos dois lugares* na marcação pela ID do Aplicativo que você copiou ao registrar seu suplemento. Os sinais "$" não fazem parte da ID, portanto não os inclua. Essa é a mesma ID usada para a ClientID e a Audience no web.config.

  > [!NOTE]
  > O valor **Recurso** é o**URI da ID de aplicativo** que você definiu quando registrou o suplemento. A seção **Scopes** só será usada para gerar uma caixa de diálogo de consentimento se o suplemento for vendido no AppSource.

1. Salve e feche o arquivo.

### <a name="setup-for-single-tenant"></a>Configuração para locatário único

Se você escolher "Somente contas neste diretório organizacional" para **TIPOS DE CONTA COM SUPORTE** ao registrar o suplemento, você execute estas etapas adicionais de configuração:

1. Volte para o Portal do Azure e abra a lâmina **Visão geral** do registro do suplemento. Copie a **ID de diretório (locatário)**.

1. Em Web.config, substitua o "comum" no valor de **ida:Authority** pela GUID copiada na etapa anterior. Ao terminar, o valor deverá ser similar a este: `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`.

1. Salve e feche o web.config.

## <a name="code-the-client-side"></a>Codificar o lado do cliente

1. Abra o arquivo HomeES6.js na pasta **Scripts**. Ele já apresenta alguns códigos:

    * Um polyfill que atribui o objeto Office.Promise ao objeto de janela global, para que o suplemento possa ser executado quando o Office estiver usando o Internet Explorer para a interface de usuário. (Para obter mais detalhes, confira [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).)
    * Uma atribuição ao método `Office.initialize` que, por sua vez, atribui um manipulador ao evento clicar do botão `getGraphAccessTokenButton`.
    * Um método `showResult` que exibirá os dados retornados do Microsoft Graph (ou uma mensagem de erro) na parte inferior do painel de tarefas.
    * Um método `logErrors` que registrará erros de console que não são destinados ao usuário final.
    * O código implementa o sistema de autorização de fallback que o suplemento usará em situações em que o SSO não é compatível ou gera um erro.

1. Abaixo da atribuição a `Office.initialize`, adicione o código a seguir. Observe o seguinte sobre este código:

    * O processamento de erros no suplemento às vezes tentará novamente obter um token de acesso automaticamente, usando um conjunto diferente de opções. A variável de contador `retryGetAccessToken` é usada para garantir que o usuário não seja trocado repetidas vezes em tentativas falhas de obter um token.
    * A função `getGraphData` é definida com a palavra-chave ES6 `async`. Usar a sintaxe ES6 facilita o uso da API de SSO em Suplementos do Office. Esse é o único arquivo na solução que usará a sintaxe sem suporte do Internet Explorer. Colocamos "ES6" no nome do arquivo como um lembrete. A solução usa o transcompilador de tsc para transcompilar esse arquivo em ES5, para que o suplemento possa ser executado quando o Office estiver usando o Internet Explorer para a interface do usuário. (Veja o arquivo tsconfig.json na raiz do projeto.)

    ```javascript
    var retryGetAccessToken = 0;

    async function getGraphData() {
        await getDataWithToken({ allowSignInPrompt: true, forMSGraphAccess: true });
    }
    ```

1. Abaixo da função `getGraphData`, adicione a função a seguir. Observe que você criará a função `handleClientSideErrors` em uma etapa posterior.

    ```javascript
    async function getDataWithToken() {
        try {

            // TODO 1: Get the bootstrap token and send it to the server to exchange
            //         for an access token to Microsoft Graphn and then get the data
            //         from Microsoft Graph.

        }
        catch (exception) {
            if (exception.code) {
                handleClientSideErrors(exception);
            }
            else {
                showResult(["EXCEPTION: " + JSON.stringify(exception)]);
            }
        }
    }
    ```

1. Substitua `TODO 1` pelo seguinte. Sobre este código, observe:

    * `getAccessToken` instrui o Office a obter um token de bootstrap do Azure AD e retornar ao suplemento.
    * `allowSignInPrompt` indica ao Office para solicitar que o usuário entre caso ele ainda não esteja conectado ao Office.
    * `forMSGraphAccess` instrui o Office que o suplemento pretende trocar o token de bootstrap por um token de acesso ao Micrsoft Graph, em vez de apenas usar o token de bootstrap como um token de ID. A configuração dessa opção dá ao Office a oportunidade de cancelar o processo de obtenção do token de bootstrap (e retornar o código de erro 13012) se o administrador de locatários do usuário não tiver concedido consentimento para o suplemento. O código do lado do cliente do suplemento pode responder ao 13012 por meio da ramificação para um sistema de autorização de fallback. Se o `forMSGraphAccess` não for usado, e o administrador não tiver concedido consentimento, o token de bootstrap será retornado, mas a tentativa de alterá-lo com o fluxo "on-behalf-of" resultará em um erro. Portanto, a opção `forMSGraphAccess` permite ao suplemento ramificar para o sistema de fallback rapidamente.
    * Você criará a função `getData` em uma etapa posterior.
    * O parâmetro `/api/values` é a URL de um controlador do lado do servidor que fará a troca de tokens e usará o token de acesso recebido para fazer a chamada para o Microsoft Graph.

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        forMSGraphAccess: true });

    getData("/api/values", bootstrapToken);
    ```

1. Abaixo da função `getGraphData`, adicione o seguinte. Sobre este código, observe:

    * Ele é usado pelos sistemas de autorização de fallback e SSO.
    * O parâmetro `relativeUrl` é um controlador do lado do servidor.
    * O parâmetro `accessToken` pode ser um token de bootstrap ou um token de acesso completo.
    * O `writeFileNamesToOfficeDocument` já faz parte do projeto.
    * Você criará a função `handleServerSideErrors` em uma última etapa.

    ```javascript
    function getData(relativeUrl, accessToken) {

        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
            .done(function (result) {
                writeFileNamesToOfficeDocument(result)
                    .then(function () {
                        showResult(["Your data has been added to the document."]);
                    })
                    .catch(function (error) {
                        showResult([JSON.stringify(error)]);
                    });
            })
            .fail(function (result) {
                handleServerSideErrors(result);
            });
    }
    ```

### <a name="handle-client-side-errors"></a>Tratar erros do lado do cliente

1. Abaixo da função `getData`, adicione a função a seguir. Observe que `error.code` é um número, normalmente no intervalo 13xxx.

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 2: Handle errors where the add-in should NOT invoke
            //         the alternative system of authorization.

            // TODO 3: Handle errors where the add-in should invoke
            //         the alternative system of authorization.

        }
    }
    ```

1. Substitua `TODO 2` pelo código a seguir. Para saber mais sobre esses erros, confira [Solucionar problemas de SSO em suplementos do Office em](troubleshoot-sso-in-office-add-ins.md).

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one
        // is logged into Office, then the first call of getAccessToken should pass the
        // `allowSignInPrompt: true` option.
        showResult(["No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."]);
        break;
    case 13002:
        // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
        // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
        showResult(["You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."]);
        break;
    case 13006:
        // Only seen in Office on the Web.
        showResult(["Office on the Web is experiencing a problem. Please sign out of Office, close the browser, and then start again."]);
        break;
    case 13008:
        // Only seen in Office on the Web.
        showResult(["Office is still working on the last operation. When it completes, try this operation again."]);
        break;
    case 13010:
        // Only seen in Office on the Web.
        showResult(["Follow the instructions to change your browser's zone configuration."]);
        break;
    ```

1. Substitua `TODO 3` pelo código a seguir. Para todos os outros erros, o suplemento ramificará para o sistema de autorização de fallback. Para mais informações sobre esses erros, confira [Solucionar problemas de SSO em Suplementos do Office](troubleshoot-sso-in-office-add-ins.md). Neste suplemento, o sistema de fallback abre uma caixa de diálogo que exige que o usuário entre, mesmo que o usuário já esteja, e use o msal.js e Implicit Flow para obter um token de acesso ao Microsoft Graph.

    ```javascript
    default:
        dialogFallback();
        break;
    ```

### <a name="handle-server-side-errors"></a>Resolver erros do lado do servidor

1. Abaixo da função `handleClientSideErrors`, adicione a função a seguir.

    ```javascript
    function handleServerSideErrors(result) {

    // TODO 4: Parse the JSON response.

    // TODO 5: Handle case where Microsoft Graph requires an additional form
    //         of authentication.

    // TODO 6: Handle other Azure AD errors

    }
    ```

1. Substitua `TODO 4` pelo seguinte. Sobre esse código, observe que as classes de erro ASP.NET foram criadas antes de haver algo como a MFA. Como um efeito colateral de como a lógica do lado do servidor lida com as solicitações de um segundo fator de autenticação, o erro do lado do servidor enviado para o cliente tem uma propriedade **Message**, mas não uma propriedade **ExceptionMessage**. Mas todos os outros erros terão uma propriedade **ExceptionMessage**, para que o código do cliente precise analisar a resposta para ambos. Uma ou outra variável será indefinida.

    ```javascript
    var message = JSON.parse(result.responseText).Message;
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    ```

1. Substitua `TODO 5` pelo seguinte. Quando o Microsoft Graph requer uma forma adicional de autenticação, ele envia o erro AADSTS50076. Isso inclui informações sobre os requisitos adicionais na propriedade **Message.Claims**. Para lidar com isso, o código faz uma segunda tentativa de obter o token de bootstrap, mas, desta vez, ele inclui a solicitação de um fator adicional, como o valor da opção `authChallenge`, que informa ao Azure AD a solicitar todos os formulários de autenticação necessários.

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
            return;
        }
    }
    ```

1. Substitua `TODO 6` pelo seguinte.

    ```javascript
    if (exceptionMessage) {

        // TODO 7: Handle case where bootstrap token has expired.

        // TODO 8: Handle all other Azure AD errors.
    }
    ```

1. Substitua `TODO 7` pelo seguinte. Observe que, em raras ocasiões, o token de bootstrap fica não vencido quando o Office o valida, mas vence no momento em que ele é enviado ao Azure AD para o Exchange. O Azure AD responderá com o erro AADSTS500133. Quando isso acontece, o código recupera a API de SSO (mas não mais de uma vez). Desta vez, o Office retorna um novo token de bootstrap não vencido.

    ```javascript
    if ((exceptionMessage.indexOf("AADSTS500133") !== -1)
        && (retryGetAccessToken <= 0)) {

        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. Substitua `TODO 8` pelo seguinte.

    ```javascript
    else {
        dialogFallback();
    }
    ```

1. Salve o arquivo.

## <a name="code-the-server-side"></a>Codifique o lado do servidor

### <a name="configure-the-owin-middleware"></a>Configurar o middleware OWIN

1. Abra o arquivo Startup.cs na raiz do projeto **Office-Add-in-ASPNET-SSO-WebAPI** e adicione o seguinte método à classe **Inicialização**. Observe que você criará o método `ConfigureAuth` em uma etapa posterior.

    ```csharp
    public void Configuration(IAppBuilder app)
    {
        ConfigureAuth(app);
    }
    ```

1. Salve e feche o arquivo.

1. Clique com botão direito do mouse na pasta **App_Start** e selecione **Adicionar > Classe**.

1. Na caixa de diálogo **Adicionar novo item** nomeie o arquivo **Startup.Auth.cs** e, em seguida, clique em **Adicionar**.

1. Encurte o nome do namespace no novo arquivo para `Office_Add_in_ASPNET_SSO_WebAPI`.

1. Verifique se todas as seguintes instruções `using` estão na parte superior do arquivo.

    ```csharp
    using Owin;
    using Microsoft.IdentityModel.Tokens;
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
        // TODO 1: Configure the validation settings

        // TODO 2: Specify the type of authorization and the discovery endpoint
        //        of the secure token service.
    }
    ```

1. Substitua o `TODO 1` pelo seguinte. Observação sobre o código:

    * O código instrui o OWIN a garantir que o público especificado no token de bootstrap que vem do host do Office deve coincidir com os valores especificados no Web.config.
    * As Contas da Microsoft têm um GUID de emissor diferente de qualquer GUID de locatário organizacional, portanto, para dar suporte a ambos os tipos de contas, não validamos o emissor.
    * Definir `SaveSigninToken` como `true` faz com que o OWIN salve o token bruto de bootstrap do host do Office. O suplemento precisa dele para obter um token de acesso para o Microsoft Graph com o fluxo "on-behalf-of".
    * Os escopos não são validados pelo middleware OWIN. Os escopos do token de bootstrap, que devem conter `access_as_user`, são validados no controlador.

    ```csharp
    TokenValidationParameters tvps = new TokenValidationParameters
    {
        ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
        ValidateIssuer = false,
        SaveSigninToken = true
    };
    ```

1. Substitua `TODO 2` pelo seguinte. Observação sobre o código:

    * O método `UseOAuthBearerAuthentication` é chamado em vez do `UseWindowsAzureActiveDirectoryBearerAuthentication` que é mais comum, porque este último não é compatível com o ponto de extremidade V2 do Azure AD.
    * A URL transmitida ao método é onde o middleware OWIN obtém instruções para conseguir a chave que precisa para verificar a assinatura no token de bootstrap recebido do host do Office. O segmento de Autoridade da URL vem do Web.config. Ele é a cadeia de caracteres "comum" ou, para um suplemento de locatário único, uma GUID.

    ```csharp
    string[] endAuthoritySegments = { "oauth2/v2.0" };
    string[] parsedAuthority = ConfigurationManager.AppSettings["ida:Authority"].Split(endAuthoritySegments, System.StringSplitOptions.None);
    string wellKnownURL = parsedAuthority[0] + "v2.0/.well-known/openid-configuration";

    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
    {
        AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider(wellKnownURL))
    });
    ```

1. Salve e feche o arquivo.

### <a name="create-the-apivalues-controller"></a>Criar o controlador /api/values

1. Abra o arquivo **Controllers\ValueController.cs**. Esse controlador é usado quando o sistema SSO obtém um token de bootstrap com êxito. Ele não é usado como parte do sistema de autorização de fallback. Esse sistema usou o AzureADAuthController que foi criado para você.

1. Verifique se as seguintes instruções `using` estão na parte superior do arquivo.

    ```csharp
    using Microsoft.Identity.Client;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    ```

1. Logo acima da linha que declara o `ValuesController`, adicione o atributo `[Authorize]`. Isso garante que seu suplemento executará o processo de autorização configurado no último procedimento sempre que um método controlador for chamado. Apenas os chamadores com um token de acesso válido para o seu suplemento podem invocar os métodos do controlador.

1. Adicione o método a seguir ao `ValuesController`. Observe que é o valor de retorno é `Task<HttpResponseMessage>` em vez de `Task<IEnumerable<string>>`, como seria mais comum para um método `GET api/values`. Este é o efeito colateral deste fato que a lógica de autorização do OAuth deve estar no controlador, em fez de em um filtro ASP.NET. Algumas condições de erro na lógica exigem que um objeto de resposta HTTP seja enviado para o cliente do suplemento.

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO 1: Validate the scopes of the bootstrap token.

        // TODO 2: Assemble all the information that is needed to get a
        //        token for Microsoft Graph using the on-behalf-of flow.

        // TODO 3: Get the access token for Microsoft Graph.

        // TODO 4: Use the token to call Microsoft Graph.
    }
    ```

1. Substitua `TODO1` pelo seguinte código para validar que os escopos especificados no token incluam `access_as_user`. Observe que o segundo parâmetro do método `SendErrorToClient` é um objeto **Exception**. Nesse caso, o código passa `null` porque incluir o objeto **Exception** bloqueia a inclusão da propriedade **Message** na resposta HTTP que é gerada.


    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (!(addinScopes.Contains("access_as_user")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    }
    ```

1. Substitua `TODO 2` pelo seguinte código para montar todas as informações necessárias para obter um token do Microsoft Graph usando o fluxo "on behalf of". Sobre este código, observe:

    * Seu suplemento não está mais desempenhando o papel de um recurso (ou público) para o qual o host do Office e o usuário precisam de acesso. Agora, ele mesmo é um cliente que precisa de acesso ao Microsoft Graph. `ConfidentialClientApplication` é o objeto "client context" da MSAL.
    * A partir da MSAL.NET 3.x.x, o `bootstrapContext` é apenas o token de bootstrap em si.
    * A Autoridade vem do Web.config. Ela é a cadeia de caracteres "comum" ou, para um suplemento de locatário único, uma GUID.
    * A MSAL exige os escopos `openid` e `offline_access` para funcionar, mas ela lança um erro se o código solicitá-los de forma redundante. Ela também lançará um erro se o seu código solicitar o `profile`, que realmente é usado apenas quando o aplicativo host do Office recebe o token para o aplicativo Web do seu suplemento. Então, apenas `Files.Read.All` é explicitamente solicitado.

    ```csharp
    string bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext.ToString();
    UserAssertion userAssertion = new UserAssertion(bootstrapContext);

    var cca = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ida:ClientID"])
                                                    .WithRedirectUri("https://localhost:44355")
                                                    .WithClientSecret(ConfigurationManager.AppSettings["ida:Password"])
                                                    .WithAuthority(ConfigurationManager.AppSettings["ida:Authority"])
                                                    .Build();

    string[] graphScopes = { "https://graph.microsoft.com/Files.Read.All" };
    ```

1. Substitua `TODO 3` pelo código a seguir. Observação sobre este código:

    * O método `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` procurará primeiro no cache da MSAL, que está na memória, para fazer a correspondência com o token de acesso. Somente se não houver um, ele iniciará o fluxo "on behalf of" com o ponto de extremidade V2 do Azure AD.
    * Quaisquer exceções que não forem do tipo `MsalServiceException` são intencionalmente não detectadas, e, portanto, se propagarão para o cliente como mensagens `500 Server Error`.

    ```csharp
    AcquireTokenOnBehalfOfParameterBuilder parameterBuilder = null;
    AuthenticationResult authResult = null;
    try
    {
        parameterBuilder = cca.AcquireTokenOnBehalfOf(graphScopes, userAssertion);
        authResult = await parameterBuilder.ExecuteAsync();
    }
    catch (MsalServiceException e)
    {
        // TODO 3a: Handle request for multi-factor authentication.

        // TODO 3b: Handle lack of consent and invalid scope (permission).

        // TODO 3c: Handle all other MsalServiceExceptions.
    }
    ```

1. Substitua `TODO 3a` pelo código a seguir. Sobre este código, observe:

    * Se a autenticação multifator for exigida pelo recurso Microsoft Graph e o usuário ainda não a tiver fornecido, o Azure AD retornará "400 Bad Request" com o erro `AADSTS50076` e uma propriedade **Declarações**. O MSAL exibe **MsalUiRequiredException** (que herda de **MsalServiceException**) com essas informações.
    * O valor da propriedade **Declarações** deve ser passado para o cliente, que deve passá-lo para o host do Office, que, por sua vez, o incluirá em um pedido para um novo token de bootstrap. O Azure AD solicitará ao usuário todas as formas de autenticação necessárias.
    * As APIs que criam respostas HTTP a partir de exceções não conhecem a propriedade **Claims**, portanto, elas não a incluem no objeto de resposta. É necessário criar manualmente uma mensagem que inclua esse recurso. Uma propriedade **Message** personalizada, no entanto, impede a criação de uma propriedade **ExceptionMessage**, portanto, a única maneira de obter a ID de erro `AADSTS50076` para o cliente é adicioná-la à **Message** personalizada. O JavaScript no cliente precisará descobrir se uma resposta tem uma **Message** ou **ExceptionMessage** para saber qual ler.
    * A mensagem personalizada é formatada como JSON para que o JavaScript do cliente possa analisá-la com métodos de objeto `JSON` JavaScript conhecidos.

    ```csharp
    if (e.Message.StartsWith("AADSTS50076"))
    {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. Substitua `TODO 3b` pelo código a seguir. Sobre este código, observe:

    * Se a chamada para o Azure AD contiver pelo menos um escopo (permissão) que não tenha sido consentido pelo usuário ou por um administrador de locatários (ou se o consentimento foi revogado), o Azure AD retornará "400 Solicitação Incorreta" com o erro `AADSTS65001`. O MSAL exibe um **MsalUiRequiredException** com essas informações.
    *  Se a chamada para o Azure AD contiver pelo menos um escopo que Azure AD não reconhece, o Azure AD retornará "400 Solicitação Incorreta" com o erro `AADSTS70011`. O MSAL exibe um **MsalUiRequiredException** com essas informações.
    *  A descrição completa é incluída porque 70011 é retornado em outras condições e ele deverá ser processado neste suplemento somente quando significar que há um escopo inválido.
    *  O objeto **MsalUiRequiredException** é passado para `SendErrorToClient`. Isso garante que uma propriedade **ExceptionMessage** contendo as informações de erro seja incluída na resposta HTTP.

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001")) || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. Substitua `TODO 3c` pelo seguinte código para lidar com todas as outras **MsalServiceException**s. Conforme observado anteriormente,

    ```csharp
    else
    {
        throw e;
    }
    ```

1. Substitua `TODO 4` pelo código a seguir. O método `GraphApiHelper.GetOneDriveFileNames`, que foi criado para você, faz a solicitação de dados ao Microsoft Graph e inclui o token de acesso.

    ```csharp
    return await GraphApiHelper.GetOneDriveFileNames(authResult.AccessToken);
    ```

1. Salve e feche o arquivo.

## <a name="run-the-solution"></a>Executar a solução

1. Abra o arquivo de solução do Visual Studio.
1. No menu **Build**, selecione **Solução Limpa**. Quando terminar, abra o menu **Build** novamente e selecione **Solução de Build**.
1. No **Gerenciador de Soluções**, selecione o nó de projeto **Office-Add-in-ASPNET-SSO** (não o nó da solução principal e não o projeto cujo nome termina em "WebAPI").
1. No painel **Propriedades**, abra o menu suspenso **Iniciar documento** e escolha uma das três opções (Excel, Word ou PowerPoint).

    ![Escolha o aplicativo host do Office desejado: Excel, PowerPoint ou Word](../images/SelectHost.JPG)

1. Pressione F5.
1. No aplicativo do Office, na faixa de opções **Home**, selecione **Mostrar suplemento** no grupo **SSO ASP.NET** para abrir o suplemento do painel de tarefas.
1. Clique no botão **Definir Nome de Arquivos do One Drive**. Se você estiver conectado ao Office com uma conta corporativa ou de estudante (Office 365) ou com uma conta da Microsoft e o SSO estiver funcionando conforme o esperado, os primeiros 10 nomes de arquivos e pastas no OneDrive for Business serão exibidos no painel de tarefas. Se você não estiver conectado ou se você estiver em um cenário que não tem suporte para SSO, ou se o SSO não estiver funcionando por nenhum motivo, você será solicitado a fazer logon. Depois de entrar, os nomes de arquivos e pastas serão exibidos.
