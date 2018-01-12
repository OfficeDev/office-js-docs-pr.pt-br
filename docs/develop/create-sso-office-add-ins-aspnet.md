# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on"></a>Criar um Suplemento do Office com ASP.NET que usa logon único

Quando os usuários estão conectados ao Office, o seu suplemento pode usar as mesmas credenciais para permitir que os usuários acessem vários aplicativos sem exigir que eles entrem uma segunda vez. Confira uma visão geral no artigo [Habilitar o SSO em um Suplemento do Office](../../docs/develop/sso-in-office-add-ins.md).

Este artigo apresenta o processo passo a passo de habilitação do logon único (SSO) em um suplemento que foi criado com ASP.NET, OWIN e com a Biblioteca de Autenticação da Microsoft (MSAL) para .NET.

> **Observação:** Para ler um artigo semelhante sobre um suplemento baseado em Node.js, confira [Criar um Suplemento do Office com Node.js que usa o logon único](../../docs/develop/create-sso-office-add-ins-nodejs.md).

## <a name="prerequisites"></a>Pré-requisitos

* A versão mais recente disponível do Visual Studio 2017 Preview.

>**Observação:** A versão mais recente do Visual Studio 2017 Preview atualmente não é compatível com a marcação do manifesto do suplemento que é exigido para SSO. Os procedimentos a seguir fornecem detalhes sobre como resolver esse problema.

* Office 2016, versão 1708, build 8424.nnnn ou posterior (a versão de assinatura do Office 365, às vezes chamada de "Clique para Executar"). Você talvez precise ser um participante do programa Office Insider para obter essa versão. Para obter mais informações, confira a página [Seja um Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1).

## <a name="set-up-the-starter-project"></a>Configurar o projeto inicial

1. Clone ou baixe o repositório em [SSO com Suplemento ASPNET do Office](https://github.com/officedev/office-add-in-aspnet-sso).

1. Abra a pasta **Before** (antes) e abra o arquivo .sln no Visual Studio. Esse é um projeto inicial. A interface do usuário e outros aspectos do suplemento que não estão diretamente ligados ao SSO ou à autorização já estão prontos.

    > Observação: Há também uma versão concluída do exemplo no mesmo repositório. Essa versão apresenta como seria o suplemento quando concluídos os procedimentos apresentados neste artigo, com exceção de que o projeto concluído traz comentários de códigos que seriam redundantes neste artigo. Para usar a versão concluída, apenas abra o arquivo *.sln e siga as instruções apresentadas neste artigo, mas pule as seções **Codificar o lado do cliente** e **Codificar o lado do servidor**.

1. Depois que o projeto for aberto, compile-o no Visual Studio. Isso instalará os pacotes listados no arquivo packages.config. Esse procedimento poderá levar entre alguns segundos e alguns minutos dependendo de quantos pacotes estiverem no cache de pacote local do computador.

    > **Importante!** O packages.config na raiz do projeto de API Web especifica a versão `1.1.1-alpha0393` do Microsoft.Identity.Client, a biblioteca de MSAL. Você deve verificar se esta versão (ou posterior) é instalada depois de pressionar F5 pela primeira vez: No menu **Ferramentas**, navegue até **Gerenciador de Pacote Nuget** > **Gerenciar Pacotes Nuget para a solução** > **Instalado**. Role até **Microsoft.Identity.Client** para ver a versão instalada. Se ela for anterior a `1.1.1-alpha0393` (ou não aparecer na lista **Instalado**), navegue até **Gerenciador de Pacote Nuget** > **Console do Gerenciador de Pacotes**. No console, execute o comando `Install-Package Microsoft.Identity.Client -Version 1.1.1-alpha0393 -Source https://www.myget.org/F/aad-clients-nightly/api/v3/index.json`.

1. Depois que o projeto for completamente compilado, pressione F5. O PowerPoint será aberto e haverá um grupo **SSO ASP.NET** na faixa de opções **Página Inicial**.

1. Pressione o botão **Mostrar Suplemento** nesse grupo para ver a interface do usuário do suplemento no painel de tarefas. O botão no painel de tarefas ainda não está conectado.
2. No Visual Studio, interrompa o depurador.

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Registrar o suplemento com o ponto de extremidade v2.0 do Azure AD

1. Acesse [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com).

1. Entre com as credenciais de administrador em sua locação do Office 365. Por exemplo, MeuNome@contoso.onmicrosoft.com

1. Clique em **Adicionar um aplicativo**.

1. Quando for solicitado, use "Office-Add-in-ASPNET-SSO" como o nome do aplicativo e, em seguida, pressione **Criar aplicativo**.

1. Quando a página de configuração do aplicativo abrir, copie a **ID do aplicativo** e salve-a. Você irá usá-la em um procedimento posterior.

    > **Observação**: Essa ID é o valor "audience" (público) quando outros aplicativos, como o aplicativo host do Office (por exemplo, PowerPoint, Word, Excel), buscam o acesso autorizado ao aplicativo. Também é a "ID do cliente" do aplicativo quando ela, por sua vez, busca o acesso autorizado ao Microsoft Graph.

1. Na seção **Segredos do Aplicativo**, pressione **Gerar Nova Senha**. Uma caixa de diálogo pop-up abrirá e uma nova senha (também chamada de "segredo do aplicativo") será mostrada. *Copie a senha imediatamente e salve-a com a ID do aplicativo.* Você precisará dela em um procedimento posterior. Feche a caixa de diálogo.

1. Na seção **Plataformas**, clique em **Adicionar plataforma**.

1. Na caixa de diálogo que abrir, selecione **API Web**.

1. Um **URI da ID do aplicativo** foi gerado do formulário “api://{App ID GUID}”. Insira a cadeia de caracteres “localhost:44355/” entre as barras duplas e o GUID. A ID inteira deve ser `api://localhost:44355/{App ID GUID}`. (A parte do domínio do nome do **Escopo** logo abaixo do **URI da ID do aplicativo** será automaticamente alterada para que haja correspondência. Ela deve ser assim: `api://localhost:44355/{App ID GUID}/access_as_user`.)

1. Na seção **Aplicativos pré-autorizados**, você identifica os aplicativos que deseja autorizar para o aplicativo da Web do seu suplemento. Cada uma das seguintes IDs precisa ser pré-autorizada. Cada vez que você inserir uma, uma nova caixa de texto vazia aparece. (Insira apenas o GUID.)
 * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
 * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
 * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)

1. Abra o menu suspenso do **Escopo** ao lado de cada **ID do aplicativo** e marque a caixa para `api://localhost:44355/{App ID GUID}/access_as_user`.

1. Próximo ao topo da seção **Plataformas**, clique em **Adicionar Plataforma** novamente e selecione **Web**.

1. Na nova seção **Web** em **Plataformas**, insira o seguinte como uma **URL de redirecionamento**: `https://localhost:44355`.

    > Observação: Até o presente momento, a plataforma **API Web** às vezes desaparece da seção **Plataformas**, especialmente se a página for atualizada depois que a plataforma **Web** é adicionada *e a página de registro é salva*. Para garantir que sua plataforma **API Web** ainda faz parte do registro, clique no botão **Editar Manifesto do Aplicativo** próximo à parte inferior da página. Você deve ver a cadeia de caracteres `api://localhost:44355/{App ID GUID}` na propriedade **identifierUris** do manifesto. Também haverá uma propriedade **oauth2Permissions** cuja subpropriedade **value** tem o valor `access_as_user`.

1. Role para baixo até a seção **Permissões do Microsoft Graph**, na subseção **Permissões Delegadas**. Use o botão **Adicionar** para abrir a caixa de diálogo **Selecionar Permissões**.

1. Na caixa de diálogo, marque as caixas das seguintes permissões (algumas já podem estar marcadas por padrão). Somente a primeira é realmente exigida pelo suplemento propriamente dito, mas a biblioteca MSAL usada pelo código de servidor exige `offline_access` e `openid`. A permissão `profile` é necessária para que o host do Office obtenha um token no aplicativo Web do seu suplemento.
 * Files.Read.All
 * offline_access
 * openid
 * profile

1. Clique em **OK** no final da caixa de diálogo.

1. Clique em **Salvar** na parte inferior da página de registro.

## <a name="grant-admin-consent-to-the-add-in"></a>Conceder consentimento de administrador para o suplemento

> **Observação:** Este procedimento só é necessário quando você está desenvolvendo o suplemento. Quando o seu suplemento de produção é implantado na Loja do Office ou em um catálogo de suplementos, os usuários confiarão individualmente nele ou um administrador concordará pela organização na instalação.

1. Se o suplemento não estiver em execução no Visual Studio, pressione **F5** para executá-lo. Ele precisa estar em execução no IIS para que este procedimento seja concluído sem problemas.

1. Na cadeia de caracteres a seguir, substitua o espaço reservado "{application_ID}" pela ID do Aplicativo que você copiou quando registrou seu suplemento:  `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. Cole a URL resultante na barra de endereços do navegador e acesse-a.

1. Quando for solicitado, entre com as credenciais de administrador em sua locação do Office 365.

1. Em seguida, será solicitado que você conceda permissão para seu suplemento acessar os dados do Microsoft Graph. Clique em **Aceitar**.

1. A guia ou janela do navegador é, então, redirecionada para a **URL de redirecionamento** que você especificou ao registrar o suplemento. Sendo assim, a pagina inicial do suplemento abrirá no navegador.

2. Na barra de endereços do navegador, você verá um parâmetro de consulta de locatário "tenant" com um valor GUID. Esta é a ID da sua locação do Office 365. Copie e salve esse valor. Você irá usá-lo em uma etapa posterior.

3. Feche a janela ou a guia.

1. Interrompa o depurador no Visual Studio.

## <a name="configure-the-add-in"></a>Configurar o suplemento

1. Na seguinte cadeia de caracteres, substitua o espaço reservado "{tenant_ID}" pela ID de locatário do Office 365 obtida anteriormente. Se por algum motivo, você não obteve a ID anteriormente, use um dos métodos descritos em [Localizar a ID de locatário do Office 365](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) para obtê-la.

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

1. No Visual Studio, abra o Web.config. Existem algumas chaves na seção **appSettings** às quais você precisa atribuir valores.

1. Use a cadeia de caracteres construída na etapa 1 como o valor para a chave denominada "ida:Issuer". Não deixe espaços em branco no valor.

1. Atribua os seguintes valores para as chaves correspondentes:

|Chave|Valor|
|:-----|:-----|
|ida:ClientID|A ID do aplicativo obtida ao registrar o suplemento.|
|ida:Audience|A ID do aplicativo obtida ao registrar o suplemento.|
|ida:Password|A senha obtida ao registrar o suplemento.|


Veja a seguir um exemplo de como as quatro chaves que você alterou devem se parecer. *Observe que as chaves ClientID e Audience são iguais*. Você também pode usar uma única chave para ambos os fins, mas sua marcação web.config será mais reutilizável se mantê-la separada porque ela não é sempre a mesma. Além disso, ter chaves separadas reforça a ideia de que seu suplemento é tanto um recurso de OAuth - em relação a um host do Office - e um cliente OAuth - em relação ao Microsoft Graph.

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />
    ```

> **Observação:** Não altere as demais configurações na seção **appSettings**.

1. Salve e feche o arquivo.

1. Na raiz do projeto, abra o arquivo do manifesto do suplemento "Office-Add-in-ASPNET-SSO.xml".

1. Role até o final do arquivo.

1. Logo acima da marca de fim `</VersionOverrides>`, você encontrará a marcação a seguir:

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:44355/{application_GUID here}<Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. Substitua o espaço reservado "{application_GUID here}" *nos dois lugares* na marcação pela ID do Aplicativo que você copiou ao registrar seu suplemento. O símbolo "{}" não faz parte da ID, portanto não o inclua. Essa é a mesma ID usada para a ClientID e a Audience no web.config.

    > **Observação**:
    >* O valor de **Resource** é o **URI da ID do Aplicativo** que você definiu quando adicionou a plataforma API Web no registro do suplemento.
    >* A seção **Scopes** só será usada para gerar uma caixa de diálogo de consentimento se o suplemento for vendido na Office Store.

1. Abra a guia **Avisos** da **Lista de Erros** no Visual Studio. Se houver um aviso que `<WebApplicationInfo>` não é um filho válido de `<VersionOverrides>`, sua versão do Visual Studio 2017 Preview não reconhecerá a marcação SSO. Para solucionar esse problema, faça o seguinte para um suplemento do Word, Excel ou PowerPoint. Se você estiver trabalhando com um suplemento do Outlook, confira a solução abaixo.

   - **Solução alternativa para Word, Excel e PowerPoint**

   > 1. Comente a seção `<WebApplicationInfo>` do manifesto logo acima do final de `</VersionOverrides>`.

   > 2. Pressione F5 para iniciar uma sessão de depuração. Isso criará uma cópia do manifesto na seguinte pasta (que pode ser acessada mais facilmente pelo **Gerenciador de Arquivos** do que pelo Visual Studio):`Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`

   > 3. Na cópia do manifesto, remova a sintaxe do comentário em torno da seção `<WebApplicationInfo>`.

   > 4. Salve a cópia do manifesto.

   > 5. Agora você precisa evitar que o Visual Studio substitua a cópia do manifesto na próxima vez que pressionar F5. Clique com botão direito do mouse no nó da solução na parte superior do **Gerenciador de Soluções** (não nos nós do projeto).

   > 6. Escolha **Propriedades** no menu de contexto e uma caixa de diálogo **Páginas de Propriedades da Solução** será aberta.

   > 7. Expanda **Propriedades da Configuração** e escolha **Configuração**.

   > 8. Desmarque **Criar** e **Implantar** na linha do projeto **Office-Add-in-ASPNET-SSO** (*não* o projeto **Office-Add-in-ASPNET-SSO-WebAPI**).

   > 9. Pressione **OK** para fechar a caixa de diálogo.

   - **Solução alternativa para Outlook**

   > 1. Em sua máquina de desenvolvimento, localize o `MailAppVersionOverridesV1_1.xsd` existente. Ele deve estar localizado no diretório de instalação do Visual Studio em `./Xml/Schemas/{lcid}`. Por exemplo, em uma instalação típica do VS 2017 de 32 bits em um sistema em inglês (EUA), o caminho completo seria `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.

   > 2. Renomeie o arquivo existente para `MailAppVersionOverridesV1_1.old`.

   > 3. Copie essa versão modificada do arquivo para a pasta: [Esquema MailAppVersionOverrides modificado](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)

1. Salve e feche o arquivo de manifesto principal no Visual Studio.

## <a name="code-the-client-side"></a>Codificar o lado do cliente

1. Abra o arquivo Home.js da pasta **Scripts**. Ele já apresenta alguns códigos:
    * Uma atribuição ao método `Office.initialize` que, por sua vez, atribui um manipulador ao evento clicar do botão `getGraphAccessTokenButton`.
    * Um método `showResult` que exibirá os dados retornados do Microsoft Graph (ou uma mensagem de erro) na parte inferior do painel de tarefas.

1. Abaixo da atribuição a `Office.initialize`, adicione o código a seguir. Observe o seguinte sobre este código:

    * O `getAccessTokenAsync` é a nova API no Office.js que permite que um suplemento solicite ao aplicativo host do Office (Excel, PowerPoint, Word, etc.) um token de acesso para o suplemento (para o usuário conectado ao Office). O aplicativo host do Office, por sua vez, solicita o token ao ponto de extremidade 2 do Azure AD. Uma vez que você previamente autorizou o host do Office para o seu suplemento ao registrá-lo, o Azure AD enviará o token.
    * Se nenhum usuário estiver conectado ao Office, o host do Office solicitará que o usuário se conecte.
    * O parâmetro de opções configura o `forceConsent` como falso. Dessa forma, não será solicitado que o usuário consinta o acesso ao host do Office para seu suplemento.

    ```js
    function getOneDriveFiles() {
        getDataWithToken({ forceConsent: false });
    }

    function getDataWithToken(options) {
        Office.context.auth.getAccessTokenAsync(options,
            function (result) {
                if (result.status === "succeeded") {
                    TODO1: Use the access token to get Microsoft Graph data.
                }
                else {
                    console.log("Code: " + result.error.code);
                    console.log("Message: " + result.error.message);
                    console.log("name: " + result.error.name);
                    document.getElementById("getGraphAccessTokenButton").disabled = true;
                }
            });
    }
    ```

1. Substitua TODO1 pelas linhas a seguir. Você criará o método `getData` e a rota "/api/values" do lado do servidor nas etapas posteriores. Uma URL relativa é usada para o ponto de extremidade porque ela deve ser hospedada no mesmo domínio que seu suplemento.

    ```js
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. Abaixo do método `getOneDriveFiles`, adicione o seguinte. Este método utilitário chama um ponto de extremidade da API Web especificado e transmite a ela o mesmo token de acesso que aplicativo host do Office usou para obter acesso ao seu suplemento. No lado do servidor, esse token de acesso será usado no fluxo "on behalf of" (em nome de) para obter um token de acesso para o Microsoft Graph.

    ```js
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET",
        })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            TODO2: Handle errors and the case where Microsoft Graph
                   requires additional form of authentication.
        });
    }
    ```

1. Substitua TODO2 pelas seguintes linhas. Observe o seguinte sobre este código:

    * Quando a falha acontecer em razão de o Microsoft Graph exigir uma forma de autenticação adicional, o `exceptionMessage` será uma cadeia JSON contendo "capolids". Nesse caso, o host do Office precisa obter um novo token.  
    * A mensagem de exceção informa o AAD para solicitar ao usuário todas as formas de autenticação requeridas, portanto, ela deve ser passada para o host do Office, que, por sua vez, a passa para o AAD quando ele pedir um novo token.
    * A opção `authChallenge` é o método de passar esta cadeia para o host do Office.
    * Se o erro for algo diferente de uma solicitação de autenticação adicional, ele será logado no console.

    ```js
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    if (exceptionMessage.indexOf("capolids") !== -1) {
        getDataWithToken({ authChallenge: exceptionMessage });
    } else {
        console.log(result.error);
    }
    ```

1. Salve e feche o arquivo.

## <a name="code-the-server-side"></a>Codificar o lado do servidor

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

    ```
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

    ```
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. Substitua TODO3 pelo seguinte código. Observação:

    * O código instrui o OWIN a garantir que o emissor de token e audiência especificado no token de acesso que vem do host do Office (e é transmitido pela chamada de `getData` do lado do cliente) deve coincidir com os valores especificados no Web.config.
    * Definir `SaveSigninToken` como `true` faz com que o OWIN salve o token bruto do host do Office. O suplemento precisa dele para obter um token de acesso para o Microsoft Graph com o fluxo "on behalf of".
    * Os escopos não são validados pelo middleware OWIN. Os escopos do token de acesso, que devem conter `access_as_user`, são validados no controlador.

    ```
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. Substitua TODO4 pelo seguinte código. Observação:

    * O método `UseOAuthBearerAuthentication` é chamado em vez do `UseWindowsAzureActiveDirectoryBearerAuthentication` que é mais comum, porque este último não é compatível com o ponto de extremidade V2 do Azure AD.
    * A URL de descoberta transmitida ao método é onde o middleware OWIN obtém instruções para conseguir a chave que precisa para verificar a assinatura no token de acesso recebido do host do Office.

    ```
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
            {
                AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
            });
    ```

1. Salve e feche o arquivo.

### <a name="create-the-apivalues-controller"></a>Criar o controlador /api/values

1. Abra o arquivo **Controllers\ValueController.cs**.

2. Verifique se as seguintes instruções `using` estão na parte superior do arquivo.

    ```
    using Microsoft.Identity.Client;
    using System;
    using System.IdentityModel.Tokens;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web;
    using System.Web.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    using Office_Add_in_ASPNET_SSO_WebAPI.Models;
    ```

3. Logo acima da linha que declara o `ValuesController`, adicione o atributo `[Authorize]`. Isso garante que seu suplemento executará o processo de autorização que você configurou no último procedimento sempre que um método de controle for chamado. Somente os chamadores com um token de acesso válido para o seu suplemento podem invocar os métodos do controlador.

4. Adicione o método a seguir ao `ValuesController`:

    ```
    // GET api/values
    public async Task<IEnumerable<string>> Get()
    {
        // TODO5: Validate the scopes of the access token.
    }
    ```

5. Substitua TODO5 pelo seguinte código para validar que os escopos especificados no token incluam `access_as_user`.

    ```
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (addinScopes.Contains("access_as_user"))
    {
        // TODO6: Assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.
        // TODO7: Get the access token for Microsoft Graph.
        // TODO8: Get the names of files and folders in OneDrive by using the Microsoft Graph API.
        // TODO9: Remove excess information from the data and send the data to the client.
    }
    return new string[] { "Error", "Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user." };
    ```

> **Observação:** Você deve usar apenas o escopo `access_as_user` para autorizar a API que lida com o fluxo Em Nome De para os suplementos do Office. Outras APIs em seu serviço devem ter seus próprios requisitos de escopo. Isso limita o que pode ser acessado com os tokens que o Office adquire.

6. Substitua TODO6 pelo código a seguir. Observação:
    * Ele transforma o token de acesso bruto recebido do host do Office em um objeto de `UserAssertion` que será transmitido para outro método.
    * Seu suplemento não está mais desempenhando o papel de um recurso (ou público) para o qual o host do Office e o usuário precisam de acesso. Agora, ele mesmo é um cliente que precisa de acesso ao Microsoft Graph. `ConfidentialClientApplication` é o objeto "client context" da MSAL.
    * O terceiro parâmetro para o construtor `ConfidentialClientApplication` é uma URL de redirecionamento que não é realmente usada no fluxo "on behalf of", mas usar a URL correta é uma boa prática. O quarto e o quinto parâmetros podem ser usados para definir um armazenamento persistente que permitiria a reutilização de tokens não expirados em diferentes sessões com o suplemento. Este exemplo não implementa nenhum armazenamento persistente.
    * A MSAL exige os escopos `openid` e `offline_access` para funcionar, mas ela lança um erro se o código solicitá-los de forma redundante. Ela também lançará um erro se o seu código solicitar o `profile`, que realmente é usado apenas quando o aplicativo host do Office recebe o token para o aplicativo Web do seu suplemento. Então, apenas `Files.Read.All` é explicitamente solicitado.

    ```
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

7. Substitua TODO7 pelo seguinte código. Observação:

    * O método `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` procurará primeiro no cache da MSAL, que está na memória, para fazer a correspondência com o token de acesso. Somente se não houver um, ele iniciará o fluxo "on behalf of" com o ponto de extremidade V2 do Azure AD.
    * Se a autenticação multi-fator for requerida pelo recurso MS Graph e o usuário ainda não a tiver fornecido, o AAD lançará uma exceção contendo uma propriedade de Declarações.
    * O valor da propriedade de Declarações deve ser passado para o cliente, que o passará para o host do Office, que, em seguida, o incluirá em um pedido para um novo token. O AAD solicitará ao usuário todas as formas de autenticação requeridas.
    * Quaisquer exceções que não forem do tipo `MsalUiRequiredException` são intencionalmente não detectadas, e, portanto, se propagarão para o cliente.

    ```
    AuthenticationResult result = null;
    try
    {
        result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
    }
    catch (MsalUiRequiredException e)
    {        
        if (String.IsNullOrEmpty(e.Claims))
        {
            throw e;
        }
        else
        {
            throw new HttpException(e.Claims);
        }   
    }
    ```

8. Substitua TODO8 pelo seguinte código. Observação:

    * As classes `GraphApiHelper` e `ODataHelper` são definidas nos arquivos da pasta **Helpers**. A classe `OneDriveItem` é definida em um arquivo da pasta **Models**. A discussão detalhada dessas classes não é relevante para a autorização ou o SSO, portanto, está fora do escopo deste artigo.
    * O desempenho é aprimorado ao se solicitar ao Microsoft Graph apenas os dados que são realmente necessários. Desse modo, o código usa um parâmetro de consulta ` $select` para especificar que desejamos somente a propriedade de nome, e usa um parâmetro `$top` para especificar que desejamos somente os três primeiros nomes de pasta ou de arquivo.

    ```
    var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
    var getFilesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
    ```

9. Substitua TODO9 pelo seguinte código. Observe que, embora o código acima solicite somente a propriedade *name* dos itens do OneDrive, o Microsoft Graph sempre inclui a propriedade *eTag* para os itens do OneDrive. Para reduzir a carga enviada para o cliente, o código a seguir reconstrói os resultados apenas com os nomes dos itens.

    ```
    List<string> itemNames = new List<string>();
    foreach (OneDriveItem item in getFilesResult)
    {
      itemNames.Add(item.Name);
    }                    
    return itemNames;
    ```

## <a name="run-the-add-in"></a>Executar o suplemento

1. Certifique-se de ter alguns arquivos no seu OneDrive para que você possa verificar os resultados.

1. No Visual Studio, pressione F5. O PowerPoint será aberto e haverá um grupo **SSO ASP.NET** na faixa de opções **Página Inicial**.

1. Pressione o botão **Mostrar Suplemento** nesse grupo para ver a interface do usuário do suplemento no painel de tarefas.

1. Pressione o botão **Obter Meus Arquivos do OneDrive**. Se você não estiver conectado ao Office, você será solicitado a entrar.
    > **Observação:** Se você entrou no Office com uma ID diferente e se alguns aplicativos do Office que estavam abertos no momento continuam abertos, o Office pode não alterar de forma confiável sua ID, mesmo que pareça ter feito isso no PowerPoint. Se isso acontecer, a chamada para o Microsoft Graph pode falhar ou os dados da ID anterior podem ser retornados. Para evitar isso, certifique-se de *fechar todos os outros aplicativos do Office* antes de pressionar **Obter meus arquivos do OneDrive**.

1. Depois de entrar, será exibida uma lista de seus arquivos e suas pastas no OneDrive, abaixo do botão. Esse procedimento pode levar mais de 15 segundos, principalmente na primeira vez.
