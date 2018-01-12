# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a>Criar um Suplemento do Office com Node.js que usa logon único

Os usuários podem entrar no Office, e o Suplemento Web do Office pode aproveitar esse processo de entrada para autorizá-los a acessar seu suplemento e o Microsoft Graph sem exigir que os eles façam logon uma segunda vez. Para obter uma visão geral, confira o artigo [Habilitar o SSO em um Suplemento do Office](../../docs/develop/sso-in-office-add-ins.md).

Este artigo apresenta o processo passo a passo de habilitação do logon único (SSO) em um suplemento que foi criado com Node.js e express. 

> **Observação:** Para ler um artigo semelhante sobre um suplemento baseado em ASP.NET, confira [Criar um Suplemento do Office com ASP.NET que usa o logon único](../../docs/develop/create-sso-office-add-ins-aspnet.md).

## <a name="prerequisites"></a>Pré-requisitos

* [Node e npm](https://nodejs.org/en/), versão 6.9.4 ou posterior.
* [Git Bash](https://git-scm.com/downloads) (ou outro cliente Git.)
* TypeScript, versão 2.2.2 ou posterior
* Office 2016, versão 1708, build 8424.nnnn ou posterior (a versão de assinatura do Office 365, às vezes chamada de "Clique para Executar"). Você talvez precise ser um participante do programa Office Insider para obter essa versão. Para obter mais informações, confira a página [Seja um Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1).

## <a name="set-up-the-starter-project"></a>Configurar o projeto inicial

1. Clone ou baixe o repositório em [SSO com Suplemento NodeJS do Office](https://github.com/officedev/office-add-in-nodejs-sso). 


    > **Observação:** Há duas versões do exemplo: 
    > 
    > * A pasta **Before** (antes) traz um projeto inicial. A interface do usuário e outros aspectos do suplemento que não estão diretamente ligados ao SSO ou à autorização já estão prontos. As próximas seções deste artigo apresentam uma orientação passo a passo para concluir o projeto. 
    > * A versão **Completed** (concluído) do exemplo apresenta como seria o suplemento quando concluídos os procedimentos apresentados neste artigo, com exceção de que o projeto concluído traz comentários de códigos que seriam redundantes neste artigo. Para usar a versão concluída, apenas siga as instruções apresentadas neste artigo, substituindo "Before" por "Completed" e pulando as seções **Codificar o lado do cliente** e **Codificar o lado do servidor**.

1. Abra um console Git bash na pasta **Before**.

2. Insira `npm install` no console para instalar todas as dependências discriminadas no arquivo package.json.

3. Insira `npm run build ` no console para compilar o projeto. 
     > Observação: Talvez você veja alguns erros de build informando que algumas variáveis estão declaradas mas não são usadas. Ignore esses erros. Eles são um efeito colateral, pois na versão "Before" do exemplo estão faltando alguns códigos que serão adicionados posteriormente.

## <a name="register-the-add-in-with-azure-ad-v2-endpoint"></a>Registrar o suplemento com o ponto de extremidade V2 do Azure AD

1. Acesse [https://apps.dev.microsoft.com](https://apps.dev.microsoft.com) . 

1. Entre com as credenciais de administrador em sua locação do Office 365. Por exemplo, MeuNome@contoso.onmicrosoft.com

1. Clique em **Adicionar um aplicativo**.

1. Quando for solicitado, use "Office-Add-in-NodeJS-SSO" como o nome do aplicativo e, em seguida, pressione **Criar aplicativo**.

1. Quando a página de configuração do aplicativo abrir, copie a **ID do aplicativo** e salve-a. Você irá usá-la em um procedimento posterior. 

    > Observação: Essa ID é o valor "audience" (público) quando outros aplicativos, como o aplicativo host do Office (por exemplo, PowerPoint, Word, Excel), buscam o acesso autorizado ao aplicativo. Também é a "ID do cliente" do aplicativo quando ela, por sua vez, busca o acesso autorizado ao Microsoft Graph.

1. Na seção **Segredos do Aplicativo**, pressione **Gerar Nova Senha**. Uma caixa de diálogo pop-up abrirá e uma nova senha (também chamada de "segredo do aplicativo") será mostrada. *Copie a senha imediatamente e salve-a com a ID do aplicativo.* Você precisará dela em um procedimento posterior. Feche a caixa de diálogo.

1. Na seção **Plataformas**, clique em **Adicionar plataforma**. 

1. Na caixa de diálogo que abrir, selecione **API Web**.

1. Um **URI da ID do aplicativo** foi gerado do formulário “api://{App ID GUID}”. Insira a cadeia de caracteres “localhost:3000” entre as barras duplas e o GUID. A ID inteira deve ser `api://localhost:3000/{App ID GUID}`. (A parte do domínio do nome do **Escopo** logo abaixo do **URI da ID do aplicativo** será automaticamente alterada para que haja correspondência. Ela deve ser assim: `api://localhost:3000/{App ID GUID}/access_as_user`.)

1. Esta etapa e a seguinte concede, ao aplicativo host do Office, o acesso ao seu suplemento. Na seção **Aplicativos pré-autorizados**, você identifica os aplicativos que deseja autorizar para o aplicativo da Web do seu suplemento. Cada uma das seguintes IDs precisa ser pré-autorizada. Cada vez que você inserir uma, uma nova caixa de texto vazia aparece. (Insira apenas o GUID.)

 * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
 * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
 * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online) 

1. Abra o menu suspenso do **Escopo** ao lado de cada **ID do aplicativo** e marque a caixa para `api://localhost:44355/{App ID GUID}/access_as_user`.

1. Próximo ao topo da seção **Plataformas**, clique em **Adicionar Plataforma** novamente e selecione **Web**.

1. Na nova seção **Web** em **Plataformas**, insira o seguinte como uma **URL de redirecionamento**: `https://localhost:3000`. 

    > Observação: Até o presente momento, a plataforma **API Web** às vezes desaparece da seção **Plataformas**, especialmente se a página for atualizada depois que a plataforma **Web** é adicionada *e a página de registro é salva*. Para garantir que sua plataforma **API Web** ainda faz parte do registro, clique no botão **Editar Manifesto do Aplicativo** próximo à parte inferior da página. Você deve ver a cadeia de caracteres `api://localhost:3000/{App ID GUID}` na propriedade **identifierUris** do manifesto. Também haverá uma propriedade **oauth2Permissions** cuja subpropriedade **value** tem o valor `access_as_user`.

1. Role para baixo até a seção **Permissões do Microsoft Graph**, na subseção **Permissões Delegadas**. Use o botão **Adicionar** para abrir a caixa de diálogo **Selecionar Permissões**.

1. Na caixa de diálogo, marque as caixas das seguintes permissões: 
    * Files.Read.All
    * profile

1. Clique em **OK** no final da caixa de diálogo.

1. Clique em **Salvar** na parte inferior da página de registro.

## <a name="grant-admin-consent-to-the-add-in"></a>Conceder consentimento de administrador para o suplemento

> **Observação:** Este procedimento só é necessário quando você está desenvolvendo o suplemento. Quando o seu suplemento de produção é implantado na Loja do Office ou em um catálogo de suplementos, os usuários confiarão individualmente nele ao instalá-lo.

1. Na cadeia de caracteres a seguir, substitua o espaço reservado "{application_ID}" pela ID do Aplicativo que você copiou quando registrou seu suplemento.

    `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. Cole a URL resultante na barra de endereços do navegador e acesse-a.

1. Quando for solicitado, entre com as credenciais de administrador em sua locação do Office 365.

1. Em seguida, será solicitado que você conceda permissão para seu suplemento acessar os dados do Microsoft Graph. Clique em **Aceitar**. 

1. A guia ou janela do navegador é, então, redirecionada para a **URL de redirecionamento** que você especificou ao registrar o suplemento. Portanto, se o suplemento estiver sendo executado, a página inicial do suplemento abrirá no navegador. Se o suplemento não estiver em execução, você receberá um erro informando que o recurso no localhost:3000 não pode ser encontrado ou aberto. *Mas o fato de ter ocorrido a tentativa de redirecionamento significa que o processo de consentimento de administração foi concluído com êxito*. Assim, independentemente de se a página inicial abriu ou se você recebeu o erro, prossiga para a próxima etapa.

2. Na barra de endereços do navegador, você verá um parâmetro de consulta de locatário "tenant" com um valor GUID. Esta é a ID da sua locação do Office 365. Copie e salve esse valor. Você irá usá-lo em uma etapa posterior.

3. Feche a janela ou a guia.

## <a name="configure-the-add-in"></a>Configurar o suplemento

1. Em seu editor de códigos, abra o arquivo src\server.ts. Perto da parte superior, há uma chamada para um construtor de uma classe `AuthModule`. Há alguns parâmetros de cadeia de caracteres no construtor aos quais você precisa atribuir valores.

2. Na propriedade `client_id`, substitua o espaço reservado `{client GUID}` pela ID do aplicativo que você salvou ao registrar o suplemento. Quando terminar, deverá haver apenas um GUID entre aspas simples. Não deverá haver nenhum caractere "{}"

3. Na propriedade `client_secret`, substitua o espaço reservado `{client secret}` pelo segredo do aplicativo que você salvou ao registrar o suplemento.

4. Na propriedade `audience`, substitua o espaço reservado `{audience GUID}` pela ID do aplicativo que você salvou ao registrar o suplemento. (Exatamente o mesmo valor que você atribuiu à propriedade `client_id`.)
  
3. Na cadeia de caracteres atribuída à propriedade `issuer`, você verá o espaço reservado *{O365 tenant GUID}*. Substitua-o pela ID de locação do Office 365 que você salvou no final do último procedimento. Se por algum motivo, você não obteve a ID anteriormente, use um dos métodos descritos em [Localizar a ID de locatário do Office 365](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) para obtê-la. Quando terminar, o valor da propriedade `issuer` deve ser algo parecido com isto:

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

1. Substitua o espaço reservado "{application_GUID here}" *nos dois lugares* na marcação pela ID do Aplicativo que você copiou ao registrar seu suplemento. (O símbolo "{}" não faz parte da ID, portanto não o inclua.) Essa é a mesma ID usada para a ClientID e a Audience no web.config.

    >Observação: 
    >
    >* O valor de **Resource** é o **URI da ID do Aplicativo** que você definiu quando adicionou a plataforma API Web no registro do suplemento.
    >* A seção **Scopes** só será usada para gerar uma caixa de diálogo de consentimento se o suplemento for vendido na Office Store.

1. Salve e feche o arquivo.

## <a name="code-the-client-side"></a>Codificar o lado do cliente

1. Abra o arquivo program.js da pasta **public**. Ele já apresenta alguns códigos:

    * Uma atribuição ao método `Office.initialize` que, por sua vez, atribui um manipulador ao evento clicar do botão `getGraphAccessTokenButton`.
    * Um método `showResult` que exibirá os dados retornados do Microsoft Graph (ou uma mensagem de erro) na parte inferior do painel de tarefas.

1. Abaixo da atribuição a `Office.initialize`, adicione o código a seguir. Observe o seguinte sobre este código: 

    * A função `getDataWithoutAuthChallenge` é chamada em uma primeira tentativa de usar o fluxo Em Nome De. O pressuposto é que a autenticação de um único fator é tudo o que é necessário. Você adicionará o código em uma etapa posterior para lidar com o caso em que a autenticação multi-fator é necessária.
    * O `getAccessTokenAsync` é a nova API no Office.js que permite que um suplemento solicite ao aplicativo host do Office (Excel, PowerPoint, Word, etc.) um token de acesso para o suplemento (para o usuário conectado ao Office). O aplicativo host do Office, por sua vez, solicita o token ao ponto de extremidade 2.0 do Azure AD. Uma vez que você previamente autorizou o host do Office para o seu suplemento ao registrá-lo, o Azure AD enviará o token. 
     * Se nenhum usuário estiver conectado ao Office, o host do Office solicitará que o usuário se conecte. 
     * O parâmetro de opções configura o `forceConsent` como falso. Dessa forma, não será solicitado que o usuário consinta o acesso ao host do Office para seu suplemento.

    ```js
    function getOneDriveItems() {
        getDataWithoutAuthChallenge();
    }   
    
    function getDataWithoutAuthChallenge() {       
        Office.context.auth.getAccessTokenAsync({forceConsent: false},
            function (result) {
                if (result.status === "succeeded") {
                    // TODO1: Use the access token to get Microsoft Graph data.
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

1. Substitua TODO1 pelas linhas a seguir. Você criará o método `getData` e a rota “/api/onedriveitems” do lado do servidor nas etapas posteriores. Uma URL relativa é usada para o ponto de extremidade porque ela deve ser hospedada no mesmo domínio que seu suplemento.

    ```
    accessToken = result.value;
    getData("/api/onedriveitems", accessToken);
    ```

1. Abaixo do método `getOneDriveFiles`, adicione o seguinte. Este método utilitário chama um ponto de extremidade da API Web especificado e transmite a ela o mesmo token de acesso que aplicativo host do Office usou para obter acesso ao seu suplemento. No lado do servidor, esse token de acesso será usado no fluxo "on behalf of" (em nome de) para obter um token de acesso para o Microsoft Graph. 

    ```
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET",
        })
        .done(function (result) {
            TODO2: Display data and handle demand for multi-factor authentication.
        })
        .fail(function (result) {
            console.log(result.error);
       });
    }
    ```

1. Substitua TODO2 pelo código a seguir. Sobre este código, observe:
    * Se o objetivo do Microsoft Graph solicitar fator(es) de autenticação adicional(ais), o resultado não será dados. Será um JSON de Declarações dizendo ao AAD quais fatores adicionais o usuário deve fornecer. Nesse caso, o cliente deve iniciar um novo logon que transfere essa sequência de Declarações para o AAD, a fim de que este último forneça os prompts necessários.
    * Se o resultado for o JSON de Declarações, então, ele conterá a sequência "capolids".
    * Você criará a função `getDataUsingAuthChallenge` em uma última etapa.

    ```
    if (result[0].indexOf('capolids') !== -1) {                
        result[0] = JSON.parse(result[0])
        getDataUsingAuthChallenge(result[0]);
    } else {  
        showResult(result);
    }
    ```

1. Adicione a seguinte função ao arquivo logo abaixo da função `getData`. Sobre esta função, nota:
    * A função é usada quando o AAD solicitou fator(es) de autenticação adicional(ais). 
    * A função aciona um segundo logon no qual o usuário será solicitado a fornecer fator(es) de autenticação adicional(ais). 
    * A opção `authChallenge` contém uma sequência que informa ao AAD qual(quais) fator(es) ele deve solicitar. O host do Office transfere essa sequência para o AAD quando ele solicita o token de suplemento ao seu suplemento.

    ```
    function getDataUsingAuthChallenge(authChallengeString) {       
        Office.context.auth.getAccessTokenAsync({authChallenge: authChallengeString},
            function (result) {
                if (result.status === "succeeded") {
                    accessToken = result.value;
                    getData("/api/onedriveitems", accessToken);
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

1. Salve e feche o arquivo.

## <a name="code-the-server-side"></a>Codificar o lado do servidor

Há dois arquivos do lado do servidor que precisam ser modificados. 
- O src\auth.js fornece funções auxiliares de autorização. Ele já tem membros genéricos que são usados em uma variedade de fluxos de autorização. É preciso adicionar funções a esse arquivo para implementar o fluxo "on behalf of".
- O arquivo de src\server.js tem os membros básicos necessários para executar um servidor e o middleware do express. É necessário adicionar funções a ele que ajudam a API Web e a página inicial a obterem os dados do Microsoft Graph.

### <a name="create-a-method-to-exchange-tokens"></a>Criar um método para troca de tokens

1. Abra o arquivo \src\auth.ts. Adicione o método abaixo à classe `AuthModule`. Observe o seguinte sobre este código:
    * O parâmetro jwt é o token de acesso ao aplicativo. No fluxo de "on behalf of" (em nome de), ele é trocado com AAD por um token de acesso ao recurso.
    * O parâmetro scopes (escopos) tem um valor padrão, mas neste exemplo será substituído pelo código de chamada.
    * O parâmetro de recurso é opcional. Não deve ser usado quando o STS é o ponto de extremidade V2 do AAD. Este último infere o recurso dos escopos e retorna um erro se um recurso é enviado na Solicitação HTTP. 
    

    ```
    private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        try {
            // TODO3: Construct the parameters that will be sent in the body of the 
            //        HTTP Request to the STS that starts the "on behalf of" flow.
            // TODO4: Send the request to the STS.
            // TODO5: Process the response and persist the access token to resource.
        }
        catch (exception) {
            throw new UnauthorizedError('Unable to obtain an access token to the resource' 
                                        + JSON.stringify(exception), 
                                        exception);
        }
    }
    ```

2. Substitua TODO3 pelo código a seguir. Sobre este código, observe:
    * Um STS com suporte para o fluxo "on behalf of" espera determinados pares de valor/propriedade no corpo da solicitação HTTP. Esse código constrói um objeto que se tornará o corpo da solicitação. 
    * Uma propriedade de recurso é adicionada ao corpo se, e somente se, um recurso é transmitido para o método.

    ```
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

3. Substitua TODO4 pelo código a seguir que envia a solicitação HTTP para o ponto de extremidade do token do STS.

    ```
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    }); 
    ```

4. Substitua TODO5 pelo código a seguir. Observe que o código persiste no token de acesso ao recurso, e é a hora de expiração, além de retorná-lo. O código de chamada pode evitar chamadas desnecessárias ao STS reutilizando um token de acesso não expirado ao recurso. Você verá como fazer isso na próxima seção.

    ```
    if (res.status !== 200) {
        TODO6: Handle failure and the case where AAD asks for additional
               authentication factors.
    }
    const json = await res.json();
    // Persist the token and it's expiration time.
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken; 
    ```

5. Substitua TODO6 pelo código a seguir. Sobre este código, observe:

    * Existem configurações do Azure Active Directory nas quais o usuário precisa fornecer fator(es) de autenticação adicional(ais) para acessar alguns objetivos do Microsoft Graph (por exemplo, o OneDrive), mesmo que o usuário possa fazer login no Office apenas com uma senha. Nesse caso, o AAD enviará uma resposta que tenha uma propriedade `Claims`. 
    * Este valor `Claims` precisa ser passado de volta para o cliente, que deve iniciar um segundo login para o usuário e incluir o valor `Claims` na chamada para o AAD. O AAD solicitará ao usuário que forneça o(s) fator(es) adicional(ais).
    * Por precaução, o código limpa o cache de todos os tokens de acesso que foram obtidos quando o usuário fez o login com apenas uma senha.  

    ```
    const exception = await res.json();
    // Check if AAD is the STS.
    if (this.stsDomain === 'https://login.microsoftonline.com') {
        if (JSON.stringify(exception.claims)) {                       
            ServerStorage.clear();
            return JSON.stringify(exception.claims);    
        } else {                    
            throw exception;
        }
    }
    else {                    
        throw exception;
    }
    ```

5. Salve o arquivo, mas não o feche.

### <a name="create-a-method-to-get-access-to-the-resource-using-the-on-behalf-of-flow"></a>Criar um método para obter acesso ao recurso usando o fluxo "on behalf of"

1. Ainda no arquivo src/auth.ts, adicione o método abaixo à classe `AuthModule`. Observe o seguinte sobre este código:
    * Os comentários acima sobre os parâmetros para o método `exchangeForToken` aplicam-se aos parâmetros deste método também.
    * O método primeiro verifica o armazenamento persistente para um token de acesso ao recurso que não expirou e não vai expirar no próximo minuto. Ele chama o método `exchangeForToken` que você criou na última seção somente se necessário.

    ```
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

### <a name="create-the-endpoints-that-will-serve-the-add-ins-home-page-and-data"></a>Criar os pontos de extremidade que servirão aos dados e à página inicial do suplemento

1. Abra o arquivo src\server.ts. 

2. Adicione o método a seguir na parte inferior do arquivo. Esse método servirá à página inicial do suplemento. O manifesto do suplemento especifica a URL da página inicial.

    ```
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    })); 
    ```

3. Adicione o método a seguir na parte inferior do arquivo. Este método lidará com todas as solicitações para a API `onedriveitems`.
    ```
    app.get('/api/onedriveitems', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token 
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage 
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Send to the client only the data that it actually needs.
    })); 
    ```

4. Substitua TODO7 pelo seguinte código que valida o token de acesso recebido do aplicativo host do Office. O método `verifyJWT` é definido no arquivo src\auth.ts. Ele sempre valida a audiência e o emissor. Usamos o parâmetro opcional para especificar que também desejamos que ele verifique se o escopo no token de acesso é `access_as_user`. Esta é a única permissão ao suplemento que o usuário e o host do Office precisam para obter um token de acesso para o Microsoft Graph por meio do fluxo "on behalf of". 

    ```
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    ```

> **Observação:** Você deve usar apenas o escopo `access_as_user` para autorizar a API que lida com o fluxo Em Nome De para os suplementos do Office. Outras APIs em seu serviço devem ter seus próprios requisitos de escopo. Isso limita o que pode ser acessado com os tokens que o Office adquire.

5. Substitua TODO8 pelo código a seguir. Observe o seguinte sobre este código:

    * A chamada para `acquireTokenOnBehalfOf` não inclui um parâmetro de recurso porque construímos o objeto `AuthModule` (`auth`) com o ponto de extremidade V2.0 do AAD que não oferece suporte à propriedade de recurso.
    * O segundo parâmetro da chamada especifica as permissões que o suplemento precisará para obter uma lista dos arquivos e das pastas do usuário no OneDrive. (A permissão `profile` não é solicitada, porque só é necessária quando o host do Office obtém o token de acesso ao seu suplemento, e não quando você está negociando nesse token para um token de acesso para o Microsoft Graph.)
    * Se a resposta for uma sequência contendo 'capolids', então trata-se de uma mensagem de declarações do AAD informando que a autenticação multi-fator é necessária. A mensagem é passada para o cliente, que a usa para iniciar um segundo login. A sequência informa ao AAD qual(quais) fator(es) de autenticação adicional(ais) o usuário deve fornecer.

    ```
    let graphToken = null;
    const tokenAcquisitionResponse = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    if (tokenAcquisitionResponse.includes('capolids')) {
        const claims: string[] = [];
        claims.push(tokenAcquisitionResponse);
        return res.json(claims);
    } else {
        // The response is the token to Microsoft Graph itself. Rename it so remaining code
        // is self-documenting.
        graphToken = tokenAcquisitionResponse;
    }
    ```

6. Substitua TODO9 pela seguinte linha. Observe o seguinte sobre este código:

    * A classe MSGraphHelper é definida no src\msgraph-helper.ts. 
    * Podemos minimizar os dados que devem ser retornados especificando que só queremos a propriedade de nome e somente os três primeiros itens.

    `const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");`

7. Substitua TODO10 pelo código a seguir. Observe que o Microsoft Graph retorna alguns metadados OData e uma propriedade **eTag** para cada item, mesmo se `name` é a única propriedade solicitada. O código envia somente os nomes de item para o cliente.

    ```
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

8. Salve e feche o arquivo.

## <a name="deploy-the-add-in"></a>Implantar o suplemento

Agora é preciso informar ao Office onde encontrar o suplemento.

1. Crie um compartilhamento de rede ou [compartilhe uma pasta na rede](https://technet.microsoft.com/en-us/library/cc770880.aspx).

2. Coloque uma cópia do arquivo de manifesto Office-Add-in-NodeJS-SSO.xml, da raiz do projeto, dentro da pasta compartilhada.

3. Inicie o PowerPoint e abra um documento.

4. Escolha a guia **Arquivo** e, então, **Opções**.

5. Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.

6. Escolha **Catálogos de Suplementos Confiáveis**.

7. No campo **URL do Catálogo**, insira o caminho de rede para o compartilhamento de pasta que contém o arquivo Office-Add-in-NodeJS-SSO.xml e escolha **Adicionar Catálogo**.

8. Selecione a caixa de seleção **Mostrar no Menu** e, em seguida, escolha **OK**.

9. Uma mensagem será exibida para informá-lo de que suas configurações serão aplicadas na próxima vez que você iniciar o Microsoft Office. Feche o PowerPoint.

## <a name="build-and-run-the-project"></a>Criar e executar o projeto

Há duas maneiras de criar e executar o projeto dependendo se você estiver ou não usando o Visual Studio Code. Em ambas as maneiras, o projeto cria e recria automaticamente e entra novamente em execução quando você faz alterações no código.

1. Se não estiver usando o Visual Studio Code: 
 1. Abra um nó terminal e vá até a pasta raiz do projeto.
 2. No terminal, insira **npm run build**. 
 3. Abra um segundo nó terminal e vá até a pasta raiz do projeto.
 4. No terminal, insira **npm run start**.

2. Se estiver usando o VS Code:
 1. Abra o projeto no VS Code.
 2. Pressione Ctrl+Shift+B para compilar o projeto.
 3. Pressione F5 para executar o projeto em uma sessão de depuração.


## <a name="add-the-add-in-to-an-office-document"></a>Adicionar o suplemento em um documento do Office

1. Reinicie o PowerPoint, abra ou crie uma apresentação. 

2. Na guia **Desenvolvedor** no PowerPoint, escolha **Meus Suplementos**.

3. Selecione a guia **PASTA COMPARTILHADA**.

4. Escolha **Exemplo de SSO NodeJS**e selecione **OK**.

5. Na faixa de opções **Página Inicial**, há um novo grupo chamado **SSO NodeJS** com um botão com o rótulo **Mostrar Suplemento** e um ícone. 

## <a name="test-the-add-in"></a>Testar o suplemento

1. Certifique-se de ter alguns arquivos no seu OneDrive para que você possa verificar os resultados.

2. Clique no botão **Exibir Suplemento** para abrir o suplemento.

2. O suplemento é aberto na página inicial. Clique no botão **Obter meus arquivos do OneDrive**.

2. Se você estiver conectado ao Office, será exibida uma lista de seus arquivos e suas pastas no OneDrive, abaixo do botão. Isso poderá demorar mais de 15 segundos na primeira vez.

3. Se você não tiver entrado no Office, um pop-up será aberto e pedirá que você entre. Depois de concluir a entrada, a lista de arquivos e pastas aparecerá após alguns segundos. *Não pressione o botão uma segunda vez.*
> **Observação:** Se você entrou no Office com uma ID diferente e se alguns aplicativos do Office que estavam abertos no momento continuam abertos, o Office pode não alterar de forma confiável sua ID, mesmo que pareça ter feito isso no PowerPoint. Se isso acontecer, a chamada para o Microsoft Graph pode falhar ou os dados da ID anterior podem ser retornados. Para evitar isso, certifique-se de *fechar todos os outros aplicativos do Office* antes de pressionar **Obter meus arquivos do OneDrive**.
