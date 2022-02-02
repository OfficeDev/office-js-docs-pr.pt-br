---
title: Usar o SSO para obter a identidade do usuário in-locar
description: Chame a API getAccessToken para obter o token de ID com nome, email e informações adicionais sobre o usuário instado.
ms.date: 01/25/2022
localization_priority: Normal
ms.openlocfilehash: 2c9b3c89a154d624f99e196014c7d8024286d927
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/02/2022
ms.locfileid: "62322330"
---
# <a name="use-sso-to-get-the-identity-of-the-signed-in-user"></a>Usar o SSO para obter a identidade do usuário in-locar

Use a `getAccessToken` API para obter um token de acesso que contém a identidade do usuário atual Office. O token de acesso também é um token de ID porque contém declarações de identidade sobre o usuário inscrevado, como seu nome e email. Você também pode usar o token de ID para identificar o usuário ao chamar seus próprios serviços Web. Para chamar`getAccessToken`, você deve configurar seu Office de usuário para usar o SSO com Office.

Neste artigo, você criará um Office que obtém o token de ID e exibe o nome, o email e a ID exclusiva do usuário no painel de tarefas.

> [!NOTE]
> O SSO Office e `getAccessToken` a API não funcionam em todos os cenários. Sempre implemente uma caixa de diálogo de fallback para entrar no usuário quando o SSO não estiver disponível. Para obter mais informações, [consulte Authenticate and authorize with the Office dialog API](auth-with-office-dialog-api.md).

## <a name="create-an-app-registration"></a>Criar um registro de aplicativo

Para usar o SSO com Office, você precisa criar um registro de aplicativo no portal do Azure para que o plataforma de identidade da Microsoft possa fornecer serviços de autenticação e autorização para seu Office Add-in e seus usuários.

1. Para registrar seu aplicativo, acesse a página [Portal do Azure - Registros de aplicativos](https://go.microsoft.com/fwlink/?linkid=2083908) .

1. Entre com as **_credenciais de_** administrador no seu Microsoft 365 de adoção. Por exemplo, MeuNome@contoso.onmicrosoft.com.

1. Selecione **Novo registro**. Na página **Registrar um aplicativo**, defina os valores da seguinte forma.

   - Defina **Nome** para `Office-Add-in-SSO`.
   - Defina **Tipos de conta com suporte** para **Contas em qualquer diretório organizacional e contas pessoais da Microsoft (por exemplo, Skype, Xbox, Outlook.com)**.
   - De definir o tipo de aplicativo como **Web** e, em seguida, definir **URI de redirecionamento** como `https://localhost:[port]/dialog.html`. Substitua `[port]` pelo número de porta correto para seu aplicativo Web. Se você criou o complemento usando yo office, o número da porta normalmente é 3000 e encontrado no arquivo package.json. Se você criou o add-in com Visual Studio 2019, a porta será encontrada na propriedade **URL SSL** do projeto Web.
   - Escolha **Registrar**.

1. Na página **Office-Add-in-SSO**, copie e salve os valores da **ID** do Aplicativo (cliente) e da **ID de Diretório (locatário**). Use ambos os valores nos procedimentos posteriores.

   > [!NOTE]
   > Essa **ID de Aplicativo (cliente)** é o valor "público" quando outros aplicativos, como o aplicativo cliente Office (por exemplo, PowerPoint, Word, Excel), procuram acesso autorizado ao aplicativo. Também é a "ID do cliente" do aplicativo quando ela, por sua vez, busca acesso autorizado ao Microsoft Graph.

1. Selecione **Autenticação** em **Gerenciar**. Na seção **Concessão implícita** , habilita as caixas de seleção para **token do Access** e **token de ID**.

1. Na parte superior da página, selecione **Salvar**.

1. Selecionar **Expor uma API** em **Gerenciar**. Selecione o link **Definir** . Isso gerará o URI de ID do aplicativo no formato `api://[app-id-guid]`, onde `[app-id-guid]` é a **ID do Aplicativo (cliente**).

1. Na ID gerada, insira `localhost:[port]/` (observe a barra de avanço "/" anexada ao final) entre as barras de avanço duplo e o GUID. Substitua `[port]` pelo número de porta correto para seu aplicativo Web. Se você criou o complemento usando yo office, o número da porta normalmente é 3000 e encontrado no arquivo package.json. Se você criou o add-in com Visual Studio 2019, a porta será encontrada na propriedade **URL SSL** do projeto Web.
   Quando terminar, a ID inteira deverá ter o formulário `api://localhost:[port]/[app-id-guid]`; por exemplo `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

1. Selecione o botão **Adicionar um escopo**. No painel que se abre, insira `access_as_user` como o **Nome de escopo**.

1. Definir **Quem pode consentir?** aos **Administradores e usuários**.

1. Preencha os campos para configurar os prompts de consentimento do administrador e do usuário com valores apropriados `access_as_user` para o escopo que permite que o aplicativo cliente Office use as APIs da Web do seu complemento com os mesmos direitos do usuário atual. Sugestões:

   - **Nome de exibição de consentimento do** administrador: Office pode atuar como usuário.
   - **Descrição de autorização de administrador:** Permite ao Office chamar os APIs de suplemento da web com os mesmos direitos que o usuário atual.
   - **Nome de exibição de consentimento do** usuário: Office pode agir como você.
   - **Descrição do** consentimento do usuário: Office para chamar as APIs da Web do complemento com os mesmos direitos que você tem.

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

   a. Selecione **Adicionar um botão de aplicativo** cliente e, no painel que será aberto, `[app-id-guid]` desmarcar a ID do aplicativo (cliente) e marque a caixa para `api://localhost:44355/[app-id-guid]/access_as_user`.

   b. Selecione **Adicionar aplicativo**.

1. Selecione **Permissões para API** em **Gerenciar** e selecione **Adicionar uma permissão**. No painel que se abre, escolha **Microsoft Graph** e, em seguida, escolha **Permissões delegadas**.

1. Use a caixa de pesquisa **Selecionar permissões** para procurar as permissões que o seu suplemento precisa. Pesquise e selecione a **permissão de** perfil. A `profile` permissão é necessária para que o aplicativo Office para obter um token para seu aplicativo Web de complemento.

   - perfil

   > [!NOTE]
   > A permissão `User.Read` pode já estar listada por padrão. É uma boa prática não pedir permissões desnecessárias, por isso recomendamos desmarcar a caixa para essa permissão se o suplemento não precisar dela.

1. Selecione a **botão Adicionar** seleção na parte inferior do painel.

1. Na mesma página, escolha o botão **Conceder consentimento do \<tenant-name\>** administrador e selecione **Sim** para a confirmação exibida.

## <a name="create-the-office-add-in"></a>Criar o Suplemento do Office

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Inicie Visual Studio 2019 e escolha **Criar um novo projeto**.
1. Pesquise e selecione **o modelo de projeto Excel web add-in**. Depois clique em **Próximo**. Observação: o SSO funciona com qualquer aplicativo Office, mas para este artigo funcionará com Excel.
1. Insira um nome de projeto, como **sso-display-user-info** e escolha **Criar**. Você pode deixar os outros campos em valores padrão.
1. Na caixa **de diálogo Escolher o tipo de** complemento, selecione **Adicionar nova funcionalidade** ao Excel e escolha **Concluir**.

O projeto é criado e conterá dois projetos na solução.

- **sso-display-user-info**: contém o manifesto e os detalhes para o sideload do Excel.
- **sso-display-user-infoWeb**: o projeto ASP.NET que hospeda as páginas da Web para o complemento.

# <a name="yo-office"></a>[yo office](#tab/yooffice)

Certifique-se de configurar [seu ambiente de desenvolvimento](../overview/set-up-your-dev-environment.md).

1. Para criar o projeto, digite o seguinte comando.

   ```command line
   yo office --projectType taskpane --name 'sso-display-user-info' --host excel --js true
   ```

O projeto é criado em uma nova pasta chamada **sso-display-user-info**.

---

## <a name="configure-the-manifest"></a>Configurar o manifesto

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. No **Explorador de Soluções** , abra **sso-display-user-info > sso-display-user-infoManifest > sso-display-user-info.xml**

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. Em Visual Studio código, abra o **arquivomanifest.xml**.

---

1. Perto da parte inferior do manifesto há um elemento de `</Resources>` fechamento. Insira o XML a seguir logo abaixo do `</Resources>` elemento, mas antes do elemento de `</VersionOverrides>` fechamento. Para Office outros Outlook, adicione a marcação ao final da `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` seção. Para o Outlook, adicione a marcação no final da seção `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.

   ```xml
   <WebApplicationInfo>
       <Id>[application-id]</Id>
       <Resource>api://localhost:[port]/[application-id]</Resource>
       <Scopes>
           <Scope>openid</Scope>
           <Scope>user.read</Scope>
           <Scope>profile</Scope>
       </Scopes>
   </WebApplicationInfo>
   ```

1. Substitua `[port]` pelo número de porta correto do seu projeto. Se você criou o complemento usando yo office, o número da porta normalmente é 3000 e encontrado no arquivo package.json. Se você criou o add-in com Visual Studio 2019, a porta será encontrada na propriedade **URL SSL** do projeto Web.
1. Substitua ambos `[application-id]` os espaço reservados pela ID do aplicativo real do registro do aplicativo.
1. Salve o arquivo.

O XML inserido contém os seguintes elementos e informações.

- **WebApplicationInfo** – o pai dos seguintes elementos.
- **ID** - O ID do cliente do suplemento Este é um ID do aplicativo que você obtém como parte do registro do suplemento. Confira [Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0](register-sso-add-in-aad-v2.md).
- **Resource** – A URL do suplemento. Esse é o mesmo URI (incluindo o protocolo `api:`) que você usou ao registrar o suplemento no AAD. Parte do domínio deste URI deve corresponder ao domínio, incluindo quaisquer subdomínios, usados nos URLs na seção `<Resources>` do manifesto do suplemento e o URI deve terminar com o ID do cliente no `<Id>`.
- **Scopes** – O pai de uma ou mais elementos **Scope**.
- **Scope** – Especifica uma permissão que seu suplemento precisa para o AAD. As permissões `profile` e `openID` são sempre necessárias e podem ser as únicas permissões necessárias, se o suplemento não acessar o Microsoft Graph. Se isso acontecer, você também precisa de elementos **Escopo** para as permissões necessárias do Microsoft Graph; por exemplo, `User.Read`, `Mail.Read`. Bibliotecas que você usa no seu código para acessar o Microsoft Graph pode precisar de permissões adicionais. Por exemplo, a biblioteca de autenticação da Microsoft (MSAL) para .NET requer a permissão `offline_access`. Para saber mais, confira [autorizar o Microsoft Graph de um suplemento do Office](authorize-to-microsoft-graph.md).

## <a name="add-the-jwt-decode-package"></a>Adicionar o pacote jwt-decode

Você pode chamar a `getAccessToken` API para obter o token de ID Office. Primeiro, permite adicionar o pacote jwt-decode para facilitar a decodificar e exibir o token de ID.

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Abra a Visual Studio solução.
1. No menu, escolha **Ferramentas > NuGet Gerenciador de Pacotes > Gerenciador de Pacotes Console**.
1. Insira o seguinte comando no **console Gerenciador de Pacotes.**

   `Install-Package jwt-decode -Projectname sso-display-user-infoWeb`

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. Em uma janela de terminal/console, vá para a pasta raiz do projeto do seu complemento.
1. Insira o seguinte comando

   `npm install jwt-decode`

---

## <a name="add-ui-to-the-task-pane"></a>Adicionar interface do usuário ao painel de tarefas

Precisamos modificar o painel de tarefas para que ele possa exibir as informações do usuário que obteremos do token de ID.

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Abra o Home.html arquivo.
1. Adicione a seguinte marca de script à `<head>` seção da página. Isso incluirá o pacote jwt-decode que adicionamos anteriormente.

   ```html
   <script src="Scripts/jwt-decode-2.2.0.js" type="text/javascript"></script>
   ```

1. Substitua a `<body>` seção pelo SEGUINTE HTML.

   ```html
   <body>
     <h1>Welcome</h1>
     <p>
       Sign in to Office, then choose the <b>Get ID Token</b> button to see your
       ID token information.
     </p>
     <button id="getIDToken">Get ID Token</button>
     <div>
       <span id="userInfo"></span>
     </div>
   </body>
   ```

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. Abra o **arquivo src/taskpane/taskpane.html** .
1. Substitua a `<body>` seção pelo SEGUINTE HTML.

   ```html
   <body>
     <h1>Welcome</h1>
     <p>
       Sign in to Office, then choose the <b>Get ID Token</b> button to see your
       ID token information.
     </p>
     <button id="getIDToken">Get ID Token</button>
     <div>
       <span id="userInfo"></span>
     </div>
   </body>
   ```

---

## <a name="call-the-getaccesstoken-api"></a>Chamar a API getAccessToken

A etapa final é obter o token de ID chamando `getAccessToken`.

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Abra o **Home.js** arquivo.
1. Substitua todo o conteúdo do arquivo pelo código a seguir.

   ```javascript
   (function () {
     "use strict";

     // The initialize function must be run each time a new page is loaded.
     Office.initialize = function (reason) {
       $(document).ready(function () {
         $("#getIDToken").click(getIDToken);
       });
     };

     async function getIDToken() {
       try {
         let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
           allowSignInPrompt: true,
         });
         let userToken = jwt_decode(userTokenEncoded);
         document.getElementById("userInfo").innerHTML =
           "name: " +
           userToken.name +
           "<br>email: " +
           userToken.preferred_username +
           "<br>id: " +
           userToken.oid;
         console.log(userToken);
       } catch (error) {
         document.getElementById("userInfo").innerHTML =
           "An error occurred. <br>Name: " +
           error.name +
           "<br>Code: " +
           error.code +
           "<br>Message: " +
           error.message;
         console.log(error);
       }
     }
   })();
   ```

1. Salve o arquivo.

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. Abra o **arquivo src/taskpane/taskpane.js** .
1. Substitua todo o conteúdo do arquivo pelo código a seguir.

   ```javascript
   import jwt_decode from "jwt-decode";

   Office.onReady((info) => {
     if (info.host === Office.HostType.Excel) {
       document.getElementById("getIDToken").onclick = getIDToken;
     }
   });

   async function getIDToken() {
     try {
       let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
         allowSignInPrompt: true,
       });
       let userToken = jwt_decode(userTokenEncoded);
       document.getElementById("userInfo").innerHTML =
         "name: " +
         userToken.name +
         "<br>email: " +
         userToken.preferred_username +
         "<br>id: " +
         userToken.oid;
       console.log(userToken);
     } catch (error) {
       document.getElementById("userInfo").innerHTML =
         "An error occurred. <br>Name: " +
         error.name +
         "<br>Code: " +
         error.code +
         "<br>Message: " +
         error.message;
       console.log(error);
     }
   }
   ```

1. Salve o arquivo.

---

## <a name="run-the-add-in"></a>Execute o suplemento

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Escolha **Depurar > Iniciar a Depuração** ou pressione **F5**.

# <a name="yo-office"></a>[yo office](#tab/yooffice)

Execute `npm start` a partir da linha de comando.

---

1. Quando Excel, entre no Office com a mesma conta de locatário que você usou para criar o registro do aplicativo.
1. Na faixa **de opções Home** , escolha **Mostrar o Taskpane** para abrir o complemento.
1. No painel de tarefas do complemento, escolha **Obter token de ID**.

O complemento exibirá o nome, o email e a ID da conta com a que você se inscreveu.

> [!NOTE]
> Se você encontrar algum erro, revise as etapas de registro neste artigo para o registro do aplicativo. Perder um detalhe ao configurar o registro do aplicativo é uma causa comum de problemas de trabalho com o SSO. Se você ainda não conseguir fazer com que o add-in seja executado com êxito, consulte Solucionar problemas de mensagens de erro para [SSO (login único).](troubleshoot-sso-in-office-add-ins.md).

## <a name="see-also"></a>Confira também

[Usando declarações para identificar confiávelmente um usuário (ID de assunto e objeto)](/azure/active-directory/develop/id-tokens#using-claims-to-reliably-identify-a-user-subject-and-object-id)
