---
title: Usar o SSO para obter a identidade do usuário conectado
description: Chame a API getAccessToken para obter o token de ID com nome, email e informações adicionais sobre o usuário conectado.
ms.date: 02/16/2022
localization_priority: Normal
ms.openlocfilehash: 5416c469a15d7eda9333f511c61e2cff1a901018
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660064"
---
# <a name="use-sso-to-get-the-identity-of-the-signed-in-user"></a>Usar o SSO para obter a identidade do usuário conectado

Use a `getAccessToken` API para obter um token de acesso que contém a identidade do usuário atual conectado ao Office. O token de acesso também é um token de ID porque contém declarações de identidade sobre o usuário conectado, como seu nome e email. Você também pode usar o token de ID para identificar o usuário ao chamar seus próprios serviços Web. Para ligar `getAccessToken` , você deve configurar seu Suplemento do Office para usar o SSO com o Office.

Neste artigo, você criará um Suplemento do Office que obtém o token de ID e exibe o nome, o email e a ID exclusiva do usuário no painel de tarefas.

> [!NOTE]
> O SSO com o `getAccessToken` Office e a API não funciona em todos os cenários. Sempre implemente uma caixa de diálogo de fallback para conectar o usuário quando o SSO não estiver disponível. Para obter mais informações, consulte [Autenticar e autorizar com a API de caixa de diálogo do Office](auth-with-office-dialog-api.md).

## <a name="create-an-app-registration"></a>Criar um registro de aplicativo

Para usar o SSO com o Office, você precisa criar um registro de aplicativo no portal do Azure para que o plataforma de identidade da Microsoft possa fornecer serviços de autenticação e autorização para seu Suplemento do Office e seus usuários.

1. Para registrar seu aplicativo, vá para a [portal do Azure - Registros de aplicativo](https://go.microsoft.com/fwlink/?linkid=2083908) página.

1. Entre com as **_credenciais de_** administrador no locatário do Microsoft 365. Por exemplo, MeuNome@contoso.onmicrosoft.com.

1. Selecione **Novo registro**. Na página **Registrar um aplicativo**, defina os valores da seguinte forma.

   - Defina **Nome** para `Office-Add-in-SSO`.
   - Defina **Tipos de conta com suporte** para **Contas em qualquer diretório organizacional e contas pessoais da Microsoft (por exemplo, Skype, Xbox, Outlook.com)**.
   - Defina o tipo de aplicativo como **Web** e defina **o URI de Redirecionamento** como `https://localhost:[port]/dialog.html`. Substitua `[port]` pelo número da porta correto para seu aplicativo Web. Se você criou o suplemento usando o yo office, o número da porta normalmente é 3000 e encontrado no arquivo package.json. Se você criou o suplemento com o Visual Studio 2019, a porta será encontrada na propriedade **de URL SSL** do projeto Web.
   - Escolha **Registrar**.

1. Na página **office-add-in-SSO** , copie e salve os valores para a **ID** do Aplicativo (cliente) e a **ID do Diretório (locatário**). Use ambos os valores nos procedimentos posteriores.

   > [!NOTE]
   > Essa **ID** de Aplicativo (cliente) é o valor de "público", quando outros aplicativos, como o aplicativo cliente do Office (por exemplo, PowerPoint, Word, Excel), buscam acesso autorizado ao aplicativo. Também é a "ID do cliente" do aplicativo quando ela, por sua vez, busca acesso autorizado ao Microsoft Graph.

1. Selecione **Autenticação** em **Gerenciar**. Na seção **Concessão implícita** , habilite as caixas de seleção para **token de acesso** e **token de ID**.

1. Na parte superior da página, selecione **Salvar**.

1. Selecionar **Expor uma API** em **Gerenciar**. Selecione **o link** Definir. Isso gerará o URI da ID do Aplicativo no formulário `api://[app-id-guid]`, onde está `[app-id-guid]` a **ID do Aplicativo (cliente**).

1. Na ID gerada, insira `localhost:[port]/` (observe a barra "/" acrescentada ao final) entre as barras duplas e o GUID. Substitua `[port]` pelo número da porta correto para seu aplicativo Web. Se você criou o suplemento usando o yo office, o número da porta normalmente é 3000 e encontrado no arquivo package.json. Se você criou o suplemento com o Visual Studio 2019, a porta será encontrada na propriedade **de URL SSL** do projeto Web.
   Quando terminar, a ID inteira deverá ter o formulário `api://localhost:[port]/[app-id-guid]`; por exemplo `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.

1. Selecione o botão **Adicionar um escopo**. No painel que é aberto, insira `access_as_user` como o **\<Scope\>** nome.

1. Definir **Quem pode consentir?** aos **Administradores e usuários**.

1. Preencha os campos para configurar os prompts de consentimento do administrador e do usuário com valores apropriados `access_as_user` para o escopo que permite que o aplicativo cliente do Office use as APIs Web do suplemento com os mesmos direitos que o usuário atual. Sugestões:

   - **Administração nome de exibição de consentimento**: o Office pode atuar como o usuário.
   - **Descrição de autorização de administrador:** Permite ao Office chamar os APIs de suplemento da web com os mesmos direitos que o usuário atual.
   - **Nome de exibição de consentimento do** usuário: o Office pode agir como você.
   - **Descrição de** consentimento do usuário: habilite o Office para chamar as APIs Web do suplemento com os mesmos direitos que você tem.

1. Verifique se o **Estado** está definido como **Habilitado**.

1. Selecione **Adicionar escopo**.

   > [!NOTE]
   > A parte do domínio **\<Scope\>** do nome exibida logo abaixo do campo de texto deve corresponder automaticamente ao URI da ID do Aplicativo que você definiu anteriormente, `/access_as_user` com acrescentado ao final; por exemplo, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.

1. Na seção **Aplicativos cliente autorizados** , insira a ID a seguir para pré-autorizar todos os pontos de extremidade de aplicativo do Microsoft Office.

   - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Todos os pontos de extremidade de aplicativo do Microsoft Office)

    > [!NOTE]
    > A `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` ID autoriza previamente o Office em todas as plataformas a seguir. Como alternativa, você pode inserir um subconjunto adequado das IDs a seguir se, por algum motivo, quiser negar a autorização ao Office em algumas plataformas. Basta deixar de fora as IDs das plataformas das quais você deseja reprisar a autorização. Os usuários do suplemento nessas plataformas não poderão chamar suas APIs Web, mas outras funcionalidades no suplemento ainda funcionarão.
    >
    > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    > - `93d53678-613d-4013-afc1-62e9e444a0a5`(Office na Web)
    > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3`(Outlook na Web)

1. Selecione o **botão Adicionar um** aplicativo cliente e, em seguida, no painel que é aberto, `[app-id-guid]` defina a ID do aplicativo (cliente) e marque a caixa para `api://localhost:44355/[app-id-guid]/access_as_user`.

1. Selecione **Adicionar aplicativo**.

1. Selecione **Permissões para API** em **Gerenciar** e selecione **Adicionar uma permissão**. No painel que se abre, escolha **Microsoft Graph** e, em seguida, escolha **Permissões delegadas**.

1. Use a caixa de pesquisa **Selecionar permissões** para procurar as permissões que o seu suplemento precisa. Pesquise e selecione a **permissão de** perfil. A `profile` permissão é necessária para que o aplicativo do Office obtenha um token para seu aplicativo Web de suplemento.

   - perfil

   > [!NOTE]
   > A permissão `User.Read` pode já estar listada por padrão. É uma boa prática não pedir permissões desnecessárias, por isso recomendamos desmarcar a caixa para essa permissão se o suplemento não precisar dela.

1. Selecione a **botão Adicionar** seleção na parte inferior do painel.

1. Na mesma página, escolha o botão **Conceder consentimento \<tenant-name\>** do administrador e, em seguida, selecione **Sim** para a confirmação exibida.

## <a name="create-the-office-add-in"></a>Criar o Suplemento do Office

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Inicie o Visual Studio 2019 e escolha **Criar um novo projeto**.
1. Pesquise e selecione o **modelo de projeto do** Suplemento Web do Excel. Depois clique em **Próximo**. Observação: o SSO funciona com qualquer aplicativo do Office, mas para este artigo funcionará com o Excel.
1. Insira um nome de projeto, como **sso-display-user-info** e escolha **Criar**. Você pode deixar os outros campos com valores padrão.
1. Na caixa **de diálogo Escolher o tipo de** suplemento, selecione **Adicionar nova funcionalidade ao Excel** e escolha **Concluir**.

O projeto é criado e conterá dois projetos na solução.

- **sso-display-user-info**: contém o manifesto e os detalhes para sideload do suplemento para o Excel.
- **sso-display-user-infoWeb**: o projeto ASP.NET que hospeda as páginas da Web para o suplemento.

# <a name="yo-office"></a>[yo escritório](#tab/yooffice)

Verifique se você [configurou seu ambiente de desenvolvimento](../overview/set-up-your-dev-environment.md).

1. Para criar o projeto, digite o seguinte comando.

   ```command line
   yo office --projectType taskpane --name 'sso-display-user-info' --host excel --js true
   ```

O projeto é criado em uma nova pasta chamada **sso-display-user-info**.

---

## <a name="configure-the-manifest"></a>Configurar o manifesto

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. No **Gerenciador de Soluções** **sso-display-user-info > sso-display-user-infoManifest > sso-display-user-info.xml**

# <a name="yo-office"></a>[yo escritório](#tab/yooffice)

1. No Visual Studio Code, abra o **manifest.xml** arquivo.

---

1. Próximo à parte inferior do manifesto há um elemento de `</Resources>` fechamento. Insira o XML a seguir logo abaixo do `</Resources>` elemento, mas antes do elemento de `</VersionOverrides>` fechamento. Para aplicativos do Office diferentes do Outlook, adicione a marcação ao final da `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` seção. Para o Outlook, adicione a marcação no final da seção `<VersionOverrides ... xsi:type="VersionOverridesV1_1">`.

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

1. Substitua `[port]` pelo número da porta correto para seu projeto. Se você criou o suplemento usando o yo office, o número da porta normalmente é 3000 e encontrado no arquivo package.json. Se você criou o suplemento com o Visual Studio 2019, a porta será encontrada na propriedade **de URL SSL** do projeto Web.
1. Substitua ambos `[application-id]` os espaços reservados pela ID do aplicativo real do registro do aplicativo.
1. Salve o arquivo.

O XML inserido contém os seguintes elementos e informações.

- **\<WebApplicationInfo\>** - O pai dos elementos a seguir.
- **\<Id\>** - A ID do cliente do suplemento Esta é uma ID de aplicativo que você obtém como parte do registro do suplemento. Confira [Registrar um Suplemento do Office que usa SSO com o ponto de extremidade do Azure AD v2.0](register-sso-add-in-aad-v2.md).
- **\<Resource\>** - A URL do suplemento. Esse é o mesmo URI (incluindo o protocolo `api:`) que você usou ao registrar o suplemento no AAD. A parte de domínio desse URI deve corresponder ao domínio, incluindo quaisquer subdomínios, usados nas URLs **\<Resources\>** na seção do manifesto do suplemento e o URI deve terminar com a ID **\<Id\>** do cliente no .
- **\<Scopes\>** - O pai de um ou mais **\<Scope\>** elementos.
- **\<Scope\>** – Especifica uma permissão que o suplemento precisa para o AAD. As permissões `profile` e `openID` são sempre necessárias e podem ser as únicas permissões necessárias, se o suplemento não acessar o Microsoft Graph. Se isso ocorrer, você também precisará de **\<Scope\>** elementos para as permissões necessárias do Microsoft Graph; por exemplo, `User.Read`, `Mail.Read`. Bibliotecas que você usa no seu código para acessar o Microsoft Graph pode precisar de permissões adicionais. Por exemplo, a biblioteca de autenticação da Microsoft (MSAL) para .NET requer a permissão `offline_access`. Para saber mais, confira [autorizar o Microsoft Graph de um suplemento do Office](authorize-to-microsoft-graph.md).

## <a name="add-the-jwt-decode-package"></a>Adicionar o pacote jwt-decode

Você pode chamar a `getAccessToken` API para obter o token de ID do Office. Primeiro, vamos adicionar o pacote jwt-decode para facilitar a decodificar e exibir o token de ID.

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. Abra a solução do Visual Studio.
1. No menu, escolha **Ferramentas > Gerenciador de Pacotes NuGet > Console do Gerenciador de Pacotes**.
1. Insira o comando a seguir no **Console do Gerenciador de Pacotes**.

   `Install-Package jwt-decode -Projectname sso-display-user-infoWeb`

# <a name="yo-office"></a>[yo escritório](#tab/yooffice)

1. Em uma janela de terminal/console, vá para a pasta raiz do seu projeto de suplemento.
1. Insira o comando a seguir

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

1. Substitua a `<body>` seção pelo HTML a seguir.

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

# <a name="yo-office"></a>[yo escritório](#tab/yooffice)

1. Abra o **arquivo src/taskpane/taskpane.html** .
1. Substitua a `<body>` seção pelo HTML a seguir.

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

# <a name="yo-office"></a>[yo escritório](#tab/yooffice)

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

1. Escolha **Depurar > Iniciar Depuração** ou pressione **F5**.

# <a name="yo-office"></a>[yo escritório](#tab/yooffice)

Execute `npm start` na linha de comando.

---

1. Quando o Excel for iniciado, entre no Office com a mesma conta de locatário usada para criar o registro de aplicativo.
1. Na faixa **de opções** Página Inicial, **escolha Mostrar Painel de Tarefas** para abrir o suplemento.
1. No painel de tarefas do suplemento, escolha **Obter token de ID**.

O suplemento exibirá o nome, o email e a ID da conta com a qual você entrou.

> [!NOTE]
> Se você encontrar erros, examine as etapas de registro neste artigo para o registro do aplicativo. Perder um detalhe ao configurar o registro do aplicativo é uma causa comum de problemas ao trabalhar com o SSO. Se você ainda não conseguir fazer com que o suplemento seja executado com êxito, consulte Solucionar problemas de mensagens de erro para [SSO (](troubleshoot-sso-in-office-add-ins.md)logon único).

## <a name="see-also"></a>Confira também

[Usando declarações para identificar um usuário de forma confiável (Entidade e ID de Objeto)](/azure/active-directory/develop/id-tokens#using-claims-to-reliably-identify-a-user-subject-and-object-id)

