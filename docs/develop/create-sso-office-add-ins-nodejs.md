---
title: Crie um Suplemento do Office com Node.js que use logon único
description: ''
ms.date: 08/21/2019
localization_priority: Priority
ms.openlocfilehash: 65efb7b4423a2764bcc07e3105dfb87292895297
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695795"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="542e4-102">Crie um Suplemento do Office com Node.js que use logon único (prévia)</span><span class="sxs-lookup"><span data-stu-id="542e4-102">Create a Node.js Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="542e4-p101">Os usuários podem entrar no Office, e o Suplemento Web do Office pode aproveitar esse processo de entrada para autorizá-los a acessar seu suplemento e o Microsoft Graph sem exigir que os eles entrem uma segunda vez. Para obter uma visão geral, confira o artigo [Habilitar o SSO em um Suplemento do Office](sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="542e4-p101">Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="542e4-105">Este artigo apresenta o processo passo a passo de habilitação do logon único (SSO) em um suplemento que foi criado com Node.js e Express.</span><span class="sxs-lookup"><span data-stu-id="542e4-105">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with Node.js and Express.</span></span>

> [!NOTE]
> <span data-ttu-id="542e4-106">Para ler um artigo semelhante sobre um suplemento baseado em ASP.NET, confira [Criar um Suplemento do Office com ASP.NET que usa o logon único](create-sso-office-add-ins-aspnet.md).</span><span class="sxs-lookup"><span data-stu-id="542e4-106">For a similar article about an ASP.NET-based add-in, see [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="542e4-107">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="542e4-107">Prerequisites</span></span>

* <span data-ttu-id="542e4-108">[Node e npm](https://nodejs.org/en/), versão 6.9.4 ou posterior</span><span class="sxs-lookup"><span data-stu-id="542e4-108">[Node and npm](https://nodejs.org/en/), version 6.9.4 or later</span></span>

* <span data-ttu-id="542e4-109">[Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)</span><span class="sxs-lookup"><span data-stu-id="542e4-109">[Git Bash](https://git-scm.com/downloads) (or another git client)</span></span>

* <span data-ttu-id="542e4-110">TypeScript, versão 2.2.2 ou posterior</span><span class="sxs-lookup"><span data-stu-id="542e4-110">TypeScript version 2.2.2 or later</span></span>

* <span data-ttu-id="542e4-111">Office 365 (a versão de assinatura do Office).</span><span class="sxs-lookup"><span data-stu-id="542e4-111">Office 365 (the subscription version of Office).</span></span> <span data-ttu-id="542e4-112">Build e versão mensal mais recentes do canal de Participante do programa Office Insider.</span><span class="sxs-lookup"><span data-stu-id="542e4-112">Latest monthly version and build from the Insiders channel.</span></span> <span data-ttu-id="542e4-113">É necessário ingressar no programa Office Insider para obter essa versão.</span><span class="sxs-lookup"><span data-stu-id="542e4-113">You need to be an Office Insider to get this version.</span></span> <span data-ttu-id="542e4-114">Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="542e4-114">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="542e4-115">Observe que, quando um build é promovido ao Canal Semestral de produção, o suporte para recursos de visualização, como o SSO, é desativado para esse build.</span><span class="sxs-lookup"><span data-stu-id="542e4-115">Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="542e4-116">Configure o projeto inicial</span><span class="sxs-lookup"><span data-stu-id="542e4-116">Set up the starter project</span></span>

1. <span data-ttu-id="542e4-117">Clone ou baixe o repositório em [SSO com Suplemento NodeJS do Office](https://github.com/officedev/office-add-in-nodejs-sso).</span><span class="sxs-lookup"><span data-stu-id="542e4-117">Clone or download the repo at [Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso).</span></span>

    > [!NOTE]
    > <span data-ttu-id="542e4-118">Há três versões do exemplo:</span><span class="sxs-lookup"><span data-stu-id="542e4-118">There are three versions of the sample:</span></span>  
    > * <span data-ttu-id="542e4-p103">A pasta **Before** (antes) traz um projeto inicial. A interface do usuário e outros aspectos do suplemento que não estão diretamente ligados ao SSO ou à autorização já estão prontos. As próximas seções deste artigo apresentam uma orientação passo a passo para concluir o projeto.</span><span class="sxs-lookup"><span data-stu-id="542e4-p103">The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.</span></span>
    > * <span data-ttu-id="542e4-p104">A versão **Completed** (concluído) do exemplo apresenta como seria o suplemento quando concluídos os procedimentos apresentados neste artigo, com exceção de que o projeto concluído traz comentários de códigos que seriam redundantes neste artigo. Para usar a versão concluída, apenas siga as instruções apresentadas neste artigo, substituindo "Before" por "Completed" e pulando as seções **Codificar o lado do cliente** e **Codificar o lado do servidor**.</span><span class="sxs-lookup"><span data-stu-id="542e4-p104">The **Completed** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just follow the instructions in this article, but replace "Before" with "Completed" and skip the sections **Code the client side** and **Code the server** side.</span></span>
    > * <span data-ttu-id="542e4-124">A versão **Multilocatário completa** é um exemplo completo que ofereça suporte para multilocação.</span><span class="sxs-lookup"><span data-stu-id="542e4-124">The **Completed Multitenant** version is a completed sample that supports multitenancy.</span></span> <span data-ttu-id="542e4-125">Explore este exemplo, se você pretende oferecer suporte para contas da Microsoft de domínios diferentes com SSO.</span><span class="sxs-lookup"><span data-stu-id="542e4-125">Explore this sample if you intend to support Microsoft accounts from different domains with SSO.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="542e4-126">Independentemente de qual versão você usa, será necessário confiar em um certificado para um host local.</span><span class="sxs-lookup"><span data-stu-id="542e4-126">Regardless of which version you use, you will need to trust a certificate for the localhost. See the "IMPORTANT" note in the Readme of the repo.</span></span> <span data-ttu-id="542e4-127">Siga [essas instruções para instalar certificados autoassinados](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md), exceto que as pastas `certs` de cada uma das versões neste repositório estão na pasta `/src`, não na pasta raiz.</span><span class="sxs-lookup"><span data-stu-id="542e4-127">Follow [these instructions for installing self-signed certificates](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md), except that the `certs` folders for each of the versions in this repo is in the `/src` folder, not the root folder.</span></span>

1. <span data-ttu-id="542e4-128">Abra um console Git bash na pasta **Before**.</span><span class="sxs-lookup"><span data-stu-id="542e4-128">Open a Git bash console in the **Before** folder.</span></span>

1. <span data-ttu-id="542e4-129">Insira `npm install` no console para instalar todas as dependências discriminadas no arquivo package.json.</span><span class="sxs-lookup"><span data-stu-id="542e4-129">Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.</span></span>

1. <span data-ttu-id="542e4-130">Insira `npm run build` no console para compilar o projeto.</span><span class="sxs-lookup"><span data-stu-id="542e4-130">Enter `npm run build` in the console to build the project.</span></span>

    > [!NOTE]
    > <span data-ttu-id="542e4-p107">Talvez você veja alguns erros de build informando que algumas variáveis estão declaradas mas não são usadas. Ignore esses erros. Eles são um efeito colateral, pois na versão "Before" do exemplo estão faltando alguns códigos que serão adicionados posteriormente.</span><span class="sxs-lookup"><span data-stu-id="542e4-p107">You may see some build errors saying that some variables are declared but not used. Ignore these errors. They are a side effect of the fact that the "Before" version of the sample is missing some code that will be added later.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="542e4-134">Registre o suplemento com o ponto de extremidade v2.0 do Azure AD</span><span class="sxs-lookup"><span data-stu-id="542e4-134">Register the add-in with Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="542e4-135">As instruções a seguir são escritas de forma geral, elas podem ser usadas em vários locais.</span><span class="sxs-lookup"><span data-stu-id="542e4-135">The following instruction are written generically so they can be used in multiple places.</span></span> <span data-ttu-id="542e4-136">Para este artigo faça o seguinte:</span><span class="sxs-lookup"><span data-stu-id="542e4-136">For this article do the following:</span></span>

- <span data-ttu-id="542e4-137">Substitua o espaço reservado **$ADD-IN-NAME$** por `Office-Add-in-NodeJS-SSO`.</span><span class="sxs-lookup"><span data-stu-id="542e4-137">Replace the placeholder **$ADD-IN-NAME$** with `Office-Add-in-NodeJS-SSO`.</span></span>
- <span data-ttu-id="542e4-138">Substitua o espaço reservado **$FQDN-WITHOUT-PROTOCOL$** por `localhost:3000`.</span><span class="sxs-lookup"><span data-stu-id="542e4-138">Replace the placeholder **$FQDN-WITHOUT-PROTOCOL$** with `localhost:3000`.</span></span>
- <span data-ttu-id="542e4-139">Quando você especificar permissões na caixa de diálogo **Selecionar permissões**, marque as caixas das seguintes permissões.</span><span class="sxs-lookup"><span data-stu-id="542e4-139">When you specify permissions in the **Select Permissions** dialog, check the boxes for the following permissions.</span></span> <span data-ttu-id="542e4-140">Somente a primeira permissão é realmente necessária pelo suplemento em si, mas a permissão `profile` é necessária para que o host do Office obtenha um token no aplicativo Web do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="542e4-140">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office host to get a token to your add-in web application.</span></span>
  * <span data-ttu-id="542e4-141">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="542e4-141">Files.Read.All</span></span>
  * <span data-ttu-id="542e4-142">profile</span><span class="sxs-lookup"><span data-stu-id="542e4-142">profile</span></span>

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]


## <a name="grant-administrator-consent-to-the-add-in"></a><span data-ttu-id="542e4-143">Conceder consentimento do administrador ao suplemento</span><span class="sxs-lookup"><span data-stu-id="542e4-143">Grant administrator consent to the add-in</span></span>

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a><span data-ttu-id="542e4-144">Configurar o suplemento</span><span class="sxs-lookup"><span data-stu-id="542e4-144">Configure the add-in</span></span>

1. <span data-ttu-id="542e4-p110">Em seu editor de códigos, abra o arquivo src\server.ts. Perto da parte superior, há uma chamada para um construtor de uma classe `AuthModule`. Há alguns parâmetros de cadeia de caracteres no construtor aos quais você precisa atribuir valores.</span><span class="sxs-lookup"><span data-stu-id="542e4-p110">In your code editor, open the src\server.ts file. Near the top there is a call to a constructor of an `AuthModule` class. There are some string parameters in the constructor to which you need to assign values.</span></span>

1. <span data-ttu-id="542e4-p111">Na propriedade `client_id`, substitua o espaço reservado `{client GUID}` pela ID do aplicativo que você salvou ao registrar o suplemento. Quando terminar, deverá haver apenas um GUID entre aspas simples. Não deverá haver nenhum caractere "{}"</span><span class="sxs-lookup"><span data-stu-id="542e4-p111">For the `client_id` property, replace the placeholder `{client GUID}` with the application ID that you saved when you registered the add-in. When you are done, there should just be a GUID in single quotation marks. There should not be any "{}" characters.</span></span>

1. <span data-ttu-id="542e4-151">Na propriedade `client_secret`, substitua o espaço reservado `{client secret}` pelo segredo do aplicativo que você salvou ao registrar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="542e4-151">For the `client_secret` property, replace the placeholder `{client secret}` with the application secret that you saved when you registered the add-in.</span></span>

1. <span data-ttu-id="542e4-p112">Na propriedade `audience`, substitua o espaço reservado `{audience GUID}` pela ID do aplicativo que você salvou ao registrar o suplemento. (Exatamente o mesmo valor que você atribuiu à propriedade `client_id`.)</span><span class="sxs-lookup"><span data-stu-id="542e4-p112">For the `audience` property, replace the placeholder `{audience GUID}` with the application ID that you saved when you registered the add-in. (The very same value that you assigned to the `client_id` property.)</span></span>
  
1. <span data-ttu-id="542e4-154">Na cadeia de caracteres atribuída à propriedade `issuer`, você verá o espaço reservado *{O365 tenant GUID}*.</span><span class="sxs-lookup"><span data-stu-id="542e4-154">In the string assigned to the `issuer` property, you will see the placeholder *{O365 tenant GUID}*.</span></span> <span data-ttu-id="542e4-155">Substitua pela ID de locatário do Office 365.</span><span class="sxs-lookup"><span data-stu-id="542e4-155">Replace this with the Office 365 tenancy ID.</span></span> <span data-ttu-id="542e4-156">Se você não copiou a ID de locatário quando você registrou o suplemento com AAD, use um dos métodos em [Encontrar sua ID de locatário do Office 365](/onedrive/find-your-office-365-tenant-id) para obtê-la.</span><span class="sxs-lookup"><span data-stu-id="542e4-156">If you didn't copy the tenancy ID when you registered the add-in with AAD, use one of the methods in [Find your Office 365 tenant ID](/onedrive/find-your-office-365-tenant-id) to obtain it.</span></span> <span data-ttu-id="542e4-157">Quando terminar, o valor da propriedade `issuer` deve ser algo parecido com isto:</span><span class="sxs-lookup"><span data-stu-id="542e4-157">When you are done, the `issuer` property value should look something like this:</span></span>

    `https://login.microsoftonline.com/12345678-1234-1234-1234-123456789012/v2.0`

1. <span data-ttu-id="542e4-p114">Não altere os demais parâmetros no construtor `AuthModule`. Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="542e4-p114">Leave the other parameters in the `AuthModule` constructor unchanged. Save and close the file.</span></span>

1. <span data-ttu-id="542e4-160">Na raiz do projeto, abra o arquivo do manifesto do suplemento "Office-Add-in-NodeJS-SSO.xml".</span><span class="sxs-lookup"><span data-stu-id="542e4-160">In the root of the project, open the add-in manifest file “Office-Add-in-NodeJS-SSO.xml”.</span></span>

1. <span data-ttu-id="542e4-161">Role até o final do arquivo.</span><span class="sxs-lookup"><span data-stu-id="542e4-161">Scroll to the bottom of the file.</span></span>

1. <span data-ttu-id="542e4-162">Logo acima da marca de fim `</VersionOverrides>`, você encontrará a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="542e4-162">Just above the end `</VersionOverrides>` tag, you will find the following markup:</span></span>

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

1. <span data-ttu-id="542e4-p115">Substitua o espaço reservado "{application_GUID aqui}" *nos dois lugares*, na marcação, pela ID do Aplicativo que você copiou ao registrar seu suplemento. Os "{}" não fazem parte da ID, portanto, não os inclua. Essa é a mesma ID usada para ClientID e Audience no web.config.</span><span class="sxs-lookup"><span data-stu-id="542e4-p115">Replace the placeholder “{application_GUID here}” *in both places* in the markup with the Application ID that you copied when you registered your add-in. (The "{}" are not part of the ID, so don't include them.) This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > * <span data-ttu-id="542e4-165">O valor de **Resource** é o **URI da ID do Aplicativo** que você definiu quando adicionou a plataforma API Web no registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="542e4-165">The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.</span></span>
    > * <span data-ttu-id="542e4-166">A seção **Scopes** só será usada para gerar uma caixa de diálogo de consentimento se o suplemento for vendido no AppSource.</span><span class="sxs-lookup"><span data-stu-id="542e4-166">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="542e4-167">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="542e4-167">Save and close the file.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="542e4-168">Codificar o lado do cliente</span><span class="sxs-lookup"><span data-stu-id="542e4-168">Code the client side</span></span>

1. <span data-ttu-id="542e4-p116">Abra o arquivo program.js da pasta **public**. Ele já apresenta alguns códigos:</span><span class="sxs-lookup"><span data-stu-id="542e4-p116">Open the program.js file in the **public** folder. It already has some code in it:</span></span>

    * <span data-ttu-id="542e4-171">Uma atribuição ao método `Office.initialize` que, por sua vez, atribui um manipulador ao evento clicar do botão `getGraphAccessTokenButton`.</span><span class="sxs-lookup"><span data-stu-id="542e4-171">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="542e4-172">Um método `showResult` que exibirá os dados retornados do Microsoft Graph (ou uma mensagem de erro) na parte inferior do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="542e4-172">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="542e4-173">Um método `logErrors` que registrará erros de console que não são destinados ao usuário final.</span><span class="sxs-lookup"><span data-stu-id="542e4-173">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>

1. <span data-ttu-id="542e4-p117">Abaixo da atribuição a `Office.initialize`, adicione o código a seguir. Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="542e4-p117">Below the assignment to `Office.initialize`, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="542e4-p118">O processamento de erros no suplemento às vezes tentará novamente obter um token de acesso automaticamente, usando um conjunto diferente de opções. A variável de contador `timesGetOneDriveFilesHasRun` e as variáveis sinalizador `triedWithoutForceConsent` e `timesMSGraphErrorReceived` são usadas para garantir que o usuário não seja trocado repetidas vezes em tentativas falhas de obter um token.</span><span class="sxs-lookup"><span data-stu-id="542e4-p118">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options. The counter variable `timesGetOneDriveFilesHasRun`, and the flag variables `triedWithoutForceConsent` and `timesMSGraphErrorReceived` are used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span>
    * <span data-ttu-id="542e4-p119">Você criará um método `getDataWithToken` na próxima etapa, mas observe que ele define uma opção chamada `forceConsent` como `false`. Trataremos mais disso na etapa seguinte.</span><span class="sxs-lookup"><span data-stu-id="542e4-p119">You create the `getDataWithToken` method in the next step, but note that it sets an option called `forceConsent` to `false`. More about that in the next step.</span></span>

    ```js
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;
    var timesMSGraphErrorReceived = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }
    ```

1. <span data-ttu-id="542e4-p120">Abaixo do método `getOneDriveFiles`, adicione o código a seguir. Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="542e4-p120">Below the `getOneDriveFiles` method, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="542e4-182">O [getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) é a nova API no Office.js que permite que um suplemento solicite ao aplicativo host do Office (Excel, PowerPoint, Word etc.) um token de acesso ao suplemento (para o usuário conectado ao Office).</span><span class="sxs-lookup"><span data-stu-id="542e4-182">The [getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office).</span></span> <span data-ttu-id="542e4-183">O aplicativo host do Office, por sua vez, solicita o token ao ponto de extremidade 2.0 do Azure AD.</span><span class="sxs-lookup"><span data-stu-id="542e4-183">The Office host application, in turn, asks the Azure AD 2.0 endpoint for the token.</span></span> <span data-ttu-id="542e4-184">Uma vez que você previamente autorizou o host do Office para o seu suplemento ao registrá-lo, o Azure AD enviará o token.</span><span class="sxs-lookup"><span data-stu-id="542e4-184">Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token.</span></span>
    * <span data-ttu-id="542e4-185">Se nenhum usuário estiver conectado ao Office, o host do Office solicitará que o usuário se conecte.</span><span class="sxs-lookup"><span data-stu-id="542e4-185">If no user is signed into Office, the Office host will prompt the user to sign in.</span></span>
    * <span data-ttu-id="542e4-p122">O parâmetro de opções configura o `forceConsent` como `false`. Dessa forma, não será solicitado que o usuário consinta o acesso ao host do Office ao seu suplemento sempre que ele o usar. Na primeira vez que o usuário tiver o suplemento, a chamada de `getAccessTokenAsync` falhará, mas lógica de processamento de erros que você adicionará em uma etapa posterior será automaticamente chamada com a opção `forceConsent` definida como `true` e o usuário será solicitado a consentir, mas somente essa primeira vez.</span><span class="sxs-lookup"><span data-stu-id="542e4-p122">The options parameter sets `forceConsent` to `false`, so the user will not be prompted to consent to giving the Office host access to your add-in every time she or he uses the add-in. The first time the user runs the add-in, the call of `getAccessTokenAsync` will fail, but error-handling logic that you add in a later step will automatically re-call with the `forceConsent` option set to `true` and the user will be prompted to consent, but only that first time.</span></span>
    * <span data-ttu-id="542e4-188">Você criará o método `handleClientSideErrors` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="542e4-188">You will create the `handleClientSideErrors` method in a later step.</span></span>

    ```js
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

1. <span data-ttu-id="542e4-p123">Substitua TODO1 pelas linhas a seguir. Você criará o método `getData` e a rota "/api/values" do lado do servidor nas etapas posteriores. Uma URL relativa é usada para o ponto de extremidade porque ela deve ser hospedada no mesmo domínio que seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="542e4-p123">Replace the TODO1 with the following lines. You create the `getData` method and the server-side “/api/values” route in later steps. A relative URL is used for the endpoint because it must be hosted on the same domain as your add-in.</span></span>

    ```js
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. <span data-ttu-id="542e4-p124">Abaixo do método `getOneDriveFiles`, adicione o seguinte. Observe isto sobre este código:</span><span class="sxs-lookup"><span data-stu-id="542e4-p124">Below the `getOneDriveFiles` method, add the following. About this code, note:</span></span>

    * <span data-ttu-id="542e4-p125">Este método utilitário chama um ponto de extremidade da API Web especificado e transmite a ela o mesmo token de acesso que aplicativo host do Office usou para obter acesso ao seu suplemento. No lado do servidor, esse token de acesso será usado no fluxo "on behalf of" (em nome de) para obter um token de acesso para o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="542e4-p125">This method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the “on behalf of” flow to obtain an access token to Microsoft Graph.</span></span>
    * <span data-ttu-id="542e4-196">Você criará o método `handleServerSideErrors` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="542e4-196">You will create the `handleServerSideErrors` method in a later step.</span></span>

    ```js
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

### <a name="create-the-error-handling-methods"></a><span data-ttu-id="542e4-197">Crie os métodos de processamento de erros</span><span class="sxs-lookup"><span data-stu-id="542e4-197">Create the error-handling methods</span></span>

1. <span data-ttu-id="542e4-p126">Abaixo do método `getData`, adicione o método a seguir. Esse método processará os erros no cliente do suplemento quando o host do Office não puder obter um token de acesso para o serviço Web do suplemento. Esses erros são relatados com um código de erro, portanto, o método usa uma instrução `switch` para distingui-los.</span><span class="sxs-lookup"><span data-stu-id="542e4-p126">Below the `getData` method, add the following method. This method will handle errors in the add-in's client when the Office host is unable to obtain an access token to the add-in's web service. These errors are reported with an error code, so the method uses a `switch` statement to distinguish them.</span></span>

    ```js
    function handleClientSideErrors(result) {

        switch (result.error.code) {

            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor.

            // TODO3: Handle the case where the user's sign-in or consent was aborted.

            // TODO4: Handle the case where the user is logged in with an account that is neither work or school,
            //        nor Microsoft Account.

            // TODO5: Handle the case where the Office host has not been authorized to the add-in's web service or
            //        the user has not granted the service permission to their `profile`.

            // TODO6: Handle an unspecified error from the Office host.

            // TODO7: Handle the case where the Office host cannot get an access token to the add-ins
            //        web service/application.

            // TODO8: Handle the case where the user triggered an operation that calls `getAccessTokenAsync`
            //        before a previous call of it completed.

            // TODO9: Handle the case where the add-in does not support forcing consent.

            // TODO10: Log all other client errors.
        }
    }
    ```

1. <span data-ttu-id="542e4-p127">Substitua `TODO2` pelo código a seguir. O erro 13001 ocorre quando o usuário não está conectado ou quando ele cancela, sem responder, uma solicitação para fornecer um segundo fator de autenticação. Em ambos os casos, o código executará novamente o método `getDataWithToken` e definirá uma opção para forçar uma solicitação de entrada.</span><span class="sxs-lookup"><span data-stu-id="542e4-p127">Replace `TODO2` with the following code. Error 13001 occurs when the user is not logged in, or the user cancelled, without responding, a prompt to provide a 2nd authentication factor. In either case, the code re-runs the `getDataWithToken` method and sets an option to force a sign-in prompt.</span></span>

    ```js
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. <span data-ttu-id="542e4-p128">Substitua `TODO3` pelo código a seguir. O erro 13002 ocorre quando a entrada ou o consentimento do usuário é anulado. Peça que o usuário tente novamente, mas não mais de uma vez.</span><span class="sxs-lookup"><span data-stu-id="542e4-p128">Replace `TODO3` with the following code. Error 13002 occurs when user's sign-in or consent was aborted. Ask the user to try again but no more than once again.</span></span>

    ```js
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }
        break;
    ```

1. <span data-ttu-id="542e4-207">Substitua `TODO4` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="542e4-207">Replace `TODO4` with the following code.</span></span> <span data-ttu-id="542e4-208">O erro 13003 ocorre quando o usuário está conectado com uma conta que não é corporativa, de estudante, nem da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="542e4-208">Error 13003 occurs when user is logged in with an account that is neither work or school, nor Microsoft Account.</span></span> <span data-ttu-id="542e4-209">Peça que o usuário saia e entre novamente com um tipo de conta suportado.</span><span class="sxs-lookup"><span data-stu-id="542e4-209">Ask the user to sign-out and then in again with a supported account type.</span></span>

    ```js
    case 13003:
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;
    ```

    > [!NOTE]
    > <span data-ttu-id="542e4-210">O erro 13004 não é processado neste método, pois eles ocorre apenas em desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="542e4-210">Error 13004 is not handled in this method because it should only occur in development.</span></span> <span data-ttu-id="542e4-211">Não é possível corrigi-lo pelo código de tempo de execução e não seria útil reportá-lo a um usuário final.</span><span class="sxs-lookup"><span data-stu-id="542e4-211">It cannot be fixed by runtime code and there would be no point in reporting it to an end user.</span></span>

1. <span data-ttu-id="542e4-212">Substitua `TODO5` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="542e4-212">Replace `TODO5` with the following code.</span></span> <span data-ttu-id="542e4-213">O erro 13005 ocorre quando o Office não tem autorização para o serviço Web do suplemento ou o usuário não concedeu permissão ao serviço para o respectivo `profile`.</span><span class="sxs-lookup"><span data-stu-id="542e4-213">Error 13005 occurs when Office has not been authorized to the add-in's web service or the user has not granted the service permission to their `profile`.</span></span>

    ```js
    case 13005:
        getDataWithToken({ forceConsent: true });
        break;
    ```

1. <span data-ttu-id="542e4-p132">Substitua `TODO6` pelo seguinte código. O Erro 13006 ocorre quando houve um erro não especificado no host do Office, que pode indicar a instabilidade do host. Peça ao usuário para reiniciar o Office.</span><span class="sxs-lookup"><span data-stu-id="542e4-p132">Replace `TODO6` with the following code. Error 13006 occurs when there has been an unspecified error in the Office host that may indicate that the host is in an unstable state. Ask the user to restart Office.</span></span>

    ```js
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;
    ```

1. <span data-ttu-id="542e4-p133">Substitua `TODO7` pelo código a seguir. O erro 13007 ocorre quando algo deu errado com a interação do host do Office com o AAD de forma que o host não pode obter um token de acesso para o serviço Web/aplicativo dos suplementos. É possível que esse seja um problema de rede temporário. Peça que o usuário tente novamente mais tarde.</span><span class="sxs-lookup"><span data-stu-id="542e4-p133">Replace `TODO7` with the following code. Error 13007 occurs when something has gone wrong with the Office host's interaction with AAD so the host cannot get an access token to the add-ins web service/application. This may be a temporary network issue. Ask the user to try again later.</span></span>

    ```js
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;
    ```

1. <span data-ttu-id="542e4-221">Substitua `TODO8` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="542e4-221">Replace `TODO8` with the following code.</span></span> <span data-ttu-id="542e4-222">O erro 13008 ocorre quando o usuário aciona uma operação que chama o `getAccessTokenAsync` antes que uma chamada anterior dele seja concluída.</span><span class="sxs-lookup"><span data-stu-id="542e4-222">Error 13008 occurs when the user triggered an operation that calls `getAccessTokenAsync` before a previous call of it completed.</span></span>

    ```js
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```

1. <span data-ttu-id="542e4-p135">Substitua `TODO9` pelo código a seguir. O erro 13009 ocorre quando o suplemento não permite forçar consentimento, mas `getAccessTokenAsync` foi chamado com a opção `forceConsent` definida como `true`. Normalmente, quando isso acontece, o código deve ser reexecutar `getAccessTokenAsync` automaticamente com a opção de consentimento definida como `false`. No entanto, em alguns casos, chamar o método com `forceConsent` definido como `true` é uma resposta automática para um erro em uma chamada para o método com a opção definida como `false`. Nesse caso, o código não deve tentar novamente, mas, em vez disso, ele deve solicitar que o usuário saia e entre novamente.</span><span class="sxs-lookup"><span data-stu-id="542e4-p135">Replace `TODO9` with the following code. Error 13009 occurs when the add-in does not support forcing consent, but `getAccessTokenAsync` was called with the `forceConsent` option set to `true`. In the usual case when this happens the code should automatically re-run `getAccessTokenAsync` with the consent option set to `false`. However, in some cases, calling the method with `forceConsent` set to `true` was itself an automatic response to an error in a call to the method with the option set to `false`. In that case, the code should not try again, but instead it should advise the user to sign out and sign in again.</span></span>

    ```js
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;


1. Replace `TODO10` with the following code.

    ```js
    default:
        logError(result);
        break;
    ```  

1. <span data-ttu-id="542e4-p136">Abaixo do método `handleClientSideErrors`, adicione o seguinte método. Esse método processará os erros no serviço Web do suplemento quando algo der errado na execução do fluxo on-behalf-of ou ao obter dados do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="542e4-p136">Below the `handleClientSideErrors` method, add the following method. This method will handle errors in the add-in's web service when something goes wrong in executing the on-behalf-of flow or in getting data from Microsoft Graph.</span></span>

    ```js
    function handleServerSideErrors(result) {

        // TODO11: Handle the case where AAD asks for an additional form of authentication.

        // TODO12: Handle the case where consent has not been granted, or has been revoked.

        // TODO13: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow

        // TODO14: Handle the case where the token that the add-in's client-side sends to its
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO15: Handle the case where the token sent to Microsoft Graph in the request for
        //         data is expired or invalid.

        // TODO16: Log all other server errors.
    }
    ```

1. <span data-ttu-id="542e4-p137">Substitua `TODO11` pelo código a seguir. Observação sobre este código:</span><span class="sxs-lookup"><span data-stu-id="542e4-p137">Replace `TODO11` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="542e4-p138">Existem configurações do Azure Active Directory nas quais o usuário precisa fornecer fator(es) de autenticação adicional(ais) para acessar alguns objetivos do Microsoft Graph (por exemplo, o OneDrive), mesmo que o usuário possa fazer login no Office apenas com uma senha. Nesse caso, o AAD enviará uma resposta com o erro 50076, que tem uma propriedade `Claims`.</span><span class="sxs-lookup"><span data-stu-id="542e4-p138">There are configurations of Azure Active Directory in which the user is required to provide additional authentication factor(s) to access some Microsoft Graph targets (e.g., OneDrive), even if the user can sign on to Office with just a password. In that case, AAD will send a response, with error 50076, that has a `Claims` property.</span></span>
    * <span data-ttu-id="542e4-p139">O host do Office deve obter um novo token com o valor **Claims** como a opção `authChallenge`. Isso instrui o AAD a solicitar ao usuário todas as formas de autenticação requeridas.</span><span class="sxs-lookup"><span data-stu-id="542e4-p139">The Office host should get a new token with the **Claims** value as the `authChallenge` option. This tells AAD to prompt the user for all required forms of authentication.</span></span>

    ```js
    if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 50076){
        getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
    }
    ```

1. <span data-ttu-id="542e4-p140">Substitua `TODO12` pelo seguinte código *logo abaixo da última chave de fechamento do código adicionado na etapa anterior*. Observação sobre esse código:</span><span class="sxs-lookup"><span data-stu-id="542e4-p140">Replace `TODO12` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="542e4-238">O erro 65001 significa que o consentimento para acessar o Microsoft Graph não foi concedido (ou foi revogado) para uma ou mais permissões.</span><span class="sxs-lookup"><span data-stu-id="542e4-238">Error 65001 means that consent to access Microsoft Graph was not granted (or was revoked) for one or more permissions.</span></span>
    * <span data-ttu-id="542e4-239">O suplemento deverá obter um novo token com a opção `forceConsent` definida como `true`.</span><span class="sxs-lookup"><span data-stu-id="542e4-239">The add-in should get a new token with the `forceConsent` option set to `true`.</span></span>

    ```js
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 65001){
        getDataWithToken({ forceConsent: true });
    }
    ```

1. <span data-ttu-id="542e4-p141">Substitua `TODO13` pelo seguinte código *logo abaixo da última chave de fechamento do código adicionado na etapa anterior*. Observação sobre esse código:</span><span class="sxs-lookup"><span data-stu-id="542e4-p141">Replace `TODO13` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="542e4-p142">O erro 70011 significa que um escopo inválido (permissão) foi solicitado. O suplemento deverá relatar o erro.</span><span class="sxs-lookup"><span data-stu-id="542e4-p142">Error 70011 means that an invalid scope (permission) has been requested. The add-in should report the error.</span></span>
    * <span data-ttu-id="542e4-244">O código registra qualquer outro erro com um número de erro do AAD.</span><span class="sxs-lookup"><span data-stu-id="542e4-244">The code logs any other error with an AAD error number.</span></span>

    ```js
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 70011){
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. <span data-ttu-id="542e4-p143">Substitua `TODO14` pelo seguinte código *logo abaixo da última chave de fechamento do código adicionado na etapa anterior*. Observação sobre esse código:</span><span class="sxs-lookup"><span data-stu-id="542e4-p143">Replace `TODO14` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="542e4-247">Código de servidor criado em uma etapa posterior enviará a mensagem terminada em `... expected access_as_user` se a o escopo `access_as_user` (permissão) não for o token de acesso que o cliente do suplemento enviar para o ADD para ser usado no fluxo on-behalf-of.</span><span class="sxs-lookup"><span data-stu-id="542e4-247">Server-side code that you create in a later step will send the message that ends with `... expected access_as_user` if the `access_as_user` scope (permission) is not in the access token that the add-in's client sends to AAD to be used in the on-behalf-of flow.</span></span>
    * <span data-ttu-id="542e4-248">O suplemento deverá relatar o erro.</span><span class="sxs-lookup"><span data-stu-id="542e4-248">The add-in should report the error.</span></span>

    ```js
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('expected access_as_user') !== -1){
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. <span data-ttu-id="542e4-p144">Substitua `TODO15` pelo seguinte código *logo abaixo da última chave de fechamento do código adicionado na etapa anterior*. Observação sobre esse código:</span><span class="sxs-lookup"><span data-stu-id="542e4-p144">Replace `TODO15` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="542e4-251">É improvável que um token expirado ou inválido seja enviado para o Microsoft Graph, mas, se isso acontecer, o código de servidor que você criará em uma etapa posterior terminará com a cadeia de caracteres `Microsoft Graph error`.</span><span class="sxs-lookup"><span data-stu-id="542e4-251">It is unlikely that an expired or invalid token will be sent to Microsoft Graph; but if it does happen, the server-side code that you will create in a later step will end with the string `Microsoft Graph error`.</span></span>
    * <span data-ttu-id="542e4-p145">Nesse caso, o suplemento deverá iniciar o processo de autenticação completo ao redefinir o contador `timesGetOneDriveFilesHasRun` e as variáveis de sinalizador `timesGetOneDriveFilesHasRun` e, em seguida, chamando novamente o método de identificador de botão. No entanto, isso deve ser feito apenas uma vez. Se isso acontecer novamente, o erro deve ser apenas registrado.</span><span class="sxs-lookup"><span data-stu-id="542e4-p145">In this case, the add-in should start the entire authentication process over by resetting the `timesGetOneDriveFilesHasRun` counter and `timesGetOneDriveFilesHasRun` flag variables, and then re-calling the button handler method. But it should do this only once. If it happens again, it should just log the error.</span></span>
    * <span data-ttu-id="542e4-255">O código registra o erro se isso acontecer duas vezes em sequência.</span><span class="sxs-lookup"><span data-stu-id="542e4-255">The code logs the error if it happens twice in succession.</span></span>

    ```js
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('Microsoft Graph error') !== -1) {
        if (!timesMSGraphErrorReceived) {
            timesMSGraphErrorReceived = true;
            timesGetOneDriveFilesHasRun = 0;
            triedWithoutForceConsent = false;
            getOneDriveFiles();
        } else {
            logError(result);
        }
    }
    ```

1. <span data-ttu-id="542e4-256">Substitua `TODO16` pelo seguinte código *logo abaixo da última chave de fechamento do código adicionado na etapa anterior*.</span><span class="sxs-lookup"><span data-stu-id="542e4-256">Replace `TODO16` with the following code *just below the last closing brace of the code you added in the previous step*.</span></span>

    ```js
    else {
        logError(result);
    }
    ```

## <a name="code-the-server-side"></a><span data-ttu-id="542e4-257">Codifique o lado do servidor</span><span class="sxs-lookup"><span data-stu-id="542e4-257">Code the server side</span></span>

<span data-ttu-id="542e4-258">Há dois arquivos do lado do servidor que precisam ser modificados.</span><span class="sxs-lookup"><span data-stu-id="542e4-258">There are two server-side files that need to be modified.</span></span>

- <span data-ttu-id="542e4-p146">O src\auth.js fornece funções auxiliares de autorização. Ele já tem membros genéricos que são usados em uma variedade de fluxos de autorização. É preciso adicionar funções a esse arquivo para implementar o fluxo "on behalf of".</span><span class="sxs-lookup"><span data-stu-id="542e4-p146">The src\auth.js provides authorization helper functions. It already has generic members that are used in a variety of authorization flows. We need to add functions to it that implement the "on behalf of" flow.</span></span>
- <span data-ttu-id="542e4-p147">O arquivo de src\server.js tem os membros básicos necessários para executar um servidor e o middleware do express. É necessário adicionar funções a ele que ajudam a API Web e a página inicial a obterem os dados do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="542e4-p147">The src\server.js file has the basic members need to run a server and express middleware. We need to add functions to it that serve the home page and a Web API for obtaining Microsoft Graph data.</span></span>

### <a name="create-a-method-to-exchange-tokens"></a><span data-ttu-id="542e4-264">Criar um método para troca de tokens</span><span class="sxs-lookup"><span data-stu-id="542e4-264">Create a method to exchange tokens</span></span>

1. <span data-ttu-id="542e4-p148">Abra o arquivo \src\auth.ts. Adicione o método abaixo à classe `AuthModule`. Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="542e4-p148">Open the \src\auth.ts file. Add the method below to the `AuthModule` class. Note the following about this code:</span></span>

    * <span data-ttu-id="542e4-p149">O parâmetro `jwt` é o token de acesso ao aplicativo. No fluxo de "on behalf of" (em nome de), ele é trocado com AAD por um token de acesso ao recurso.</span><span class="sxs-lookup"><span data-stu-id="542e4-p149">The `jwt` parameter is the access token to the application. In the "on behalf of" flow, it is exchanged with AAD for an access token to the resource.</span></span>
    * <span data-ttu-id="542e4-270">O parâmetro scopes (escopos) tem um valor padrão, mas neste exemplo será substituído pelo código de chamada.</span><span class="sxs-lookup"><span data-stu-id="542e4-270">The scopes parameter has a default value, but in this sample it will be overridden by the calling code.</span></span>
    * <span data-ttu-id="542e4-271">O parâmetro de recurso é opcional.</span><span class="sxs-lookup"><span data-stu-id="542e4-271">The resource parameter is optional.</span></span> <span data-ttu-id="542e4-272">Ele não deverá ser usado quando o [STS (Secure Token Service)](/previous-versions/windows-identity-foundation/ee748490(v=msdn.10)) for o ponto de extremidade do AAD V 2.0.</span><span class="sxs-lookup"><span data-stu-id="542e4-272">It should not be used when the [Secure Token Service (STS)](/previous-versions/windows-identity-foundation/ee748490(v=msdn.10)) is the AAD V 2.0 endpoint.</span></span> <span data-ttu-id="542e4-273">O ponto de extremidade V 2.0 infere o recurso dos escopos e retorna um erro se um recurso é enviado na Solicitação HTTP.</span><span class="sxs-lookup"><span data-stu-id="542e4-273">The V 2.0 endpoint infers the resource from the scopes and it returns an error if a resource is sent in the HTTP Request.</span></span>
    * <span data-ttu-id="542e4-p151">Gerar uma exceção no bloco `catch` *não* causará o envio imediato do "500 Erro Interno do Servidor" para o cliente. Chamar o código no arquivo server.js acionará essa exceção e a transformará em uma mensagem de erro que será enviada para o cliente.</span><span class="sxs-lookup"><span data-stu-id="542e4-p151">Throwing an exception in the `catch` block will *not* cause an immediate "500 Internal Server Error" to be sent to the client. Calling code in the server.js file will catch this exception and turn it into an error message that is sent to the client.</span></span>

        ```typescript
        private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
            try {
                // TODO3: Construct the parameters that will be sent in the body of the
                //        HTTP Request to the STS that starts the "on behalf of" flow.
                // TODO4: Send the request to the STS.
                // TODO5: Catch errors from the STS and relay them to the client.
                // TODO6: Process the response and persist the access token to resource.
            }
            catch (exception) {
                throw new UnauthorizedError('Unable to obtain an access token to the resource'
                                            + JSON.stringify(exception),
                                            exception);
            }
        }
        ```

1. <span data-ttu-id="542e4-p152">Substitua `TODO3` pelo código a seguir. Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="542e4-p152">Replace `TODO3` with the following code. About this code, note:</span></span>
    * <span data-ttu-id="542e4-p153">Um STS com suporte para o fluxo "on behalf of" espera determinados pares de valor/propriedade no corpo da solicitação HTTP. Esse código constrói um objeto que se tornará o corpo da solicitação.</span><span class="sxs-lookup"><span data-stu-id="542e4-p153">An STS that supports the "on behalf of" flow expects certain property/value pairs in the body of the HTTP request. This code constructs an object that will become the body of the request.</span></span>
    * <span data-ttu-id="542e4-280">Uma propriedade de recurso é adicionada ao corpo se, e somente se, um recurso é transmitido para o método.</span><span class="sxs-lookup"><span data-stu-id="542e4-280">A resource property is added to the body if, and only if, a resource was passed to the method.</span></span>

        ```typescript
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

1. <span data-ttu-id="542e4-281">Substitua `TODO4` pelo código a seguir que envia a solicitação HTTP para o ponto de extremidade do token do STS.</span><span class="sxs-lookup"><span data-stu-id="542e4-281">Replace `TODO4` with the following code which sends the HTTP request to the token endpoint of the STS.</span></span>

    ```typescript
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    });
    ```

1. <span data-ttu-id="542e4-p154">Substitua `TODO5` pelo código a seguir. Observe que gerar uma exceção *não* causará o envio imediato do "500 Erro Interno do Servidor" para o cliente. Chamar o código no arquivo server.js acionará essa exceção e a transformará em uma mensagem de erro que será enviada para o cliente.</span><span class="sxs-lookup"><span data-stu-id="542e4-p154">Replace `TODO5` with the following code. Note that throwing an exception will *not* cause an immediate "500 Internal Server Error" to be sent to the client. Calling code in the server.js file will catch this exception and turn it into an error message that is sent to the client.</span></span>

    ```typescript
     if (res.status !== 200) {
        const exception = await res.json();
        throw exception;
    }
    ```

1. <span data-ttu-id="542e4-p155">Substitua `TODO6` pelo código a seguir. Observe que o código persiste no token de acesso ao recurso, e é a hora de expiração, além de retorná-lo. O código de chamada pode evitar chamadas desnecessárias ao STS reutilizando um token de acesso não expirado ao recurso. Você verá como fazer isso na próxima seção.</span><span class="sxs-lookup"><span data-stu-id="542e4-p155">Replace `TODO6` with the following code. Note that the code persists the access token to the resource, and it's expiration time, in addition to returning it. Calling code can avoid unnecessary calls to the STS by reusing an unexpired access token to the resource. You'll see how to do that in the next section.</span></span>

    ```typescript  
    const json = await res.json();
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken;
    ```

1. <span data-ttu-id="542e4-289">Salve o arquivo, mas não o feche.</span><span class="sxs-lookup"><span data-stu-id="542e4-289">Save the file, but don't close it.</span></span>

### <a name="create-a-method-to-get-access-to-the-resource-using-the-on-behalf-of-flow"></a><span data-ttu-id="542e4-290">Criar um método para obter acesso ao recurso usando o fluxo "on behalf of"</span><span class="sxs-lookup"><span data-stu-id="542e4-290">Create a method to get access to the resource using the "on behalf of" flow</span></span>

1. <span data-ttu-id="542e4-p156">Ainda no arquivo src/auth.ts, adicione o método abaixo à classe `AuthModule`. Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="542e4-p156">Still in src/auth.ts, add the method below to the `AuthModule` class. Note the following about this code:</span></span>

    * <span data-ttu-id="542e4-293">Os comentários acima sobre os parâmetros para o método `exchangeForToken` aplicam-se aos parâmetros deste método também.</span><span class="sxs-lookup"><span data-stu-id="542e4-293">The comments above about the parameters to the the `exchangeForToken` method apply to the parameters of this method as well.</span></span>
    * <span data-ttu-id="542e4-p157">O método primeiro verifica o armazenamento persistente para um token de acesso ao recurso que não expirou e não vai expirar no próximo minuto. Ele chama o método `exchangeForToken` que você criou na última seção somente se necessário.</span><span class="sxs-lookup"><span data-stu-id="542e4-p157">The method first checks the persistent storage for an access token to the resource that has not expired and is not going to expire in the next minute. It calls the `exchangeForToken` method you created in the last section only if it needs to.</span></span>

    ```typescript
    async acquireTokenOnBehalfOf(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        const resourceTokenExpirationTime = ServerStorage.retrieve('ResourceTokenExpiresAt');
        if (moment().add(1, 'minute').diff(await resourceTokenExpirationTime) < 1 ) {
            return ServerStorage.retrieve('ResourceToken');
        } else if (resource) {
            return this.exchangeForToken(jwt, scopes, resource);
        } else {
            return this.exchangeForToken(jwt, scopes);
        }
    }
    ```

1. <span data-ttu-id="542e4-296">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="542e4-296">Save and close the file.</span></span>

### <a name="create-the-endpoints-that-will-serve-the-add-ins-home-page-and-data"></a><span data-ttu-id="542e4-297">Criar os pontos de extremidade que servirão aos dados e à página inicial do suplemento</span><span class="sxs-lookup"><span data-stu-id="542e4-297">Create the endpoints that will serve the add-in's home page and data</span></span>

1. <span data-ttu-id="542e4-298">Abra o arquivo src\server.ts.</span><span class="sxs-lookup"><span data-stu-id="542e4-298">Open the src\server.ts file.</span></span>

1. <span data-ttu-id="542e4-p158">Adicione o método a seguir na parte inferior do arquivo. Esse método servirá à página inicial do suplemento. O manifesto do suplemento especifica a URL da página inicial.</span><span class="sxs-lookup"><span data-stu-id="542e4-p158">Add the following method to the bottom of the file. This method will serve the add-in's home page. The add-in manifest specifies the home page URL.</span></span>

    ```typescript
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    }));
    ```

1. <span data-ttu-id="542e4-p159">Adicione o método a seguir na parte inferior do arquivo. Este método lidará com todas as solicitações para a API `values`.</span><span class="sxs-lookup"><span data-stu-id="542e4-p159">Add the following method to bottom of the file. This method will handle any requests for the `values` API.</span></span>

    ```typescript
    app.get('/api/values', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Relay any errors from Microsoft Graph to the client.
        // TODO11: Send to the client only the data that it actually needs.
    }));
    ```

1. <span data-ttu-id="542e4-304">Substitua `TODO7` pelo seguinte código que valida o token de acesso recebido do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="542e4-304">Replace `TODO7` with the following code which validates the access token received from the Office host application.</span></span> <span data-ttu-id="542e4-305">O método `verifyJWT` é definido no arquivo src\auth.ts.</span><span class="sxs-lookup"><span data-stu-id="542e4-305">The `verifyJWT` method is defined in the src\auth.ts file.</span></span> <span data-ttu-id="542e4-306">Ele sempre valida a audiência e o emissor.</span><span class="sxs-lookup"><span data-stu-id="542e4-306">It always validates the audience and the issuer.</span></span> <span data-ttu-id="542e4-307">Usamos o parâmetro opcional para especificar que também desejamos que ele verifique se o escopo no token de acesso é `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="542e4-307">We use the optional parameter to specify that we also want it to verify that the scope in the access token is `access_as_user`.</span></span> <span data-ttu-id="542e4-308">Esta é a única permissão ao suplemento que o usuário e o host do Office precisam para obter um token de acesso para o Microsoft Graph por meio do fluxo "on behalf of".</span><span class="sxs-lookup"><span data-stu-id="542e4-308">This is the only permission to the add-in that the user and the Office host need in order to get an access token to Microsoft Graph by means of the "on behalf" flow.</span></span>

    ```typescript
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' });
    ```

    > [!NOTE]
    > <span data-ttu-id="542e4-309">Você deve usar apenas o escopo `access_as_user` para autorizar a API que lida com o fluxo Em Nome De para os suplementos do Office. Outras APIs em seu serviço devem ter seus próprios requisitos de escopo.</span><span class="sxs-lookup"><span data-stu-id="542e4-309">You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office Add-ins. Other APIs in your service should have their own scope requirements.</span></span> <span data-ttu-id="542e4-310">Isso limita o que pode ser acessado com os tokens que o Office adquire.</span><span class="sxs-lookup"><span data-stu-id="542e4-310">This limits what can be accessed with the tokens that Office acquires.</span></span>

1. <span data-ttu-id="542e4-p162">Substitua `TODO8` pelo código a seguir. Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="542e4-p162">Replace `TODO8` with the following code. Note the following about this code:</span></span>

    * <span data-ttu-id="542e4-313">A chamada para `acquireTokenOnBehalfOf` não inclui um parâmetro de recurso porque construímos o objeto `AuthModule` (`auth`) com o ponto de extremidade V2.0 do AAD que não oferece suporte à propriedade de recurso.</span><span class="sxs-lookup"><span data-stu-id="542e4-313">The call to `acquireTokenOnBehalfOf` does not include a resource parameter because we constructed the `AuthModule` object (`auth`) with the AAD V2.0 endpoint which does not support a resource property.</span></span>
    * <span data-ttu-id="542e4-p163">O segundo parâmetro da chamada especifica as permissões que o suplemento precisará para obter uma lista dos arquivos e das pastas do usuário no OneDrive. (A permissão `profile` não é solicitada, porque só é necessária quando o host do Office obtém o token de acesso ao seu suplemento, e não quando você está negociando nesse token para um token de acesso para o Microsoft Graph.)</span><span class="sxs-lookup"><span data-stu-id="542e4-p163">The second parameter of the call specifies the permissions the add-in will need to get a list of the user's files and folders on OneDrive. (The `profile` permission is not requested because it is only needed when the Office host gets the access token to your add-in, not when you are trading in that token for an access token to Microsoft Graph.)</span></span>

    ```typescript
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    ```

1. <span data-ttu-id="542e4-p164">Substitua `TODO9` pela linha a seguir. Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="542e4-p164">Replace `TODO9` with the following line. Note the following about this code:</span></span>

    * <span data-ttu-id="542e4-318">A classe MSGraphHelper é definida no src\msgraph-helper.ts.</span><span class="sxs-lookup"><span data-stu-id="542e4-318">The MSGraphHelper class is defined in src\msgraph-helper.ts.</span></span>
    * <span data-ttu-id="542e4-319">Podemos minimizar os dados que devem ser retornados especificando que só queremos a propriedade de nome e somente os três primeiros itens.</span><span class="sxs-lookup"><span data-stu-id="542e4-319">We minimize the data that must be returned by specifying that we only want the name property and only the first 3 items.</span></span>

    ```typescript
    const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");
    ```

1. <span data-ttu-id="542e4-p165">Substitua `TODO10` pelo código a seguir. Observe que esse código processa erros "401 Não Autorizado" do Microsoft Graph que indicariam um token expirado ou inválido. É muito improvável que isso aconteça, pois a lógica persistente do token impede essa situação. (Confira a seção **Criar um método para obter acesso ao recurso usando o fluxo "on behalf of"** acima.) Se isso acontecer, o código transmitirá o erro para o cliente com "Erro do Microsoft Graph" no nome do erro. (Confira o método `handleClientSideErrors` que você criou no arquivo program.js em uma etapa anterior.) O código adicionado ao arquivo ODataHelper.js em uma etapa posterior ajuda a processar erros do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="542e4-p165">Replace `TODO10` with the following code. Note that this code handles '401 Unauthorized" errors from Microsoft Graph which would indicate an expired or invalid token. It is very unlikely that this would ever happen since the token persisting logic should prevent it. (See the section **Create a method to get access to the resource using the "on behalf of" flow** above.) If it does happen, this code will relay the error to the client with "Microsoft Graph error" in the error name. (See the `handleClientSideErrors` method that you created in the program.js file in an earlier step.) Code that you add to the ODataHelper.js file in a later step helps process errors from Microsoft Graph.</span></span>

    ```typescript
    if (graphData.code) {
        if (graphData.code === 401) {
            throw new UnauthorizedError('Microsoft Graph error', graphData);
        }
    }
    ```


1. <span data-ttu-id="542e4-p166">Substitua `TODO11` pelo código a seguir. Observe que o Microsoft Graph retorna alguns metadados OData e uma propriedade **eTag** para cada item, mesmo se `name` é a única propriedade solicitada. O código envia somente os nomes de item para o cliente.</span><span class="sxs-lookup"><span data-stu-id="542e4-p166">Replace `TODO11` with the following code. Note that Microsoft Graph returns some OData metadata and an **eTag** property for every item, even if `name` is the only property requested. The code sends only the item names to the client.</span></span>

    ```typescript
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

1. <span data-ttu-id="542e4-328">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="542e4-328">Save and close the file.</span></span>

### <a name="add-response-handling-to-the-odatahelper"></a><span data-ttu-id="542e4-329">Adicione processamento de respostas ao ODataHelper</span><span class="sxs-lookup"><span data-stu-id="542e4-329">Add response handling to the ODataHelper</span></span>

1. <span data-ttu-id="542e4-p167">Abra o arquivo src\odata-helper.ts. O arquivo está quase pronto. O que está ausente é o corpo do retorno de chamada para o identificador do evento “end” da solicitação. Substitua o `TODO` pelo código a seguir. Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="542e4-p167">Open the file src\odata-helper.ts. The file is almost complete. What's missing is the body of the callback to the handler for the request "end" event. Replace the `TODO` with the following code. About this code note:</span></span>

    * <span data-ttu-id="542e4-p168">A resposta do ponto de extremidade OData pode ser um erro, por exemplo, 401, se o ponto de extremidade exigir um token de acesso e ele for inválido ou estiver expirado. Uma mensagem de erro é ainda um *mensagem*, não um erro, nas chamadas de `https.get`, portanto, a linha `on('error', reject)` no final do `https.get` não é acionada. Portanto, o código distingue mensagens de sucesso (200) de mensagens de erro e envia um objeto JSON para o chamador com o OData solicitado ou informações de erro.</span><span class="sxs-lookup"><span data-stu-id="542e4-p168">The response from the OData endpoint might be an error, say a 401 if the endpoint requires an access token and it was invalid or expired. But an error message is still a *message*, not an error in the call of `https.get`, so the `on('error', reject)` line at the end of `https.get` isn't triggered. So, the code distinguishes success (200) messages from error messages and sends a JSON object to the caller with either the requested OData or error information.</span></span>

    ```typescript
    var error;
    if (response.statusCode === 200) {
        // TODO1: Return the data to the caller and resolve the Promise.
    } else {
       // TODO2: Return an error object to the caller and resolve the Promise.
    }
    ```

1. <span data-ttu-id="542e4-p169">Substitua `TODO1` pelo código a seguir. Observe que o código pressupõe que os dados retornados são JSON.</span><span class="sxs-lookup"><span data-stu-id="542e4-p169">Replace `TODO1` with the following code. Note that the code assumes the data is returned as JSON.</span></span>

    ```typescript
    let parsedBody = JSON.parse(body);
    resolve(parsedBody);
    ```

1. <span data-ttu-id="542e4-p170">Substitua `TODO2` pelo código a seguir. Observação sobre este código:</span><span class="sxs-lookup"><span data-stu-id="542e4-p170">Replace `TODO2` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="542e4-p171">Uma resposta de erro de uma fonte de OData sempre terá um statusCode e, normalmente, um statusMessage. Algumas fontes de OData também adicionam uma propriedade de erro ao corpo da mensagem com mais informações, como uma solicitação interna ou, mais especificamente, um código e uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="542e4-p171">An error response from an OData source will always have a statusCode and usually a statusMessage. Some OData sources also add an error property to the body with further information, such as an inner, or more specific, code and message.</span></span>
    * <span data-ttu-id="542e4-p172">O objeto Promise é resolvido, não rejeitado. O `https.get` é executado quando um serviço Web chama um ponto de extremidade OData de servidor para servidor. No entanto, essa chamada chega no contexto de uma chamada de um cliente para uma Web API do serviço Web. A solicitação "externa" do cliente para o serviço Web nunca é concluída se essa solicitação "interna" for rejeitada. Além disso, a solicitação com o objeto `Error` personalizado é necessária se o chamador de `http.get` precisar transmitir erros do ponto de extremidade OData para o cliente.</span><span class="sxs-lookup"><span data-stu-id="542e4-p172">The Promise object is resolved, not rejected. The `https.get` runs when a web service calls an OData endpoint server-to-server. But that call comes in the context of a call from a client to a web API in the web service. The "outer" request from the client to the web service never completes if this "inner" request is rejected. Also, resolving the request with the custom `Error` object is required if the caller of `http.get` needs to relay errors from the OData endpoint to the client.</span></span>

    ```typescript
    error = new Error();
    error.code = response.statusCode;
    error.message = response.statusMessage;

    // The error body sometimes includes an empty space
    // before the first character, remove it or it causes an error.
    body = body.trim();
    error.bodyCode = JSON.parse(body).error.code;
    error.bodyMessage = JSON.parse(body).error.message;
    resolve(error);
    ```

1. <span data-ttu-id="542e4-349">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="542e4-349">Save and close the file.</span></span>

## <a name="deploy-the-add-in"></a><span data-ttu-id="542e4-350">Implantar o suplemento</span><span class="sxs-lookup"><span data-stu-id="542e4-350">Deploy the add-in</span></span>

<span data-ttu-id="542e4-351">Agora é preciso informar ao Office onde encontrar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="542e4-351">Now you need to let Office know where to find the add-in.</span></span>

1. <span data-ttu-id="542e4-352">Crie um compartilhamento de rede ou [compartilhe uma pasta na rede](/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc770880(v=ws.11)).</span><span class="sxs-lookup"><span data-stu-id="542e4-352">Create a network share, or [share a folder to the network](/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc770880(v=ws.11)).</span></span>

1. <span data-ttu-id="542e4-353">Coloque uma cópia do arquivo de manifesto Office-Add-in-NodeJS-SSO.xml, da raiz do projeto, dentro da pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="542e4-353">Place a copy of the Office-Add-in-NodeJS-SSO.xml manifest file, from the root of the project, into the shared folder.</span></span>

1. <span data-ttu-id="542e4-354">Inicie o PowerPoint e abra um documento.</span><span class="sxs-lookup"><span data-stu-id="542e4-354">Launch PowerPoint and open a document.</span></span>

1. <span data-ttu-id="542e4-355">Escolha a guia **Arquivo** e, então, **Opções**.</span><span class="sxs-lookup"><span data-stu-id="542e4-355">Choose the **File** tab, and then choose **Options**.</span></span>

1. <span data-ttu-id="542e4-356">Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="542e4-356">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>

1. <span data-ttu-id="542e4-357">Escolha **Catálogos de Suplementos Confiáveis**.</span><span class="sxs-lookup"><span data-stu-id="542e4-357">Choose **Trusted Add-ins Catalogs**.</span></span>

1. <span data-ttu-id="542e4-358">No campo **URL do Catálogo**, insira o caminho de rede para o compartilhamento de pasta que contém o arquivo Office-Add-in-NodeJS-SSO.xml e escolha **Adicionar Catálogo**.</span><span class="sxs-lookup"><span data-stu-id="542e4-358">In the **Catalog Url** field, enter the network path to the folder share that contains Office-Add-in-NodeJS-SSO.xml, and then choose **Add Catalog**.</span></span>

1. <span data-ttu-id="542e4-359">Selecione a caixa de seleção **Mostrar no Menu** e, em seguida, escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="542e4-359">Select the **Show in Menu** check box, and then choose **OK**.</span></span>

1. <span data-ttu-id="542e4-p173">Uma mensagem será exibida para informá-lo de que suas configurações serão aplicadas na próxima vez que você iniciar o Microsoft Office. Feche o PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="542e4-p173">A message is displayed to inform you that your settings will be applied the next time you start Microsoft Office. Close PowerPoint.</span></span>

## <a name="build-and-run-the-project"></a><span data-ttu-id="542e4-362">Criar e executar o projeto</span><span class="sxs-lookup"><span data-stu-id="542e4-362">Build and run the project</span></span>

<span data-ttu-id="542e4-p174">Há duas maneiras de criar e executar o projeto dependendo se você estiver ou não usando o Visual Studio Code. Em ambas as maneiras, o projeto cria e recria automaticamente e entra novamente em execução quando você faz alterações no código.</span><span class="sxs-lookup"><span data-stu-id="542e4-p174">There are two ways to build and run the project depending on whether you are using Visual Studio Code. For both ways, the project builds and automatically rebuilds and reruns when you make changes to the code.</span></span>

1. <span data-ttu-id="542e4-365">Se não estiver usando o Visual Studio Code:</span><span class="sxs-lookup"><span data-stu-id="542e4-365">If you are not using Visual Studio Code:</span></span>
   1. <span data-ttu-id="542e4-366">Abra um nó terminal e vá até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="542e4-366">Open a node terminal and navigate to the root folder of the project.</span></span>
   1. <span data-ttu-id="542e4-367">No terminal, insira **npm run build**.</span><span class="sxs-lookup"><span data-stu-id="542e4-367">In the terminal, enter **npm run build**.</span></span>
   1. <span data-ttu-id="542e4-368">Abra um segundo nó terminal e vá até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="542e4-368">Open a second node terminal and navigate to the root folder of the project.</span></span>
   1. <span data-ttu-id="542e4-369">No terminal, insira **npm run start**.</span><span class="sxs-lookup"><span data-stu-id="542e4-369">In the terminal, enter **npm run start**.</span></span>

1. <span data-ttu-id="542e4-370">Se estiver usando o VS Code:</span><span class="sxs-lookup"><span data-stu-id="542e4-370">If you are using VS Code:</span></span>
   1. <span data-ttu-id="542e4-371">Abra o projeto no VS Code.</span><span class="sxs-lookup"><span data-stu-id="542e4-371">Open the project in VS Code.</span></span>
   1. <span data-ttu-id="542e4-372">Pressione Ctrl+Shift+B para compilar o projeto.</span><span class="sxs-lookup"><span data-stu-id="542e4-372">Press CTRL-SHIFT-B to build the project.</span></span>
   1. <span data-ttu-id="542e4-373">Pressione **F5** para executar o projeto em uma sessão de depuração.</span><span class="sxs-lookup"><span data-stu-id="542e4-373">Press **F5** to run the project in a debugging session.</span></span>


## <a name="add-the-add-in-to-an-office-document"></a><span data-ttu-id="542e4-374">Adicionar o suplemento em um documento do Office</span><span class="sxs-lookup"><span data-stu-id="542e4-374">Add the add-in to an Office document</span></span>

1. <span data-ttu-id="542e4-375">Reinicie o PowerPoint, abra ou crie uma apresentação.</span><span class="sxs-lookup"><span data-stu-id="542e4-375">Restart PowerPoint and open or create a presentation.</span></span>

1. <span data-ttu-id="542e4-376">Se a guia **Desenvolvedor** não estiver visível na faixa de opções, habilite-a através das seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="542e4-376">If the **Developer** tab is not visible on the ribbon, enable it with the following steps:</span></span>
   1. <span data-ttu-id="542e4-377">Navegue até **Arquivo** | **Opções** | **Personalizar faixa de opções**.</span><span class="sxs-lookup"><span data-stu-id="542e4-377">Navigate to **File** | **Options** | **Customize Ribbon**.</span></span>
   1. <span data-ttu-id="542e4-378">Clique na caixa de seleção para habilitar o **Desenvolvedor** na árvore de nomes de controle do lado direito da página **Personalizar faixa de opções**.</span><span class="sxs-lookup"><span data-stu-id="542e4-378">Click the check box to enable **Developer** in the tree of control names on the right of the **Customize Ribbon** page.</span></span>
   1. <span data-ttu-id="542e4-379">Pressione **OK**.</span><span class="sxs-lookup"><span data-stu-id="542e4-379">Press **OK**.</span></span>

1. <span data-ttu-id="542e4-380">Na guia **Desenvolvedor** no PowerPoint, escolha **Meus Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="542e4-380">On the **Developer** tab in PowerPoint, choose **My Add-ins**.</span></span>

1. <span data-ttu-id="542e4-381">Selecione a guia **PASTA COMPARTILHADA**.</span><span class="sxs-lookup"><span data-stu-id="542e4-381">Select the **SHARED FOLDER** tab.</span></span>

1. <span data-ttu-id="542e4-382">Escolha **Exemplo de SSO NodeJS**e selecione **OK**.</span><span class="sxs-lookup"><span data-stu-id="542e4-382">Choose **SSO NodeJS Sample**, and then select **OK**.</span></span>

1. <span data-ttu-id="542e4-383">Na faixa de opções **Página Inicial**, há um novo grupo chamado **SSO NodeJS** com um botão com o rótulo **Mostrar Suplemento** e um ícone.</span><span class="sxs-lookup"><span data-stu-id="542e4-383">On the **Home** ribbon is a new group called **SSO NodeJS** with a button labeled **Show Add-in** and an icon.</span></span>

## <a name="test-the-add-in"></a><span data-ttu-id="542e4-384">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="542e4-384">Test the add-in</span></span>

1. <span data-ttu-id="542e4-385">Certifique-se de ter alguns arquivos no seu OneDrive para que você possa verificar os resultados.</span><span class="sxs-lookup"><span data-stu-id="542e4-385">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

1. <span data-ttu-id="542e4-386">Clique no botão **Exibir Suplemento** para abrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="542e4-386">Click **Show Add-in** button to open the add-in.</span></span>

1. <span data-ttu-id="542e4-p175">O suplemento é aberto na página inicial. Clique no botão **Obter Meus Arquivos do OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="542e4-p175">The add-in opens with a Welcome page. Click the **Get My Files from OneDrive** button.</span></span>

1. <span data-ttu-id="542e4-p176">Se você estiver conectado ao Office, será exibida uma lista de seus arquivos e suas pastas no OneDrive, abaixo do botão. Isso poderá demorar mais de 15 segundos na primeira vez.</span><span class="sxs-lookup"><span data-stu-id="542e4-p176">If you are are signed into Office, a list of your files and folders on OneDrive will appear below the button. This may take more than 15 seconds the first time.</span></span>

1. <span data-ttu-id="542e4-391">Se você não tiver entrado no Office, um pop-up será aberto e pedirá que você entre.</span><span class="sxs-lookup"><span data-stu-id="542e4-391">If you are not signed into Office, a popup will open and prompt you to sign in.</span></span> <span data-ttu-id="542e4-392">Depois de concluir a entrada, a lista de arquivos e pastas aparecerá após alguns segundos.</span><span class="sxs-lookup"><span data-stu-id="542e4-392">After you have completed the sign-in, the list of your files and folders will appear after a few seconds.</span></span> <span data-ttu-id="542e4-393">*Você não deve pressionar o botão uma segunda vez.*</span><span class="sxs-lookup"><span data-stu-id="542e4-393">*You should not press the button a second time.*</span></span>

> [!NOTE]
> <span data-ttu-id="542e4-p178">Se você entrou no Office com uma ID diferente e se alguns aplicativos do Office que estavam abertos no momento continuam abertos, o Office pode não alterar de forma confiável sua ID, mesmo que pareça ter feito isso no PowerPoint. Se isso acontecer, a chamada para o Microsoft Graph pode falhar ou os dados da ID anterior podem ser retornados. Para evitar isso, certifique-se de *fechar todos os outros aplicativos do Office* antes de pressionar **Obter meus arquivos do OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="542e4-p178">If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint. If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned. To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.</span></span>
