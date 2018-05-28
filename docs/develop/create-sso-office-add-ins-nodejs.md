---
title: Criar um Suplemento do Office com Node.js que usa logon ?nico
description: 23/01/2018
ms.openlocfilehash: 4086471bec2ded671e1b3eafebc4fe69e9818344
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="adb27-103">Crie um Suplemento do Office com Node.js que use logon ?nico (pr?via)</span><span class="sxs-lookup"><span data-stu-id="adb27-103">Create a Node.js Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="adb27-p101">Os usu?rios podem entrar no Office, e o Suplemento Web do Office pode aproveitar esse processo de entrada para autoriz?-los a acessar seu suplemento e o Microsoft Graph sem exigir que os eles entrem uma segunda vez. Para obter uma vis?o geral, confira o artigo [Habilitar o SSO em um Suplemento do Office](sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="adb27-p101">Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="adb27-106">Este artigo apresenta o processo passo a passo de habilita??o do logon ?nico (SSO) em um suplemento que foi criado com Node.js e Express.</span><span class="sxs-lookup"><span data-stu-id="adb27-106">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with Node.js and Express.</span></span> 

> [!NOTE]
> <span data-ttu-id="adb27-107">Para ler um artigo semelhante sobre um suplemento baseado em ASP.NET, confira [Criar um Suplemento do Office com ASP.NET que usa o logon ?nico](create-sso-office-add-ins-aspnet.md).</span><span class="sxs-lookup"><span data-stu-id="adb27-107">For a similar article about an ASP.NET-based add-in, see [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="adb27-108">Pr?-requisitos</span><span class="sxs-lookup"><span data-stu-id="adb27-108">Prerequisites</span></span>

* <span data-ttu-id="adb27-109">[Node e npm](https://nodejs.org/en/), vers?o 6.9.4 ou posterior</span><span class="sxs-lookup"><span data-stu-id="adb27-109">[Node and npm](https://nodejs.org/en/), version 6.9.4 or later</span></span>

* <span data-ttu-id="adb27-110">[Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)</span><span class="sxs-lookup"><span data-stu-id="adb27-110">[Git Bash](https://git-scm.com/downloads) (or another git client)</span></span>

* <span data-ttu-id="adb27-111">TypeScript, vers?o 2.2.2 ou posterior</span><span class="sxs-lookup"><span data-stu-id="adb27-111">TypeScript version 2.2.2 or later</span></span>

* <span data-ttu-id="adb27-112">Office 2016, vers?o 1708, build 8424.nnnn ou posterior (a vers?o de assinatura do Office 365, ?s vezes chamada de "Clique para Executar")</span><span class="sxs-lookup"><span data-stu-id="adb27-112">Office 2016, Version 1708, build 8424.nnnn or later (the Office 365 subscription version, sometimes called ?Click to Run?)</span></span>

  <span data-ttu-id="adb27-p102">Talvez seja necess?rio ser um Office Insider para obter essa vers?o. Para saber mais, confira [Seja um Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="adb27-p102">You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="adb27-115">Configure o projeto inicial</span><span class="sxs-lookup"><span data-stu-id="adb27-115">Set up the starter project</span></span>

1. <span data-ttu-id="adb27-116">Clone ou baixe o reposit?rio em [SSO com Suplemento NodeJS do Office](https://github.com/officedev/office-add-in-nodejs-sso).</span><span class="sxs-lookup"><span data-stu-id="adb27-116">Clone or download the repo at [Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="adb27-117">H? duas vers?es do exemplo:</span><span class="sxs-lookup"><span data-stu-id="adb27-117">There are two versions of the sample:</span></span>  
    > * <span data-ttu-id="adb27-p103">A pasta **Before** (antes) traz um projeto inicial. A interface do usu?rio e outros aspectos do suplemento que n?o est?o diretamente ligados ao SSO ou ? autoriza??o j? est?o prontos. As pr?ximas se??es deste artigo apresentam uma orienta??o passo a passo para concluir o projeto.</span><span class="sxs-lookup"><span data-stu-id="adb27-p103">The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.</span></span> 
    > * <span data-ttu-id="adb27-p104">A vers?o **Completed** (conclu?do) do exemplo apresenta como seria o suplemento quando conclu?dos os procedimentos apresentados neste artigo, com exce??o de que o projeto conclu?do traz coment?rios de c?digos que seriam redundantes neste artigo. Para usar a vers?o conclu?da, apenas siga as instru??es apresentadas neste artigo, substituindo "Before" por "Completed" e pulando as se??es **Codificar o lado do cliente** e **Codificar o lado do servidor**.</span><span class="sxs-lookup"><span data-stu-id="adb27-p104">The **Completed** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just follow the instructions in this article, but replace "Before" with "Completed" and skip the sections **Code the client side** and **Code the server** side.</span></span>

2. <span data-ttu-id="adb27-123">Abra um console Git bash na pasta **Before**.</span><span class="sxs-lookup"><span data-stu-id="adb27-123">Open a Git bash console in the **Before** folder.</span></span>

3. <span data-ttu-id="adb27-124">Insira `npm install` no console para instalar todas as depend?ncias discriminadas no arquivo package.json.</span><span class="sxs-lookup"><span data-stu-id="adb27-124">Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.</span></span>

4. <span data-ttu-id="adb27-125">Insira `npm run build ` no console para compilar o projeto.</span><span class="sxs-lookup"><span data-stu-id="adb27-125">Enter `npm run build ` in the console to build the project.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="adb27-p105">Talvez voc? veja alguns erros de build informando que algumas vari?veis est?o declaradas mas n?o s?o usadas. Ignore esses erros. Eles s?o um efeito colateral, pois na vers?o "Before" do exemplo est?o faltando alguns c?digos que ser?o adicionados posteriormente.</span><span class="sxs-lookup"><span data-stu-id="adb27-p105">You may see some build errors saying that some variables are declared but not used. Ignore these errors. They are a side effect of the fact that the "Before" version of the sample is missing some code that will be added later.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="adb27-129">Registre o suplemento com o ponto de extremidade v2.0 do Azure AD</span><span class="sxs-lookup"><span data-stu-id="adb27-129">Register the add-in with Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="adb27-130">As instru??es a seguir foram escritas de modo gen?rico para que possam ser usadas em diversos lugares.</span><span class="sxs-lookup"><span data-stu-id="adb27-130">The following instruction are written generically so they can be used in multiple places.</span></span> <span data-ttu-id="adb27-131">Para este artigo, fa?a o seguinte:</span><span class="sxs-lookup"><span data-stu-id="adb27-131">For this ariticle do the following:</span></span>
- <span data-ttu-id="adb27-132">Substitua o espa?o reservado **$ADD-IN-NAME$** por `?Office-Add-in-NodeJS-SSO`.</span><span class="sxs-lookup"><span data-stu-id="adb27-132">Replace the placeholder **$ADD-IN-NAME$** with `?Office-Add-in-NodeJS-SSO`.</span></span>
- <span data-ttu-id="adb27-133">Substitua o espa?o reservado **$FQDN-WITHOUT-PROTOCOL$** por `localhost:3000`.</span><span class="sxs-lookup"><span data-stu-id="adb27-133">Replace the placeholder **$FQDN-WITHOUT-PROTOCOL$** with `localhost:3000`.</span></span>
- <span data-ttu-id="adb27-134">Quando voc? especificar permiss?es na caixa de di?logo **Selecionar Permiss?es**, marque as caixas para as permiss?es a seguir.</span><span class="sxs-lookup"><span data-stu-id="adb27-134">When you specify permissions in the **Select Permissions** dialog, check the boxes for the following permissions.</span></span> <span data-ttu-id="adb27-135">Apenas a primeira ? realmente necess?ria pelo seu suplemento; mas a `profile` permiss?o ? necess?ria para que o host do Office obtenha um token para seu suplemento de aplicativo da Web.</span><span class="sxs-lookup"><span data-stu-id="adb27-135">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office host to get a token to your add-in web application.</span></span>
    * <span data-ttu-id="adb27-136">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="adb27-136">Files.Read.All</span></span>
    * <span data-ttu-id="adb27-137">perfil</span><span class="sxs-lookup"><span data-stu-id="adb27-137">profile</span></span>

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]


## <a name="grant-administrator-consent-to-the-add-in"></a><span data-ttu-id="adb27-138">Conceder autoriza??o do administrador ao suplemento</span><span class="sxs-lookup"><span data-stu-id="adb27-138">Details are at: Grant administrator consent to the add-in</span></span>

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a><span data-ttu-id="adb27-139">Configure o suplemento</span><span class="sxs-lookup"><span data-stu-id="adb27-139">Configure the add-in</span></span>

1. <span data-ttu-id="adb27-p108">Em seu editor de c?digos, abra o arquivo src\server.ts. Perto da parte superior, h? uma chamada para um construtor de uma classe `AuthModule`. H? alguns par?metros de cadeia de caracteres no construtor aos quais voc? precisa atribuir valores.</span><span class="sxs-lookup"><span data-stu-id="adb27-p108">In your code editor, open the src\server.ts file. Near the top there is a call to a constructor of an `AuthModule` class. There are some string parameters in the constructor to which you need to assign values.</span></span>

2. <span data-ttu-id="adb27-143">Na propriedade `client_id`, substitua o espa?o reservado `{client GUID}` pelo ID do aplicativo que voc? salvou ao registrar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="adb27-143">For the `client_id` property, replace the placeholder `{client GUID}` with the application secret that you saved when you registered the add-in.</span></span> <span data-ttu-id="adb27-144">Ao terminar, deve haver apenas um GUID entre aspas simples.</span><span class="sxs-lookup"><span data-stu-id="adb27-144">When you are done, there should just be a GUID in single quotation marks.</span></span> <span data-ttu-id="adb27-145">N?o deve haver nenhum caractere "{}".</span><span class="sxs-lookup"><span data-stu-id="adb27-145">There should not be any "{}" characters.</span></span>

3. <span data-ttu-id="adb27-146">Na propriedade `client_secret`, substitua o espa?o reservado `{client secret}` pelo segredo do aplicativo que voc? salvou ao registrar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="adb27-146">For the `client_secret` property, replace the placeholder `{client secret}` with the application secret that you saved when you registered the add-in.</span></span>

4. <span data-ttu-id="adb27-p110">Na propriedade `audience`, substitua o espa?o reservado `{audience GUID}` pela ID do aplicativo que voc? salvou ao registrar o suplemento. (Exatamente o mesmo valor que voc? atribuiu ? propriedade `client_id`.)</span><span class="sxs-lookup"><span data-stu-id="adb27-p110">For the `audience` property, replace the placeholder `{audience GUID}` with the application ID that you saved when you registered the add-in. (The very same value that you assigned to the `client_id` property.)</span></span>
  
3. <span data-ttu-id="adb27-149">Na sequ?ncia atribu?da ? propriedade `issuer`, voc? ver? o espa?o reservado *{O365 tenant GUID}*.</span><span class="sxs-lookup"><span data-stu-id="adb27-149">In the string assigned to the `issuer` property, you will see the placeholder *{O365 tenant GUID}*.</span></span> <span data-ttu-id="adb27-150">Substitua-o pelo ID de loca??o do Office 365.</span><span class="sxs-lookup"><span data-stu-id="adb27-150">Replace this with the Office 365 tenancy ID.</span></span> <span data-ttu-id="adb27-151">Use um dos m?todos em [Encontre seu ID de locat?rio do Office 365](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) para obt?-lo.</span><span class="sxs-lookup"><span data-stu-id="adb27-151">Use one of the methods in [Find your Office 365 tenant ID](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) to obtain it.</span></span> <span data-ttu-id="adb27-152">Quando voc? terminar, o `issuer` valor da propriedade deve ser algo como isto:</span><span class="sxs-lookup"><span data-stu-id="adb27-152">When you are done, the `issuer` property value should look something like this:</span></span>

    `https://login.microsoftonline.com/12345678-1234-1234-1234-123456789012/v2.0`

1. <span data-ttu-id="adb27-153">N?o altere os demais par?metros no construtor `AuthModule`.</span><span class="sxs-lookup"><span data-stu-id="adb27-153">Leave the other parameters in the `AuthModule` constructor unchanged.</span></span> <span data-ttu-id="adb27-154">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="adb27-154">Save and close the file.</span></span>

1. <span data-ttu-id="adb27-155">Na raiz do projeto, abra o arquivo do manifesto do suplemento "Office-Add-in-NodeJS-SSO.xml".</span><span class="sxs-lookup"><span data-stu-id="adb27-155">In the root of the project, open the add-in manifest file ?Office-Add-in-NodeJS-SSO.xml?.</span></span>

1. <span data-ttu-id="adb27-156">Role at? o final do arquivo.</span><span class="sxs-lookup"><span data-stu-id="adb27-156">Scroll to the bottom of the file.</span></span>

1. <span data-ttu-id="adb27-157">Logo acima da marca de fim `</VersionOverrides>`, voc? encontrar? a marca??o a seguir:</span><span class="sxs-lookup"><span data-stu-id="adb27-157">Just above the end `</VersionOverrides>` tag, you will find the following markup:</span></span>

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

1. <span data-ttu-id="adb27-158">Substitua o espa?o reservado "{application_GUID here}" *nos dois lugares* na marca??o pelo ID do Aplicativo que voc? copiou ao registrar seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="adb27-158">Replace the placeholder ?{application_GUID here}? *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="adb27-159">(Os "{}"n?o fazem parte do ID, portanto, n?o inclua-os.) Esse ? o mesmo ID usado para o ClientID e Audience no web.config.</span><span class="sxs-lookup"><span data-stu-id="adb27-159">(The "{}" are not part of the ID, so don't include them.) This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > * <span data-ttu-id="adb27-160">O valor de **Resource** ? o **URI da ID do Aplicativo** que voc? definiu quando adicionou a plataforma API Web no registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="adb27-160">The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.</span></span>
    > * <span data-ttu-id="adb27-161">A se??o **Scopes** s? ser? usada para gerar uma caixa de di?logo de consentimento se o suplemento for vendido no AppSource.</span><span class="sxs-lookup"><span data-stu-id="adb27-161">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="adb27-162">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="adb27-162">Save and close the file.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="adb27-163">Codificar o lado do cliente</span><span class="sxs-lookup"><span data-stu-id="adb27-163">Code the client side</span></span>

1. <span data-ttu-id="adb27-p114">Abra o arquivo program.js da pasta **public**. Ele j? apresenta alguns c?digos:</span><span class="sxs-lookup"><span data-stu-id="adb27-p114">Open the program.js file in the **public** folder. It already has some code in it:</span></span>

    * <span data-ttu-id="adb27-166">Uma atribui??o ao m?todo `Office.initialize` que, por sua vez, atribui um manipulador ao evento clicar do bot?o `getGraphAccessTokenButton`.</span><span class="sxs-lookup"><span data-stu-id="adb27-166">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="adb27-167">Um m?todo `showResult` que exibir? os dados retornados do Microsoft Graph (ou uma mensagem de erro) na parte inferior do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="adb27-167">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="adb27-168">Um m?todo `logErrors` que registrar? erros de console que n?o s?o destinados ao usu?rio final.</span><span class="sxs-lookup"><span data-stu-id="adb27-168">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>

11. <span data-ttu-id="adb27-p115">Abaixo da atribui??o a `Office.initialize`, adicione o c?digo a seguir. Observe o seguinte sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="adb27-p115">Below the assignment to `Office.initialize`, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="adb27-171">O processamento de erros no suplemento ?s vezes tentar? novamente obter um token de acesso automaticamente, usando um conjunto diferente de op??es.</span><span class="sxs-lookup"><span data-stu-id="adb27-171">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options.</span></span> <span data-ttu-id="adb27-172">A vari?vel de contador `timesGetOneDriveFilesHasRun` e as vari?veis sinalizador `triedWithoutForceConsent` e `timesMSGraphErrorReceived` s?o usadas para garantir que o usu?rio n?o seja trocado repetidas vezes em tentativas falhas de obter um token.</span><span class="sxs-lookup"><span data-stu-id="adb27-172">The counter variable `timesGetOneDriveFilesHasRun`, and the flag variables `triedWithoutForceConsent` and `timesMSGraphErrorReceived` are used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span> 
    * <span data-ttu-id="adb27-p117">Voc? criar? um m?todo `getDataWithToken` na pr?xima etapa, mas observe que ele define uma op??o chamada `forceConsent` como `false`. Trataremos mais disso na etapa seguinte.</span><span class="sxs-lookup"><span data-stu-id="adb27-p117">You create the `getDataWithToken` method in the next step, but note that it sets an option called `forceConsent` to `false`. More about that in the next step.</span></span>

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;
    var timesMSGraphErrorReceived = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }   
    ```

1. <span data-ttu-id="adb27-p118">Abaixo do m?todo `getOneDriveFiles`, adicione o c?digo a seguir. Observe o seguinte sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="adb27-p118">Below the `getOneDriveFiles` method, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="adb27-p119">O `getAccessTokenAsync` ? a nova API no Office.js que permite que um suplemento solicite ao aplicativo host do Office (Excel, PowerPoint, Word, etc.) um token de acesso para o suplemento (para o usu?rio conectado ao Office). O aplicativo host do Office, por sua vez, solicita o token ao ponto de extremidade 2.0 do Azure AD. Uma vez que voc? previamente autorizou o host do Office para o seu suplemento ao registr?-lo, o Azure AD enviar? o token.</span><span class="sxs-lookup"><span data-stu-id="adb27-p119">The `getAccessTokenAsync` is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office). The Office host application, in turn, asks the Azure AD 2.0 endpoint for the token. Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token.</span></span>
    * <span data-ttu-id="adb27-180">Se nenhum usu?rio estiver conectado ao Office, o host do Office solicitar? que o usu?rio se conecte.</span><span class="sxs-lookup"><span data-stu-id="adb27-180">If no user is signed into Office, the Office host will prompt the user to sign in.</span></span>
    * <span data-ttu-id="adb27-181">O par?metro de op??es configura o `forceConsent` como `false`. Dessa forma, n?o ser? solicitado que o usu?rio consinta o acesso ao host do Office ao seu suplemento sempre que ele o usar.</span><span class="sxs-lookup"><span data-stu-id="adb27-181">The options parameter sets `forceConsent` to `false`, so the user will not be prompted to consent to giving the Office host access to your add-in every time she or he uses the add-in.</span></span> <span data-ttu-id="adb27-182">Na primeira vez que o usu?rio tiver o suplemento, a chamada de `getAccessTokenAsync` falhar?, mas l?gica de processamento de erros que voc? adicionar? em uma etapa posterior ser? automaticamente chamada com a op??o `forceConsent` definida como `true` e o usu?rio ser? solicitado a consentir, mas somente essa primeira vez.</span><span class="sxs-lookup"><span data-stu-id="adb27-182">The first time the user runs the add-in, the call of `getAccessTokenAsync` will fail, but error-handling logic that you add in a later step will automatically re-call with the `forceConsent` option set to `true` and the user will be prompted to consent, but only that first time.</span></span>
    * <span data-ttu-id="adb27-183">Voc? criar? o m?todo `handleClientSideErrors` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="adb27-183">You will create the `handleClientSideErrors` method in a later step.</span></span>

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

1. <span data-ttu-id="adb27-p121">Substitua TODO1 pelas linhas a seguir. Voc? criar? o m?todo `getData` e a rota "/api/values" do lado do servidor nas etapas posteriores. Uma URL relativa ? usada para o ponto de extremidade porque ela deve ser hospedada no mesmo dom?nio que seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="adb27-p121">Replace the TODO1 with the following lines. You create the `getData` method and the server-side ?/api/values? route in later steps. A relative URL is used for the endpoint because it must be hosted on the same domain as your add-in.</span></span>

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. <span data-ttu-id="adb27-p122">Abaixo do m?todo `getOneDriveFiles`, adicione o seguinte. Observe isto sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="adb27-p122">Below the `getOneDriveFiles` method, add the following. About this code, note:</span></span>

    * <span data-ttu-id="adb27-p123">Este m?todo utilit?rio chama um ponto de extremidade da API Web especificado e transmite a ela o mesmo token de acesso que aplicativo host do Office usou para obter acesso ao seu suplemento. No lado do servidor, esse token de acesso ser? usado no fluxo "on behalf of" (em nome de) para obter um token de acesso para o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="adb27-p123">This method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the ?on behalf of? flow to obtain an access token to Microsoft Graph.</span></span>
    * <span data-ttu-id="adb27-191">Voc? criar? o m?todo `handleServerSideErrors` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="adb27-191">You will create the `handleServerSideErrors` method in a later step.</span></span>

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

### <a name="create-the-error-handling-methods"></a><span data-ttu-id="adb27-192">Crie os m?todos de processamento de erros</span><span class="sxs-lookup"><span data-stu-id="adb27-192">Create the error-handling methods</span></span>

1. <span data-ttu-id="adb27-193">Abaixo do m?todo `getData`, adicione o m?todo a seguir.</span><span class="sxs-lookup"><span data-stu-id="adb27-193">Below the `getData` method, add the following method.</span></span> <span data-ttu-id="adb27-194">Esse m?todo processar? os erros no cliente do suplemento quando o host do Office n?o puder obter um token de acesso para o servi?o Web do suplemento.</span><span class="sxs-lookup"><span data-stu-id="adb27-194">This method will handle errors in the add-in's client when the Office host is unable to obtain an access token to the add-in's web service.</span></span> <span data-ttu-id="adb27-195">Esses erros s?o relatados com um c?digo de erro, portanto, o m?todo usa uma instru??o `switch` para distingui-los.</span><span class="sxs-lookup"><span data-stu-id="adb27-195">These errors are reported with an error code, so the method uses a `switch` statement to distinguish them.</span></span>

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

1. <span data-ttu-id="adb27-196">Substitua `TODO2` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="adb27-196">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="adb27-197">O erro 13001 ocorre quando o usu?rio n?o est? conectado ou quando ele cancela, sem responder, uma solicita??o para fornecer um segundo fator de autentica??o.</span><span class="sxs-lookup"><span data-stu-id="adb27-197">Error 13001 occurs when the user is not logged in, or the user cancelled, without responding, a prompt to provide a 2nd authentication factor.</span></span> <span data-ttu-id="adb27-198">Em ambos os casos, o c?digo executar? novamente o m?todo `getDataWithToken` e definir? uma op??o para for?ar uma solicita??o de entrada.</span><span class="sxs-lookup"><span data-stu-id="adb27-198">In either case, the code re-runs the `getDataWithToken` method and sets an option to force a sign-in prompt.</span></span>

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. <span data-ttu-id="adb27-199">Substitua `TODO3` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="adb27-199">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="adb27-200">O erro 13002 ocorre quando a entrada ou o consentimento do usu?rio ? anulado.</span><span class="sxs-lookup"><span data-stu-id="adb27-200">Error 13002 occurs when user's sign-in or consent was aborted.</span></span> <span data-ttu-id="adb27-201">Pe?a que o usu?rio tente novamente, mas n?o mais de uma vez.</span><span class="sxs-lookup"><span data-stu-id="adb27-201">Ask the user to try again but no more than once again.</span></span>

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }          
        break; 
    ```

1. <span data-ttu-id="adb27-202">Substitua `TODO4` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="adb27-202">Replace `TODO4` with the following code.</span></span> <span data-ttu-id="adb27-203">O erro 13003 ocorre quando o usu?rio est? conectado com uma conta que n?o ? corporativa, de estudante nem da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="adb27-203">Error 13003 occurs when user is logged in with an account that is neither work or school, nor Micrososoft Account.</span></span> <span data-ttu-id="adb27-204">Pe?a que o usu?rio saia e entre novamente com um tipo de conta suportado.</span><span class="sxs-lookup"><span data-stu-id="adb27-204">Ask the user to sign-out and then in again with a supported account type.</span></span>

    ```javascript
    case 13003: 
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;   
    ```

    > [!NOTE]
    > <span data-ttu-id="adb27-205">Os erros 13004 e 13005 n?o s?o processados neste m?todo, pois eles s? ocorrem em desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="adb27-205">Errors 13004 and 13005 are not handled in this method because they should only occur in development.</span></span> <span data-ttu-id="adb27-206">Eles n?o podem ser corrigidos pelo c?digo de tempo de execu??o e n?o seria ?til report?-lo a um usu?rio final.</span><span class="sxs-lookup"><span data-stu-id="adb27-206">They cannot be fixed by runtime code and there would be no point in reporting them to an end user.</span></span>

1. <span data-ttu-id="adb27-p129">Substitua `TODO5` pelo seguinte c?digo. O Erro 13006 ocorre quando houve um erro n?o especificado no host do Office, que pode indicar a instabilidade do host. Pe?a ao usu?rio para reiniciar o Office.</span><span class="sxs-lookup"><span data-stu-id="adb27-p129">Replace `TODO5` with the following code. Error 13006 occurs when there has been an unspecified error in the Office host that may indicate that the host is in an unstable state. Ask the user to restart Office.</span></span>

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;        
    ```

1. <span data-ttu-id="adb27-210">Substitua `TODO6` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="adb27-210">Replace `TODO6` with the following code.</span></span> <span data-ttu-id="adb27-211">O erro 13007 ocorre quando algo deu errado com a intera??o do host do Office com o AAD de forma que o host n?o pode obter um token de acesso para o servi?o Web/aplicativo dos suplementos.</span><span class="sxs-lookup"><span data-stu-id="adb27-211">Error 13007 occurs when something has gone wrong with the Office host's interaction with AAD so the host cannot get an access token to the add-ins web service/application.</span></span> <span data-ttu-id="adb27-212">? poss?vel que esse seja um problema de rede tempor?rio.</span><span class="sxs-lookup"><span data-stu-id="adb27-212">This may be a temporary network issue.</span></span> <span data-ttu-id="adb27-213">Pe?a que o usu?rio tente novamente mais tarde.</span><span class="sxs-lookup"><span data-stu-id="adb27-213">Ask the user to try again later.</span></span>

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;      
    ```

1. <span data-ttu-id="adb27-p131">Substitua `TODO7` pelo c?digo a seguir. O Erro 13008 ocorre quando o usu?rio aciona uma opera??o que chama `getAccessTokenAsync` antes que uma chamada anterior dele seja conclu?da.</span><span class="sxs-lookup"><span data-stu-id="adb27-p131">Replace `TODO7` with the following code. Error 13008 occurs when the user tiggered an operation that calls `getAccessTokenAsync` before a previous call of it completed.</span></span>

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```      

1. <span data-ttu-id="adb27-216">Substitua `TODO8` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="adb27-216">Replace `TODO8` with the following code.</span></span> <span data-ttu-id="adb27-217">O erro 13009 ocorre quando o suplemento n?o permite for?ar consentimento, mas `getAccessTokenAsync` foi chamado com a op??o `forceConsent` definida como `true`.</span><span class="sxs-lookup"><span data-stu-id="adb27-217">Error 13009 occurs when the add-in does not support forcing consent, but `getAccessTokenAsync` was called with the `forceConsent` option set to `true`.</span></span> <span data-ttu-id="adb27-218">Normalmente, quando isso acontece, o c?digo deve ser reexecutar `getAccessTokenAsync` automaticamente com a op??o de consentimento definida como `false`.</span><span class="sxs-lookup"><span data-stu-id="adb27-218">In the usual case when this happens the code should automatically re-run `getAccessTokenAsync` with the consent option set to `false`.</span></span> <span data-ttu-id="adb27-219">No entanto, em alguns casos, chamar o m?todo com `forceConsent` definido como `true` ? uma resposta autom?tica para um erro em uma chamada para o m?todo com a op??o definida como `false`.</span><span class="sxs-lookup"><span data-stu-id="adb27-219">However, in some cases, calling the method with `forceConsent` set to `true` was itself an automatic response to an error in a call to the method with the option set to `false`.</span></span> <span data-ttu-id="adb27-220">Nesse caso, o c?digo n?o deve tentar novamente, mas, em vez disso, ele deve solicitar que o usu?rio saia e entre novamente.</span><span class="sxs-lookup"><span data-stu-id="adb27-220">In that case, the code should not try again, but instead it should advise the user to sign out and sign in again.</span></span>

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```      
    
1. <span data-ttu-id="adb27-221">Substitua `TODO9` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="adb27-221">Replace `TODO9` with the following code.</span></span>

    ```javascript
    default:
        logError(result);
        break;
    ```  

1. <span data-ttu-id="adb27-p133">Abaixo do m?todo `handleClientSideErrors`, adicione o seguinte m?todo. Esse m?todo processar? os erros no servi?o Web do suplemento quando algo der errado na execu??o do fluxo on-behalf-of ou ao obter dados do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="adb27-p133">Below the `handleClientSideErrors` method, add the following method. This method will handle errors in the add-in's web service when something goes wrong in executing the on-behalf-of flow or in getting data from Microsoft Graph.</span></span>

    ```javascript
    function handleServerSideErrors(result) {
    
        // TODO10: Handle the case where AAD asks for an additional form of authentication.

        // TODO11: Handle the case where consent has not been granted, or has been revoked.

        // TODO12: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow

        // TODO13: Handle the case where the token that the add-in's client-side sends to it's 
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO14: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        // TODO15: Log all other server errors.
    }
    ```

1. <span data-ttu-id="adb27-p134">Substitua `TODO10` pelo c?digo a seguir. Observa??o sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="adb27-p134">Replace `TODO10` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="adb27-p135">Existem configura??es do Azure Active Directory nas quais o usu?rio precisa fornecer fator(es) de autentica??o adicional(ais) para acessar alguns objetivos do Microsoft Graph (por exemplo, o OneDrive), mesmo que o usu?rio possa fazer login no Office apenas com uma senha. Nesse caso, o AAD enviar? uma resposta com o erro 50076, que tem uma propriedade `Claims`.</span><span class="sxs-lookup"><span data-stu-id="adb27-p135">There are configurations of Azure Active Directory in which the user is required to provide additional authentication factor(s) to access some Microsoft Graph targets (e.g., OneDrive), even if the user can sign on to Office with just a password. In that case, AAD will send a response, with error 50076, that has a `Claims` property.</span></span> 
    * <span data-ttu-id="adb27-228">O host do Office deve obter um novo token com o valor **Claims** como a op??o `authChallenge`.</span><span class="sxs-lookup"><span data-stu-id="adb27-228">The Office host should get a new token with the **Claims** value as the `authChallenge` option.</span></span> <span data-ttu-id="adb27-229">Isso instrui o AAD a solicitar ao usu?rio todas as formas de autentica??o requeridas.</span><span class="sxs-lookup"><span data-stu-id="adb27-229">This tells AAD to prompt the user for all required forms of authentication.</span></span> 

    ```javascript
    if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 50076){
        getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
    }
    ```

1. <span data-ttu-id="adb27-p137">Substitua `TODO11` pelo seguinte c?digo *logo abaixo da ?ltima chave de fechamento do c?digo adicionado na etapa anterior*. Observa??o sobre esse c?digo:</span><span class="sxs-lookup"><span data-stu-id="adb27-p137">Replace `TODO11` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="adb27-232">O erro 65001 significa que o consentimento para acessar o Microsoft Graph n?o foi concedido (ou foi revogado) para uma ou mais permiss?es.</span><span class="sxs-lookup"><span data-stu-id="adb27-232">Error 65001 means that consent to access Microsoft Graph was not granted (or was revoked) for one or more permissions.</span></span> 
    * <span data-ttu-id="adb27-233">O suplemento dever? obter um novo token com a op??o `forceConsent` definida como `true`.</span><span class="sxs-lookup"><span data-stu-id="adb27-233">The add-in should get a new token with the `forceConsent` option set to `true`.</span></span>

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 65001){
        showResult(['Please grant consent to this add-in to access your Microsoft Graph data.']);        
        /*
            THE FORCE CONSENT OPTION IS NOT AVAILABLE IN DURING PREVIEW. WHEN SSO FOR
            OFFICE ADD-INS IS RELEASED, REMOVE THE showResult LINE ABOVE AND UNCOMMENT
            THE FOLLOWING LINE.
        */
        // getDataWithToken({ forceConsent: true });
    }
    ```

1. <span data-ttu-id="adb27-p138">Substitua `TODO12` pelo seguinte c?digo *logo abaixo da ?ltima chave de fechamento do c?digo adicionado na etapa anterior*. Observa??o sobre esse c?digo:</span><span class="sxs-lookup"><span data-stu-id="adb27-p138">Replace `TODO12` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="adb27-236">O erro 70011 significa que um escopo inv?lido (permiss?o) foi solicitado.</span><span class="sxs-lookup"><span data-stu-id="adb27-236">Error 70011 means that an invalid scope (permission) has been requested.</span></span> <span data-ttu-id="adb27-237">O suplemento dever? relatar o erro.</span><span class="sxs-lookup"><span data-stu-id="adb27-237">The add-in should report the error.</span></span>
    * <span data-ttu-id="adb27-238">O c?digo registra qualquer outro erro com um n?mero de erro do AAD.</span><span class="sxs-lookup"><span data-stu-id="adb27-238">The code logs any other error with an AAD error number.</span></span>

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 70011){
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. <span data-ttu-id="adb27-p140">Substitua `TODO13` pelo seguinte c?digo *logo abaixo da ?ltima chave de fechamento do c?digo adicionado na etapa anterior*. Observa??o sobre esse c?digo:</span><span class="sxs-lookup"><span data-stu-id="adb27-p140">Replace `TODO13` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="adb27-241">C?digo de servidor criado em uma etapa posterior enviar? a mensagem terminada em `... expected access_as_user` se a o escopo `access_as_user` (permiss?o) n?o for o token de acesso que o cliente do suplemento enviar para o ADD para ser usado no fluxo on-behalf-of.</span><span class="sxs-lookup"><span data-stu-id="adb27-241">Server-side code that you create in a later step will send the message that ends with `... expected access_as_user` if the `access_as_user` scope (permission) is not in the access token that the add-in's client sends to AAD to be used in the on-behalf-of flow.</span></span>
    * <span data-ttu-id="adb27-242">O suplemento dever? relatar o erro.</span><span class="sxs-lookup"><span data-stu-id="adb27-242">The add-in should report the error.</span></span>

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('expected access_as_user') !== -1){
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. <span data-ttu-id="adb27-p141">Substitua `TODO14` pelo seguinte c?digo *logo abaixo da ?ltima chave de fechamento do c?digo adicionado na etapa anterior*. Observa??o sobre esse c?digo:</span><span class="sxs-lookup"><span data-stu-id="adb27-p141">Replace `TODO14` with the following code *just below the last closing brace of the code you added in the previous step*. Note about this code:</span></span>

    * <span data-ttu-id="adb27-245">? improv?vel que um token expirado ou inv?lido seja enviado para o Microsoft Graph, mas, se isso acontecer, o c?digo de servidor que voc? criar? em uma etapa posterior terminar? com a cadeia de caracteres `Microsoft Graph error`.</span><span class="sxs-lookup"><span data-stu-id="adb27-245">It is unlikely that an expired or invalid token will be sent to Microsoft Graph; but if it does happen, the server-side code that you will create in a later step will end with the string `Microsoft Graph error`.</span></span>
    * <span data-ttu-id="adb27-246">Nesse caso, o suplemento dever? iniciar o processo de autentica??o completo ao redefinir o contador `timesGetOneDriveFilesHasRun` e as vari?veis de sinalizador `timesGetOneDriveFilesHasRun` e, em seguida, chamando novamente o m?todo de identificador de bot?o.</span><span class="sxs-lookup"><span data-stu-id="adb27-246">In this case, the add-in should start the entire authentication process over by resetting the `timesGetOneDriveFilesHasRun` counter and `timesGetOneDriveFilesHasRun` flag variables, and then re-calling the button handler method.</span></span> <span data-ttu-id="adb27-247">No entanto, isso deve ser feito apenas uma vez.</span><span class="sxs-lookup"><span data-stu-id="adb27-247">But it should do this only once.</span></span> <span data-ttu-id="adb27-248">Se isso acontecer novamente, o erro deve ser apenas registrado.</span><span class="sxs-lookup"><span data-stu-id="adb27-248">If it happens again, it should just log the error.</span></span>
    * <span data-ttu-id="adb27-249">O c?digo registra o erro se isso acontecer duas vezes em sequ?ncia.</span><span class="sxs-lookup"><span data-stu-id="adb27-249">The code logs the error if it happens twice in succession.</span></span>

    ```javascript
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

1. <span data-ttu-id="adb27-250">Substitua `TODO15` pelo seguinte c?digo *logo abaixo da ?ltima chave de fechamento do c?digo adicionado na etapa anterior*.</span><span class="sxs-lookup"><span data-stu-id="adb27-250">Replace `TODO15` with the following code *just below the last closing brace of the code you added in the previous step*.</span></span>

    ```javascript
    else {
        logError(result);
    }
    ```

## <a name="code-the-server-side"></a><span data-ttu-id="adb27-251">Codifique o lado do servidor</span><span class="sxs-lookup"><span data-stu-id="adb27-251">Code the server side</span></span>

<span data-ttu-id="adb27-252">H? dois arquivos do lado do servidor que precisam ser modificados.</span><span class="sxs-lookup"><span data-stu-id="adb27-252">There are two server-side files that need to be modified.</span></span> 
- <span data-ttu-id="adb27-p143">O src\auth.js fornece fun??es auxiliares de autoriza??o. Ele j? tem membros gen?ricos que s?o usados em uma variedade de fluxos de autoriza??o. ? preciso adicionar fun??es a esse arquivo para implementar o fluxo "on behalf of".</span><span class="sxs-lookup"><span data-stu-id="adb27-p143">The src\auth.js provides authorization helper functions. It already has generic members that are used in a variety of authorization flows. We need to add functions to it that implement the "on behalf of" flow.</span></span>
- <span data-ttu-id="adb27-p144">O arquivo de src\server.js tem os membros b?sicos necess?rios para executar um servidor e o middleware do express. ? necess?rio adicionar fun??es a ele que ajudam a API Web e a p?gina inicial a obterem os dados do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="adb27-p144">The src\server.js file has the basic members need to run a server and express middleware. We need to add functions to it that serve the home page and a Web API for obtaining Microsoft Graph data.</span></span>

### <a name="create-a-method-to-exchange-tokens"></a><span data-ttu-id="adb27-258">Criar um m?todo para troca de tokens</span><span class="sxs-lookup"><span data-stu-id="adb27-258">Create a method to exchange tokens</span></span>

1. <span data-ttu-id="adb27-p145">Abra o arquivo \src\auth.ts. Adicione o m?todo abaixo ? classe `AuthModule`. Observe o seguinte sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="adb27-p145">Open the \src\auth.ts file. Add the method below to the `AuthModule` class. Note the following about this code:</span></span>

    * <span data-ttu-id="adb27-p146">O par?metro `jwt` ? o token de acesso ao aplicativo. No fluxo de "on behalf of" (em nome de), ele ? trocado com AAD por um token de acesso ao recurso.</span><span class="sxs-lookup"><span data-stu-id="adb27-p146">The `jwt` parameter is the access token to the application. In the "on behalf of" flow, it is exchanged with AAD for an access token to the resource.</span></span>
    * <span data-ttu-id="adb27-264">O par?metro scopes (escopos) tem um valor padr?o, mas neste exemplo ser? substitu?do pelo c?digo de chamada.</span><span class="sxs-lookup"><span data-stu-id="adb27-264">The scopes parameter has a default value, but in this sample it will be overridden by the calling code.</span></span>
    * <span data-ttu-id="adb27-p147">O par?metro de recurso ? opcional. N?o deve ser usado quando o STS ? o ponto de extremidade V 2.0 do AAD. ele infere o recurso dos escopos e retorna um erro se um recurso ? enviado na Solicita??o HTTP.</span><span class="sxs-lookup"><span data-stu-id="adb27-p147">The resource parameter is optional. It should not be used when the STS is the AAD V 2.0 endpoint. The V 2.0 endpoint infers the resource from the scopes and it returns an error if a resource is sent in the HTTP Request.</span></span> 
    * <span data-ttu-id="adb27-268">Gerar uma exce??o no bloco `catch` *n?o* causar? o envio imediato do "500 Erro Interno do Servidor" para o cliente.</span><span class="sxs-lookup"><span data-stu-id="adb27-268">Throwing an exception in the `catch` block will *not* cause an immediate "500 Internal Server Error" to be sent to the client.</span></span> <span data-ttu-id="adb27-269">Chamar o c?digo no arquivo server.js acionar? essa exce??o e a transformar? em uma mensagem de erro que ser? enviada para o cliente.</span><span class="sxs-lookup"><span data-stu-id="adb27-269">Calling code in the server.js file will catch this exception and turn it into an error message that is sent to the client.</span></span>

        ```javascript
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

2. <span data-ttu-id="adb27-p149">Substitua `TODO3` pelo c?digo a seguir. Sobre este c?digo, observe:</span><span class="sxs-lookup"><span data-stu-id="adb27-p149">Replace `TODO3` with the following code. About this code, note:</span></span>
    * <span data-ttu-id="adb27-p150">Um STS com suporte para o fluxo "on behalf of" espera determinados pares de valor/propriedade no corpo da solicita??o HTTP. Esse c?digo constr?i um objeto que se tornar? o corpo da solicita??o.</span><span class="sxs-lookup"><span data-stu-id="adb27-p150">An STS that supports the "on behalf of" flow expects certain property/value pairs in the body of the HTTP request. This code constructs an object that will become the body of the request.</span></span> 
    * <span data-ttu-id="adb27-274">Uma propriedade de recurso ? adicionada ao corpo se, e somente se, um recurso ? transmitido para o m?todo.</span><span class="sxs-lookup"><span data-stu-id="adb27-274">A resource property is added to the body if, and only if, a resource was passed to the method.</span></span>

        ```javascript
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

3. <span data-ttu-id="adb27-275">Substitua `TODO4` pelo c?digo a seguir que envia a solicita??o HTTP para o ponto de extremidade do token do STS.</span><span class="sxs-lookup"><span data-stu-id="adb27-275">Replace `TODO4` with the following code which sends the HTTP request to the token endpoint of the STS.</span></span>

    ```javascript
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    }); 
    ```

4. <span data-ttu-id="adb27-276">Substitua `TODO5` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="adb27-276">Replace `TODO5` with the following code.</span></span> <span data-ttu-id="adb27-277">Observe que gerar uma exce??o *n?o* causar? o envio imediato do "500 Erro Interno do Servidor" para o cliente.</span><span class="sxs-lookup"><span data-stu-id="adb27-277">Note that throwing an exception will *not* cause an immediate "500 Internal Server Error" to be sent to the client.</span></span> <span data-ttu-id="adb27-278">Chamar o c?digo no arquivo server.js acionar? essa exce??o e a transformar? em uma mensagem de erro que ser? enviada para o cliente.</span><span class="sxs-lookup"><span data-stu-id="adb27-278">Calling code in the server.js file will catch this exception and turn it into an error message that is sent to the client.</span></span>

    ```javascript
     if (res.status !== 200) {
        const exception = await res.json();
        throw exception;                
    } 
    ```

5. <span data-ttu-id="adb27-p152">Substitua `TODO6` pelo c?digo a seguir. Observe que o c?digo persiste no token de acesso ao recurso, e ? a hora de expira??o, al?m de retorn?-lo. O c?digo de chamada pode evitar chamadas desnecess?rias ao STS reutilizando um token de acesso n?o expirado ao recurso. Voc? ver? como fazer isso na pr?xima se??o.</span><span class="sxs-lookup"><span data-stu-id="adb27-p152">Replace `TODO6` with the following code. Note that the code persists the access token to the resource, and it's expiration time, in addition to returning it. Calling code can avoid unnecessary calls to the STS by reusing an unexpired access token to the resource. You'll see how to do that in the next section.</span></span>

    ```javascript  
    const json = await res.json();
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken; 
    ```

6. <span data-ttu-id="adb27-283">Salve o arquivo, mas n?o o feche.</span><span class="sxs-lookup"><span data-stu-id="adb27-283">Save the file, but don't close it.</span></span>

### <a name="create-a-method-to-get-access-to-the-resource-using-the-on-behalf-of-flow"></a><span data-ttu-id="adb27-284">Criar um m?todo para obter acesso ao recurso usando o fluxo "on behalf of"</span><span class="sxs-lookup"><span data-stu-id="adb27-284">Create a method to get access to the resource using the "on behalf of" flow</span></span>

1. <span data-ttu-id="adb27-p153">Ainda no arquivo src/auth.ts, adicione o m?todo abaixo ? classe `AuthModule`. Observe o seguinte sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="adb27-p153">Still in src/auth.ts, add the method below to the `AuthModule` class. Note the following about this code:</span></span>

    * <span data-ttu-id="adb27-287">Os coment?rios acima sobre os par?metros para o m?todo `exchangeForToken` aplicam-se aos par?metros deste m?todo tamb?m.</span><span class="sxs-lookup"><span data-stu-id="adb27-287">The comments above about the parameters to the the `exchangeForToken` method apply to the parameters of this method as well.</span></span>
    * <span data-ttu-id="adb27-p154">O m?todo primeiro verifica o armazenamento persistente para um token de acesso ao recurso que n?o expirou e n?o vai expirar no pr?ximo minuto. Ele chama o m?todo `exchangeForToken` que voc? criou na ?ltima se??o somente se necess?rio.</span><span class="sxs-lookup"><span data-stu-id="adb27-p154">The method first checks the persistent storage for an access token to the resource that has not expired and is not going to expire in the next minute. It calls the `exchangeForToken` method you created in the last section only if it needs to.</span></span>

    ```javascript
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

2. <span data-ttu-id="adb27-290">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="adb27-290">Save and close the file.</span></span>

### <a name="create-the-endpoints-that-will-serve-the-add-ins-home-page-and-data"></a><span data-ttu-id="adb27-291">Criar os pontos de extremidade que servir?o aos dados e ? p?gina inicial do suplemento</span><span class="sxs-lookup"><span data-stu-id="adb27-291">Create the endpoints that will serve the add-in's home page and data</span></span>

1. <span data-ttu-id="adb27-292">Abra o arquivo src\server.ts.</span><span class="sxs-lookup"><span data-stu-id="adb27-292">Open the src\server.ts file.</span></span> 

2. <span data-ttu-id="adb27-p155">Adicione o m?todo a seguir na parte inferior do arquivo. Esse m?todo servir? ? p?gina inicial do suplemento. O manifesto do suplemento especifica a URL da p?gina inicial.</span><span class="sxs-lookup"><span data-stu-id="adb27-p155">Add the following method to the bottom of the file. This method will serve the add-in's home page. The add-in manifest specifies the home page URL.</span></span>

    ```javascript
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    })); 
    ```

3. <span data-ttu-id="adb27-p156">Adicione o m?todo a seguir na parte inferior do arquivo. Este m?todo lidar? com todas as solicita??es para a API `onedriveitems`.</span><span class="sxs-lookup"><span data-stu-id="adb27-p156">Add the following method to bottom of the file. This method will handle any requests for the `onedriveitems` API.</span></span>
    ```javascript
    app.get('/api/onedriveitems', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token 
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage 
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Relay any errors from Microsoft Graph to the client.
        // TODO11: Send to the client only the data that it actually needs.
    })); 
    ```

4. <span data-ttu-id="adb27-p157">Substitua `TODO7` pelo seguinte c?digo que valida o token de acesso recebido do aplicativo host do Office. O m?todo `verifyJWT` ? definido no arquivo src\auth.ts. Ele sempre valida a audi?ncia e o emissor. Usamos o par?metro opcional para especificar que tamb?m desejamos que ele verifique se o escopo no token de acesso ? `access_as_user`. Esta ? a ?nica permiss?o ao suplemento que o usu?rio e o host do Office precisam para obter um token de acesso para o Microsoft Graph por meio do fluxo "on behalf of".</span><span class="sxs-lookup"><span data-stu-id="adb27-p157">Replace `TODO7` with the following code which validates the access token received from the Office host application. The `verifyJWT` method is defined in the src\auth.ts file. It always validates the audience and the issuer. We use the optional parameter to specify that we also want it to verify that the scope in the access token is `access_as_user`. This is the only permisison to the add-in that the user and the Office host need in order to get an access token to Microsoft Graph by means of the "on behalf" flow.</span></span> 

    ```javascript
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    ```

    > [!NOTE]
    > <span data-ttu-id="adb27-p158">Voc? deve usar apenas o escopo `access_as_user` para autorizar a API que lida com o fluxo Em Nome De para os suplementos do Office. Outras APIs em seu servi?o devem ter seus pr?prios requisitos de escopo. Isso limita o que pode ser acessado com os tokens que o Office adquire.</span><span class="sxs-lookup"><span data-stu-id="adb27-p158">You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office add-ins. Other APIs in your service should have their own scope requirements. This limits what can be accessed with the tokens that Office acquires.</span></span>

5. <span data-ttu-id="adb27-p159">Substitua `TODO8` pelo c?digo a seguir. Observe o seguinte sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="adb27-p159">Replace `TODO8` with the following code. Note the following about this code:</span></span>

    * <span data-ttu-id="adb27-307">A chamada para `acquireTokenOnBehalfOf` n?o inclui um par?metro de recurso porque constru?mos o objeto `AuthModule` (`auth`) com o ponto de extremidade V2.0 do AAD que n?o oferece suporte ? propriedade de recurso.</span><span class="sxs-lookup"><span data-stu-id="adb27-307">The call to `acquireTokenOnBehalfOf` does not include a resource parameter because we constructed the `AuthModule` object (`auth`) with the AAD V2.0 endpoint which does not support a resource property.</span></span>
    * <span data-ttu-id="adb27-308">O segundo par?metro da chamada especifica as permiss?es que o suplemento precisar? para obter uma lista dos arquivos e das pastas do usu?rio no OneDrive.</span><span class="sxs-lookup"><span data-stu-id="adb27-308">The second parameter of the call specifies the permissions the add-in will need to get a list of the user's files and folders on OneDrive.</span></span> <span data-ttu-id="adb27-309">(A permiss?o `profile` n?o ? solicitada, porque s? ? necess?ria quando o host do Office obt?m o token de acesso ao seu suplemento, e n?o quando voc? est? negociando nesse token para um token de acesso para o Microsoft Graph.)</span><span class="sxs-lookup"><span data-stu-id="adb27-309">(The `profile` permission is not requested because it is only needed when the Office host gets the access token to your add-in, not when you are trading in that token for an access token to Microsoft Graph.)</span></span>

    ```javascript
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    ```

6. <span data-ttu-id="adb27-p161">Substitua `TODO9` pela linha a seguir. Observe o seguinte sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="adb27-p161">Replace `TODO9` with the following line. Note the following about this code:</span></span>

    * <span data-ttu-id="adb27-312">A classe MSGraphHelper ? definida no src\msgraph-helper.ts.</span><span class="sxs-lookup"><span data-stu-id="adb27-312">The MSGraphHelper class is defined in src\msgraph-helper.ts.</span></span> 
    * <span data-ttu-id="adb27-313">Podemos minimizar os dados que devem ser retornados especificando que s? queremos a propriedade de nome e somente os tr?s primeiros itens.</span><span class="sxs-lookup"><span data-stu-id="adb27-313">We minimize the data that must be returned by specifying that we only want the name property and only the first 3 items.</span></span>

    `const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");`

7. <span data-ttu-id="adb27-314">Substitua `TODO10` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="adb27-314">Replace `TODO10` with the following code.</span></span> <span data-ttu-id="adb27-315">Observe que esse c?digo processa erros "401 N?o Autorizado" do Microsoft Graph que indicariam um token expirado ou inv?lido.</span><span class="sxs-lookup"><span data-stu-id="adb27-315">Note that this code handles '401 Unauthorized" errors from Microsoft Graph which would indicate an expired or invalid token.</span></span> <span data-ttu-id="adb27-316">? muito improv?vel que isso aconte?a, pois a l?gica persistente do token impede essa situa??o.</span><span class="sxs-lookup"><span data-stu-id="adb27-316">It is very unlikely that this would ever happen since the token persisting logic should prevent it.</span></span> <span data-ttu-id="adb27-317">(Confira a se??o **Criar um m?todo para obter acesso ao recurso usando o fluxo "on behalf of"** acima.) Se isso acontecer, o c?digo transmitir? o erro para o cliente com "Erro do Microsoft Graph" no nome do erro.</span><span class="sxs-lookup"><span data-stu-id="adb27-317">(See the section **Create a method to get access to the resource using the "on behalf of" flow** above.) If it does happen, this code will relay the error to the client with "Microsoft Graph error" in the error name.</span></span> <span data-ttu-id="adb27-318">(Confira o m?todo `handleClientSideErrors` que voc? criou no arquivo program.js em uma etapa anterior.) O c?digo adicionado ao arquivo ODataHelper.js em uma etapa posterior ajuda a processar erros do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="adb27-318">(See the `handleClientSideErrors` method that you created in the program.js file in an earlier step.) Code that you add to the ODataHelper.js file in a later step helps process errors from Microsoft Graph.</span></span>

    ```javascript
    if (graphData.code) {
        if (graphData.code === 401) {
            throw new UnauthorizedError('Microsoft Graph error', graphData);
        }
    }
    ```


1. <span data-ttu-id="adb27-p163">Substitua `TODO11` pelo c?digo a seguir. Observe que o Microsoft Graph retorna alguns metadados OData e uma propriedade **eTag** para cada item, mesmo se `name` ? a ?nica propriedade solicitada. O c?digo envia somente os nomes de item para o cliente.</span><span class="sxs-lookup"><span data-stu-id="adb27-p163">Replace `TODO11` with the following code. Note that Microsoft Graph returns some OData metadata and an **eTag** property for every item, even if `name` is the only property requested. The code sends only the item names to the client.</span></span>

    ```javascript
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

8. <span data-ttu-id="adb27-322">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="adb27-322">Save and close the file.</span></span>

### <a name="add-response-handling-to-the-odatahelper"></a><span data-ttu-id="adb27-323">Adicione processamento de respostas ao ODataHelper</span><span class="sxs-lookup"><span data-stu-id="adb27-323">Add response handling to the ODataHelper</span></span>

1. <span data-ttu-id="adb27-324">Abra o arquivo src\odata-helper.ts.</span><span class="sxs-lookup"><span data-stu-id="adb27-324">Open the file src\odata-helper.ts.</span></span> <span data-ttu-id="adb27-325">O arquivo est? quase pronto.</span><span class="sxs-lookup"><span data-stu-id="adb27-325">The file is almost complete.</span></span> <span data-ttu-id="adb27-326">O que est? ausente ? o corpo do retorno de chamada para o identificador do evento ?end? da solicita??o.</span><span class="sxs-lookup"><span data-stu-id="adb27-326">What's missing is the body of the callback to the handler for the request "end" event.</span></span> <span data-ttu-id="adb27-327">Substitua o `TODO` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="adb27-327">Replace the `TODO` with the following code.</span></span> <span data-ttu-id="adb27-328">Sobre este c?digo, observe:</span><span class="sxs-lookup"><span data-stu-id="adb27-328">About this code note:</span></span>

    * <span data-ttu-id="adb27-329">A resposta do ponto de extremidade OData pode ser um erro, por exemplo, 401, se o ponto de extremidade exigir um token de acesso e ele for inv?lido ou estiver expirado.</span><span class="sxs-lookup"><span data-stu-id="adb27-329">The response from the OData endpoint might be an error, say a 401 if the endpoint requires an access token and it was invalid or expired.</span></span> <span data-ttu-id="adb27-330">Uma mensagem de erro ? ainda um *mensagem*, n?o um erro, nas chamadas de `https.get`, portanto, a linha `on('error', reject)` no final do `https.get` n?o ? acionada.</span><span class="sxs-lookup"><span data-stu-id="adb27-330">But an error message is still a *message*, not an error in the call of `https.get`, so the `on('error', reject)` line at the end of `https.get` isn't triggered.</span></span> <span data-ttu-id="adb27-331">Portanto, o c?digo distingue mensagens de sucesso (200) de mensagens de erro e envia um objeto JSON para o chamador com o OData solicitado ou informa??es de erro.</span><span class="sxs-lookup"><span data-stu-id="adb27-331">So, the code distinguishes success (200) messages from error messages and sends a JSON object to the caller with either the requested OData or error information.</span></span>

    ```javascript
    var error;
    if (response.statusCode === 200) {
        // TODO1: Return the data to the caller and resolve the Promise.
    } else {
       // TODO2: Return an error object to the caller and resolve the Promise.
    }
    ```

1.  <span data-ttu-id="adb27-p166">Substitua `TODO1` pelo c?digo a seguir. Observe que o c?digo pressup?e que os dados retornados s?o JSON.</span><span class="sxs-lookup"><span data-stu-id="adb27-p166">Replace `TODO1` with the following code. Note that the code assumes the data is returned as JSON.</span></span>

    ```javascript
    let parsedBody = JSON.parse(body);
    resolve(parsedBody);
    ```

1.  <span data-ttu-id="adb27-p167">Substitua `TODO2` pelo c?digo a seguir. Observa??o sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="adb27-p167">Replace `TODO2` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="adb27-336">Uma resposta de erro de uma fonte de OData sempre ter? um statusCode e, normalmente, um statusMessage.</span><span class="sxs-lookup"><span data-stu-id="adb27-336">An error response from an OData source will always have a statusCode and usually a statusMessage.</span></span> <span data-ttu-id="adb27-337">Algumas fontes de OData tamb?m adicionam uma propriedade de erro ao corpo da mensagem com mais informa??es, como uma solicita??o interna ou, mais especificamente, um c?digo e uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="adb27-337">Some OData sources also add an error property to the body with further information, such as an inner, or more specific, code and message.</span></span>
    * <span data-ttu-id="adb27-338">O objeto Promise ? resolvido, n?o rejeitado.</span><span class="sxs-lookup"><span data-stu-id="adb27-338">The Promise object is resolved, not rejected.</span></span> <span data-ttu-id="adb27-339">O `https.get` ? executado quando um servi?o Web chama um ponto de extremidade OData de servidor para servidor.</span><span class="sxs-lookup"><span data-stu-id="adb27-339">The `https.get` runs when a web service calls an OData endpoint server-to-server.</span></span> <span data-ttu-id="adb27-340">No entanto, essa chamada chega no contexto de uma chamada de um cliente para uma Web API do servi?o Web.</span><span class="sxs-lookup"><span data-stu-id="adb27-340">But that call comes in the context of a call from a client to a web API in the web service.</span></span> <span data-ttu-id="adb27-341">A solicita??o "externa" do cliente para o servi?o Web nunca ? conclu?da se essa solicita??o "interna" for rejeitada.</span><span class="sxs-lookup"><span data-stu-id="adb27-341">The "outer" request from the client to the web service never completes if this "inner" request is rejected.</span></span> <span data-ttu-id="adb27-342">Al?m disso, a solicita??o com o objeto `Error` personalizado ? necess?ria se o chamador de `http.get` precisar transmitir erros do ponto de extremidade OData para o cliente.</span><span class="sxs-lookup"><span data-stu-id="adb27-342">Also, resolving the request with the custom `Error` object is required if the caller of `http.get` needs to relay errors from the OData endpoint to the client.</span></span>

    ```javascript
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

1. <span data-ttu-id="adb27-343">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="adb27-343">Save and close the file.</span></span>

## <a name="deploy-the-add-in"></a><span data-ttu-id="adb27-344">Implantar o suplemento</span><span class="sxs-lookup"><span data-stu-id="adb27-344">Deploy the add-in</span></span>

<span data-ttu-id="adb27-345">Agora ? preciso informar ao Office onde encontrar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="adb27-345">Now you need to let Office know where to find the add-in.</span></span>

1. <span data-ttu-id="adb27-346">Crie um compartilhamento de rede ou [compartilhe uma pasta na rede](https://technet.microsoft.com/en-us/library/cc770880.aspx).</span><span class="sxs-lookup"><span data-stu-id="adb27-346">Create a network share, or [share a folder to the network](https://technet.microsoft.com/en-us/library/cc770880.aspx).</span></span>

2. <span data-ttu-id="adb27-347">Coloque uma c?pia do arquivo de manifesto Office-Add-in-NodeJS-SSO.xml, da raiz do projeto, dentro da pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="adb27-347">Place a copy of the Office-Add-in-NodeJS-SSO.xml manifest file, from the root of the project, into the shared folder.</span></span>

3. <span data-ttu-id="adb27-348">Inicie o PowerPoint e abra um documento.</span><span class="sxs-lookup"><span data-stu-id="adb27-348">Launch PowerPoint and open a document.</span></span>

4. <span data-ttu-id="adb27-349">Escolha a guia **Arquivo** e, ent?o, **Op??es**.</span><span class="sxs-lookup"><span data-stu-id="adb27-349">Choose the **File** tab, and then choose **Options**.</span></span>

5. <span data-ttu-id="adb27-350">Escolha **Central de Confiabilidade**, e escolha o bot?o **Configura??es da Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="adb27-350">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>

6. <span data-ttu-id="adb27-351">Escolha **Cat?logos de Suplementos Confi?veis**.</span><span class="sxs-lookup"><span data-stu-id="adb27-351">Choose **Trusted Add-ins Catalogs**.</span></span>

7. <span data-ttu-id="adb27-352">No campo **URL do Cat?logo**, insira o caminho de rede para o compartilhamento de pasta que cont?m o arquivo Office-Add-in-NodeJS-SSO.xml e escolha **Adicionar Cat?logo**.</span><span class="sxs-lookup"><span data-stu-id="adb27-352">In the **Catalog Url** field, enter the network path to the folder share that contains Office-Add-in-NodeJS-SSO.xml, and then choose **Add Catalog**.</span></span>

8. <span data-ttu-id="adb27-353">Selecione a caixa de sele??o **Mostrar no Menu** e, em seguida, escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="adb27-353">Select the **Show in Menu** check box, and then choose **OK**.</span></span>

9. <span data-ttu-id="adb27-p170">Uma mensagem ser? exibida para inform?-lo de que suas configura??es ser?o aplicadas na pr?xima vez que voc? iniciar o Microsoft Office. Feche o PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="adb27-p170">A message is displayed to inform you that your settings will be applied the next time you start Microsoft Office. Close PowerPoint.</span></span>

## <a name="build-and-run-the-project"></a><span data-ttu-id="adb27-356">Criar e executar o projeto</span><span class="sxs-lookup"><span data-stu-id="adb27-356">Build and run the project</span></span>

<span data-ttu-id="adb27-p171">H? duas maneiras de criar e executar o projeto dependendo se voc? estiver ou n?o usando o Visual Studio Code. Em ambas as maneiras, o projeto cria e recria automaticamente e entra novamente em execu??o quando voc? faz altera??es no c?digo.</span><span class="sxs-lookup"><span data-stu-id="adb27-p171">There are two ways to build and run the project depending on whether you are using Visual Studio Code. For both ways, the project builds and automatically rebuilds and reruns when you make changes to the code.</span></span>

1. <span data-ttu-id="adb27-359">Se n?o estiver usando o Visual Studio Code:</span><span class="sxs-lookup"><span data-stu-id="adb27-359">If you are not using Visual Studio Code:</span></span> 
 1. <span data-ttu-id="adb27-360">Abra um n? terminal e v? at? a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="adb27-360">Open a node terminal and navigate to the root folder of the project.</span></span>
 2. <span data-ttu-id="adb27-361">No terminal, insira **npm run build**.</span><span class="sxs-lookup"><span data-stu-id="adb27-361">In the terminal, enter **npm run build**.</span></span> 
 3. <span data-ttu-id="adb27-362">Abra um segundo n? terminal e v? at? a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="adb27-362">Open a second node terminal and navigate to the root folder of the project.</span></span>
 4. <span data-ttu-id="adb27-363">No terminal, insira **npm run start**.</span><span class="sxs-lookup"><span data-stu-id="adb27-363">In the terminal, enter **npm run start**.</span></span>

2. <span data-ttu-id="adb27-364">Se estiver usando o VS Code:</span><span class="sxs-lookup"><span data-stu-id="adb27-364">If you are using VS Code:</span></span>
 1. <span data-ttu-id="adb27-365">Abra o projeto no VS Code.</span><span class="sxs-lookup"><span data-stu-id="adb27-365">Open the project in VS Code.</span></span>
 2. <span data-ttu-id="adb27-366">Pressione Ctrl+Shift+B para compilar o projeto.</span><span class="sxs-lookup"><span data-stu-id="adb27-366">Press CTRL-SHIFT-B to build the project.</span></span>
 3. <span data-ttu-id="adb27-367">Pressione F5 para executar o projeto em uma sess?o de depura??o.</span><span class="sxs-lookup"><span data-stu-id="adb27-367">Press F5 to run the project in a debugging session.</span></span>


## <a name="add-the-add-in-to-an-office-document"></a><span data-ttu-id="adb27-368">Adicionar o suplemento em um documento do Office</span><span class="sxs-lookup"><span data-stu-id="adb27-368">Add the add-in to an Office document</span></span>

1. <span data-ttu-id="adb27-369">Reinicie o PowerPoint, abra ou crie uma apresenta??o.</span><span class="sxs-lookup"><span data-stu-id="adb27-369">Restart PowerPoint and open or create a presentation.</span></span> 

2. <span data-ttu-id="adb27-370">Na guia **Desenvolvedor** no PowerPoint, escolha **Meus Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="adb27-370">On the **Developer** tab in PowerPoint, choose **My Add-ins**.</span></span>

3. <span data-ttu-id="adb27-371">Selecione a guia **PASTA COMPARTILHADA**.</span><span class="sxs-lookup"><span data-stu-id="adb27-371">Select the **SHARED FOLDER** tab.</span></span>

4. <span data-ttu-id="adb27-372">Escolha **Exemplo de SSO NodeJS**e selecione **OK**.</span><span class="sxs-lookup"><span data-stu-id="adb27-372">Choose **SSO NodeJS Sample**, and then select **OK**.</span></span>

5. <span data-ttu-id="adb27-373">Na faixa de op??es **P?gina Inicial**, h? um novo grupo chamado **SSO NodeJS** com um bot?o com o r?tulo **Mostrar Suplemento** e um ?cone.</span><span class="sxs-lookup"><span data-stu-id="adb27-373">On the **Home** ribbon is a new group called **SSO NodeJS** with a button labeled **Show Add-in** and an icon.</span></span> 

## <a name="test-the-add-in"></a><span data-ttu-id="adb27-374">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="adb27-374">Test the add-in</span></span>

1. <span data-ttu-id="adb27-375">Certifique-se de ter alguns arquivos no seu OneDrive para que voc? possa verificar os resultados.</span><span class="sxs-lookup"><span data-stu-id="adb27-375">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

2. <span data-ttu-id="adb27-376">Clique no bot?o **Exibir Suplemento** para abrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="adb27-376">Click **Show Add-in** button to open the add-in.</span></span>

2. <span data-ttu-id="adb27-p172">O suplemento ? aberto na p?gina inicial. Clique no bot?o **Obter Meus Arquivos do OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="adb27-p172">The add-in opens with a Welcome page. Click the **Get My Files from OneDrive** button.</span></span>

2. <span data-ttu-id="adb27-p173">Se voc? estiver conectado ao Office, ser? exibida uma lista de seus arquivos e suas pastas no OneDrive, abaixo do bot?o. Isso poder? demorar mais de 15 segundos na primeira vez.</span><span class="sxs-lookup"><span data-stu-id="adb27-p173">If you are are signed into Office, a list of your files and folders on OneDrive will appear below the button. This may take more than 15 seconds the first time.</span></span>

3. <span data-ttu-id="adb27-p174">Se voc? n?o tiver entrado no Office, um pop-up ser? aberto e pedir? que voc? entre. Depois de concluir a entrada, a lista de arquivos e pastas aparecer? ap?s alguns segundos. *N?o pressione o bot?o uma segunda vez.*</span><span class="sxs-lookup"><span data-stu-id="adb27-p174">If you are not signed into Office, a popup will open and prompt you to sign in. After you have completed the sign-in, the list of your files and folders will appear after a few seconds. *You do not press the button a second time.*</span></span>

> [!NOTE]
> <span data-ttu-id="adb27-384">Se voc? entrou no Office com uma ID diferente e se alguns aplicativos do Office que estavam abertos no momento continuam abertos, o Office pode n?o alterar de forma confi?vel sua ID, mesmo que pare?a ter feito isso no PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="adb27-384">If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint.</span></span> <span data-ttu-id="adb27-385">Se isso acontecer, a chamada para o Microsoft Graph pode falhar ou os dados da ID anterior podem ser retornados.</span><span class="sxs-lookup"><span data-stu-id="adb27-385">If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned.</span></span> <span data-ttu-id="adb27-386">Para evitar isso, certifique-se de *fechar todos os outros aplicativos do Office* antes de pressionar **Obter meus arquivos do OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="adb27-386">To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.</span></span>
