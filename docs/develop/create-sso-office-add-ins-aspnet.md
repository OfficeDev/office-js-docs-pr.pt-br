---
title: Criar um Suplemento do Office com ASP.NET que usa logon ?nico
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 6a1c8ea7a8634d701a43e08fd8bb9c5f9c1863cd
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="6b508-102">Criar um Suplemento do Office com ASP.NET que use logon ?nico (visualiza??o)</span><span class="sxs-lookup"><span data-stu-id="6b508-102">Create an ASP.NET Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="6b508-p101">Quando os usu?rios est?o conectados ao Office, o seu suplemento pode usar as mesmas credenciais para permitir que os usu?rios acessem v?rios aplicativos sem exigir que eles entrem uma segunda vez. Para obter uma vis?o geral, consulte [Habilitar o SSO em um Suplemento do Office](sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="6b508-p101">When users are signed in to Office, your add-in can use the same credentials to permit users to access multiple applications without requiring them to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="6b508-105">Este artigo apresenta o processo passo a passo de habilita??o do logon ?nico (SSO) em um suplemento que foi criado com ASP.NET, OWIN e com a Biblioteca de Autentica??o da Microsoft (MSAL) para .NET.</span><span class="sxs-lookup"><span data-stu-id="6b508-105">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with ASP.NET, OWIN, and Microsoft Authentication Library (MSAL) for .NET.</span></span>

> [!NOTE]
> <span data-ttu-id="6b508-106">Para ler um artigo semelhante sobre um suplemento baseado em Node.js, confira [Criar um Suplemento do Office com Node.js que use logon ?nico](create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="6b508-106">For a similar article about a Node.js-based add-in, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="6b508-107">Pr?-requisitos</span><span class="sxs-lookup"><span data-stu-id="6b508-107">Prerequisites</span></span>

* <span data-ttu-id="6b508-108">A vers?o mais recente dispon?vel do Visual Studio 2017 Preview.</span><span class="sxs-lookup"><span data-stu-id="6b508-108">The latest available version of Visual Studio 2017 Preview.</span></span>

* <span data-ttu-id="6b508-p102">Office 2016, vers?o 1708, build 8424.nnnn ou posterior (a vers?o de assinatura do Office 365, ?s vezes chamada de "Clique para Executar"). Voc? talvez precise ser um participante do programa Office Insider para obter essa vers?o. Para obter mais informa??es, confira a p?gina [Seja um Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="6b508-p102">Office 2016, Version 1708, build 8424.nnnn or later (the Office 365 subscription version, sometimes called ?Click to Run?). You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="6b508-112">Configure o projeto inicial</span><span class="sxs-lookup"><span data-stu-id="6b508-112">Set up the starter project</span></span>

1. <span data-ttu-id="6b508-113">Clone ou baixe o reposit?rio em [SSO com Suplemento ASPNET do Office](https://github.com/officedev/office-add-in-aspnet-sso).</span><span class="sxs-lookup"><span data-stu-id="6b508-113">Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span></span>

1. <span data-ttu-id="6b508-p103">Abra a pasta **Before** (antes) e abra o arquivo .sln no Visual Studio. Esse ? um projeto inicial. A interface do usu?rio e outros aspectos do suplemento que n?o est?o diretamente ligados ao SSO ou ? autoriza??o j? est?o prontos.</span><span class="sxs-lookup"><span data-stu-id="6b508-p103">Open the **Before** folder and open the .sln file in Visual Studio. This is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6b508-p104">H? tamb?m uma vers?o conclu?da do exemplo no mesmo reposit?rio. Essa vers?o apresenta como seria o suplemento quando conclu?dos os procedimentos apresentados neste artigo, com exce??o de que o projeto conclu?do traz coment?rios de c?digos que seriam redundantes neste artigo. Para usar a vers?o conclu?da, apenas abra o arquivo `sln` e siga as instru??es apresentadas neste artigo, mas pule as se??es **Codificar o lado do cliente** e **Codificar o lado do servidor**.</span><span class="sxs-lookup"><span data-stu-id="6b508-p104">There is also a completed version of the sample in the same repo. It is just like the add-in that you would have if you completed the procedures in this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just open the `sln` file and follow the instructions in this article, but skip the sections **Code the client side** and **Code the server** side.</span></span>

1. <span data-ttu-id="6b508-p105">Depois que o projeto for aberto, compile-o no Visual Studio, que instalar? os pacotes listados no arquivo packages.config. Esse procedimento poder? levar entre alguns segundos e alguns minutos dependendo de quantos pacotes estiverem no cache de pacote local do computador.</span><span class="sxs-lookup"><span data-stu-id="6b508-p105">After the project opens, build it in Visual Studio, which will install the packages listed in the packages.config file. This can take a few seconds to several minutes depending on how many of the packages are in the computer's local package cache.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6b508-122">Voc? receber? um erro sobre o namespace Identity.</span><span class="sxs-lookup"><span data-stu-id="6b508-122">You will get an error about the Identity namespace.</span></span> <span data-ttu-id="6b508-123">Este ? um efeito colateral de um problema de configura??o que voc? corrigir? no pr?ximo passo.</span><span class="sxs-lookup"><span data-stu-id="6b508-123">This is a side effect of a configuration issue that you will fix with the next step.</span></span> <span data-ttu-id="6b508-124">O importante ? que os pacotes estejam instalados.</span><span class="sxs-lookup"><span data-stu-id="6b508-124">The important thing is that the packages are installed.</span></span>

1. <span data-ttu-id="6b508-125">Atualmente, a vers?o da biblioteca MSAL (Microsoft.Identity.Client) necess?ria para SSO (vers?o `1.1.1-alpha0393`) n?o faz parte do cat?logo padr?o de nuget, portanto, n?o est? listada no package.config e deve ser instalada separadamente.</span><span class="sxs-lookup"><span data-stu-id="6b508-125">Currently, the version of the MSAL library (Microsoft.Identity.Client) that you need for SSO (version `1.1.1-alpha0393`) is not part of the standard nuget catalog, so it is not listed in the package.config, and it must be installed separately.</span></span> 

   > 1. <span data-ttu-id="6b508-126">No menu **Ferramentas**, navegue at? **Nuget Package Manager** > **Console do Gerenciador de Pacotes**.</span><span class="sxs-lookup"><span data-stu-id="6b508-126">On the **Tools** menu, navigate to **Nuget Package Manager** > **Package Manager Console**.</span></span> 

   > 2. <span data-ttu-id="6b508-127">No console, execute o seguinte comando:</span><span class="sxs-lookup"><span data-stu-id="6b508-127">At the console, run the following command.</span></span> <span data-ttu-id="6b508-128">Pode levar um minuto ou mais para concluir, mesmo com uma conex?o r?pida ? Internet.</span><span class="sxs-lookup"><span data-stu-id="6b508-128">It may take a minute or more to complete even with a fast Internet connection.</span></span> <span data-ttu-id="6b508-129">Quando terminar, voc? deve ver **Microsoft.Identity.Client 1.1.1-alpha0393' instalado com sucesso...** perto do final da sa?da no console.</span><span class="sxs-lookup"><span data-stu-id="6b508-129">When it finishes you should see **Successfully installed 'Microsoft.Identity.Client 1.1.1-alpha0393' ...** near the end of the output in the console.</span></span>

   >    `Install-Package Microsoft.Identity.Client -Version 1.1.1-alpha0393 -Source https://www.myget.org/F/aad-clients-nightly/api/v3/index.json`

   > 3. <span data-ttu-id="6b508-p108">No **Explorador de solu??es**, clique com o bot?o direito em **Refer?ncias**. Confirme que o **Microsoft.Identity.Client** est? listado. Se n?o estiver, ou se houver um ?cone de aviso na entrada dele, exclua a entrada e use o assistente do Visual Studio Add Reference para adicionar uma refer?ncia ? montagem em **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.1-alpha0393\lib\net45\Microsoft.Identity.Client.dll**</span><span class="sxs-lookup"><span data-stu-id="6b508-p108">In **Solution Explorer**, right-click **References**. Verify that **Microsoft.Identity.Client** is listed. If it is not or there is a warning icon on its entry, delete the entry and then use the Visual Studio Add Reference wizard to add a reference to the assembly at **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.1-alpha0393\lib\net45\Microsoft.Identity.Client.dll**</span></span>

1. <span data-ttu-id="6b508-133">Crie o projeto pela segunda vez.</span><span class="sxs-lookup"><span data-stu-id="6b508-133">Build the project a second time.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="6b508-134">Registre o suplemento com o ponto de extremidade do Azure AD v2.0</span><span class="sxs-lookup"><span data-stu-id="6b508-134">Register the add-in with Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="6b508-135">As instru??es a seguir foram escritas de modo gen?rico para que possam ser usadas em diversos lugares.</span><span class="sxs-lookup"><span data-stu-id="6b508-135">The following instruction are written generically so they can be used in multiple places.</span></span> <span data-ttu-id="6b508-136">Para este artigo, fa?a o seguinte:</span><span class="sxs-lookup"><span data-stu-id="6b508-136">For this ariticle do the following:</span></span>
- <span data-ttu-id="6b508-137">Substitua o espa?o reservado **$ADD-IN-NAME$** por `Office-Add-in-ASPNET-SSO`.</span><span class="sxs-lookup"><span data-stu-id="6b508-137">Replace the placeholder **$ADD-IN-NAME$** with `Office-Add-in-ASPNET-SSO`.</span></span>
- <span data-ttu-id="6b508-138">Substitua o espa?o reservado **$FQDN-WITHOUT-PROTOCOL$** por `localhost:44355`.</span><span class="sxs-lookup"><span data-stu-id="6b508-138">Replace the placeholder **$FQDN-WITHOUT-PROTOCOL$** with `localhost:44355`.</span></span>
- <span data-ttu-id="6b508-139">Quando voc? especifica permiss?es no di?logo **Selecionar Permiss?es**, marque as caixas para as permiss?es a seguir.</span><span class="sxs-lookup"><span data-stu-id="6b508-139">When you specify permissions in the **Select Permissions** dialog, check the boxes for the following permissions.</span></span> <span data-ttu-id="6b508-140">Somente a primeira ? realmente exigida pelo suplemento propriamente dito, mas a biblioteca MSAL usada pelo c?digo de servidor exige `offline_access` e `openid`.</span><span class="sxs-lookup"><span data-stu-id="6b508-140">Only the first is really required by your add-in itself; but the MSAL library that the server-side code uses requires `offline_access` and `openid`.</span></span> <span data-ttu-id="6b508-141">A permiss?o `profile` ? necess?ria para que o host do Office obtenha um token no aplicativo Web do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="6b508-141">The `profile` permission is required for the Office host to get a token to your add-in web application.</span></span>
    * <span data-ttu-id="6b508-142">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="6b508-142">Files.Read.All</span></span>
    * <span data-ttu-id="6b508-143">offline_access</span><span class="sxs-lookup"><span data-stu-id="6b508-143">offline_access</span></span>
    * <span data-ttu-id="6b508-144">openid</span><span class="sxs-lookup"><span data-stu-id="6b508-144">openid</span></span>
    * <span data-ttu-id="6b508-145">perfil</span><span class="sxs-lookup"><span data-stu-id="6b508-145">profile</span></span>


[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="grant-administrator-consent-to-the-add-in"></a><span data-ttu-id="6b508-146">Conceder autoriza??o do administrador ao suplemento</span><span class="sxs-lookup"><span data-stu-id="6b508-146">Details are at: Grant administrator consent to the add-in</span></span>

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a><span data-ttu-id="6b508-147">Configurar o suplemento</span><span class="sxs-lookup"><span data-stu-id="6b508-147">Configure the add-in</span></span>

1. <span data-ttu-id="6b508-148">Na cadeia de caracteres a seguir, substitua o espa?o reservado "{tenant_ID}" pelo ID de locat?rio do Office 365.</span><span class="sxs-lookup"><span data-stu-id="6b508-148">In the following string, replace the placeholder ?{tenant_ID}? with your Office 365 tenant ID.</span></span> <span data-ttu-id="6b508-149">Use um dos m?todos em [Encontre seu ID de locat?rio do Office 365](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) para obt?-lo.</span><span class="sxs-lookup"><span data-stu-id="6b508-149">Use one of the methods in [Find your Office 365 tenant ID](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) to obtain it.</span></span>

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

2. <span data-ttu-id="6b508-150">No Visual Studio, abra o Web.config. Existem algumas chaves na se??o **appSettings** ?s quais voc? precisa atribuir valores.</span><span class="sxs-lookup"><span data-stu-id="6b508-150">In Visual Studio, open the web.config. There are some keys in the **appSettings** section to which you need to assign values.</span></span>

3. <span data-ttu-id="6b508-p112">Use a cadeia de caracteres constru?da na etapa 1 como o valor para a chave denominada "ida:Issuer". N?o deixe espa?os em branco no valor.</span><span class="sxs-lookup"><span data-stu-id="6b508-p112">Use the string you constructed in step 1 as the value to the key named ?ida:Issuer?. Be sure there are no blank spaces in the value.</span></span>

4. <span data-ttu-id="6b508-153">Atribua os seguintes valores para as chaves correspondentes:</span><span class="sxs-lookup"><span data-stu-id="6b508-153">Assign the following values to the corresponding keys:</span></span>

    |<span data-ttu-id="6b508-154">Chave</span><span class="sxs-lookup"><span data-stu-id="6b508-154">Key</span></span>|<span data-ttu-id="6b508-155">Valor</span><span class="sxs-lookup"><span data-stu-id="6b508-155">Value</span></span>|
    |:-----|:-----|
    |<span data-ttu-id="6b508-156">ida:ClientID</span><span class="sxs-lookup"><span data-stu-id="6b508-156">ida:ClientID</span></span>|<span data-ttu-id="6b508-157">A ID do aplicativo obtida ao registrar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="6b508-157">The application ID you obtained when you registered the add-in.</span></span>|
    |<span data-ttu-id="6b508-158">ida:Audience</span><span class="sxs-lookup"><span data-stu-id="6b508-158">ida:Audience</span></span>|<span data-ttu-id="6b508-159">A ID do aplicativo obtida ao registrar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="6b508-159">The application ID you obtained when you registered the add-in.</span></span>|
    |<span data-ttu-id="6b508-160">ida:Password</span><span class="sxs-lookup"><span data-stu-id="6b508-160">ida:Password</span></span>|<span data-ttu-id="6b508-161">A senha obtida ao registrar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="6b508-161">TThe password you obtained when you registered the add-in.</span></span>|

   <span data-ttu-id="6b508-p113">Veja a seguir um exemplo de como as quatro chaves que voc? alterou devem se parecer. *Observe que as chaves ClientID e Audience s?o iguais*. Voc? tamb?m pode usar uma ?nica chave para ambos os fins, mas sua marca??o web.config ? mais reutiliz?vel se for mantida separada, pois ela n?o ? sempre a mesma. Al?m disso, ter chaves separadas refor?a a ideia de que seu suplemento ? tanto um recurso de OAuth, em rela??o a um host do Office, e um cliente OAuth, em rela??o ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="6b508-p113">The following is an example of what the four keys you changed should look like. *Note that ClientID and Audience are the same*. You can also use a single key for both purposes, but your web.config markup is more reusable if you keep them separate because they aren't always the same. Also, having separate keys reinforces the idea that your add-in is both an OAuth resource, relative to the Office host, and an OAuth client, relative to Microsoft Graph.</span></span>

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />
    
    ```

   > [!NOTE]
   > <span data-ttu-id="6b508-166">N?o altere as demais configura??es na se??o **appSettings**.</span><span class="sxs-lookup"><span data-stu-id="6b508-166">Leave the other settings in the **appSettings** section unchanged.</span></span>

1. <span data-ttu-id="6b508-167">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6b508-167">Save and close the file.</span></span>

1. <span data-ttu-id="6b508-168">Na raiz do projeto, abra o arquivo do manifesto do suplemento "Office-Add-in-ASPNET-SSO.xml".</span><span class="sxs-lookup"><span data-stu-id="6b508-168">In the add-in project, open the add-in manifest file ?Office-Add-in-ASPNET-SSO.xml?.</span></span>

1. <span data-ttu-id="6b508-169">Role at? o final do arquivo.</span><span class="sxs-lookup"><span data-stu-id="6b508-169">Scroll to the bottom of the file.</span></span>

1. <span data-ttu-id="6b508-170">Logo acima da marca de fim `</VersionOverrides>`, voc? encontrar? a marca??o a seguir:</span><span class="sxs-lookup"><span data-stu-id="6b508-170">Just above the end `</VersionOverrides>` tag, you'll find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:44355/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="6b508-171">Substitua o espa?o reservado "{application_GUID here}" *nos dois lugares* na marca??o pela ID do Aplicativo que voc? copiou ao registrar seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="6b508-171">Replace the placeholder ?{application_GUID here}? *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="6b508-172">Os "{}" n?o fazem parte do ID, portanto n?o os inclua.</span><span class="sxs-lookup"><span data-stu-id="6b508-172">The "{}" are not part of the ID, so do not include them.</span></span> <span data-ttu-id="6b508-173">Essa ? a mesma ID usada para a ClientID e a Audience no web.config.</span><span class="sxs-lookup"><span data-stu-id="6b508-173">This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > * <span data-ttu-id="6b508-174">O valor de **Resource** ? o **URI da ID do Aplicativo** que voc? definiu quando adicionou a plataforma API Web no registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6b508-174">The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.</span></span>
    > * <span data-ttu-id="6b508-175">A se??o **Scopes** s? ser? usada para gerar uma caixa de di?logo de consentimento se o suplemento for vendido no AppSource.</span><span class="sxs-lookup"><span data-stu-id="6b508-175">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="6b508-176">Abra a guia **Avisos** da **Lista de Erros** no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="6b508-176">Open the **Warnings** tab of the **Error List** in Visual Studio.</span></span> <span data-ttu-id="6b508-177">Se houver um aviso informando que `<WebApplicationInfo>` n?o ? um filho v?lido de `<VersionOverrides>`, sua vers?o do Visual Studio 2017 Preview n?o reconhecer? a marca??o SSO.</span><span class="sxs-lookup"><span data-stu-id="6b508-177">If there is a warning that `<WebApplicationInfo>` is not a valid child of `<VersionOverrides>`, your version of Visual Studio 2017 Preview does not  recognize the SSO markup.</span></span> <span data-ttu-id="6b508-178">Para solucionar esse problema, fa?a o seguinte para um suplemento do Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="6b508-178">As a workaround, do the following for a Word, Excel, or PowerPoint add-in.</span></span> <span data-ttu-id="6b508-179">Se voc? estiver trabalhando com um suplemento do Outlook, confira a solu??o abaixo.</span><span class="sxs-lookup"><span data-stu-id="6b508-179">(If you are working with an Outlook add-in see the workaround below.)</span></span>

   - <span data-ttu-id="6b508-180">**Solu??o alternativa para Word, Excel e PowerPoint**</span><span class="sxs-lookup"><span data-stu-id="6b508-180">**Workaround for Word, Excel, and Powerpoint**</span></span>

        1. <span data-ttu-id="6b508-181">Comente a se??o `<WebApplicationInfo>` do manifesto logo acima do final de `</VersionOverrides>`.</span><span class="sxs-lookup"><span data-stu-id="6b508-181">Comment out the `<WebApplicationInfo>` section from the manifest just above the end of `</VersionOverrides>`.</span></span>

        2. <span data-ttu-id="6b508-p116">Pressione F5 para iniciar uma sess?o de depura??o. Isso criar? uma c?pia do manifesto na seguinte pasta (que pode ser acessada mais facilmente pelo **Gerenciador de Arquivos** do que pelo Visual Studio): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`</span><span class="sxs-lookup"><span data-stu-id="6b508-p116">Press F5 to start a debugging session. This will create a copy of the manifest in the following folder (which is easier to access in **File Explorer** than in Visual Studio): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`</span></span>

        3. <span data-ttu-id="6b508-184">Na c?pia do manifesto, remova a sintaxe do coment?rio em torno da se??o `<WebApplicationInfo>`.</span><span class="sxs-lookup"><span data-stu-id="6b508-184">In the copy of the manifest, remove the comment syntax around the `<WebApplicationInfo>` section.</span></span>

        4. <span data-ttu-id="6b508-185">Salve a c?pia do manifesto.</span><span class="sxs-lookup"><span data-stu-id="6b508-185">Save the copy of the manifest.</span></span>

        5. <span data-ttu-id="6b508-p117">Agora, ? preciso evitar que o Visual Studio substitua a c?pia do manifesto quando voc? terminar na pr?xima vez que pressionar F5. Clique com bot?o direito do mouse no n? da solu??o na parte superior do **Gerenciador de Solu??es** (n?o nos n?s do projeto).</span><span class="sxs-lookup"><span data-stu-id="6b508-p117">Now you must prevent Visual Studio from overwriting the copy of the manifest the next time you press F5. Right-click the solution node at the very top of **Solution Explorer** (not either of the project nodes).</span></span>

        6. <span data-ttu-id="6b508-188">Escolha **Propriedades** no menu de contexto e uma caixa de di?logo **P?ginas de Propriedades da Solu??o** ser? aberta.</span><span class="sxs-lookup"><span data-stu-id="6b508-188">Select **Properties** from the context menu and a **Solution Property Pages** dialog box opens.</span></span>

        7. <span data-ttu-id="6b508-189">Expanda **Propriedades da Configura??o** e escolha **Configura??o**.</span><span class="sxs-lookup"><span data-stu-id="6b508-189">Expand **Configuration Properties** and select **Configuration**.</span></span>

        8. <span data-ttu-id="6b508-190">Desmarque **Criar** e **Implantar** na linha do projeto **Office-Add-in-ASPNET-SSO** (*n?o* o projeto **Office-Add-in-ASPNET-SSO-WebAPI**).</span><span class="sxs-lookup"><span data-stu-id="6b508-190">Deselect **Build** and **Deploy** in the row for the **Office-Add-in-ASPNET-SSO** project (*not* the **Office-Add-in-ASPNET-SSO-WebAPI** project).</span></span>

        9. <span data-ttu-id="6b508-191">Pressione **OK** para fechar a caixa de di?logo.</span><span class="sxs-lookup"><span data-stu-id="6b508-191">Press **OK** to close the dialog box.</span></span>

   - <span data-ttu-id="6b508-192">**Solu??o alternativa para Outlook**</span><span class="sxs-lookup"><span data-stu-id="6b508-192">**Workaround for Outlook**</span></span>

        1. <span data-ttu-id="6b508-193">Em sua m?quina de desenvolvimento, localize o `MailAppVersionOverridesV1_1.xsd` existente.</span><span class="sxs-lookup"><span data-stu-id="6b508-193">On your development machine, locate the existing `MailAppVersionOverridesV1_1.xsd`.</span></span> <span data-ttu-id="6b508-194">Ele deve estar localizado no diret?rio de instala??o do Visual Studio em `./Xml/Schemas/{lcid}`.</span><span class="sxs-lookup"><span data-stu-id="6b508-194">This should be located in your Visual Studio installation directory under `./Xml/Schemas/{lcid}`.</span></span> <span data-ttu-id="6b508-195">Por exemplo, em uma instala??o t?pica do VS 2017 de 32 bits em um sistema em ingl?s (EUA), o caminho completo seria `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.</span><span class="sxs-lookup"><span data-stu-id="6b508-195">For example, on a typical installation of VS 2017 32-bit on an English (US) system, the full path would be `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.</span></span>

        2. <span data-ttu-id="6b508-196">Renomeie o arquivo existente para `MailAppVersionOverridesV1_1.old`.</span><span class="sxs-lookup"><span data-stu-id="6b508-196">Rename the existing file to `MailAppVersionOverridesV1_1.old`.</span></span>

        3. <span data-ttu-id="6b508-197">Copie essa vers?o modificada do arquivo para a pasta: [Esquema MailAppVersionOverrides modificado](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)</span><span class="sxs-lookup"><span data-stu-id="6b508-197">Copy this modified version of the file into the folder: [Modified MailAppVersionOverrides Schema](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)</span></span>

1. <span data-ttu-id="6b508-198">Salve e feche o arquivo de manifesto principal no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="6b508-198">Save and close the main manifest file in Visual Studio.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="6b508-199">Codificar o lado do cliente</span><span class="sxs-lookup"><span data-stu-id="6b508-199">Code the client side</span></span>

1. <span data-ttu-id="6b508-p119">Abra o arquivo Home.js da pasta **Scripts**. Ele j? apresenta alguns c?digos:</span><span class="sxs-lookup"><span data-stu-id="6b508-p119">Open the Home.js file in the **Scripts** folder. It already has some code in it:</span></span>
    * <span data-ttu-id="6b508-202">Uma atribui??o ao m?todo `Office.initialize` que, por sua vez, atribui um manipulador ao evento clicar do bot?o `getGraphAccessTokenButton`.</span><span class="sxs-lookup"><span data-stu-id="6b508-202">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="6b508-203">Um m?todo `showResult` que exibir? os dados retornados do Microsoft Graph (ou uma mensagem de erro) na parte inferior do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="6b508-203">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="6b508-204">Um m?todo `logErrors` que registrar? erros de console que n?o s?o destinados ao usu?rio final.</span><span class="sxs-lookup"><span data-stu-id="6b508-204">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>

1. <span data-ttu-id="6b508-p120">Abaixo da atribui??o a `Office.initialize`, adicione o c?digo a seguir. Observe o seguinte sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="6b508-p120">Below the assignment to `Office.initialize`, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="6b508-207">O processamento de erros no suplemento ?s vezes tentar? novamente obter um token de acesso automaticamente, usando um conjunto diferente de op??es.</span><span class="sxs-lookup"><span data-stu-id="6b508-207">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options.</span></span> <span data-ttu-id="6b508-208">A vari?vel de contador `timesGetOneDriveFilesHasRun` e a vari?veis de sinalizador `triedWithoutForceConsent` s?o usadas para garantir que o usu?rio n?o seja trocado repetidas vezes em tentativas falhas de obter um token.</span><span class="sxs-lookup"><span data-stu-id="6b508-208">The counter variable `timesGetOneDriveFilesHasRun`, and the flag variable `triedWithoutForceConsent` are used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span> 
    * <span data-ttu-id="6b508-p122">Voc? criar? um m?todo `getDataWithToken` na pr?xima etapa, mas observe que ele define uma op??o chamada `forceConsent` como `false`. Trataremos mais disso na etapa seguinte.</span><span class="sxs-lookup"><span data-stu-id="6b508-p122">You create the `getDataWithToken` method in the next step, but note that it sets an option called `forceConsent` to `false`. More about that in the next step.</span></span>

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }   
    ```

1. <span data-ttu-id="6b508-p123">Abaixo do m?todo `getOneDriveFiles`, adicione o c?digo a seguir. Observe o seguinte sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="6b508-p123">Below the `getOneDriveFiles` method, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="6b508-p124">O `getAccessTokenAsync` ? a nova API no Office.js que permite que um suplemento solicite ao aplicativo host do Office (Excel, PowerPoint, Word, etc.) um token de acesso para o suplemento (para o usu?rio conectado ao Office). O aplicativo host do Office, por sua vez, solicita o token ao ponto de extremidade 2.0 do Azure AD. Uma vez que voc? previamente autorizou o host do Office para o seu suplemento ao registr?-lo, o Azure AD enviar? o token.</span><span class="sxs-lookup"><span data-stu-id="6b508-p124">The `getAccessTokenAsync` is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office). The Office host application, in turn, asks the Azure AD 2.0 endpoint for the token. Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token.</span></span>
    * <span data-ttu-id="6b508-216">Se nenhum usu?rio estiver conectado ao Office, o host do Office solicitar? que o usu?rio se conecte.</span><span class="sxs-lookup"><span data-stu-id="6b508-216">If no user is signed into Office, the Office host will prompt the user to sign in.</span></span>
    * <span data-ttu-id="6b508-217">O par?metro de op??es configura o `forceConsent` como `false`. Dessa forma, n?o ser? solicitado que o usu?rio consinta o acesso ao host do Office ao seu suplemento sempre que ele o usar.</span><span class="sxs-lookup"><span data-stu-id="6b508-217">The options parameter sets `forceConsent` to `false`, so the user will not be prompted to consent to giving the Office host access to your add-in every time she or he uses the add-in.</span></span> <span data-ttu-id="6b508-218">Na primeira vez que o usu?rio tiver o suplemento, a chamada de `getAccessTokenAsync` falhar?, mas l?gica de processamento de erros que voc? adicionar? em uma etapa posterior ser? automaticamente chamada com a op??o `forceConsent` definida como `true` e o usu?rio ser? solicitado a consentir, mas somente essa primeira vez.</span><span class="sxs-lookup"><span data-stu-id="6b508-218">The first time the user runs the add-in, the call of `getAccessTokenAsync` will fail, but error-handling logic that you add in a later step will automatically re-call with the `forceConsent` option set to `true` and the user will be prompted to consent, but only that first time.</span></span>
    * <span data-ttu-id="6b508-219">Voc? criar? o m?todo `handleClientSideErrors` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="6b508-219">You will create the `handleClientSideErrors` method in a later step.</span></span>

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

1. <span data-ttu-id="6b508-p126">Substitua TODO1 pelas linhas a seguir. Voc? criar? o m?todo `getData` e a rota "/api/values" do lado do servidor nas etapas posteriores. Uma URL relativa ? usada para o ponto de extremidade porque ela deve ser hospedada no mesmo dom?nio que seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="6b508-p126">Replace the TODO1 with the following lines. You create the `getData` method and the server-side ?/api/values? route in later steps. A relative URL is used for the endpoint because it must be hosted on the same domain as your add-in.</span></span>

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. <span data-ttu-id="6b508-p127">Abaixo do m?todo `getOneDriveFiles`, adicione o seguinte. Observe isto sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="6b508-p127">Below the `getOneDriveFiles` method, add the following. About this code, note:</span></span>

    * <span data-ttu-id="6b508-p128">Este m?todo utilit?rio chama um ponto de extremidade da API Web especificado e transmite a ela o mesmo token de acesso que aplicativo host do Office usou para obter acesso ao seu suplemento. No lado do servidor, esse token de acesso ser? usado no fluxo "on behalf of" (em nome de) para obter um token de acesso para o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="6b508-p128">This method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the ?on behalf of? flow to obtain an access token to Microsoft Graph.</span></span>
    * <span data-ttu-id="6b508-227">Voc? criar? o m?todo `handleServerSideErrors` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="6b508-227">You will create the `handleServerSideErrors` method in a later step.</span></span>

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

### <a name="create-the-error-handling-methods"></a><span data-ttu-id="6b508-228">Crie os m?todos de processamento de erros</span><span class="sxs-lookup"><span data-stu-id="6b508-228">Create the error-handling methods</span></span>

1. <span data-ttu-id="6b508-229">Abaixo do m?todo `getData`, adicione o m?todo a seguir.</span><span class="sxs-lookup"><span data-stu-id="6b508-229">Below the `getData` method, add the following method.</span></span> <span data-ttu-id="6b508-230">Esse m?todo processar? os erros no cliente do suplemento quando o host do Office n?o puder obter um token de acesso para o servi?o Web do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6b508-230">This method will handle errors in the add-in's client when the Office host is unable to obtain an access token to the add-in's web service.</span></span> <span data-ttu-id="6b508-231">Esses erros s?o relatados com um c?digo de erro, portanto, o m?todo usa uma instru??o `switch` para distingui-los.</span><span class="sxs-lookup"><span data-stu-id="6b508-231">These errors are reported with an error code, so the method uses a `switch` statement to distinguish them.</span></span>

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

1. <span data-ttu-id="6b508-232">Substitua `TODO2` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="6b508-232">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="6b508-233">O erro 13001 ocorre quando o usu?rio n?o est? conectado ou quando ele cancela, sem responder, uma solicita??o para fornecer um segundo fator de autentica??o.</span><span class="sxs-lookup"><span data-stu-id="6b508-233">Error 13001 occurs when the user is not logged in, or the user cancelled, without responding, a prompt to provide a 2nd authentication factor.</span></span> <span data-ttu-id="6b508-234">Em ambos os casos, o c?digo executar? novamente o m?todo `getDataWithToken` e definir? uma op??o para for?ar uma solicita??o de entrada.</span><span class="sxs-lookup"><span data-stu-id="6b508-234">In either case, the code re-runs the `getDataWithToken` method and sets an option to force a sign-in prompt.</span></span>

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. <span data-ttu-id="6b508-235">Substitua `TODO3` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="6b508-235">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="6b508-236">O erro 13002 ocorre quando a entrada ou o consentimento do usu?rio ? anulado.</span><span class="sxs-lookup"><span data-stu-id="6b508-236">Error 13002 occurs when user's sign-in or consent was aborted.</span></span> <span data-ttu-id="6b508-237">Pe?a que o usu?rio tente novamente, mas n?o mais de uma vez.</span><span class="sxs-lookup"><span data-stu-id="6b508-237">Ask the user to try again but no more than once again.</span></span>

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }          
        break; 
    ```

1. <span data-ttu-id="6b508-238">Substitua `TODO4` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="6b508-238">Replace `TODO4` with the following code.</span></span> <span data-ttu-id="6b508-239">O erro 13003 ocorre quando o usu?rio est? conectado com uma conta que n?o ? corporativa, de estudante nem da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="6b508-239">Error 13003 occurs when user is logged in with an account that is neither work or school, nor Micrososoft Account.</span></span> <span data-ttu-id="6b508-240">Pe?a que o usu?rio saia e entre novamente com um tipo de conta suportado.</span><span class="sxs-lookup"><span data-stu-id="6b508-240">Ask the user to sign-out and then in again with a supported account type.</span></span>

    ```javascript
    case 13003: 
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;   
    ```

    > [!NOTE]
    > <span data-ttu-id="6b508-241">Os erros 13004 e 13005 n?o s?o processados neste m?todo, pois eles s? ocorrem em desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="6b508-241">Errors 13004 and 13005 are not handled in this method because they should only occur in development.</span></span> <span data-ttu-id="6b508-242">Eles n?o podem ser corrigidos pelo c?digo de tempo de execu??o e n?o seria ?til report?-lo a um usu?rio final.</span><span class="sxs-lookup"><span data-stu-id="6b508-242">They cannot be fixed by runtime code and there would be no point in reporting them to an end user.</span></span>

1. <span data-ttu-id="6b508-p134">Substitua `TODO5` pelo seguinte c?digo. O Erro 13006 ocorre quando houve um erro n?o especificado no host do Office, que pode indicar a instabilidade do host. Pe?a ao usu?rio para reiniciar o Office.</span><span class="sxs-lookup"><span data-stu-id="6b508-p134">Replace `TODO5` with the following code. Error 13006 occurs when there has been an unspecified error in the Office host that may indicate that the host is in an unstable state. Ask the user to restart Office.</span></span>

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;        
    ```

1. <span data-ttu-id="6b508-246">Substitua `TODO6` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="6b508-246">Replace `TODO6` with the following code.</span></span> <span data-ttu-id="6b508-247">O erro 13007 ocorre quando algo deu errado com a intera??o do host do Office com o AAD de forma que o host n?o pode obter um token de acesso para o servi?o Web/aplicativo dos suplementos.</span><span class="sxs-lookup"><span data-stu-id="6b508-247">Error 13007 occurs when something has gone wrong with the Office host's interaction with AAD so the host cannot get an access token to the add-ins web service/application.</span></span> <span data-ttu-id="6b508-248">? poss?vel que esse seja um problema de rede tempor?rio.</span><span class="sxs-lookup"><span data-stu-id="6b508-248">This may be a temporary network issue.</span></span> <span data-ttu-id="6b508-249">Pe?a que o usu?rio tente novamente mais tarde.</span><span class="sxs-lookup"><span data-stu-id="6b508-249">Ask the user to try again later.</span></span>

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;      
    ```

1. <span data-ttu-id="6b508-p136">Substitua `TODO7` pelo c?digo a seguir. O Erro 13008 ocorre quando o usu?rio aciona uma opera??o que chama `getAccessTokenAsync` antes que uma chamada anterior dele seja conclu?da.</span><span class="sxs-lookup"><span data-stu-id="6b508-p136">Replace `TODO7` with the following code. Error 13008 occurs when the user tiggered an operation that calls `getAccessTokenAsync` before a previous call of it completed.</span></span>

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```      

1. <span data-ttu-id="6b508-252">Substitua `TODO8` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="6b508-252">Replace `TODO8` with the following code.</span></span> <span data-ttu-id="6b508-253">O erro 13009 ocorre quando o suplemento n?o permite for?ar consentimento, mas `getAccessTokenAsync` foi chamado com a op??o `forceConsent` definida como `true`.</span><span class="sxs-lookup"><span data-stu-id="6b508-253">Error 13009 occurs when the add-in does not support forcing consent, but `getAccessTokenAsync` was called with the `forceConsent` option set to `true`.</span></span> <span data-ttu-id="6b508-254">Normalmente, quando isso acontece, o c?digo deve ser reexecutar `getAccessTokenAsync` automaticamente com a op??o de consentimento definida como `false`.</span><span class="sxs-lookup"><span data-stu-id="6b508-254">In the usual case when this happens the code should automatically re-run `getAccessTokenAsync` with the consent option set to `false`.</span></span> <span data-ttu-id="6b508-255">No entanto, em alguns casos, chamar o m?todo com `forceConsent` definido como `true` ? uma resposta autom?tica para um erro em uma chamada para o m?todo com a op??o definida como `false`.</span><span class="sxs-lookup"><span data-stu-id="6b508-255">However, in some cases, calling the method with `forceConsent` set to `true` was itself an automatic response to an error in a call to the method with the option set to `false`.</span></span> <span data-ttu-id="6b508-256">Nesse caso, o c?digo n?o deve tentar novamente, mas, em vez disso, ele deve solicitar que o usu?rio saia e entre novamente.</span><span class="sxs-lookup"><span data-stu-id="6b508-256">In that case, the code should not try again, but instead it should advise the user to sign out and sign in again.</span></span>

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```      
    
1. <span data-ttu-id="6b508-257">Substitua `TODO9` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="6b508-257">Replace `TODO9` with the following code.</span></span>

    ```javascript
    default:
        logError(result);
        break;
    ```  


1. <span data-ttu-id="6b508-p138">Abaixo do m?todo `handleClientSideErrors`, adicione o seguinte m?todo. Esse m?todo processar? os erros no servi?o Web do suplemento quando algo der errado na execu??o do fluxo on-behalf-of ou ao obter dados do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="6b508-p138">Below the `handleClientSideErrors` method, add the following method. This method will handle errors in the add-in's web service when something goes wrong in executing the on-behalf-of flow or in getting data from Microsoft Graph.</span></span>

    ```javascript
    function handleServerSideErrors(result) {
    
        // TODO10: Parse the JSON response.

        // TODO11: Handle the case where AAD asks for an additional form of authentication.

        // TODO12: Handle the case where consent has not been granted, or has been revoked.

        // TODO13: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow.

        // TODO14: Handle the case where the token that the add-in's client-side sends to it's 
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO15: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        // TODO16: Log all other server errors.
    }
    ```

1. <span data-ttu-id="6b508-260">Substitua `TODO10` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="6b508-260">Replace `TODO10` with the following code.</span></span> <span data-ttu-id="6b508-261">Observe que, para a maioria dos erros `4xx` que o servi?o Web do suplemento passar? para o suplemento do lado do cliente, haver? uma propriedade **ExceptionMessage** em resposta com o n?mero de erro AADSTS (Azure Active Directory Secure Token Service) al?m de outros dados.</span><span class="sxs-lookup"><span data-stu-id="6b508-261">Note that for most of the `4xx` errors that the add-in's web service will pass to the add-in's client-side, there will be an **ExceptionMessage** property in the response that contains the AADSTS (Azure Active Directory Secure Token Service) error number as well as other data.</span></span> <span data-ttu-id="6b508-262">No entanto, quando AAD envia uma mensagem para o servi?o Web do suplemento solicitando um fator de autentica??o adicional, a mensagem cont?m uma propriedade **Claims** especial que especifica (com um n?mero de c?digo) qual fator adicional ? necess?rio.</span><span class="sxs-lookup"><span data-stu-id="6b508-262">However, when AAD sends a message to the add-in's web service asking for an additonal authentication factor, the message contains a special **Claims** property that specifies (with a code number) what additional factor is needed.</span></span> <span data-ttu-id="6b508-263">As APIs ASP.NET que criam e enviam respostas HTTP para clientes n?o conhecem a propriedade **Claims**, portanto, elas n?o a incluem no objeto Response.</span><span class="sxs-lookup"><span data-stu-id="6b508-263">The ASP.NET APIs that create and send HTTP Responses to clients do not know about this **Claims** property, so they do not include it in the Response object.</span></span> <span data-ttu-id="6b508-264">O c?digo de servidor que ser? criado em uma etapa posterior lidar? com isso adicionando manualmente o valor **Claims** no objeto Response.</span><span class="sxs-lookup"><span data-stu-id="6b508-264">Server-side code that you will create in a later step will cope with this by manually adding the **Claims** value to the Response object.</span></span> <span data-ttu-id="6b508-265">Esse valor ser? salvo na propriedade **Message**, portanto, o c?digo tamb?m precisar? analisar essa propriedade.</span><span class="sxs-lookup"><span data-stu-id="6b508-265">This value will be in the **Message** property, so the code needs to parse out that property as well.</span></span>

    ```javascript
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    var message = JSON.parse(result.responseText).Message;
    ```

1. <span data-ttu-id="6b508-p140">Substitua `TODO11` pelo c?digo a seguir. Observa??o sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="6b508-p140">Replace `TODO11` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="6b508-268">O erro 50076 ocorre quando o Microsoft Graph requer uma forma adicional de autentica??o.</span><span class="sxs-lookup"><span data-stu-id="6b508-268">Error 50076 occurs when Microsoft Graph requires an additional form of authentication.</span></span>
    * <span data-ttu-id="6b508-269">O host do Office deve obter um novo token com o valor **Claims** como a op??o `authChallenge`.</span><span class="sxs-lookup"><span data-stu-id="6b508-269">The Office host should get a new token with the **Claims** value as the `authChallenge` option.</span></span> <span data-ttu-id="6b508-270">Isso instrui o AAD a solicitar ao usu?rio todas as formas de autentica??o requeridas.</span><span class="sxs-lookup"><span data-stu-id="6b508-270">This tells AAD to prompt the user for all required forms of authentication.</span></span> 

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
        }
    }    
    ```

1. <span data-ttu-id="6b508-p142">Substitua `TODO12` pelo c?digo a seguir. Observa??o sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="6b508-p142">Replace `TODO12` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="6b508-273">O erro 65001 significa que o consentimento para acessar o Microsoft Graph n?o foi concedido (ou foi revogado) para uma ou mais permiss?es.</span><span class="sxs-lookup"><span data-stu-id="6b508-273">Error 65001 means that consent to access Microsoft Graph was not granted (or was revoked) for one or more permissions.</span></span> 
    * <span data-ttu-id="6b508-274">O suplemento dever? obter um novo token com a op??o `forceConsent` definida como `true`.</span><span class="sxs-lookup"><span data-stu-id="6b508-274">The add-in should get a new token with the `forceConsent` option set to `true`.</span></span>

    ```javascript
    if (exceptionMessage.indexOf('AADSTS65001') !== -1) {
        showResult(['Please grant consent to this add-in to access your Microsoft Graph data.']);        
        /*
            THE FORCE CONSENT OPTION IS NOT AVAILABLE IN DURING PREVIEW. WHEN SSO FOR
            OFFICE ADD-INS IS RELEASED, REMOVE THE showResult LINE ABOVE AND UNCOMMENT
            THE FOLLOWING LINE.
        */
       // getDataWithToken({ forceConsent: true });
    }    
    ```

1. <span data-ttu-id="6b508-p143">Substitua `TODO13` pelo c?digo a seguir. Observa??o sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="6b508-p143">Replace `TODO13` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="6b508-p144">O Erro 70011 tem muitos significados. O que importa para este suplemento ? quando ele significa que um escopo inv?lido (permiss?o) foi solicitado, ent?o o c?digo verifica a descri??o completa do erro, n?o apenas o n?mero.</span><span class="sxs-lookup"><span data-stu-id="6b508-p144">Error 70011 has multiple meanings. The one that matters to this add-in is when it means that an invalid scope (permission) has been requested, so the code checks for the full error description, not just the number.</span></span>
    * <span data-ttu-id="6b508-279">O suplemento dever? relatar o erro.</span><span class="sxs-lookup"><span data-stu-id="6b508-279">The add-in should report the error.</span></span>

    ```javascript
     else if (exceptionMessage.indexOf("AADSTS70011: The provided value for the input parameter 'scope' is not valid.") !== -1) {
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }    
    ```

1. <span data-ttu-id="6b508-p145">Substitua `TODO14` pelo c?digo a seguir. Observa??o sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="6b508-p145">Replace `TODO14` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="6b508-282">C?digo de servidor criado em uma etapa posterior enviar? a mensagem `Missing access_as_user` se o escopo `access_as_user` (permiss?o) n?o for o token de acesso que o cliente do suplemento enviar para o ADD para ser usado no fluxo on-behalf-of.</span><span class="sxs-lookup"><span data-stu-id="6b508-282">Server-side code that you create in a later step will send the message `Missing access_as_user` if the `access_as_user` scope (permission) is not in the access token that the add-in's client sends to AAD to be used in the on-behalf-of flow.</span></span>
    * <span data-ttu-id="6b508-283">O suplemento dever? relatar o erro.</span><span class="sxs-lookup"><span data-stu-id="6b508-283">The add-in should report the error.</span></span>

    ```javascript
    else if (exceptionMessage.indexOf('Missing access_as_user.') !== -1) {
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }    
    ```

1. <span data-ttu-id="6b508-p146">Substitua `TODO15` pelo c?digo a seguir. Observa??o sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="6b508-p146">Replace `TODO15` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="6b508-p147">A biblioteca de identidade que voc? usar? no c?digo do lado do servidor (Biblioteca de Autentica??o da Microsoft - MSAL) deve garantir que nenhum token inv?lido ou expirado seja enviado para o Microsoft Graph. Contudo, se isso ocorrer, o erro retornado para servi?o Web do suplemento do Microsoft Graph ter? o c?digo `InvalidAuthenticationToken`. O c?digo do lado do servidor que voc? criar? em uma etapa futura transmitir? essa mensagem ao cliente do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6b508-p147">The identity library that you will be using in the server-side code (Microsoft Authentication Library - MSAL) should ensure that no expired or invalid token is sent to Microsoft Graph; but if it does happen, the error that is returned to the add-in's web service from Microsoft Graph has the code `InvalidAuthenticationToken`. Server-side code you will create in a latter step will relay this message to the add-in's client.</span></span>
    * <span data-ttu-id="6b508-288">Nesse caso, o suplemento dever? iniciar o processo de autentica??o completo ao redefinir o contador e as vari?veis de sinalizador e, em seguida, chamando novamente o m?todo de identificador de bot?o.</span><span class="sxs-lookup"><span data-stu-id="6b508-288">In this case, the add-in should start the entire authentication process over by resetting the counter and flag varibles, and then re-calling the button handler method.</span></span>

    ```javascript
    // If the token sent to MS Graph is expired or invalid, start the whole process over.
    else if (result.code === 'InvalidAuthenticationToken') {
        timesGetOneDriveFilesHasRun = 0;
        triedWithoutForceConsent = false;
        getOneDriveFiles();
    }    
    ```

1. <span data-ttu-id="6b508-289">Substitua `TODO16` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="6b508-289">Replace `TODO16` with the following code.</span></span>

    ```javascript
    else {
        logError(result);
    }    
    ```

1. <span data-ttu-id="6b508-290">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6b508-290">Save and close the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="6b508-291">Codifique o lado do servidor</span><span class="sxs-lookup"><span data-stu-id="6b508-291">Code the server side</span></span>

### <a name="configure-the-owin-middleware"></a><span data-ttu-id="6b508-292">Configurar o middleware OWIN</span><span class="sxs-lookup"><span data-stu-id="6b508-292">Configure the OWIN middleware</span></span>

1. <span data-ttu-id="6b508-293">Abra o arquivo Startup.cs na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="6b508-293">Open the Startup.cs file in the root of the project.</span></span>

1. <span data-ttu-id="6b508-p148">Adicione a palavra-chave `partial` para a declara??o da classe Startup, se ainda n?o estiver l?. A linha dever? ser assim:</span><span class="sxs-lookup"><span data-stu-id="6b508-p148">Add the keyword `partial` to the declaration of the Startup class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="6b508-p149">Adicione a linha a seguir ao corpo do m?todo `Configuration`. Voc? criar? o m?todo `ConfigureAuth` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="6b508-p149">Add the following line to the body of the `Configuration` method. You create the `ConfigureAuth` method in a later step.</span></span>

    `ConfigureAuth(app);`

1. <span data-ttu-id="6b508-298">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6b508-298">Save and close the file.</span></span>

1. <span data-ttu-id="6b508-299">Clique com bot?o direito do mouse na pasta **App_Start** e selecione **Adicionar > Classe**.</span><span class="sxs-lookup"><span data-stu-id="6b508-299">Right-click the **App_Start** folder and select **Add > Class**.</span></span>

1. <span data-ttu-id="6b508-300">Na caixa de di?logo **Adicionar novo item** nomeie o arquivo **Startup.Auth.cs** e, em seguida, clique em **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="6b508-300">In the **Add new item** dialog name the file **Startup.Auth.cs** and then click **Add**.</span></span>

1. <span data-ttu-id="6b508-301">Encurte o nome do namespace no novo arquivo para `Office_Add_in_ASPNET_SSO_WebAPI`.</span><span class="sxs-lookup"><span data-stu-id="6b508-301">Shorten the namespace name in the new file to `Office_Add_in_ASPNET_SSO_WebAPI`.</span></span>

1. <span data-ttu-id="6b508-302">Verifique se todas as seguintes instru??es `using` est?o na parte superior do arquivo.</span><span class="sxs-lookup"><span data-stu-id="6b508-302">Ensure that all of the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Owin;
    using System.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. <span data-ttu-id="6b508-p150">Adicione a palavra-chave `partial` ? declara??o da classe `Startup`, se ainda n?o estiver l?. A linha dever? ser assim:</span><span class="sxs-lookup"><span data-stu-id="6b508-p150">Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="6b508-p151">Adicione o m?todo a seguir ? classe `Startup`. Este m?todo especifica como o middleware OWIN validar? os tokens de acesso que s?o transmitidos a ele do m?todo `getData` no arquivo Home.js do lado do cliente. O processo de autoriza??o ? disparado sempre que um ponto de extremidade da API Web decorado com o atributo `[Authorize]` ? chamado.</span><span class="sxs-lookup"><span data-stu-id="6b508-p151">Add the following method to the `Startup` class. This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file. The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.</span></span>

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. <span data-ttu-id="6b508-308">Substitua TODO3 pelo seguinte c?digo.</span><span class="sxs-lookup"><span data-stu-id="6b508-308">Replace the TODO3 with the following.</span></span> <span data-ttu-id="6b508-309">Observa??o sobre o c?digo:</span><span class="sxs-lookup"><span data-stu-id="6b508-309">Note about this code:</span></span>

    * <span data-ttu-id="6b508-310">O c?digo instrui o OWIN a garantir que o emissor de token e audi?ncia especificado no token de acesso que vem do host do Office (e ? transmitido pela chamada de `getData` do lado do cliente) deve coincidir com os valores especificados no Web.config.</span><span class="sxs-lookup"><span data-stu-id="6b508-310">The code instructs OWIN to ensure that the audience and token issuer specified in the access token that comes from the Office host (and is passed on by the client-side call of `getData`) must match the values specified in the web.config.</span></span>
    * <span data-ttu-id="6b508-p153">Definir `SaveSigninToken` como `true` faz com que o OWIN salve o token bruto do host do Office. O suplemento precisa dele para obter um token de acesso para o Microsoft Graph com o fluxo "on behalf of".</span><span class="sxs-lookup"><span data-stu-id="6b508-p153">Setting `SaveSigninToken` to `true` causes OWIN to save the raw token from the Office host. The add-in needs it to obtain an access token to Microsoft Graph with the ?on behalf of? flow.</span></span>
    * <span data-ttu-id="6b508-p154">Os escopos n?o s?o validados pelo middleware OWIN. Os escopos do token de acesso, que devem conter `access_as_user`, s?o validados no controlador.</span><span class="sxs-lookup"><span data-stu-id="6b508-p154">Scopes are not validated by the OWIN middleware. The scopes of the access token, which should include `access_as_user`, is validated in the controller.</span></span>

    ```csharp
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. <span data-ttu-id="6b508-p155">Substitua TODO4 pelo seguinte. Observa??o sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="6b508-p155">Replace TODO4 with the following. Note about this code:</span></span>

    * <span data-ttu-id="6b508-317">O m?todo `UseOAuthBearerAuthentication` ? chamado em vez do `UseWindowsAzureActiveDirectoryBearerAuthentication` que ? mais comum, porque este ?ltimo n?o ? compat?vel com o ponto de extremidade V2 do Azure AD.</span><span class="sxs-lookup"><span data-stu-id="6b508-317">The method `UseOAuthBearerAuthentication` is called instead of the more common `UseWindowsAzureActiveDirectoryBearerAuthentication` because the latter is not compatible with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="6b508-318">A URL de descoberta transmitida ao m?todo ? onde o middleware OWIN obt?m instru??es para conseguir a chave que precisa para verificar a assinatura no token de acesso recebido do host do Office.</span><span class="sxs-lookup"><span data-stu-id="6b508-318">The discovery URL that is passed to the method is where the OWIN middleware obtains instructions for getting the key it needs to verify the signature on the access token received from the Office host.</span></span>

    ```csharp
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
        {
            AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
        });
    ```

1. <span data-ttu-id="6b508-319">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6b508-319">Save and close the file.</span></span>

### <a name="create-the-apivalues-controller"></a><span data-ttu-id="6b508-320">Criar o controlador /api/values</span><span class="sxs-lookup"><span data-stu-id="6b508-320">Create the /api/values controller</span></span>

1. <span data-ttu-id="6b508-321">Abra o arquivo **Controllers\ValueController.cs**.</span><span class="sxs-lookup"><span data-stu-id="6b508-321">Open the file **Controllers\ValueController.cs**.</span></span>

2. <span data-ttu-id="6b508-322">Verifique se as seguintes instru??es `using` est?o na parte superior do arquivo.</span><span class="sxs-lookup"><span data-stu-id="6b508-322">Ensure that the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Microsoft.Identity.Client;
    using System.IdentityModel.Tokens;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    using Office_Add_in_ASPNET_SSO_WebAPI.Models;
    ```

3. <span data-ttu-id="6b508-p156">Logo acima da linha que declara o `ValuesController`, adicione o atributo `[Authorize]`. Isso garante que seu suplemento executar? o processo de autoriza??o configurado no ?ltimo procedimento sempre que um m?todo controlador for chamado. Apenas os chamadores com um token de acesso v?lido para o seu suplemento podem invocar os m?todos do controlador.</span><span class="sxs-lookup"><span data-stu-id="6b508-p156">Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6b508-326">Um servi?o da ASP.NET MVC Web API de produ??o deve ter l?gica personalizada para o fluxo on-behalf-of em uma ou mais classes [FilterAttribute](https://msdn.microsoft.com/en-us/library/system.web.http.filters(v=vs.108).aspx) personalizadas.</span><span class="sxs-lookup"><span data-stu-id="6b508-326">A production ASP.NET MVC Web API service should have custom logic for the on-behalf-of flow in one or more custom [FilterAttribute](https://msdn.microsoft.com/en-us/library/system.web.http.filters(v=vs.108).aspx) classes.</span></span> <span data-ttu-id="6b508-327">Este exemplo educacional coloca a l?gica no controlador de principal para que o fluxo de autoriza??o e dados busca l?gica inteiro possa ser acompanhado facilmente.</span><span class="sxs-lookup"><span data-stu-id="6b508-327">This educational sample puts the logic in the main controller so that the entire flow of the authorization and data fetching logic can be easily followed.</span></span> <span data-ttu-id="6b508-328">Isso tamb?m faz com que o exemplo fique consistente com os exemplos de padr?o de autoriza??o nos [Exemplos do Azure](https://github.com/Azure-Samples/).</span><span class="sxs-lookup"><span data-stu-id="6b508-328">This also makes the sample consistent with the pattern of authorization samples in [Azure Samples](https://github.com/Azure-Samples/).</span></span>    

4. <span data-ttu-id="6b508-329">Adicione o m?todo a seguir ao `ValuesController`.</span><span class="sxs-lookup"><span data-stu-id="6b508-329">Add the following method to the `ValuesController`.</span></span> <span data-ttu-id="6b508-330">Observe que ? o valor de retorno ? `Task<HttpResponseMessage>` em vez de `Task<IEnumerable<string>>`, como seria mais comum para um m?todo `GET api/values`.</span><span class="sxs-lookup"><span data-stu-id="6b508-330">Note that the return value is `Task<HttpResponseMessage>` instead of `Task<IEnumerable<string>>` as would be more common for a `GET api/values` method.</span></span> <span data-ttu-id="6b508-331">Este ? um efeito colateral do fato de que nossa l?gica de autoriza??o personalizada estar? no controlador: algumas condi??es de erro nessa l?gica exigem que um objeto de resposta HTTP seja enviado para o cliente do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6b508-331">This is a side effect of that fact that our custom authorization logic will be in the controller: some error conditions in that logic require that an HTTP Response object be sent to the add-in's client.</span></span> 

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO1: Validate the scopes of the access token.
    }
    ```

5. <span data-ttu-id="6b508-332">Substitua `TODO1` pelo seguinte c?digo para validar que os escopos especificados no token incluam `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="6b508-332">Replace `TODO1` with the following code to validate that the scopes that are specified in the token include `access_as_user`.</span></span>

    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (addinScopes.Contains("access_as_user"))
    {
        // TODO2: Assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.
        // TODO3: Get the access token for Microsoft Graph.
        // TODO4: Get the names of files and folders in OneDrive by using the Microsoft Graph API.
        // TODO5: Remove excess information from the data and send the data to the client.
    }
    return SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    ```

    > [!NOTE]
    > <span data-ttu-id="6b508-p159">Voc? deve usar apenas o escopo `access_as_user` para autorizar a API que lida com o fluxo Em Nome De para os suplementos do Office. Outras APIs em seu servi?o devem ter seus pr?prios requisitos de escopo. Isso limita o que pode ser acessado com os tokens que o Office adquire.</span><span class="sxs-lookup"><span data-stu-id="6b508-p159">You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office add-ins. Other APIs in your service should have their own scope requirements. This limits what can be accessed with the tokens that Office acquires.</span></span>

6. <span data-ttu-id="6b508-p160">Substitua `TODO2` pelo c?digo a seguir. Observa??o sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="6b508-p160">Replace `TODO2` with the following code. Note about this code:</span></span>
    * <span data-ttu-id="6b508-337">Ele transforma o token de acesso bruto recebido do host do Office em um objeto de `UserAssertion` que ser? transmitido para outro m?todo.</span><span class="sxs-lookup"><span data-stu-id="6b508-337">It turns the raw access token received from the Office host into a `UserAssertion` object that will be passed to another method.</span></span>
    * <span data-ttu-id="6b508-p161">Seu suplemento n?o est? mais desempenhando o papel de um recurso (ou p?blico) para o qual o host do Office e o usu?rio precisam de acesso. Agora, ele mesmo ? um cliente que precisa de acesso ao Microsoft Graph. `ConfidentialClientApplication` ? o objeto "client context" da MSAL.</span><span class="sxs-lookup"><span data-stu-id="6b508-p161">Your add-in is no longer playing the role of a resource (or audience) to which the Office host and user need access. Now it is itself a client that needs access to Microsoft Graph. `ConfidentialClientApplication` is the MSAL ?client context? object.</span></span>
    * <span data-ttu-id="6b508-p162">O terceiro par?metro para o construtor `ConfidentialClientApplication` ? uma URL de redirecionamento que n?o ? realmente usada no fluxo "on behalf of", mas usar a URL correta ? uma boa pr?tica. O quarto e o quinto par?metros podem ser usados para definir um armazenamento persistente que permitiria a reutiliza??o de tokens n?o expirados em diferentes sess?es com o suplemento. Este exemplo n?o implementa nenhum armazenamento persistente.</span><span class="sxs-lookup"><span data-stu-id="6b508-p162">The third parameter to the `ConfidentialClientApplication` constructor is a redirect URL which is not actually used in the ?on behalf of? flow, but it is a good practice to use the correct URL. The fourth and fifth parameters can be used to define a persistent store that would enable the reuse of unexpired tokens across different sessions with the add-in. This sample does not implement any persistent storage.</span></span>
    * <span data-ttu-id="6b508-344">A MSAL exige os escopos `openid` e `offline_access` para funcionar, mas ela lan?a um erro se o c?digo solicit?-los de forma redundante.</span><span class="sxs-lookup"><span data-stu-id="6b508-344">MSAL requires the `openid` and `offline_access` scopes to function, but it throws an error if your code redundantly requests them.</span></span> <span data-ttu-id="6b508-345">Ela tamb?m lan?ar? um erro se o seu c?digo solicitar o `profile`, que realmente ? usado apenas quando o aplicativo host do Office recebe o token para o aplicativo Web do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="6b508-345">It will also throw an error if your code requests `profile`, which is really only used when the Office host application gets the token to your add-in's web application.</span></span> <span data-ttu-id="6b508-346">Ent?o, apenas `Files.Read.All` ? explicitamente solicitado.</span><span class="sxs-lookup"><span data-stu-id="6b508-346">So only `Files.Read.All` is explicitly requested.</span></span>

    ```csharp
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

7. <span data-ttu-id="6b508-p164">Substitua `TODO3` pelo c?digo a seguir. Observa??o sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="6b508-p164">Replace `TODO3` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="6b508-p165">O m?todo `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` procurar? primeiro no cache da MSAL, que est? na mem?ria, para fazer a correspond?ncia com o token de acesso. Somente se n?o houver um, ele iniciar? o fluxo "on behalf of" com o ponto de extremidade V2 do Azure AD.</span><span class="sxs-lookup"><span data-stu-id="6b508-p165">The `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` method will first look in the MSAL cache, which is in memory, for a matching access token. Only if there isn't one, does it initiate the "on behalf of" flow with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="6b508-351">Se a autentica??o multi-fator for requerida pelo recurso MS Graph e o usu?rio ainda n?o a tiver fornecido, o AAD lan?ar? uma exce??o contendo uma propriedade de Declara??es.</span><span class="sxs-lookup"><span data-stu-id="6b508-351">If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will throw an exception containing a Claims property.</span></span>
    * <span data-ttu-id="6b508-p166">O valor da propriedade de Declara??es deve ser passado para o cliente, que o passar? para o host do Office, que, em seguida, o incluir? em um pedido para um novo token. O AAD solicitar? ao usu?rio todas as formas de autentica??o necess?rias.</span><span class="sxs-lookup"><span data-stu-id="6b508-p166">The Claims property value must be passed to the client which will pass it to the Office host, which will then include it in a request for a new token. AAD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="6b508-354">Quaisquer exce??es que n?o forem do tipo `MsalServiceException` s?o intencionalmente n?o detectadas, e, portanto, se propagar?o para o cliente como mensagens `500 Server Error`.</span><span class="sxs-lookup"><span data-stu-id="6b508-354">Any exceptions that are not of type `MsalServiceException` are intentionally not caught, so they will propagate to the client as `500 Server Error` messages.</span></span>

    ```csharp
    AuthenticationResult result = null;
    try
    {
        result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
    }
    catch (MsalServiceException e)
    {        
        // TODO3a: Handle request for multi-factor authentication.
        // TODO3b: Handle lack of consent.
        // TODO3c: Handle invalid scope (permission).
        // TODO3d: Handle all other MsalServiceExceptions.
    }
    ```

8. <span data-ttu-id="6b508-p167">Substitua `TODO3a` pelo c?digo a seguir. Observa??o sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="6b508-p167">Replace `TODO3a` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="6b508-p168">Se a autentica??o multifator for exigida pelo recurso MS Graph e o usu?rio ainda n?o a tiver fornecido, o AAD retornar? "400 Bad Request" com o erro AADSTS50076 e uma propriedade **Declara??es**. O MSAL lan?ar? uma **MsalUiRequiredException** (que herda de **MsalServiceException**) com essas informa??es.</span><span class="sxs-lookup"><span data-stu-id="6b508-p168">If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will return "400 Bad Request" with error AADSTS50076 and a **Claims** property. MSAL throws a **MsalUiRequiredException** (which inherits from **MsalServiceException**) with this information.</span></span> 
    * <span data-ttu-id="6b508-p169">O valor da propriedade **Declara??es** deve ser passado para o cliente, que deve pass?-lo para o host do Office, que, por sua vez, o incluir? em um pedido para um novo token. O AAD solicitar? ao usu?rio todas as formas de autentica??o necess?rias.</span><span class="sxs-lookup"><span data-stu-id="6b508-p169">The **Claims** property value must be passed to the client which should pass it to the Office host, which then includes it in a request for a new token. AAD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="6b508-361">As APIs que criam respostas HTTP a partir de exce??es n?o conhecem a propriedade **Claims**, portanto, elas n?o a incluem no objeto de resposta.</span><span class="sxs-lookup"><span data-stu-id="6b508-361">The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object.</span></span> <span data-ttu-id="6b508-362">? necess?rio criar manualmente uma mensagem que inclua esse recurso.</span><span class="sxs-lookup"><span data-stu-id="6b508-362">We have to manually create a message that includes it.</span></span> <span data-ttu-id="6b508-363">Uma propriedade **Message** personalizada, no entanto, impede a cria??o de uma propriedade **ExceptionMessage**, portanto, a ?nica maneira de obter a ID de erro `AADSTS50076` para o cliente ? adicion?-la ? **Message** personalizada.</span><span class="sxs-lookup"><span data-stu-id="6b508-363">A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**.</span></span> <span data-ttu-id="6b508-364">O JavaScript no cliente precisar? descobrir se uma resposta tem uma **Message** ou **ExceptionMessage** para saber qual ler.</span><span class="sxs-lookup"><span data-stu-id="6b508-364">JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.</span></span>
    * <span data-ttu-id="6b508-365">A mensagem personalizada ? formatada como JSON para que o JavaScript do cliente possa analis?-la com m?todos de objeto `JSON` conhecidos.</span><span class="sxs-lookup"><span data-stu-id="6b508-365">The custom message is formatted as JSON so that the client-side JavaScript can parse it with well-known `JSON` object methods.</span></span>
    * <span data-ttu-id="6b508-366">Voc? criar? o m?todo `SendErrorToClient` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="6b508-366">You will create the `SendErrorToClient` method in a later step.</span></span> <span data-ttu-id="6b508-367">? segundo par?metro ? um objeto **Exception**.</span><span class="sxs-lookup"><span data-stu-id="6b508-367">It's second parameter is an **Exception** object.</span></span> <span data-ttu-id="6b508-368">Nesse caso, o c?digo passa `null` porque incluir o objeto **Exception** bloqueia a inclus?o da propriedade **Message** na resposta HTTP que ? gerada.</span><span class="sxs-lookup"><span data-stu-id="6b508-368">In this case, the code passes `null` because including the **Exception** object blocks the inclusion of the **Message** property in the HTTP Response that is generated.</span></span>

    ```csharp
    if (e.Message.StartsWith("AADSTS50076")) {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

9. <span data-ttu-id="6b508-p172">Substitua `TODO3b` e `TODO3c` pelo c?digo a seguir. Observa??o sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="6b508-p172">Replace `TODO3b` and `TODO3c` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="6b508-371">Se a chamada para o AAD contiver pelo menos um escopo (permiss?o) que n?o tenha sido consentido pelo usu?rio ou por um administrador de locat?rios (ou se o consentimento foi revogado),</span><span class="sxs-lookup"><span data-stu-id="6b508-371">If the call to AAD contained at least one scope (permission) for which neither the user nor a tenant administrator has consented (or consent was revoked).</span></span> <span data-ttu-id="6b508-372">o AAD retornar? "400 Solicita??o Incorreta" com o erro `AADSTS65001`.</span><span class="sxs-lookup"><span data-stu-id="6b508-372">AAD will return "400 Bad Request" with error `AADSTS65001`.</span></span> <span data-ttu-id="6b508-373">O MSAL exibe um **MsalUiRequiredException** com essas informa??es.</span><span class="sxs-lookup"><span data-stu-id="6b508-373">MSAL throws a **MsalUiRequiredException** with this information.</span></span> <span data-ttu-id="6b508-374">O cliente deve chamar `getAccessTokenAsync` novamente com a op??o `{ forceConsent: true }`.</span><span class="sxs-lookup"><span data-stu-id="6b508-374">The client should re-call `getAccessTokenAsync` with the option `{ forceConsent: true }`.</span></span>
    *  <span data-ttu-id="6b508-375">Se a chamada para o AAD contiver pelo menos um escopo que AAD n?o reconhece, o AAD retornar? "400 Solicita??o Incorreta" com o erro `AADSTS70011`.</span><span class="sxs-lookup"><span data-stu-id="6b508-375">If the call to AAD contained at least one scope that AAD does not recognize, AAD returns "400 Bad Request" with error `AADSTS70011`.</span></span> <span data-ttu-id="6b508-376">O MSAL exibe um **MsalUiRequiredException** com essas informa??es.</span><span class="sxs-lookup"><span data-stu-id="6b508-376">MSAL throws a **MsalUiRequiredException** with this information.</span></span> <span data-ttu-id="6b508-377">O cliente deve informar o usu?rio.</span><span class="sxs-lookup"><span data-stu-id="6b508-377">The client should inform the user.</span></span>
    *  <span data-ttu-id="6b508-378">A descri??o completa ? inclu?da porque 70011 ? retornado em outras condi??es e ele deve ser processado nesse suplemento somente quando significar que h? um escopo inv?lido.</span><span class="sxs-lookup"><span data-stu-id="6b508-378">The entire description is included beause 70011 is returned in other conditions and we it should only be handled in this add-in when it means that there is an invalid scope.</span></span> 
    *  <span data-ttu-id="6b508-p175">O objeto **MsalUiRequiredException** ? passado para `SendErrorToClient`. Isso garante que uma propriedade **ExceptionMessage** contendo as informa??es de erro seja inclu?da na resposta HTTP.</span><span class="sxs-lookup"><span data-stu-id="6b508-p175">The **MsalUiRequiredException** object is passed to `SendErrorToClient`. This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.</span></span>
    *  <span data-ttu-id="6b508-381">N?o h? uma mensagem personalizada, portanto, `null` ? passado para o terceiro par?metro.</span><span class="sxs-lookup"><span data-stu-id="6b508-381">There is no custom message, so `null` is passed for the third parameter.</span></span>

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001"))
    || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

10. <span data-ttu-id="6b508-382">Substitua `TODO3d` pelo c?digo a seguir.</span><span class="sxs-lookup"><span data-stu-id="6b508-382">Replace `TODO3d` with the following code.</span></span> <span data-ttu-id="6b508-383">Observe que o c?digo exibe a exce??o em vez de transmiti-la em uma resposta HTTP personalizada com **HttpStatusCode.Forbidden** (401).</span><span class="sxs-lookup"><span data-stu-id="6b508-383">Note that the code rethrows the exception instead of relaying it in a custom HTTP Response with **HttpStatusCode.Forbidden** (401).</span></span> <span data-ttu-id="6b508-384">O efeito disso ? que o ASP.NET enviar? sua pr?pria resposta HTTP com o status "500 Erro de Servidor".</span><span class="sxs-lookup"><span data-stu-id="6b508-384">The effect of this is that the ASP.NET will send its own HTTP Response with status "500 Server Error".</span></span>

    ```csharp
    else
    {
        throw e;
    }  
    ```

11. <span data-ttu-id="6b508-p177">Substitua `TODO4` pelo seguinte. Observa??o sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="6b508-p177">Replace `TODO4` with the following. Note about this code:</span></span>

    * <span data-ttu-id="6b508-p178">As classes `GraphApiHelper` e `ODataHelper` s?o definidas nos arquivos da pasta **Helpers**. A classe `OneDriveItem` ? definida em um arquivo da pasta **Models**. A discuss?o detalhada dessas classes n?o ? relevante para a autoriza??o ou o SSO, portanto, est? fora do escopo deste artigo.</span><span class="sxs-lookup"><span data-stu-id="6b508-p178">The `GraphApiHelper` and `ODataHelper` classes are defined in files in the **Helpers** folder. The `OneDriveItem` class is defined in a file in the **Models** folder. Detailed discussion of these classes is not relevant to authorization or SSO, so it is out-of-scope for this article.</span></span>
    * <span data-ttu-id="6b508-390">O desempenho ? aprimorado ao se solicitar ao Microsoft Graph apenas os dados que s?o realmente necess?rios. Desse modo, o c?digo usa um par?metro de consulta ` $select` para especificar que desejamos somente a propriedade de nome, e usa um par?metro `$top` para especificar que desejamos somente os tr?s primeiros nomes de pasta ou de arquivo.</span><span class="sxs-lookup"><span data-stu-id="6b508-390">Performance is improved by asking Microsoft Graph for only the data actually needed, so the code uses a ` $select` query parameter to specify that we only want the name property, and a `$top` parameter to specify that we want only the first three folder or file names.</span></span>
    * <span data-ttu-id="6b508-391">Se o token enviado para o Microsoft Graph for inv?lido, o Microsoft Graph enviar? um erro "401 N?o Autorizado" com o c?digo "InvalidAuthenticationToken".</span><span class="sxs-lookup"><span data-stu-id="6b508-391">If the token sent to Microsoft Graph is invalid, Microsoft Graph sends a "401 Unauthorized" error with the code "InvalidAuthenticationToken".</span></span> <span data-ttu-id="6b508-392">Em seguida, o ASP.NET exibe um **RuntimeBinderException**.</span><span class="sxs-lookup"><span data-stu-id="6b508-392">ASP.NET then throws a **RuntimeBinderException**.</span></span> <span data-ttu-id="6b508-393">Isso tamb?m ocorre quando o token expira, embora o MSAL deva impedir que isso aconte?a.</span><span class="sxs-lookup"><span data-stu-id="6b508-393">This is also what happens when the token is expired, although MSAL should prevent that from ever happening.</span></span> 

    ```csharp
    var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
    IEnumerable<OneDriveItem> filesResult;
    try
    {
        filesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
    }
    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
    {
        return SendErrorToClient(HttpStatusCode.Unauthorized, e, null);                    
    }
    ```

12. <span data-ttu-id="6b508-p180">Substitua `TODO5` pelo seguinte. Observa??o sobre este c?digo:</span><span class="sxs-lookup"><span data-stu-id="6b508-p180">Replace `TODO5` with the following. Note about this code:</span></span> 

    * <span data-ttu-id="6b508-p181">Embora o c?digo acima solicite somente a propriedade *name* dos itens do OneDrive, o Microsoft Graph sempre inclui a propriedade *eTag* para os itens do OneDrive. Para reduzir a carga enviada para o cliente, o c?digo a seguir reconstr?i os resultados apenas com os nomes dos itens.</span><span class="sxs-lookup"><span data-stu-id="6b508-p181">Although the code above asked for only the *name* property of the OneDrive items, Microsoft Graph always includes the *eTag* property for OneDrive items. To reduce the payload sent to the client, the code below reconstructs the results with only the item names.</span></span>
    * <span data-ttu-id="6b508-398">A lista de tr?s pastas e arquivos do OneDrive ? enviada para o cliente como uma resposta HTTP "200 OK".</span><span class="sxs-lookup"><span data-stu-id="6b508-398">The list of three OneDrive files and folders is sent to the client as a "200 OK" HTTP Response.</span></span>

    ```csharp
    List<string> itemNames = new List<string>();
    foreach (OneDriveItem item in filesResult)
    {
        itemNames.Add(item.Name);
    }

    var requestMessage = new HttpRequestMessage();
    requestMessage.SetConfiguration(new HttpConfiguration());
    var response = requestMessage.CreateResponse<List<string>>(HttpStatusCode.OK, itemNames); 
    return response;
    ```

13. <span data-ttu-id="6b508-399">Abaixo do m?todo Get, adicione o m?todo a seguir.</span><span class="sxs-lookup"><span data-stu-id="6b508-399">Below the Get method, add the following method.</span></span> <span data-ttu-id="6b508-400">Sobre este c?digo, observe:</span><span class="sxs-lookup"><span data-stu-id="6b508-400">About this code note:</span></span>  

    * <span data-ttu-id="6b508-401">O m?todo transmite ao cliente informa??es sobre uma exce??o do servidor.</span><span class="sxs-lookup"><span data-stu-id="6b508-401">The method relays to the client information about a server-side exception.</span></span> 
    * <span data-ttu-id="6b508-402">Se a exce??o original for passada para o m?todo, o construtor HttpError incluir? informa??es do objeto de exce??o em uma propriedade **ExceptionMessage**.</span><span class="sxs-lookup"><span data-stu-id="6b508-402">If the original exception is passed to the method, then the HttpError constuctor will include information from the exception object in an **ExceptionMessage** property.</span></span>  
    * <span data-ttu-id="6b508-403">Se `null` for passado para a exce??o, o construtor HttpError incluir? o par?metro de mensagem em uma propriedade **Message** e n?o haver? uma propriedade **ExceptionMessage**.</span><span class="sxs-lookup"><span data-stu-id="6b508-403">If `null` is passed for the exception, then the HttpError constuctor will include the message parameter in a **Message** property and there is no **ExceptionMessage** property.</span></span>

    ```csharp
    private HttpResponseMessage SendErrorToClient(HttpStatusCode statusCode, Exception e, string message)
    {
        HttpError error;
        if (e != null)
        {
            error = new HttpError(e, true);
        }
        else
        {
            error = new HttpError(message);
        }
        var requestMessage = new HttpRequestMessage();
        var errorMessage = requestMessage.CreateErrorResponse(statusCode, error);
        return errorMessage;
    }        
    ```

## <a name="run-the-add-in"></a><span data-ttu-id="6b508-404">Execute o suplemento</span><span class="sxs-lookup"><span data-stu-id="6b508-404">Run the add-in</span></span>

1. <span data-ttu-id="6b508-405">Certifique-se de ter alguns arquivos no seu OneDrive para que voc? possa verificar os resultados.</span><span class="sxs-lookup"><span data-stu-id="6b508-405">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

1. <span data-ttu-id="6b508-p183">No Visual Studio, pressione F5. O PowerPoint ser? aberto e haver? um grupo **SSO ASP.NET** na faixa de op??es **P?gina Inicial**.</span><span class="sxs-lookup"><span data-stu-id="6b508-p183">In Visual Studio, press F5. PowerPoint opens and there is an **SSO ASP.NET** group on the **Home** ribbon.</span></span>

1. <span data-ttu-id="6b508-408">Pressione o bot?o **Mostrar Suplemento** nesse grupo para ver a interface do usu?rio do suplemento no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="6b508-408">Press the **Show Add-in** button in this group to see the add-in?s UI in the task pane.</span></span>

1. <span data-ttu-id="6b508-p184">Pressione o bot?o **Obter meus arquivos do OneDrive**. Se voc? n?o estiver conectado ao Office, voc? ser? solicitado a entrar.</span><span class="sxs-lookup"><span data-stu-id="6b508-p184">Press the button **Get My Files from OneDrive**. If you are not signed into Office, you'll be prompted to sign in.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="6b508-411">Se voc? entrou no Office com uma ID diferente e se alguns aplicativos do Office que estavam abertos no momento continuam abertos, o Office pode n?o alterar de forma confi?vel sua ID, mesmo que pare?a ter feito isso no PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="6b508-411">If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint.</span></span> <span data-ttu-id="6b508-412">Se isso acontecer, a chamada para o Microsoft Graph pode falhar ou os dados da ID anterior podem ser retornados.</span><span class="sxs-lookup"><span data-stu-id="6b508-412">If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned.</span></span> <span data-ttu-id="6b508-413">Para evitar isso, certifique-se de *fechar todos os outros aplicativos do Office* antes de pressionar **Obter meus arquivos do OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="6b508-413">To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.</span></span>

1. <span data-ttu-id="6b508-p186">Depois de entrar, ser? exibida uma lista de seus arquivos e suas pastas no OneDrive, abaixo do bot?o. Esse procedimento pode levar mais de 15 segundos, principalmente na primeira vez.</span><span class="sxs-lookup"><span data-stu-id="6b508-p186">After you are signed in, a list of your files and folders on OneDrive will appear below the button. This may take over 15 seconds, especially the first time.</span></span>
