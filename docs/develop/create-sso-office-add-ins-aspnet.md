---
title: Criar um Suplemento do Office com ASP.NET que use logon único
description: ''
ms.date: 04/15/2019
localization_priority: Priority
ms.openlocfilehash: bc8c2427171f06865de6c809a5d7311018fcc278
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695802"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="9819e-102">Criar um Suplemento do Office com ASP.NET que use logon único (visualização)</span><span class="sxs-lookup"><span data-stu-id="9819e-102">Create an ASP.NET Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="9819e-p101">Quando os usuários estão conectados ao Office, o seu suplemento pode usar as mesmas credenciais para permitir que os usuários acessem vários aplicativos sem exigir que eles entrem uma segunda vez. Para obter uma visão geral, consulte [Habilitar o SSO em um Suplemento do Office](sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="9819e-p101">When users are signed in to Office, your add-in can use the same credentials to permit users to access multiple applications without requiring them to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="9819e-105">Este artigo apresenta o processo passo a passo de habilitação do logon único (SSO) em um suplemento que foi criado com ASP.NET, OWIN e com a Biblioteca de Autenticação da Microsoft (MSAL) para .NET.</span><span class="sxs-lookup"><span data-stu-id="9819e-105">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with ASP.NET, OWIN, and Microsoft Authentication Library (MSAL) for .NET.</span></span>

> [!NOTE]
> <span data-ttu-id="9819e-106">Para ler um artigo semelhante sobre um suplemento baseado em Node.js, confira [Criar um Suplemento do Office com Node.js que use logon único](create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="9819e-106">For a similar article about a Node.js-based add-in, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="9819e-107">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="9819e-107">Prerequisites</span></span>

* <span data-ttu-id="9819e-108">A versão mais recente disponível do Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="9819e-108">The latest available version of Visual Studio 2017.</span></span>

* <span data-ttu-id="9819e-109">Office 365 (a versão de assinatura do Office).</span><span class="sxs-lookup"><span data-stu-id="9819e-109">Office 365 (the subscription version of Office).</span></span> <span data-ttu-id="9819e-110">Build e versão mensal mais recentes do canal de Participante do programa Office Insider.</span><span class="sxs-lookup"><span data-stu-id="9819e-110">Latest monthly version and build from the Insiders channel.</span></span> <span data-ttu-id="9819e-111">É necessário ingressar no programa Office Insider para obter essa versão.</span><span class="sxs-lookup"><span data-stu-id="9819e-111">You need to be an Office Insider to get this version.</span></span> <span data-ttu-id="9819e-112">Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="9819e-112">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="9819e-113">Observe que, quando um build é promovido ao Canal Semestral de produção, o suporte para recursos de visualização, como o SSO, é desativado para esse build.</span><span class="sxs-lookup"><span data-stu-id="9819e-113">Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="9819e-114">Configure o projeto inicial</span><span class="sxs-lookup"><span data-stu-id="9819e-114">Set up the starter project</span></span>

1. <span data-ttu-id="9819e-115">Clone ou baixe o repositório em [SSO com Suplemento ASPNET do Office](https://github.com/officedev/office-add-in-aspnet-sso).</span><span class="sxs-lookup"><span data-stu-id="9819e-115">Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span></span>

1. <span data-ttu-id="9819e-p103">Abra a pasta **Before** (antes) e abra o arquivo .sln no Visual Studio. Esse é um projeto inicial. A interface do usuário e outros aspectos do suplemento que não estão diretamente ligados ao SSO ou à autorização já estão prontos.</span><span class="sxs-lookup"><span data-stu-id="9819e-p103">Open the **Before** folder and open the .sln file in Visual Studio. This is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done.</span></span>

    > [!NOTE]
    > <span data-ttu-id="9819e-p104">Há também uma versão concluída do exemplo no mesmo repositório. Essa versão apresenta como seria o suplemento quando concluídos os procedimentos apresentados neste artigo, com exceção de que o projeto concluído traz comentários de códigos que seriam redundantes neste artigo. Para usar a versão concluída, apenas abra o arquivo `sln` e siga as instruções apresentadas neste artigo, mas pule as seções **Codificar o lado do cliente** e **Codificar o lado do servidor**.</span><span class="sxs-lookup"><span data-stu-id="9819e-p104">There is also a completed version of the sample in the same repo. It is just like the add-in that you would have if you completed the procedures in this article, except that the completed project has code comments that would be redundant with the text of this article. To use the completed version, just open the `sln` file and follow the instructions in this article, but skip the sections **Code the client side** and **Code the server** side.</span></span>

1. <span data-ttu-id="9819e-p105">Depois que o projeto for aberto, compile-o no Visual Studio, que instalará os pacotes listados no arquivo packages.config. Esse procedimento poderá levar entre alguns segundos e alguns minutos dependendo de quantos pacotes estiverem no cache de pacote local do computador.</span><span class="sxs-lookup"><span data-stu-id="9819e-p105">After the project opens, build it in Visual Studio, which will install the packages listed in the packages.config file. This can take a few seconds to several minutes depending on how many of the packages are in the computer's local package cache.</span></span>

    > [!NOTE]
    > <span data-ttu-id="9819e-p106">Você receberá um erro sobre o namespace Identity. Este é um efeito colateral de um problema de configuração que você corrigirá no próximo passo. O importante é que os pacotes estejam instalados.</span><span class="sxs-lookup"><span data-stu-id="9819e-p106">You will get an error about the Identity namespace. This is a side effect of a configuration issue that you will fix with the next step. The important thing is that the packages are installed.</span></span>

1. <span data-ttu-id="9819e-127">Atualmente, a versão da biblioteca MSAL (Microsoft.Identity.Client) necessária para SSO (versão `1.1.4-preview0002`) não faz parte do catálogo padrão de nuget, portanto, não está listada no package.config e deve ser instalada separadamente.</span><span class="sxs-lookup"><span data-stu-id="9819e-127">Currently, the version of the MSAL library (Microsoft.Identity.Client) that you need for SSO (version `1.1.4-preview0002`) is not part of the standard nuget catalog, so it is not listed in the package.config, and it must be installed separately.</span></span>

   > 1. <span data-ttu-id="9819e-128">No menu **Ferramentas**, navegue até **Nuget Package Manager** > **Console do Gerenciador de Pacotes**.</span><span class="sxs-lookup"><span data-stu-id="9819e-128">On the **Tools** menu, navigate to **Nuget Package Manager** > **Package Manager Console**.</span></span>
   > 2. <span data-ttu-id="9819e-129">No console, execute o seguinte comando:</span><span class="sxs-lookup"><span data-stu-id="9819e-129">At the console, run the following command.</span></span> <span data-ttu-id="9819e-130">Pode levar um minuto ou mais para concluir, mesmo com uma conexão rápida à Internet.</span><span class="sxs-lookup"><span data-stu-id="9819e-130">It may take a minute or more to complete even with a fast Internet connection.</span></span> <span data-ttu-id="9819e-131">Quando terminar, você deve ver **'Microsoft.Identity.Client 1.1.4-preview0002' instalado com sucesso...** perto do final da saída no console.</span><span class="sxs-lookup"><span data-stu-id="9819e-131">When it finishes you should see **Successfully installed 'Microsoft.Identity.Client 1.1.4-preview0002' ...** near the end of the output in the console.</span></span>
   >    `Install-Package Microsoft.Identity.Client -Version 1.1.4-preview0002`
   > 3. <span data-ttu-id="9819e-132">No **Explorador de soluções**, expanda **Referências** do projeto **Office-Add-in-ASPNET-SSO-WebAPI**.</span><span class="sxs-lookup"><span data-stu-id="9819e-132">In **Solution Explorer**, expand **References** of **Office-Add-in-ASPNET-SSO-WebAPI** project.</span></span> <span data-ttu-id="9819e-133">Verifique se **Microsoft.Identity.Client** está na lista.</span><span class="sxs-lookup"><span data-stu-id="9819e-133">Verify that **Microsoft.Identity.Client** is listed.</span></span> <span data-ttu-id="9819e-134">Se não estiver ou se houver um ícone de aviso na entrada, exclua a entrada e use o Assistente Visual Studio Add Reference para adicionar uma referência à montagem em **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.4-preview0002\lib\net45\Microsoft.Identity.Client.dll**</span><span class="sxs-lookup"><span data-stu-id="9819e-134">If it is not or there is a warning icon on its entry, delete the entry and then use the Visual Studio Add Reference Wizard to add a reference to the assembly at **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.4-preview0002\lib\net45\Microsoft.Identity.Client.dll**</span></span>

1. <span data-ttu-id="9819e-135">Crie o projeto pela segunda vez.</span><span class="sxs-lookup"><span data-stu-id="9819e-135">Build the project a second time.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="9819e-136">Registre o suplemento com o ponto de extremidade v2.0 do Azure AD</span><span class="sxs-lookup"><span data-stu-id="9819e-136">Register the add-in with Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="9819e-137">As instruções a seguir são escritas de forma geral, elas podem ser usadas em vários locais.</span><span class="sxs-lookup"><span data-stu-id="9819e-137">The following instruction are written generically so they can be used in multiple places.</span></span> <span data-ttu-id="9819e-138">Para este artigo faça o seguinte:</span><span class="sxs-lookup"><span data-stu-id="9819e-138">For this article do the following:</span></span>

- <span data-ttu-id="9819e-139">Substitua o espaço reservado **$ADD-IN-NAME$** por `Office-Add-in-ASPNET-SSO`.</span><span class="sxs-lookup"><span data-stu-id="9819e-139">Replace the placeholder **$ADD-IN-NAME$** with `Office-Add-in-ASPNET-SSO`.</span></span>
- <span data-ttu-id="9819e-140">Substitua o espaço reservado **$FQDN-WITHOUT-PROTOCOL$** por `localhost:44355`.</span><span class="sxs-lookup"><span data-stu-id="9819e-140">Replace the placeholder **$FQDN-WITHOUT-PROTOCOL$** with `localhost:44355`.</span></span>
- <span data-ttu-id="9819e-141">Quando você especificar permissões na caixa de diálogo **Selecionar permissões**, marque as caixas das seguintes permissões.</span><span class="sxs-lookup"><span data-stu-id="9819e-141">When you specify permissions in the **Select Permissions** dialog, check the boxes for the following permissions.</span></span> <span data-ttu-id="9819e-142">Somente a primeira é realmente exigida pelo suplemento propriamente dito, mas a biblioteca MSAL usada pelo código de servidor exige `offline_access` e `openid`.</span><span class="sxs-lookup"><span data-stu-id="9819e-142">Only the first is really required by your add-in itself; but the MSAL library that the server-side code uses requires `offline_access` and `openid`.</span></span> <span data-ttu-id="9819e-143">A permissão `profile` é necessária para que o host do Office obtenha um token no aplicativo Web do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="9819e-143">The `profile` permission is required for the Office host to get a token to your add-in web application.</span></span>
  * <span data-ttu-id="9819e-144">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="9819e-144">Files.Read.All</span></span>
  * <span data-ttu-id="9819e-145">offline_access</span><span class="sxs-lookup"><span data-stu-id="9819e-145">offline_access</span></span>
  * <span data-ttu-id="9819e-146">openid</span><span class="sxs-lookup"><span data-stu-id="9819e-146">openid</span></span>
  * <span data-ttu-id="9819e-147">profile</span><span class="sxs-lookup"><span data-stu-id="9819e-147">profile</span></span>


[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="grant-administrator-consent-to-the-add-in"></a><span data-ttu-id="9819e-148">Conceder consentimento do administrador ao suplemento</span><span class="sxs-lookup"><span data-stu-id="9819e-148">Grant administrator consent to the add-in</span></span>

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a><span data-ttu-id="9819e-149">Configurar o suplemento</span><span class="sxs-lookup"><span data-stu-id="9819e-149">Configure the add-in</span></span>

1. <span data-ttu-id="9819e-150">No seguinte cadeia, substitua o espaço reservado "{tenant_ID}" pela sua ID de locatário do Office 365.</span><span class="sxs-lookup"><span data-stu-id="9819e-150">In the following string, replace the placeholder “{tenant_ID}” with your Office 365 tenancy ID.</span></span> <span data-ttu-id="9819e-151">Se você não copiou a ID de locatário quando você registrou o suplemento com AAD, use um dos métodos em [Encontrar sua ID de locatário do Office 365](/onedrive/find-your-office-365-tenant-id) para obtê-la.</span><span class="sxs-lookup"><span data-stu-id="9819e-151">If you didn't copy the tenancy ID when you registered the add-in with AAD, use one of the methods in [Find your Office 365 tenant ID](/onedrive/find-your-office-365-tenant-id) to obtain it.</span></span>

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

1. <span data-ttu-id="9819e-152">No Visual Studio, abra o Web.config. Existem algumas chaves na seção **appSettings** às quais você precisa atribuir valores.</span><span class="sxs-lookup"><span data-stu-id="9819e-152">In Visual Studio, open the web.config. There are some keys in the **appSettings** section to which you need to assign values.</span></span>

1. <span data-ttu-id="9819e-p112">Use a cadeia de caracteres construída na etapa 1 como o valor para a chave denominada "ida:Issuer". Não deixe espaços em branco no valor.</span><span class="sxs-lookup"><span data-stu-id="9819e-p112">Use the string you constructed in step 1 as the value to the key named “ida:Issuer”. Be sure there are no blank spaces in the value.</span></span>

1. <span data-ttu-id="9819e-155">Atribua os seguintes valores para as chaves correspondentes:</span><span class="sxs-lookup"><span data-stu-id="9819e-155">Assign the following values to the corresponding keys:</span></span>

    |<span data-ttu-id="9819e-156">Chave</span><span class="sxs-lookup"><span data-stu-id="9819e-156">Key</span></span>|<span data-ttu-id="9819e-157">Valor</span><span class="sxs-lookup"><span data-stu-id="9819e-157">Value</span></span>|
    |:-----|:-----|
    |<span data-ttu-id="9819e-158">ida:ClientID</span><span class="sxs-lookup"><span data-stu-id="9819e-158">ida:ClientID</span></span>|<span data-ttu-id="9819e-159">A ID do aplicativo obtida ao registrar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="9819e-159">The application ID you obtained when you registered the add-in.</span></span>|
    |<span data-ttu-id="9819e-160">ida:Audience</span><span class="sxs-lookup"><span data-stu-id="9819e-160">ida:Audience</span></span>|<span data-ttu-id="9819e-161">A ID do aplicativo obtida ao registrar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="9819e-161">The application ID you obtained when you registered the add-in.</span></span>|
    |<span data-ttu-id="9819e-162">ida:Password</span><span class="sxs-lookup"><span data-stu-id="9819e-162">ida:Password</span></span>|<span data-ttu-id="9819e-163">A senha obtida ao registrar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="9819e-163">The password you obtained when you registered the add-in.</span></span>|

   <span data-ttu-id="9819e-p113">Veja a seguir um exemplo de como as quatro chaves que você alterou devem se parecer. *Observe que as chaves ClientID e Audience são iguais*. Você também pode usar uma única chave para ambos os fins, mas sua marcação web.config é mais reutilizável se for mantida separada, pois ela não é sempre a mesma. Além disso, ter chaves separadas reforça a ideia de que seu suplemento é tanto um recurso de OAuth, em relação a um host do Office, e um cliente OAuth, em relação ao Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="9819e-p113">The following is an example of what the four keys you changed should look like. *Note that ClientID and Audience are the same*. You can also use a single key for both purposes, but your web.config markup is more reusable if you keep them separate because they aren't always the same. Also, having separate keys reinforces the idea that your add-in is both an OAuth resource, relative to the Office host, and an OAuth client, relative to Microsoft Graph.</span></span>

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />

    ```

   > [!NOTE]
   > <span data-ttu-id="9819e-168">Não altere as demais configurações na seção **appSettings**.</span><span class="sxs-lookup"><span data-stu-id="9819e-168">Leave the other settings in the **appSettings** section unchanged.</span></span>

1. <span data-ttu-id="9819e-169">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="9819e-169">Save and close the file.</span></span>

1. <span data-ttu-id="9819e-170">Na raiz do projeto, abra o arquivo do manifesto do suplemento "Office-Add-in-ASPNET-SSO.xml".</span><span class="sxs-lookup"><span data-stu-id="9819e-170">In the add-in project, open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml”.</span></span>

1. <span data-ttu-id="9819e-171">Role até o final do arquivo.</span><span class="sxs-lookup"><span data-stu-id="9819e-171">Scroll to the bottom of the file.</span></span>

1. <span data-ttu-id="9819e-172">Logo acima da marca de fim `</VersionOverrides>`, você encontrará a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="9819e-172">Just above the end `</VersionOverrides>` tag, you'll find the following markup:</span></span>

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

1. <span data-ttu-id="9819e-p114">Substitua o espaço reservado "{application_GUID aqui}" *nos dois lugares*, na marcação, pela ID do Aplicativo que você copiou ao registrar seu suplemento. Os "{}" não fazem parte da ID, portanto, não os inclua. Essa é a mesma ID usada para ClientID e Audience no web.config.</span><span class="sxs-lookup"><span data-stu-id="9819e-p114">Replace the placeholder “{application_GUID here}” *in both places* in the markup with the Application ID that you copied when you registered your add-in. The "{}" are not part of the ID, so do not include them. This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > * <span data-ttu-id="9819e-176">O valor de **Resource** é o **URI da ID do Aplicativo** que você definiu quando adicionou a plataforma API Web no registro do suplemento.</span><span class="sxs-lookup"><span data-stu-id="9819e-176">The **Resource** value is the **Application ID URI** you set when you added the Web API platform to the registration of the add-in.</span></span>
    > * <span data-ttu-id="9819e-177">A seção **Scopes** só será usada para gerar uma caixa de diálogo de consentimento se o suplemento for vendido no AppSource.</span><span class="sxs-lookup"><span data-stu-id="9819e-177">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="9819e-178">Abra a guia **Avisos** da **Lista de Erros** no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="9819e-178">Open the **Warnings** tab of the **Error List** in Visual Studio.</span></span> <span data-ttu-id="9819e-179">Se houver um aviso que `<WebApplicationInfo>` não é um filho válido de `<VersionOverrides>`, sua versão do Visual Studio 2017 Preview não reconhecerá a marcação SSO.</span><span class="sxs-lookup"><span data-stu-id="9819e-179">If there is a warning that `<WebApplicationInfo>` is not a valid child of `<VersionOverrides>`, your version of Visual Studio 2017 Preview does not recognize the SSO markup.</span></span> <span data-ttu-id="9819e-180">Para solucionar esse problema, faça o seguinte para um suplemento do Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="9819e-180">As a workaround, do the following for a Word, Excel, or PowerPoint add-in.</span></span> <span data-ttu-id="9819e-181">Se você estiver trabalhando com um suplemento do Outlook, confira a solução abaixo.</span><span class="sxs-lookup"><span data-stu-id="9819e-181">(If you are working with an Outlook add-in see the workaround below.)</span></span>

   - <span data-ttu-id="9819e-182">**Solução alternativa para Word, Excel e PowerPoint**</span><span class="sxs-lookup"><span data-stu-id="9819e-182">**Workaround for Word, Excel, and PowerPoint**</span></span>

        1. <span data-ttu-id="9819e-183">Comente a seção `<WebApplicationInfo>` do manifesto logo acima do final de `</VersionOverrides>`.</span><span class="sxs-lookup"><span data-stu-id="9819e-183">Comment out the `<WebApplicationInfo>` section from the manifest just above the end of `</VersionOverrides>`.</span></span>

        2. <span data-ttu-id="9819e-p116">Pressione **F5** para iniciar uma sessão de depuração. Isso criará uma cópia do manifesto na seguinte pasta (que pode ser acessada mais facilmente pelo **Gerenciador de Arquivos** do que pelo Visual Studio): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`</span><span class="sxs-lookup"><span data-stu-id="9819e-p116">Press **F5** to start a debugging session. This will create a copy of the manifest in the following folder (which is easier to access in **File Explorer** than in Visual Studio): `Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`</span></span>

        3. <span data-ttu-id="9819e-186">Na cópia do manifesto, remova a sintaxe do comentário em torno da seção `<WebApplicationInfo>`.</span><span class="sxs-lookup"><span data-stu-id="9819e-186">In the copy of the manifest, remove the comment syntax around the `<WebApplicationInfo>` section.</span></span>

        4. <span data-ttu-id="9819e-187">Salve a cópia do manifesto.</span><span class="sxs-lookup"><span data-stu-id="9819e-187">Save the copy of the manifest.</span></span>

        5. <span data-ttu-id="9819e-p117">Agora, é preciso evitar que o Visual Studio substitua a cópia do manifesto quando você terminar na próxima vez que pressionar F5. Clique com botão direito do mouse no nó da solução na parte superior do **Gerenciador de Soluções** (não nos nós do projeto).</span><span class="sxs-lookup"><span data-stu-id="9819e-p117">Now you must prevent Visual Studio from overwriting the copy of the manifest the next time you press F5. Right-click the solution node at the very top of **Solution Explorer** (not either of the project nodes).</span></span>

        6. <span data-ttu-id="9819e-190">Escolha **Propriedades** no menu de contexto e uma caixa de diálogo **Páginas de Propriedades da Solução** será aberta.</span><span class="sxs-lookup"><span data-stu-id="9819e-190">Select **Properties** from the context menu and a **Solution Property Pages** dialog box opens.</span></span>

        7. <span data-ttu-id="9819e-191">Expanda **Propriedades da Configuração** e escolha **Configuração**.</span><span class="sxs-lookup"><span data-stu-id="9819e-191">Expand **Configuration Properties** and select **Configuration**.</span></span>

        8. <span data-ttu-id="9819e-192">Desmarque **Criar** e **Implantar** na linha do projeto **Office-Add-in-ASPNET-SSO** (*não* o projeto **Office-Add-in-ASPNET-SSO-WebAPI**).</span><span class="sxs-lookup"><span data-stu-id="9819e-192">Deselect **Build** and **Deploy** in the row for the **Office-Add-in-ASPNET-SSO** project (*not* the **Office-Add-in-ASPNET-SSO-WebAPI** project).</span></span>

        9. <span data-ttu-id="9819e-193">Pressione **OK** para fechar a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="9819e-193">Press **OK** to close the dialog box.</span></span>

   - <span data-ttu-id="9819e-194">**Solução alternativa para Outlook**</span><span class="sxs-lookup"><span data-stu-id="9819e-194">**Workaround for Outlook**</span></span>

        1. <span data-ttu-id="9819e-p118">Em sua máquina de desenvolvimento, localize o `MailAppVersionOverridesV1_1.xsd` existente. Ele deve estar localizado no diretório de instalação do Visual Studio em `./Xml/Schemas/{lcid}`. Por exemplo, em uma instalação típica do VS 2017 de 32 bits em um sistema em inglês (EUA), o caminho completo seria `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.</span><span class="sxs-lookup"><span data-stu-id="9819e-p118">On your development machine, locate the existing `MailAppVersionOverridesV1_1.xsd`. This should be located in your Visual Studio installation directory under `./Xml/Schemas/{lcid}`. For example, on a typical installation of VS 2017 32-bit on an English (US) system, the full path would be `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`.</span></span>

        2. <span data-ttu-id="9819e-198">Renomeie o arquivo existente para `MailAppVersionOverridesV1_1.old`.</span><span class="sxs-lookup"><span data-stu-id="9819e-198">Rename the existing file to `MailAppVersionOverridesV1_1.old`.</span></span>

        3. <span data-ttu-id="9819e-199">Copie essa versão modificada do arquivo para a pasta: [Esquema MailAppVersionOverrides modificado](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/master/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)</span><span class="sxs-lookup"><span data-stu-id="9819e-199">Copy this modified version of the file into the folder: [Modified MailAppVersionOverrides Schema](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/master/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)</span></span>

1. <span data-ttu-id="9819e-200">Salve e feche o arquivo de manifesto principal no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="9819e-200">Save and close the main manifest file in Visual Studio.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="9819e-201">Codificar o lado do cliente</span><span class="sxs-lookup"><span data-stu-id="9819e-201">Code the client side</span></span>

1. <span data-ttu-id="9819e-p119">Abra o arquivo Home.js da pasta **Scripts**. Ele já apresenta alguns códigos:</span><span class="sxs-lookup"><span data-stu-id="9819e-p119">Open the Home.js file in the **Scripts** folder. It already has some code in it:</span></span>
    * <span data-ttu-id="9819e-204">Uma atribuição ao método `Office.initialize` que, por sua vez, atribui um manipulador ao evento clicar do botão `getGraphAccessTokenButton`.</span><span class="sxs-lookup"><span data-stu-id="9819e-204">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="9819e-205">Um método `showResult` que exibirá os dados retornados do Microsoft Graph (ou uma mensagem de erro) na parte inferior do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="9819e-205">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="9819e-206">Um método `logErrors` que registrará erros de console que não são destinados ao usuário final.</span><span class="sxs-lookup"><span data-stu-id="9819e-206">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>

1. <span data-ttu-id="9819e-p120">Abaixo da atribuição a `Office.initialize`, adicione o código a seguir. Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="9819e-p120">Below the assignment to `Office.initialize`, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="9819e-p121">O processamento de erros no suplemento às vezes tentará novamente obter um token de acesso automaticamente, usando um conjunto diferente de opções. A variável de contador `timesGetOneDriveFilesHasRun` e a variáveis de sinalizador `triedWithoutForceConsent` são usadas para garantir que o usuário não seja trocado repetidas vezes em tentativas falhas de obter um token.</span><span class="sxs-lookup"><span data-stu-id="9819e-p121">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options. The counter variable `timesGetOneDriveFilesHasRun`, and the flag variable `triedWithoutForceConsent` are used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span>
    * <span data-ttu-id="9819e-p122">Você criará um método `getDataWithToken` na próxima etapa, mas observe que ele define uma opção chamada `forceConsent` como `false`. Trataremos mais disso na etapa seguinte.</span><span class="sxs-lookup"><span data-stu-id="9819e-p122">You create the `getDataWithToken` method in the next step, but note that it sets an option called `forceConsent` to `false`. More about that in the next step.</span></span>

    ```js
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }
    ```

1. <span data-ttu-id="9819e-p123">Abaixo do método `getOneDriveFiles`, adicione o código a seguir. Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="9819e-p123">Below the `getOneDriveFiles` method, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="9819e-215">O [getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) é a nova API no Office.js que permite que um suplemento solicite ao aplicativo host do Office (Excel, PowerPoint, Word etc.) um token de acesso ao suplemento (para o usuário conectado ao Office).</span><span class="sxs-lookup"><span data-stu-id="9819e-215">The [getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) is the new API in Office.js that enables an add-in to ask the Office host application (Excel, PowerPoint, Word, etc.) for an access token to the add-in (for the user signed into Office).</span></span> <span data-ttu-id="9819e-216">O aplicativo host do Office, por sua vez, solicita o token ao ponto de extremidade 2.0 do Azure AD.</span><span class="sxs-lookup"><span data-stu-id="9819e-216">The Office host application, in turn, asks the Azure AD 2.0 endpoint for the token.</span></span> <span data-ttu-id="9819e-217">Uma vez que você previamente autorizou o host do Office para o seu suplemento ao registrá-lo, o Azure AD enviará o token.</span><span class="sxs-lookup"><span data-stu-id="9819e-217">Since you preauthorized the Office host to your add-in when you registered it, Azure AD will send the token.</span></span>
    * <span data-ttu-id="9819e-218">Se nenhum usuário estiver conectado ao Office, o host do Office solicitará que o usuário se conecte.</span><span class="sxs-lookup"><span data-stu-id="9819e-218">If no user is signed into Office, the Office host will prompt the user to sign in.</span></span>
    * <span data-ttu-id="9819e-p125">O parâmetro de opções configura o `forceConsent` como `false`. Dessa forma, não será solicitado que o usuário consinta o acesso ao host do Office ao seu suplemento sempre que ele o usar. Na primeira vez que o usuário tiver o suplemento, a chamada de `getAccessTokenAsync` falhará, mas lógica de processamento de erros que você adicionará em uma etapa posterior será automaticamente chamada com a opção `forceConsent` definida como `true` e o usuário será solicitado a consentir, mas somente essa primeira vez.</span><span class="sxs-lookup"><span data-stu-id="9819e-p125">The options parameter sets `forceConsent` to `false`, so the user will not be prompted to consent to giving the Office host access to your add-in every time she or he uses the add-in. The first time the user runs the add-in, the call of `getAccessTokenAsync` will fail, but error-handling logic that you add in a later step will automatically re-call with the `forceConsent` option set to `true` and the user will be prompted to consent, but only that first time.</span></span>
    * <span data-ttu-id="9819e-221">Você criará o método `handleClientSideErrors` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="9819e-221">You will create the `handleClientSideErrors` method in a later step.</span></span>

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

1. <span data-ttu-id="9819e-p126">Substitua TODO1 pelas linhas a seguir. Você criará o método `getData` e a rota "/api/values" do lado do servidor nas etapas posteriores. Uma URL relativa é usada para o ponto de extremidade porque ela deve ser hospedada no mesmo domínio que seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="9819e-p126">Replace the TODO1 with the following lines. You create the `getData` method and the server-side “/api/values” route in later steps. A relative URL is used for the endpoint because it must be hosted on the same domain as your add-in.</span></span>

    ```js
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. <span data-ttu-id="9819e-p127">Abaixo do método `getOneDriveFiles`, adicione o seguinte. Observe isto sobre este código:</span><span class="sxs-lookup"><span data-stu-id="9819e-p127">Below the `getOneDriveFiles` method, add the following. About this code, note:</span></span>

    * <span data-ttu-id="9819e-p128">Este método utilitário chama um ponto de extremidade da API Web especificado e transmite a ela o mesmo token de acesso que aplicativo host do Office usou para obter acesso ao seu suplemento. No lado do servidor, esse token de acesso será usado no fluxo "on behalf of" (em nome de) para obter um token de acesso para o Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="9819e-p128">This method calls a specified Web API endpoint and passes it the same access token that the Office host application used to get access to your add-in. On the server-side, this access token will be used in the “on behalf of” flow to obtain an access token to Microsoft Graph.</span></span>
    * <span data-ttu-id="9819e-229">Você criará o método `handleServerSideErrors` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="9819e-229">You will create the `handleServerSideErrors` method in a later step.</span></span>

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

### <a name="create-the-error-handling-methods"></a><span data-ttu-id="9819e-230">Crie os métodos de processamento de erros</span><span class="sxs-lookup"><span data-stu-id="9819e-230">Create the error-handling methods</span></span>

1. <span data-ttu-id="9819e-p129">Abaixo do método `getData`, adicione o método a seguir. Esse método processará os erros no cliente do suplemento quando o host do Office não puder obter um token de acesso para o serviço Web do suplemento. Esses erros são relatados com um código de erro, portanto, o método usa uma instrução `switch` para distingui-los.</span><span class="sxs-lookup"><span data-stu-id="9819e-p129">Below the `getData` method, add the following method. This method will handle errors in the add-in's client when the Office host is unable to obtain an access token to the add-in's web service. These errors are reported with an error code, so the method uses a `switch` statement to distinguish them.</span></span>

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

1. <span data-ttu-id="9819e-p130">Substitua `TODO2` pelo código a seguir. O erro 13001 ocorre quando o usuário não está conectado ou quando ele cancela, sem responder, uma solicitação para fornecer um segundo fator de autenticação. Em ambos os casos, o código executará novamente o método `getDataWithToken` e definirá uma opção para forçar uma solicitação de entrada.</span><span class="sxs-lookup"><span data-stu-id="9819e-p130">Replace `TODO2` with the following code. Error 13001 occurs when the user is not logged in, or the user cancelled, without responding, a prompt to provide a 2nd authentication factor. In either case, the code re-runs the `getDataWithToken` method and sets an option to force a sign-in prompt.</span></span>

    ```js
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. <span data-ttu-id="9819e-p131">Substitua `TODO3` pelo código a seguir. O erro 13002 ocorre quando a entrada ou o consentimento do usuário é anulado. Peça que o usuário tente novamente, mas não mais de uma vez.</span><span class="sxs-lookup"><span data-stu-id="9819e-p131">Replace `TODO3` with the following code. Error 13002 occurs when user's sign-in or consent was aborted. Ask the user to try again but no more than once again.</span></span>

    ```js
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }
        break;
    ```

1. <span data-ttu-id="9819e-240">Substitua `TODO4` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="9819e-240">Replace `TODO4` with the following code.</span></span> <span data-ttu-id="9819e-241">O erro 13003 ocorre quando o usuário está conectado com uma conta que não é corporativa, de estudante nem da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="9819e-241">Error 13003 occurs when user is logged in with an account that is neither work or school, nor Microsoft account.</span></span> <span data-ttu-id="9819e-242">Peça que o usuário saia e entre novamente com um tipo de conta suportado.</span><span class="sxs-lookup"><span data-stu-id="9819e-242">Ask the user to sign-out and then in again with a supported account type.</span></span>

    ```js
    case 13003:
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;
    ```

    > [!NOTE]
    > <span data-ttu-id="9819e-243">O erro 13004 não é processado neste método, pois eles ocorre apenas em desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="9819e-243">Error 13004 is not handled in this method because it should only occur in development.</span></span> <span data-ttu-id="9819e-244">Não é possível corrigi-lo pelo código de tempo de execução e não seria útil reportá-lo a um usuário final.</span><span class="sxs-lookup"><span data-stu-id="9819e-244">It cannot be fixed by runtime code and there would be no point in reporting it to an end user.</span></span>

1. <span data-ttu-id="9819e-245">Substitua `TODO5` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="9819e-245">Replace `TODO5` with the following code.</span></span> <span data-ttu-id="9819e-246">O erro 13005 ocorre quando o Office não tem autorização para o serviço Web do suplemento ou o usuário não concedeu permissão ao serviço para o respectivo `profile`.</span><span class="sxs-lookup"><span data-stu-id="9819e-246">Error 13005 occurs when Office has not been authorized to the add-in's web service or the user has not granted the service permission to their `profile`.</span></span>

    ```js
    case 13005:
        getDataWithToken({ forceConsent: true });
        break;
    ```

1. <span data-ttu-id="9819e-p135">Substitua `TODO6` pelo seguinte código. O Erro 13006 ocorre quando houve um erro não especificado no host do Office, que pode indicar a instabilidade do host. Peça ao usuário para reiniciar o Office.</span><span class="sxs-lookup"><span data-stu-id="9819e-p135">Replace `TODO6` with the following code. Error 13006 occurs when there has been an unspecified error in the Office host that may indicate that the host is in an unstable state. Ask the user to restart Office.</span></span>

    ```js
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;
    ```

1. <span data-ttu-id="9819e-p136">Substitua `TODO7` pelo código a seguir. O erro 13007 ocorre quando algo deu errado com a interação do host do Office com o AAD de forma que o host não pode obter um token de acesso para o serviço Web/aplicativo dos suplementos. É possível que esse seja um problema de rede temporário. Peça que o usuário tente novamente mais tarde.</span><span class="sxs-lookup"><span data-stu-id="9819e-p136">Replace `TODO7` with the following code. Error 13007 occurs when something has gone wrong with the Office host's interaction with AAD so the host cannot get an access token to the add-ins web service/application. This may be a temporary network issue. Ask the user to try again later.</span></span>

    ```js
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;
    ```

1. <span data-ttu-id="9819e-254">Substitua `TODO8` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="9819e-254">Replace `TODO8` with the following code.</span></span> <span data-ttu-id="9819e-255">O erro 13008 ocorre quando o usuário aciona uma operação que chama o `getAccessTokenAsync` antes que uma chamada anterior dele seja concluída.</span><span class="sxs-lookup"><span data-stu-id="9819e-255">Error 13008 occurs when the user triggered an operation that calls `getAccessTokenAsync` before a previous call of it completed.</span></span>

    ```js
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```

1. <span data-ttu-id="9819e-p138">Substitua `TODO9` pelo código a seguir. O erro 13009 ocorre quando o suplemento não permite forçar consentimento, mas `getAccessTokenAsync` foi chamado com a opção `forceConsent` definida como `true`. Normalmente, quando isso acontece, o código deve ser reexecutar `getAccessTokenAsync` automaticamente com a opção de consentimento definida como `false`. No entanto, em alguns casos, chamar o método com `forceConsent` definido como `true` é uma resposta automática para um erro em uma chamada para o método com a opção definida como `false`. Nesse caso, o código não deve tentar novamente, mas, em vez disso, ele deve solicitar que o usuário saia e entre novamente.</span><span class="sxs-lookup"><span data-stu-id="9819e-p138">Replace `TODO9` with the following code. Error 13009 occurs when the add-in does not support forcing consent, but `getAccessTokenAsync` was called with the `forceConsent` option set to `true`. In the usual case when this happens the code should automatically re-run `getAccessTokenAsync` with the consent option set to `false`. However, in some cases, calling the method with `forceConsent` set to `true` was itself an automatic response to an error in a call to the method with the option set to `false`. In that case, the code should not try again, but instead it should advise the user to sign out and sign in again.</span></span>

    ```js
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```

1. <span data-ttu-id="9819e-261">Substitua `TODO10` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="9819e-261">Replace `TODO10` with the following code.</span></span>

    ```js
    default:
        logError(result);
        break;
    ```  


1. <span data-ttu-id="9819e-p139">Abaixo do método `handleClientSideErrors`, adicione o seguinte método. Esse método processará os erros no serviço Web do suplemento quando algo der errado na execução do fluxo on-behalf-of ou ao obter dados do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="9819e-p139">Below the `handleClientSideErrors` method, add the following method. This method will handle errors in the add-in's web service when something goes wrong in executing the on-behalf-of flow or in getting data from Microsoft Graph.</span></span>

    ```js
    function handleServerSideErrors(result) {

        // TODO11: Parse the JSON response.

        // TODO12: Handle the case where AAD asks for an additional form of authentication.

        // TODO13: Handle missing consent and scope (permission) related issues.

        // TODO14: Handle the case where the token sent to Microsoft Graph in the request for
        //         data is expired or invalid.

        // TODO15: Log all other server errors.
    }
    ```

1. <span data-ttu-id="9819e-264">Substitua `TODO11` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="9819e-264">Replace `TODO11` with the following code.</span></span> <span data-ttu-id="9819e-265">Observe que, para a maioria dos erros `4xx` que o serviço Web do suplemento passará para o suplemento do lado do cliente, haverá uma propriedade **ExceptionMessage** em resposta com o número de erro AADSTS (Azure Active Directory Secure Token Service) além de outros dados.</span><span class="sxs-lookup"><span data-stu-id="9819e-265">Note that for most of the `4xx` errors that the add-in's web service will pass to the add-in's client-side, there will be an **ExceptionMessage** property in the response that contains the AADSTS (Azure Active Directory Secure Token Service) error number as well as other data.</span></span> <span data-ttu-id="9819e-266">No entanto, quando o AAD envia uma mensagem ao serviço Web do suplemento solicitando um fator de autenticação adicional, a mensagem contém uma propriedade **Claims** especial que especifica (com um número de código) qual fator adicional é necessário.</span><span class="sxs-lookup"><span data-stu-id="9819e-266">However, when AAD sends a message to the add-in's web service asking for an additional authentication factor, the message contains a special **Claims** property that specifies (with a code number) what additional factor is needed.</span></span> <span data-ttu-id="9819e-267">As APIs ASP.NET que criam e enviam respostas HTTP para clientes não conhecem a propriedade **Claims**, portanto, elas não a incluem no objeto Response.</span><span class="sxs-lookup"><span data-stu-id="9819e-267">The ASP.NET APIs that create and send HTTP Responses to clients do not know about this **Claims** property, so they do not include it in the Response object.</span></span> <span data-ttu-id="9819e-268">O código de servidor que será criado em uma etapa posterior lidará com isso adicionando manualmente o valor **Claims** no objeto Response.</span><span class="sxs-lookup"><span data-stu-id="9819e-268">Server-side code that you will create in a later step will cope with this by manually adding the **Claims** value to the Response object.</span></span> <span data-ttu-id="9819e-269">Esse valor será salvo na propriedade **Message**, portanto, o código também precisará analisar essa propriedade.</span><span class="sxs-lookup"><span data-stu-id="9819e-269">This value will be in the **Message** property, so the code needs to parse out that property as well.</span></span>

    ```js
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    var message = JSON.parse(result.responseText).Message;
    ```

1. <span data-ttu-id="9819e-270">Substitua `TODO12` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="9819e-270">Replace `TODO12` with the following code.</span></span> <span data-ttu-id="9819e-271">Observação sobre o código:</span><span class="sxs-lookup"><span data-stu-id="9819e-271">Note about this code:</span></span>

    * <span data-ttu-id="9819e-272">O erro 50076 ocorre quando o Microsoft Graph requer uma forma adicional de autenticação.</span><span class="sxs-lookup"><span data-stu-id="9819e-272">Error 50076 occurs when Microsoft Graph requires an additional form of authentication.</span></span>
    * <span data-ttu-id="9819e-p142">O host do Office deve obter um novo token com o valor **Claims** como a opção `authChallenge`. Isso instrui o AAD a solicitar ao usuário todas as formas de autenticação requeridas.</span><span class="sxs-lookup"><span data-stu-id="9819e-p142">The Office host should get a new token with the **Claims** value as the `authChallenge` option. This tells AAD to prompt the user for all required forms of authentication.</span></span>

    ```js
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
        }
    }
    ```

1. <span data-ttu-id="9819e-275">Substitua `TODO13` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="9819e-275">Replace `TODO13` with the following code.</span></span> <span data-ttu-id="9819e-276">Substitua os três `TODO`s neste código por um bloqueio condicional *interno* nas próximas etapas.</span><span class="sxs-lookup"><span data-stu-id="9819e-276">You will replace the three `TODO`s in this code with an *inner* conditional block in the next few steps.</span></span>

    ```js
    else if (exceptionMessage) {

        // TODO13A: Handle the case where consent has not been granted, or has been revoked.

        // TODO13B: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow.

        // TODO13C: Handle the case where the token that the add-in's client-side sends to it's
        //          server-side is not valid because it is missing `access_as_user` scope (permission).
    }
  
    ```


1. <span data-ttu-id="9819e-277">Substitua `TODO13A` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="9819e-277">Replace `TODO13A` with the following code.</span></span> <span data-ttu-id="9819e-278">(Isso cria a primeira parte de um bloqueio condicional *interno*.) Observação sobre o código:</span><span class="sxs-lookup"><span data-stu-id="9819e-278">(This creates the first part of an *inner* conditional block.) Note about this code:</span></span>

    * <span data-ttu-id="9819e-279">O erro 65001 significa que o consentimento para acessar o Microsoft Graph não foi concedido (ou foi revogado) para uma ou mais permissões.</span><span class="sxs-lookup"><span data-stu-id="9819e-279">Error 65001 means that consent to access Microsoft Graph was not granted (or was revoked) for one or more permissions.</span></span>
    * <span data-ttu-id="9819e-280">O suplemento deverá obter um novo token com a opção `forceConsent` definida como `true`.</span><span class="sxs-lookup"><span data-stu-id="9819e-280">The add-in should get a new token with the `forceConsent` option set to `true`.</span></span>

    ```js
    if (exceptionMessage.indexOf('AADSTS65001') !== -1) {
       getDataWithToken({ forceConsent: true });
    }
    ```

1. <span data-ttu-id="9819e-p145">Substitua `TODO13B` pelo código a seguir. Observação sobre este código:</span><span class="sxs-lookup"><span data-stu-id="9819e-p145">Replace `TODO13B` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="9819e-p146">O Erro 70011 tem muitos significados. O que importa para este suplemento é quando ele significa que um escopo inválido (permissão) foi solicitado, então o código verifica a descrição completa do erro, não apenas o número.</span><span class="sxs-lookup"><span data-stu-id="9819e-p146">Error 70011 has multiple meanings. The one that matters to this add-in is when it means that an invalid scope (permission) has been requested, so the code checks for the full error description, not just the number.</span></span>
    * <span data-ttu-id="9819e-285">O suplemento deverá relatar o erro.</span><span class="sxs-lookup"><span data-stu-id="9819e-285">The add-in should report the error.</span></span>

    ```js
     else if (exceptionMessage.indexOf("AADSTS70011: The provided value for the input parameter 'scope' is not valid.") !== -1) {
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. <span data-ttu-id="9819e-p147">Substitua `TODO13C` pelo código a seguir. Observação sobre este código:</span><span class="sxs-lookup"><span data-stu-id="9819e-p147">Replace `TODO13C` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="9819e-288">Código de servidor criado em uma etapa posterior enviará a mensagem `Missing access_as_user` se o escopo `access_as_user` (permissão) não for o token de acesso que o cliente do suplemento enviar para o ADD para ser usado no fluxo on-behalf-of.</span><span class="sxs-lookup"><span data-stu-id="9819e-288">Server-side code that you create in a later step will send the message `Missing access_as_user` if the `access_as_user` scope (permission) is not in the access token that the add-in's client sends to AAD to be used in the on-behalf-of flow.</span></span>
    * <span data-ttu-id="9819e-289">O suplemento deverá relatar o erro.</span><span class="sxs-lookup"><span data-stu-id="9819e-289">The add-in should report the error.</span></span>

    ```js
    else if (exceptionMessage.indexOf('Missing access_as_user.') !== -1) {
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. <span data-ttu-id="9819e-290">Substitua `TODO14` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="9819e-290">Replace `TODO14` with the following code.</span></span> <span data-ttu-id="9819e-291">(Isso faz parte do bloqueio condicional *externo* e deve ser colocado imediatamente após o colchete de fechamento da estrutura que começa com `else if (exceptionMessage) {` e com o mesmo nível de recuo.) Observação sobre o código:</span><span class="sxs-lookup"><span data-stu-id="9819e-291">(This is part of the *outer* conditional block and should be immediately after the close bracket of the structure that begins with `else if (exceptionMessage) {` and at the same level of indentation.) Note about this code:</span></span>

    * <span data-ttu-id="9819e-292">A biblioteca de identidade que você estiver usando no código do servidor (Biblioteca de Autenticação da Microsoft - MSAL) deve garantir que nenhum token inválido ou expirado seja enviado para o Microsoft Graph, mas, se isso acontecer, o erro retornado para serviço Web do suplemento do Microsoft Graph terá o código `InvalidAuthenticationToken`.</span><span class="sxs-lookup"><span data-stu-id="9819e-292">The identity library that you will be using in the server-side code (Microsoft Authentication Library - MSAL) should ensure that no expired or invalid token is sent to Microsoft Graph; but if it does happen, the error that is returned to the add-in's web service from Microsoft Graph has the code `InvalidAuthenticationToken`.</span></span> <span data-ttu-id="9819e-293">O código de servidor criado em uma etapa posterior transmitirá essa mensagem para o cliente do suplemento.</span><span class="sxs-lookup"><span data-stu-id="9819e-293">Server-side code you will create in a later step will relay this message to the add-in's client.</span></span>
    * <span data-ttu-id="9819e-294">Nesse caso, o suplemento deverá iniciar o processo de autenticação completo ao redefinir o contador e as variáveis de sinalizador e, em seguida, chamando novamente o método de identificador de botão.</span><span class="sxs-lookup"><span data-stu-id="9819e-294">In this case, the add-in should start the entire authentication process over by resetting the counter and flag variables, and then re-calling the button handler method.</span></span>

    ```js
    // If the token sent to MS Graph is expired or invalid, start the whole process over.
    else if (result.code === 'InvalidAuthenticationToken') {
        timesGetOneDriveFilesHasRun = 0;
        triedWithoutForceConsent = false;
        getOneDriveFiles();
    }
    ```

1. <span data-ttu-id="9819e-295">Substitua `TODO15` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="9819e-295">Replace `TODO15` with the following code.</span></span>

    ```js
    else {
        logError(result);
    }
    ```

1. <span data-ttu-id="9819e-296">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="9819e-296">Save and close the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="9819e-297">Codifique o lado do servidor</span><span class="sxs-lookup"><span data-stu-id="9819e-297">Code the server side</span></span>

### <a name="configure-the-owin-middleware"></a><span data-ttu-id="9819e-298">Configurar o middleware OWIN</span><span class="sxs-lookup"><span data-stu-id="9819e-298">Configure the OWIN middleware</span></span>

1. <span data-ttu-id="9819e-299">Abra o arquivo Startup.cs na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="9819e-299">Open the Startup.cs file in the root of the project.</span></span>

1. <span data-ttu-id="9819e-p150">Adicione a palavra-chave `partial` para a declaração da classe Startup, se ainda não estiver lá. A linha deverá ser assim:</span><span class="sxs-lookup"><span data-stu-id="9819e-p150">Add the keyword `partial` to the declaration of the Startup class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="9819e-p151">Adicione a linha a seguir ao corpo do método `Configuration`. Você criará o método `ConfigureAuth` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="9819e-p151">Add the following line to the body of the `Configuration` method. You create the `ConfigureAuth` method in a later step.</span></span>

    `ConfigureAuth(app);`

1. <span data-ttu-id="9819e-304">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="9819e-304">Save and close the file.</span></span>

1. <span data-ttu-id="9819e-305">Clique com botão direito do mouse na pasta **App_Start** e selecione **Adicionar > Classe**.</span><span class="sxs-lookup"><span data-stu-id="9819e-305">Right-click the **App_Start** folder and select **Add > Class**.</span></span>

1. <span data-ttu-id="9819e-306">Na caixa de diálogo **Adicionar novo item** nomeie o arquivo **Startup.Auth.cs** e, em seguida, clique em **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="9819e-306">In the **Add new item** dialog name the file **Startup.Auth.cs** and then click **Add**.</span></span>

1. <span data-ttu-id="9819e-307">Encurte o nome do namespace no novo arquivo para `Office_Add_in_ASPNET_SSO_WebAPI`.</span><span class="sxs-lookup"><span data-stu-id="9819e-307">Shorten the namespace name in the new file to `Office_Add_in_ASPNET_SSO_WebAPI`.</span></span>

1. <span data-ttu-id="9819e-308">Verifique se todas as seguintes instruções `using` estão na parte superior do arquivo.</span><span class="sxs-lookup"><span data-stu-id="9819e-308">Ensure that all of the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Owin;
    using System.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. <span data-ttu-id="9819e-p152">Adicione a palavra-chave `partial` à declaração da classe `Startup`, se ainda não estiver lá. A linha deverá ser assim:</span><span class="sxs-lookup"><span data-stu-id="9819e-p152">Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="9819e-p153">Adicione o método a seguir à classe `Startup`. Este método especifica como o middleware OWIN validará os tokens de acesso que são transmitidos a ele do método `getData` no arquivo Home.js do lado do cliente. O processo de autorização é disparado sempre que um ponto de extremidade da API Web decorado com o atributo `[Authorize]` é chamado.</span><span class="sxs-lookup"><span data-stu-id="9819e-p153">Add the following method to the `Startup` class. This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file. The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.</span></span>

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. <span data-ttu-id="9819e-p154">Substitua TODO3 pelo seguinte código. Observação sobre o código:</span><span class="sxs-lookup"><span data-stu-id="9819e-p154">Replace the TODO3 with the following. Note about this code:</span></span>

    * <span data-ttu-id="9819e-316">O código instrui o OWIN a garantir que o emissor de token e audiência especificado no token de acesso que vem do host do Office (e é transmitido pela chamada de `getData` do lado do cliente) deve coincidir com os valores especificados no Web.config.</span><span class="sxs-lookup"><span data-stu-id="9819e-316">The code instructs OWIN to ensure that the audience and token issuer specified in the access token that comes from the Office host (and is passed on by the client-side call of `getData`) must match the values specified in the web.config.</span></span>
    * <span data-ttu-id="9819e-p155">Definir `SaveSigninToken` como `true` faz com que o OWIN salve o token bruto do host do Office. O suplemento precisa dele para obter um token de acesso para o Microsoft Graph com o fluxo "on behalf of".</span><span class="sxs-lookup"><span data-stu-id="9819e-p155">Setting `SaveSigninToken` to `true` causes OWIN to save the raw token from the Office host. The add-in needs it to obtain an access token to Microsoft Graph with the “on behalf of” flow.</span></span>
    * <span data-ttu-id="9819e-p156">Os escopos não são validados pelo middleware OWIN. Os escopos do token de acesso, que devem conter `access_as_user`, são validados no controlador.</span><span class="sxs-lookup"><span data-stu-id="9819e-p156">Scopes are not validated by the OWIN middleware. The scopes of the access token, which should include `access_as_user`, is validated in the controller.</span></span>

    ```csharp
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. <span data-ttu-id="9819e-p157">Substitua TODO4 pelo seguinte. Observação sobre este código:</span><span class="sxs-lookup"><span data-stu-id="9819e-p157">Replace TODO4 with the following. Note about this code:</span></span>

    * <span data-ttu-id="9819e-323">O método `UseOAuthBearerAuthentication` é chamado em vez do `UseWindowsAzureActiveDirectoryBearerAuthentication` que é mais comum, porque este último não é compatível com o ponto de extremidade V2 do Azure AD.</span><span class="sxs-lookup"><span data-stu-id="9819e-323">The method `UseOAuthBearerAuthentication` is called instead of the more common `UseWindowsAzureActiveDirectoryBearerAuthentication` because the latter is not compatible with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="9819e-324">A URL de descoberta transmitida ao método é onde o middleware OWIN obtém instruções para conseguir a chave que precisa para verificar a assinatura no token de acesso recebido do host do Office.</span><span class="sxs-lookup"><span data-stu-id="9819e-324">The discovery URL that is passed to the method is where the OWIN middleware obtains instructions for getting the key it needs to verify the signature on the access token received from the Office host.</span></span>

    ```csharp
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
        {
            AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
        });
    ```

1. <span data-ttu-id="9819e-325">Salve e feche o arquivo.</span><span class="sxs-lookup"><span data-stu-id="9819e-325">Save and close the file.</span></span>

### <a name="create-the-apivalues-controller"></a><span data-ttu-id="9819e-326">Criar o controlador /api/values</span><span class="sxs-lookup"><span data-stu-id="9819e-326">Create the /api/values controller</span></span>

1. <span data-ttu-id="9819e-327">Abra o arquivo **Controllers\ValueController.cs**.</span><span class="sxs-lookup"><span data-stu-id="9819e-327">Open the file **Controllers\ValueController.cs**.</span></span>

1. <span data-ttu-id="9819e-328">Verifique se as seguintes instruções `using` estão na parte superior do arquivo.</span><span class="sxs-lookup"><span data-stu-id="9819e-328">Ensure that the following `using` statements are at the top of the file.</span></span>

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

1. <span data-ttu-id="9819e-p158">Logo acima da linha que declara o `ValuesController`, adicione o atributo `[Authorize]`. Isso garante que seu suplemento executará o processo de autorização configurado no último procedimento sempre que um método controlador for chamado. Apenas os chamadores com um token de acesso válido para o seu suplemento podem invocar os métodos do controlador.</span><span class="sxs-lookup"><span data-stu-id="9819e-p158">Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.</span></span>

    > [!NOTE]
    > <span data-ttu-id="9819e-p159">Um serviço da ASP.NET MVC Web API de produção deve ter lógica personalizada para o fluxo on-behalf-of em uma ou mais classes **FilterAttribute** personalizadas. Este exemplo educacional coloca a lógica no controlador de principal para que o fluxo de autorização e dados busca lógica inteiro possa ser acompanhado facilmente. Isso também faz com que o exemplo fique consistente com os exemplos de padrão de autorização nos [Exemplos do Azure](https://github.com/Azure-Samples/).</span><span class="sxs-lookup"><span data-stu-id="9819e-p159">A production ASP.NET MVC Web API service should have custom logic for the on-behalf-of flow in one or more custom **FilterAttribute** classes. This educational sample puts the logic in the main controller so that the entire flow of the authorization and data fetching logic can be easily followed. This also makes the sample consistent with the pattern of authorization samples in [Azure Samples](https://github.com/Azure-Samples/).</span></span>

1. <span data-ttu-id="9819e-p160">Adicione o método a seguir ao `ValuesController`. Observe que é o valor de retorno é `Task<HttpResponseMessage>` em vez de `Task<IEnumerable<string>>`, como seria mais comum para um método `GET api/values`. Este é um efeito colateral do fato de que nossa lógica de autorização personalizada estará no controlador: algumas condições de erro nessa lógica exigem que um objeto de resposta HTTP seja enviado para o cliente do suplemento.</span><span class="sxs-lookup"><span data-stu-id="9819e-p160">Add the following method to the `ValuesController`. Note that the return value is `Task<HttpResponseMessage>` instead of `Task<IEnumerable<string>>` as would be more common for a `GET api/values` method. This is a side effect of that fact that our custom authorization logic will be in the controller: some error conditions in that logic require that an HTTP Response object be sent to the add-in's client.</span></span>

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO1: Validate the scopes of the access token.
    }
    ```

1. <span data-ttu-id="9819e-338">Substitua `TODO1` pelo seguinte código para validar que os escopos especificados no token incluam `access_as_user`.</span><span class="sxs-lookup"><span data-stu-id="9819e-338">Replace `TODO1` with the following code to validate that the scopes that are specified in the token include `access_as_user`.</span></span>

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
    > <span data-ttu-id="9819e-339">Você deve usar apenas o escopo `access_as_user` para autorizar a API que lida com o fluxo Em Nome De para os suplementos do Office. Outras APIs em seu serviço devem ter seus próprios requisitos de escopo.</span><span class="sxs-lookup"><span data-stu-id="9819e-339">You should only use the `access_as_user` scope to authorize the API that handles the on-behalf-of flow for Office Add-ins. Other APIs in your service should have their own scope requirements.</span></span> <span data-ttu-id="9819e-340">Isso limita o que pode ser acessado com os tokens que o Office adquire.</span><span class="sxs-lookup"><span data-stu-id="9819e-340">This limits what can be accessed with the tokens that Office acquires.</span></span>

1. <span data-ttu-id="9819e-p162">Substitua `TODO2` pelo código a seguir. Observação sobre este código:</span><span class="sxs-lookup"><span data-stu-id="9819e-p162">Replace `TODO2` with the following code. Note about this code:</span></span>
    * <span data-ttu-id="9819e-343">Ele transforma o token de acesso bruto recebido do host do Office em um objeto de `UserAssertion` que será transmitido para outro método.</span><span class="sxs-lookup"><span data-stu-id="9819e-343">It turns the raw access token received from the Office host into a `UserAssertion` object that will be passed to another method.</span></span>
    * <span data-ttu-id="9819e-p163">Seu suplemento não está mais desempenhando o papel de um recurso (ou público) para o qual o host do Office e o usuário precisam de acesso. Agora, ele mesmo é um cliente que precisa de acesso ao Microsoft Graph. `ConfidentialClientApplication` é o objeto "client context" da MSAL.</span><span class="sxs-lookup"><span data-stu-id="9819e-p163">Your add-in is no longer playing the role of a resource (or audience) to which the Office host and user need access. Now it is itself a client that needs access to Microsoft Graph. `ConfidentialClientApplication` is the MSAL “client context” object.</span></span>
    * <span data-ttu-id="9819e-p164">O terceiro parâmetro para o construtor `ConfidentialClientApplication` é uma URL de redirecionamento que não é realmente usada no fluxo "on behalf of", mas usar a URL correta é uma boa prática. O quarto e o quinto parâmetros podem ser usados para definir um armazenamento persistente que permitiria a reutilização de tokens não expirados em diferentes sessões com o suplemento. Este exemplo não implementa nenhum armazenamento persistente.</span><span class="sxs-lookup"><span data-stu-id="9819e-p164">The third parameter to the `ConfidentialClientApplication` constructor is a redirect URL which is not actually used in the “on behalf of” flow, but it is a good practice to use the correct URL. The fourth and fifth parameters can be used to define a persistent store that would enable the reuse of unexpired tokens across different sessions with the add-in. This sample does not implement any persistent storage.</span></span>
    * <span data-ttu-id="9819e-p165">A MSAL exige os escopos `openid` e `offline_access` para funcionar, mas ela lança um erro se o código solicitá-los de forma redundante. Ela também lançará um erro se o seu código solicitar o `profile`, que realmente é usado apenas quando o aplicativo host do Office recebe o token para o aplicativo Web do seu suplemento. Então, apenas `Files.Read.All` é explicitamente solicitado.</span><span class="sxs-lookup"><span data-stu-id="9819e-p165">MSAL requires the `openid` and `offline_access` scopes to function, but it throws an error if your code redundantly requests them. It will also throw an error if your code requests `profile`, which is really only used when the Office host application gets the token to your add-in's web application. So only `Files.Read.All` is explicitly requested.</span></span>

    ```csharp
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

1. <span data-ttu-id="9819e-p166">Substitua `TODO3` pelo código a seguir. Observação sobre este código:</span><span class="sxs-lookup"><span data-stu-id="9819e-p166">Replace `TODO3` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="9819e-p167">O método `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` procurará primeiro no cache da MSAL, que está na memória, para fazer a correspondência com o token de acesso. Somente se não houver um, ele iniciará o fluxo "on behalf of" com o ponto de extremidade V2 do Azure AD.</span><span class="sxs-lookup"><span data-stu-id="9819e-p167">The `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` method will first look in the MSAL cache, which is in memory, for a matching access token. Only if there isn't one, does it initiate the "on behalf of" flow with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="9819e-357">Se a autenticação multi-fator for requerida pelo recurso MS Graph e o usuário ainda não a tiver fornecido, o AAD lançará uma exceção contendo uma propriedade de Declarações.</span><span class="sxs-lookup"><span data-stu-id="9819e-357">If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will throw an exception containing a Claims property.</span></span>
    * <span data-ttu-id="9819e-p168">O valor da propriedade de Declarações deve ser passado para o cliente, que o passará para o host do Office, que, em seguida, o incluirá em um pedido para um novo token. O AAD solicitará ao usuário todas as formas de autenticação necessárias.</span><span class="sxs-lookup"><span data-stu-id="9819e-p168">The Claims property value must be passed to the client which will pass it to the Office host, which will then include it in a request for a new token. AAD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="9819e-360">Quaisquer exceções que não forem do tipo `MsalServiceException` são intencionalmente não detectadas, e, portanto, se propagarão para o cliente como mensagens `500 Server Error`.</span><span class="sxs-lookup"><span data-stu-id="9819e-360">Any exceptions that are not of type `MsalServiceException` are intentionally not caught, so they will propagate to the client as `500 Server Error` messages.</span></span>

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

1. <span data-ttu-id="9819e-p169">Substitua `TODO3a` pelo código a seguir. Observação sobre este código:</span><span class="sxs-lookup"><span data-stu-id="9819e-p169">Replace `TODO3a` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="9819e-p170">Se a autenticação multifator for exigida pelo recurso MS Graph e o usuário ainda não a tiver fornecido, o AAD retornará "400 Bad Request" com o erro AADSTS50076 e uma propriedade **Declarações**. O MSAL lançará uma **MsalUiRequiredException** (que herda de **MsalServiceException**) com essas informações.</span><span class="sxs-lookup"><span data-stu-id="9819e-p170">If multi-factor authentication is required by the MS Graph resource and the user has not yet provided it, AAD will return "400 Bad Request" with error AADSTS50076 and a **Claims** property. MSAL throws a **MsalUiRequiredException** (which inherits from **MsalServiceException**) with this information.</span></span> 
    * <span data-ttu-id="9819e-p171">O valor da propriedade **Declarações** deve ser passado para o cliente, que deve passá-lo para o host do Office, que, por sua vez, o incluirá em um pedido para um novo token. O AAD solicitará ao usuário todas as formas de autenticação necessárias.</span><span class="sxs-lookup"><span data-stu-id="9819e-p171">The **Claims** property value must be passed to the client which should pass it to the Office host, which then includes it in a request for a new token. AAD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="9819e-p172">As APIs que criam respostas HTTP a partir de exceções não conhecem a propriedade **Claims**, portanto, elas não a incluem no objeto de resposta. É necessário criar manualmente uma mensagem que inclua esse recurso. Uma propriedade **Message** personalizada, no entanto, impede a criação de uma propriedade **ExceptionMessage**, portanto, a única maneira de obter a ID de erro `AADSTS50076` para o cliente é adicioná-la à **Message** personalizada. O JavaScript no cliente precisará descobrir se uma resposta tem uma **Message** ou **ExceptionMessage** para saber qual ler.</span><span class="sxs-lookup"><span data-stu-id="9819e-p172">The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object. We have to manually create a message that includes it. A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**. JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.</span></span>
    * <span data-ttu-id="9819e-371">A mensagem personalizada é formatada como JSON para que o JavaScript do cliente possa analisá-la com métodos de objeto `JSON` conhecidos.</span><span class="sxs-lookup"><span data-stu-id="9819e-371">The custom message is formatted as JSON so that the client-side JavaScript can parse it with well-known `JSON` object methods.</span></span>
    * <span data-ttu-id="9819e-p173">Você criará o método `SendErrorToClient` em uma etapa posterior. É segundo parâmetro é um objeto **Exception**. Nesse caso, o código passa `null` porque incluir o objeto **Exception** bloqueia a inclusão da propriedade **Message** na resposta HTTP que é gerada.</span><span class="sxs-lookup"><span data-stu-id="9819e-p173">You will create the `SendErrorToClient` method in a later step. It's second parameter is an **Exception** object. In this case, the code passes `null` because including the **Exception** object blocks the inclusion of the **Message** property in the HTTP Response that is generated.</span></span>

    ```csharp
    if (e.Message.StartsWith("AADSTS50076")) {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. <span data-ttu-id="9819e-p174">Substitua `TODO3b` e `TODO3c` pelo código a seguir. Observação sobre este código:</span><span class="sxs-lookup"><span data-stu-id="9819e-p174">Replace `TODO3b` and `TODO3c` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="9819e-p175">Se a chamada para o AAD contiver pelo menos um escopo (permissão) que não tenha sido consentido pelo usuário ou por um administrador de locatários (ou se o consentimento foi revogado), o AAD retornará "400 Solicitação Incorreta" com o erro `AADSTS65001`. O MSAL exibe um **MsalUiRequiredException** com essas informações. O cliente deve chamar `getAccessTokenAsync` novamente com a opção `{ forceConsent: true }`.</span><span class="sxs-lookup"><span data-stu-id="9819e-p175">If the call to AAD contained at least one scope (permission) for which neither the user nor a tenant administrator has consented (or consent was revoked). AAD will return "400 Bad Request" with error `AADSTS65001`. MSAL throws a **MsalUiRequiredException** with this information. The client should re-call `getAccessTokenAsync` with the option `{ forceConsent: true }`.</span></span>
    *  <span data-ttu-id="9819e-p176">Se a chamada para o AAD contiver pelo menos um escopo que AAD não reconhece, o AAD retornará "400 Solicitação Incorreta" com o erro `AADSTS70011`. O MSAL exibe um **MsalUiRequiredException** com essas informações. O cliente deve informar o usuário.</span><span class="sxs-lookup"><span data-stu-id="9819e-p176">If the call to AAD contained at least one scope that AAD does not recognize, AAD returns "400 Bad Request" with error `AADSTS70011`. MSAL throws a **MsalUiRequiredException** with this information. The client should inform the user.</span></span>
    *  <span data-ttu-id="9819e-384">A descrição completa é incluída porque 70011 é retornado em outras condições e ele deverá ser processado neste suplemento somente quando significar que há um escopo inválido.</span><span class="sxs-lookup"><span data-stu-id="9819e-384">The entire description is included because 70011 is returned in other conditions and we it should only be handled in this add-in when it means that there is an invalid scope.</span></span>
    *  <span data-ttu-id="9819e-p177">O objeto **MsalUiRequiredException** é passado para `SendErrorToClient`. Isso garante que uma propriedade **ExceptionMessage** contendo as informações de erro seja incluída na resposta HTTP.</span><span class="sxs-lookup"><span data-stu-id="9819e-p177">The **MsalUiRequiredException** object is passed to `SendErrorToClient`. This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.</span></span>
    *  <span data-ttu-id="9819e-387">Não há uma mensagem personalizada, portanto, `null` é passado para o terceiro parâmetro.</span><span class="sxs-lookup"><span data-stu-id="9819e-387">There is no custom message, so `null` is passed for the third parameter.</span></span>

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001"))
    || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. <span data-ttu-id="9819e-p178">Substitua `TODO3d` pelo código a seguir. Observe que o código exibe a exceção em vez de transmiti-la em uma resposta HTTP personalizada com **HttpStatusCode.Forbidden** (401). O efeito disso é que o ASP.NET enviará sua própria resposta HTTP com o status "500 Erro de Servidor".</span><span class="sxs-lookup"><span data-stu-id="9819e-p178">Replace `TODO3d` with the following code. Note that the code rethrows the exception instead of relaying it in a custom HTTP Response with **HttpStatusCode.Forbidden** (401). The effect of this is that the ASP.NET will send its own HTTP Response with status "500 Server Error".</span></span>

    ```csharp
    else
    {
        throw e;
    }  
    ```

1. <span data-ttu-id="9819e-p179">Substitua `TODO4` pelo seguinte. Observação sobre este código:</span><span class="sxs-lookup"><span data-stu-id="9819e-p179">Replace `TODO4` with the following. Note about this code:</span></span>

    * <span data-ttu-id="9819e-p180">As classes `GraphApiHelper` e `ODataHelper` são definidas nos arquivos da pasta **Helpers**. A classe `OneDriveItem` é definida em um arquivo da pasta **Models**. A discussão detalhada dessas classes não é relevante para a autorização ou o SSO, portanto, está fora do escopo deste artigo.</span><span class="sxs-lookup"><span data-stu-id="9819e-p180">The `GraphApiHelper` and `ODataHelper` classes are defined in files in the **Helpers** folder. The `OneDriveItem` class is defined in a file in the **Models** folder. Detailed discussion of these classes is not relevant to authorization or SSO, so it is out-of-scope for this article.</span></span>
    * <span data-ttu-id="9819e-396">O desempenho é aprimorado ao se solicitar ao Microsoft Graph apenas os dados que são realmente necessários. Desse modo, o código usa um parâmetro de consulta `$select` para especificar que desejamos somente a propriedade de nome, e usa um parâmetro `$top` para especificar que desejamos somente os três primeiros nomes de pasta ou de arquivo.</span><span class="sxs-lookup"><span data-stu-id="9819e-396">Performance is improved by asking Microsoft Graph for only the data actually needed, so the code uses a `$select` query parameter to specify that we only want the name property, and a `$top` parameter to specify that we want only the first three folder or file names.</span></span>
    * <span data-ttu-id="9819e-p181">Se o token enviado para o Microsoft Graph for inválido, o Microsoft Graph enviará um erro "401 Não Autorizado" com o código "InvalidAuthenticationToken". Em seguida, o ASP.NET exibe um **RuntimeBinderException**. Isso também ocorre quando o token expira, embora o MSAL deva impedir que isso aconteça.</span><span class="sxs-lookup"><span data-stu-id="9819e-p181">If the token sent to Microsoft Graph is invalid, Microsoft Graph sends a "401 Unauthorized" error with the code "InvalidAuthenticationToken". ASP.NET then throws a **RuntimeBinderException**. This is also what happens when the token is expired, although MSAL should prevent that from ever happening.</span></span> 

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

1. <span data-ttu-id="9819e-p182">Substitua `TODO5` pelo seguinte. Observação sobre este código:</span><span class="sxs-lookup"><span data-stu-id="9819e-p182">Replace `TODO5` with the following. Note about this code:</span></span>

    * <span data-ttu-id="9819e-p183">Embora o código acima solicite somente a propriedade *name* dos itens do OneDrive, o Microsoft Graph sempre inclui a propriedade *eTag* para os itens do OneDrive. Para reduzir a carga enviada para o cliente, o código a seguir reconstrói os resultados apenas com os nomes dos itens.</span><span class="sxs-lookup"><span data-stu-id="9819e-p183">Although the code above asked for only the *name* property of the OneDrive items, Microsoft Graph always includes the *eTag* property for OneDrive items. To reduce the payload sent to the client, the code below reconstructs the results with only the item names.</span></span>
    * <span data-ttu-id="9819e-404">A lista de três pastas e arquivos do OneDrive é enviada para o cliente como uma resposta HTTP "200 OK".</span><span class="sxs-lookup"><span data-stu-id="9819e-404">The list of three OneDrive files and folders is sent to the client as a "200 OK" HTTP Response.</span></span>

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

1. <span data-ttu-id="9819e-p184">Abaixo do método Get, adicione o método a seguir. Sobre este código, observe:</span><span class="sxs-lookup"><span data-stu-id="9819e-p184">Below the Get method, add the following method. About this code note:</span></span>  

    * <span data-ttu-id="9819e-407">O método transmite ao cliente informações sobre uma exceção do servidor.</span><span class="sxs-lookup"><span data-stu-id="9819e-407">The method relays to the client information about a server-side exception.</span></span>
    * <span data-ttu-id="9819e-408">Se a exceção original for passada para o método, o construtor HttpError incluirá informações do objeto de exceção em uma propriedade **ExceptionMessage**.</span><span class="sxs-lookup"><span data-stu-id="9819e-408">If the original exception is passed to the method, then the HttpError constructor will include information from the exception object in an **ExceptionMessage** property.</span></span>  
    * <span data-ttu-id="9819e-409">Se `null` for passado para a exceção, o construtor HttpError incluirá o parâmetro de mensagem em uma propriedade **Message** e não haverá uma propriedade **ExceptionMessage**.</span><span class="sxs-lookup"><span data-stu-id="9819e-409">If `null` is passed for the exception, then the HttpError constructor will include the message parameter in a **Message** property and there is no **ExceptionMessage** property.</span></span>

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

## <a name="run-the-add-in"></a><span data-ttu-id="9819e-410">Execute o suplemento</span><span class="sxs-lookup"><span data-stu-id="9819e-410">Run the add-in</span></span>

1. <span data-ttu-id="9819e-411">Certifique-se de ter alguns arquivos no seu OneDrive para que você possa verificar os resultados.</span><span class="sxs-lookup"><span data-stu-id="9819e-411">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

1. <span data-ttu-id="9819e-p185">No Visual Studio, pressione F5. O PowerPoint será aberto e haverá um grupo **SSO ASP.NET** na faixa de opções **Página Inicial**.</span><span class="sxs-lookup"><span data-stu-id="9819e-p185">In Visual Studio, press F5. PowerPoint opens and there is an **SSO ASP.NET** group on the **Home** ribbon.</span></span>

1. <span data-ttu-id="9819e-414">Pressione o botão **Mostrar Suplemento** nesse grupo para ver a interface do usuário do suplemento no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="9819e-414">Press the **Show Add-in** button in this group to see the add-in’s UI in the task pane.</span></span>

1. <span data-ttu-id="9819e-p186">Pressione o botão **Obter meus arquivos do OneDrive**. Se você não estiver conectado ao Office, você será solicitado a entrar.</span><span class="sxs-lookup"><span data-stu-id="9819e-p186">Press the button **Get My Files from OneDrive**. If you are not signed into Office, you'll be prompted to sign in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="9819e-p187">Se você entrou no Office com uma ID diferente e se alguns aplicativos do Office que estavam abertos no momento continuam abertos, o Office pode não alterar de forma confiável sua ID, mesmo que pareça ter feito isso no PowerPoint. Se isso acontecer, a chamada para o Microsoft Graph pode falhar ou os dados da ID anterior podem ser retornados. Para evitar isso, certifique-se de *fechar todos os outros aplicativos do Office* antes de pressionar **Obter meus arquivos do OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="9819e-p187">If you were previously signed on to Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so in PowerPoint. If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned. To prevent this, be sure to *close all other Office applications* before you press **Get My Files from OneDrive**.</span></span>

1. <span data-ttu-id="9819e-p188">Depois de entrar, será exibida uma lista de seus arquivos e suas pastas no OneDrive, abaixo do botão. Esse procedimento pode levar mais de 15 segundos, principalmente na primeira vez.</span><span class="sxs-lookup"><span data-stu-id="9819e-p188">After you are signed in, a list of your files and folders on OneDrive will appear below the button. This may take over 15 seconds, especially the first time.</span></span>
