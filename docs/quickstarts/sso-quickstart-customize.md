---
title: Personalizar o suplemento habilitado para SSO do Node.js.
description: Saiba mais sobre como personalizar o suplemento habilitado para SSO que você criou com o gerador Yeoman.
ms.date: 07/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 8d4ab4653bcd8fccdcbecbe2e7e8e29a49a3184f
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159448"
---
# <a name="customize-your-nodejs-sso-enabled-add-in"></a><span data-ttu-id="6c2ff-103">Personalizar o suplemento habilitado para SSO do Node.js.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-103">Customize your Node.js SSO-enabled add-in</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6c2ff-104">Este artigo cria o suplemento habilitado para SSO que é criado ao concluir o [início rápido de logon único (SSO)](sso-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="6c2ff-104">This article builds upon the SSO-enabled add-in that's created by completing the [single sign-on (SSO) quick start](sso-quickstart.md).</span></span> <span data-ttu-id="6c2ff-105">Conclua o início rápido antes de ler este artigo.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-105">Please complete the quick start before reading this article.</span></span>

<span data-ttu-id="6c2ff-106">O [início rápido do SSO](sso-quickstart.md) cria um suplemento habilitado para sso que obtém as informações de perfil do usuário conectado e as grava no documento ou na mensagem.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-106">The [SSO quick start](sso-quickstart.md) creates an SSO-enabled add-in that gets the signed-in user's profile information and writes it to the document or message.</span></span> <span data-ttu-id="6c2ff-107">Neste artigo, você passará pelo processo de atualização do suplemento que você criou com o gerador Yeoman no início rápido do SSO, para Adicionar nova funcionalidade que exija permissões diferentes.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-107">In this article, you'll walk through the process of updating the add-in that you created with the Yeoman generator in the SSO quick start, to add new functionality that requires different permissions.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="6c2ff-108">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="6c2ff-108">Prerequisites</span></span>

* <span data-ttu-id="6c2ff-109">Um suplemento do Office que você criou seguindo as instruções no [início rápido de SSO](sso-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="6c2ff-109">An Office Add-in that you created by following the instructions in the [SSO quick start](sso-quickstart.md).</span></span>

* <span data-ttu-id="6c2ff-110">Pelo menos alguns arquivos e pastas armazenados no OneDrive for Business em sua assinatura do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-110">At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.</span></span>

* <span data-ttu-id="6c2ff-111">[Node.js](https://nodejs.org) (a versão mais recente de [LTS](https://nodejs.org/about/releases)).</span><span class="sxs-lookup"><span data-stu-id="6c2ff-111">[Node.js](https://nodejs.org) (the latest [LTS](https://nodejs.org/about/releases) version).</span></span>

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="review-contents-of-the-project"></a><span data-ttu-id="6c2ff-112">Revisar o conteúdo do projeto</span><span class="sxs-lookup"><span data-stu-id="6c2ff-112">Review contents of the project</span></span>

<span data-ttu-id="6c2ff-113">Vamos começar com uma revisão rápida do projeto de suplemento [criado anteriormente com o gerador Yeoman](sso-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="6c2ff-113">Let's begin with a quick review of the add-in project that you previously [created with the Yeoman generator](sso-quickstart.md).</span></span>

> [!NOTE]
> <span data-ttu-id="6c2ff-114">Em lugares onde este artigo faz referência a arquivos de script usando a extensão de arquivo **. js** , considere a extensão de arquivo **. TS** , em vez disso, se o projeto foi criado com TypeScript.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-114">In places where this article references script files using **.js** file extension, assume the **.ts** file extension instead if your project was created with TypeScript.</span></span>

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="add-new-functionality"></a><span data-ttu-id="6c2ff-115">Adicionar nova funcionalidade</span><span class="sxs-lookup"><span data-stu-id="6c2ff-115">Add new functionality</span></span>

<span data-ttu-id="6c2ff-116">O suplemento que você criou com o início rápido do SSO usa o Microsoft Graph para obter as informações de perfil do usuário conectado e grava essas informações no documento ou na mensagem.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-116">The add-in that you created with the SSO quick start uses Microsoft Graph to get the signed-in user's profile information and writes that information to the document or message.</span></span> <span data-ttu-id="6c2ff-117">Vamos alterar a funcionalidade do suplemento de forma que ele obtenha os nomes dos dez principais arquivos e pastas do OneDrive for Business do usuário conectado e grava essas informações no documento ou na mensagem.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-117">Let's change the add-in's functionality such that it gets the names of the top 10 files and folders from the signed-in user's OneDrive for Business and writes that information to the document or message.</span></span> <span data-ttu-id="6c2ff-118">Habilitar essa nova funcionalidade requer a atualização das permissões do aplicativo no Azure e a atualização do código no projeto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-118">Enabling this new functionality requires updating app permissions in Azure and updating code within the add-in project.</span></span>

### <a name="update-app-permissions-in-azure"></a><span data-ttu-id="6c2ff-119">Atualizar permissões do aplicativo no Azure</span><span class="sxs-lookup"><span data-stu-id="6c2ff-119">Update app permissions in Azure</span></span>

<span data-ttu-id="6c2ff-120">Antes que o suplemento possa ler com êxito o conteúdo do OneDrive for Business do usuário, suas informações de registro de aplicativo no Azure devem ser atualizadas com as permissões apropriadas.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-120">Before the add-in can successfully read the contents of the user's OneDrive for Business, its app registration information in Azure must be updated with the appropriate permissions.</span></span> <span data-ttu-id="6c2ff-121">Complete as etapas a seguir para conceder ao aplicativo a permissão **files. Read. All** e revogar a permissão **User. Read** , que não é mais necessária.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-121">Complete the following steps to grant the app the **Files.Read.All** permission and revoke the **User.Read** permission, which is no longer needed.</span></span>

1. <span data-ttu-id="6c2ff-122">Navegue até o [portal do Azure](https://ms.portal.azure.com/#home) e **entre usando suas credenciais de administrador do Microsoft 365**.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-122">Navigate to the [Azure portal](https://ms.portal.azure.com/#home) and **sign in using your Microsoft 365 administrator credentials**.</span></span>

2. <span data-ttu-id="6c2ff-123">Navegue até a página **registros de aplicativos** .</span><span class="sxs-lookup"><span data-stu-id="6c2ff-123">Navigate to the **App registrations** page.</span></span>
    > [!TIP]
    > <span data-ttu-id="6c2ff-124">Você pode fazer isso escolhendo o bloco de **registros do aplicativo** na home page do Azure ou usando a caixa de pesquisa na home page para localizar e escolher registros de **aplicativos**.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-124">You can do this either by choosing the **App registrations** tile on the Azure home page or by using the search box on the home page to find and choose **App registrations**.</span></span>

3. <span data-ttu-id="6c2ff-125">Na página **registros de aplicativos** , escolha o aplicativo que você criou durante o início rápido.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-125">On the **App registrations** page, choose the app that you created during the quick start.</span></span> 
    > [!TIP]
    > <span data-ttu-id="6c2ff-126">O **nome de exibição** do aplicativo corresponderá ao nome do suplemento que você especificou ao criar o projeto com o gerador Yeoman.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-126">The **Display name** of the app will match the add-in name that you specified when you created the project with the Yeoman generator.</span></span>

4. <span data-ttu-id="6c2ff-127">Na página Visão geral do aplicativo, escolha **permissões da API** sob o título **gerenciar** no lado esquerdo da página.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-127">From the app overview page, choose **API permissions** under the **Manage** heading on the left side of the page.</span></span>

5. <span data-ttu-id="6c2ff-128">Na linha **User. Read** da tabela de permissões, escolha as reticências e, em seguida, selecione **revogar consentimento do administrador** no menu exibido.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-128">In the **User.Read** row of the permissions table, choose the ellipsis and then select **Revoke admin consent** from the menu that appears.</span></span>

6. <span data-ttu-id="6c2ff-129">Selecione o botão **Sim, remover** em resposta ao prompt que é exibido.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-129">Select the **Yes, remove** button in response to the prompt that's displayed.</span></span>

7. <span data-ttu-id="6c2ff-130">Na linha **User. Read** da tabela Permissions, escolha as reticências e, em seguida, selecione **remover permissão** no menu exibido.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-130">In the **User.Read** row of the permissions table, choose the ellipsis and then select **Remove permission** from the menu that appears.</span></span>

8. <span data-ttu-id="6c2ff-131">Selecione o botão **Sim, remover** em resposta ao prompt que é exibido.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-131">Select the **Yes, remove** button in response to the prompt that's displayed.</span></span>

9. <span data-ttu-id="6c2ff-132">Selecione o botão **Adicionar uma permissão** .</span><span class="sxs-lookup"><span data-stu-id="6c2ff-132">Select the **Add a permission** button.</span></span>

10. <span data-ttu-id="6c2ff-133">No painel que é aberto, escolha **Microsoft Graph** e, em seguida, escolha **permissões delegadas**.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-133">On the panel that opens choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

11. <span data-ttu-id="6c2ff-134">No painel **solicitar permissões de API** :</span><span class="sxs-lookup"><span data-stu-id="6c2ff-134">On the **Request API permissions** panel:</span></span>

    <span data-ttu-id="6c2ff-135">a.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-135">a.</span></span> <span data-ttu-id="6c2ff-136">Em **arquivos**, selecione **arquivos. Read. All**.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-136">Under **Files**, select **Files.Read.All**.</span></span>

    <span data-ttu-id="6c2ff-137">b.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-137">b.</span></span> <span data-ttu-id="6c2ff-138">Selecione o botão **adicionar permissões** na parte inferior do painel para salvar essas alterações de permissões.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-138">Select the **Add permissions** button at the bottom of the panel to save these permissions changes.</span></span>

12. <span data-ttu-id="6c2ff-139">Selecione o botão **conceder consentimento de administrador para [nome do locatário]** .</span><span class="sxs-lookup"><span data-stu-id="6c2ff-139">Select the **Grant admin consent for [tenant name]** button.</span></span>

13. <span data-ttu-id="6c2ff-140">Selecione o botão **Sim** em resposta ao prompt exibido.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-140">Select the **Yes** button in response to the prompt that's displayed.</span></span>

### <a name="update-code-in-the-add-in-project"></a><span data-ttu-id="6c2ff-141">Atualizar código no projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="6c2ff-141">Update code in the add-in project</span></span>

<span data-ttu-id="6c2ff-142">Para permitir que o suplemento Leia o conteúdo do OneDrive for Business do usuário conectado, você precisará:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-142">To enable the add-in to read contents of the signed-in user's OneDrive for Business, you'll need to:</span></span>

- <span data-ttu-id="6c2ff-143">Atualize o código que faz referência à URL, aos parâmetros e ao escopo de acesso necessários do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-143">Update the code that references the Microsoft Graph URL, parameters, and required access scope.</span></span>

- <span data-ttu-id="6c2ff-144">Atualize o código que define a interface do usuário do painel de tarefas, para que ele descreva precisamente a nova funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-144">Update the code that defines the task pane UI, so that it accurately describes the new functionality.</span></span> 

- <span data-ttu-id="6c2ff-145">Atualize o código que analisa a resposta do Microsoft Graph e o grava no documento ou na mensagem.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-145">Update the code that parses the response from Microsoft Graph and writes it to the document or message.</span></span>

<span data-ttu-id="6c2ff-146">As etapas a seguir descrevem essas atualizações.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-146">The following steps describe these updates.</span></span>

### <a name="changes-required-for-any-type-of-add-in"></a><span data-ttu-id="6c2ff-147">Alterações necessárias para qualquer tipo de suplemento</span><span class="sxs-lookup"><span data-stu-id="6c2ff-147">Changes required for any type of add-in</span></span>

<span data-ttu-id="6c2ff-148">Conclua as seguintes etapas para o seu suplemento, para alterar a URL, os parâmetros e o escopo de acesso do Microsoft Graph e atualizar a interface do usuário do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-148">Complete the following steps for your add-in, to change the Microsoft Graph URL, parameters, and access scope, and update the taskpane UI.</span></span> <span data-ttu-id="6c2ff-149">Essas etapas são as mesmas, independentemente de qual host do Office seu suplemento está direcionado.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-149">These steps are the same, regardless of which Office host your add-in targets.</span></span>

1. <span data-ttu-id="6c2ff-150">Na **./. ENV** arquivo:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-150">In the **./.ENV** file:</span></span>

    <span data-ttu-id="6c2ff-151">a.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-151">a.</span></span> <span data-ttu-id="6c2ff-152">Substitua `GRAPH_URL_SEGMENT=/me` pelo seguinte:`GRAPH_URL_SEGMENT=/me/drive/root/children`</span><span class="sxs-lookup"><span data-stu-id="6c2ff-152">Replace `GRAPH_URL_SEGMENT=/me` with the following: `GRAPH_URL_SEGMENT=/me/drive/root/children`</span></span>

    <span data-ttu-id="6c2ff-153">b.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-153">b.</span></span> <span data-ttu-id="6c2ff-154">Substitua `QUERY_PARAM_SEGMENT=` pelo seguinte:`QUERY_PARAM_SEGMENT=?$select=name&$top=10`</span><span class="sxs-lookup"><span data-stu-id="6c2ff-154">Replace `QUERY_PARAM_SEGMENT=` with the following: `QUERY_PARAM_SEGMENT=?$select=name&$top=10`</span></span>

    <span data-ttu-id="6c2ff-155">c.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-155">c.</span></span> <span data-ttu-id="6c2ff-156">Substitua `SCOPE=User.Read` pelo seguinte:`SCOPE=Files.Read.All`</span><span class="sxs-lookup"><span data-stu-id="6c2ff-156">Replace `SCOPE=User.Read` with the following: `SCOPE=Files.Read.All`</span></span>

2. <span data-ttu-id="6c2ff-157">Em **./manifest.xml**, localize a linha `<Scope>User.Read</Scope>` próxima ao final do arquivo e substitua-a pela linha `<Scope>Files.Read.All</Scope>` .</span><span class="sxs-lookup"><span data-stu-id="6c2ff-157">In **./manifest.xml**, find the line `<Scope>User.Read</Scope>` near the end of the file and replace it with the line `<Scope>Files.Read.All</Scope>`.</span></span>

3. <span data-ttu-id="6c2ff-158">Em **./src/helpers/fallbackauthdialog.js** (ou em **./src/Helpers/fallbackauthdialog.TS** para um projeto TypeScript), localize a cadeia de caracteres `https://graph.microsoft.com/User.Read` e substitua-a pela cadeia de caracteres `https://graph.microsoft.com/Files.Read.All` , tal como `requestObj` é definida da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-158">In **./src/helpers/fallbackauthdialog.js** (or in **./src/helpers/fallbackauthdialog.ts** for a TypeScript project), find the string `https://graph.microsoft.com/User.Read` and replace it with the string `https://graph.microsoft.com/Files.Read.All`, such that `requestObj` is defined as follows:</span></span>

    ```javascript
    var requestObj = {
      scopes: [`https://graph.microsoft.com/Files.Read.All`]
    };
    ```

    ```typescript
    var requestObj: Object = {
      scopes: [`https://graph.microsoft.com/Files.Read.All`]
    };
    ```

4. <span data-ttu-id="6c2ff-159">Em **./src/taskpane/taskpane.html**, encontre o elemento `<section class="ms-firstrun-instructionstep__header">` e atualize o texto dentro desse elemento para descrever a nova funcionalidade do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-159">In **./src/taskpane/taskpane.html**, find the element `<section class="ms-firstrun-instructionstep__header">` and update the text within that element to describe the add-in's new functionality.</span></span>

    ```html
    <section class="ms-firstrun-instructionstep__header">
        <h2 class="ms-font-m">This add-in demonstrates how to use single sign-on by making a call to Microsoft
            Graph to read content from OneDrive for Business.</h2>
        <div class="ms-firstrun-instructionstep__header--image"></div>
    </section>
    ```

5. <span data-ttu-id="6c2ff-160">Em **./src/taskpane/taskpane.html**, localize e substitua as duas ocorrências da cadeia de caracteres `Get My User Profile Information` pela cadeia de caracteres `Read my OneDrive for Business` .</span><span class="sxs-lookup"><span data-stu-id="6c2ff-160">In **./src/taskpane/taskpane.html**, find and replace both occurrences of the string `Get My User Profile Information` with the string `Read my OneDrive for Business`.</span></span>

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">Click the <b>Read my OneDrive for Business</b>
            button.</span>
        <div class="clearfix"></div>
    </li>
    ```

    ```html
    <p align="center">
        <button id="getGraphDataButton" class="popupButton ms-Button ms-Button--primary"><span
                class="ms-Button-label">Read my OneDrive for Business</span></button>
    </p>
    ```

6. <span data-ttu-id="6c2ff-161">Em **./src/taskpane/taskpane.html**, localize e substitua a cadeia de caracteres `Your user profile information will be displayed in the document.` pela cadeia de caracteres `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.` .</span><span class="sxs-lookup"><span data-stu-id="6c2ff-161">In **./src/taskpane/taskpane.html**, find and replace the string `Your user profile information will be displayed in the document.` with the string `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.`.</span></span>

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.</span>
        <div class="clearfix"></div>
    </li>
    ```

7. <span data-ttu-id="6c2ff-162">Atualize o código que analisa a resposta do Microsoft Graph e o grava no documento ou na mensagem, seguindo as orientações na seção que corresponde ao seu tipo de suplemento:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-162">Update the code that parses the response from Microsoft Graph and writes it to the document or message, by following guidance in the section that corresponds to your type of add-in:</span></span>

    - [<span data-ttu-id="6c2ff-163">Alterações necessárias para um suplemento do Excel (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-163">Changes required for an Excel add-in (JavaScript)</span></span>](#changes-required-for-an-excel-add-in-javascript)
    - [<span data-ttu-id="6c2ff-164">Alterações necessárias para um suplemento do Excel (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-164">Changes required for an Excel add-in (TypeScript)</span></span>](#changes-required-for-an-excel-add-in-typescript)
    - [<span data-ttu-id="6c2ff-165">Alterações necessárias para um suplemento do Outlook (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-165">Changes required for an Outlook add-in (JavaScript)</span></span>](#changes-required-for-an-outlook-add-in-javascript)
    - [<span data-ttu-id="6c2ff-166">Alterações necessárias para um suplemento do Outlook (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-166">Changes required for an Outlook add-in (TypeScript)</span></span>](#changes-required-for-an-outlook-add-in-typescript)
    - [<span data-ttu-id="6c2ff-167">Alterações necessárias para um suplemento do PowerPoint (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-167">Changes required for a PowerPoint add-in (JavaScript)</span></span>](#changes-required-for-a-powerpoint-add-in-javascript)
    - [<span data-ttu-id="6c2ff-168">Alterações necessárias para um suplemento do PowerPoint (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-168">Changes required for a PowerPoint add-in (TypeScript)</span></span>](#changes-required-for-a-powerpoint-add-in-typescript)
    - [<span data-ttu-id="6c2ff-169">Alterações necessárias para um suplemento do Word (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-169">Changes required for a Word add-in (JavaScript)</span></span>](#changes-required-for-a-word-add-in-javascript)
    - [<span data-ttu-id="6c2ff-170">Alterações necessárias para um suplemento do Word (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-170">Changes required for a Word add-in (TypeScript)</span></span>](#changes-required-for-a-word-add-in-typescript)

### <a name="changes-required-for-an-excel-add-in-javascript"></a><span data-ttu-id="6c2ff-171">Alterações necessárias para um suplemento do Excel (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-171">Changes required for an Excel add-in (JavaScript)</span></span>

<span data-ttu-id="6c2ff-172">Se o suplemento for um suplemento do Excel que foi criado com JavaScript, faça as seguintes alterações em **./src/helpers/documentHelper.js**:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-172">If your add-in is an Excel add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="6c2ff-173">Encontre a `writeDataToOfficeDocument` função e substitua-a pela seguinte função:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-173">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToExcel(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. <span data-ttu-id="6c2ff-174">Encontre a `filterUserProfileInfo` função e substitua-a pela seguinte função:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-174">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. <span data-ttu-id="6c2ff-175">Encontre a `writeDataToExcel` função e substitua-a pela seguinte função:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-175">Find the `writeDataToExcel` function and replace it with the following function:</span></span>

    ```javascript
    function writeDataToExcel(result) {
      return Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        let data = [];
        let oneDriveInfo = filterOneDriveInfo(result);

        for (let i = 0; i < oneDriveInfo.length; i++) {
          if (oneDriveInfo[i] !== null) {
            let innerArray = [];
            innerArray.push(oneDriveInfo[i]);
            data.push(innerArray);
          }
        }

        const rangeAddress = `B5:B${5 + (data.length - 1)}`;
        const range = sheet.getRange(rangeAddress);
        range.values = data;
        range.format.autofitColumns();

        return context.sync();
      });
    }
    ```

4. <span data-ttu-id="6c2ff-176">Exclua a `writeDataToOutlook` função.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-176">Delete the `writeDataToOutlook` function.</span></span>

5. <span data-ttu-id="6c2ff-177">Exclua a `writeDataToPowerPoint` função.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-177">Delete the `writeDataToPowerPoint` function.</span></span>

6. <span data-ttu-id="6c2ff-178">Exclua a `writeDataToWord` função.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-178">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="6c2ff-179">Depois de fazer essas alterações, pule para a seção [Experimente](#try-it-out) , deste artigo, para experimentar o suplemento atualizado.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-179">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-excel-add-in-typescript"></a><span data-ttu-id="6c2ff-180">Alterações necessárias para um suplemento do Excel (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-180">Changes required for an Excel add-in (TypeScript)</span></span>

<span data-ttu-id="6c2ff-181">Se o suplemento for um suplemento do Excel que foi criado com TypeScript, abra **./src/TaskPane/TaskPane.TS**, localize a `writeDataToOfficeDocument` função e substitua-a pela seguinte função:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-181">If your add-in is an Excel add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Excel.run(function(context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
      itemNames.push(item["name"]);
    }

    for (let i = 0; i < itemNames.length; i++) {
      if (itemNames[i] !== null) {
        let innerArray = [];
        innerArray.push(itemNames[i]);
        data.push(innerArray);
      }
    }
    
    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}
```

<span data-ttu-id="6c2ff-182">Depois de fazer essas alterações, pule para a seção [Experimente](#try-it-out) , deste artigo, para experimentar o suplemento atualizado.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-182">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-outlook-add-in-javascript"></a><span data-ttu-id="6c2ff-183">Alterações necessárias para um suplemento do Outlook (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-183">Changes required for an Outlook add-in (JavaScript)</span></span>

<span data-ttu-id="6c2ff-184">Se o suplemento for um suplemento do Outlook que foi criado com JavaScript, faça as seguintes alterações em **./src/helpers/documentHelper.js**:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-184">If your add-in is an Outlook add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="6c2ff-185">Encontre a `writeDataToOfficeDocument` função e substitua-a pela seguinte função:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-185">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToOutlook(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to message. " + error.toString()));
        }
      });
    }
    ```

2. <span data-ttu-id="6c2ff-186">Encontre a `filterUserProfileInfo` função e substitua-a pela seguinte função:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-186">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. <span data-ttu-id="6c2ff-187">Encontre a `writeDataToOutlook` função e substitua-a pela seguinte função:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-187">Find the `writeDataToOutlook` function and replace it with the following function:</span></span>

    ```javascript
    function writeDataToOutlook(result) {
      let data = [];
      let oneDriveInfo = filterOneDriveInfo(result);

      for (let i = 0; i < oneDriveInfo.length; i++) {
        if (oneDriveInfo[i] !== null) {
          data.push(oneDriveInfo[i]);
        }
      }

      let objectNames = "";
      for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "<br/>";
      }

      Office.context.mailbox.item.body.setSelectedDataAsync(objectNames, { coercionType: Office.CoercionType.Html });
    }
    ```

4. <span data-ttu-id="6c2ff-188">Exclua a `writeDataToExcel` função.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-188">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="6c2ff-189">Exclua a `writeDataToPowerPoint` função.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-189">Delete the `writeDataToPowerPoint` function.</span></span>

6. <span data-ttu-id="6c2ff-190">Exclua a `writeDataToWord` função.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-190">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="6c2ff-191">Depois de fazer essas alterações, pule para a seção [Experimente](#try-it-out) , deste artigo, para experimentar o suplemento atualizado.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-191">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-an-outlook-add-in-typescript"></a><span data-ttu-id="6c2ff-192">Alterações necessárias para um suplemento do Outlook (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-192">Changes required for an Outlook add-in (TypeScript)</span></span>

<span data-ttu-id="6c2ff-193">Se o suplemento for um suplemento do Outlook que foi criado com TypeScript, abra **./src/TaskPane/TaskPane.TS**, localize a `writeDataToOfficeDocument` função e substitua-a pela seguinte função:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-193">If your add-in is an Outlook add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

```typescript
export function writeDataToOfficeDocument(result: Object): void {
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
        itemNames.push(item["name"]);
    };

    for (let i = 0; i < itemNames.length; i++) {
        if (itemNames[i] !== null) {
        data.push(itemNames[i]);
        }
    }

    let objectNames: string = "";
    for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "<br/>";
    }
    
    Office.context.mailbox.item.body.setSelectedDataAsync(objectNames, { coercionType: Office.CoercionType.Html });
}
```

<span data-ttu-id="6c2ff-194">Depois de fazer essas alterações, pule para a seção [Experimente](#try-it-out) , deste artigo, para experimentar o suplemento atualizado.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-194">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-powerpoint-add-in-javascript"></a><span data-ttu-id="6c2ff-195">Alterações necessárias para um suplemento do PowerPoint (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-195">Changes required for a PowerPoint add-in (JavaScript)</span></span>

<span data-ttu-id="6c2ff-196">Se o suplemento for um suplemento do PowerPoint que foi criado com JavaScript, faça as seguintes alterações em **./src/helpers/documentHelper.js**:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-196">If your add-in is a PowerPoint add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="6c2ff-197">Encontre a `writeDataToOfficeDocument` função e substitua-a pela seguinte função:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-197">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToPowerPoint(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. <span data-ttu-id="6c2ff-198">Encontre a `filterUserProfileInfo` função e substitua-a pela seguinte função:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-198">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. <span data-ttu-id="6c2ff-199">Encontre a `writeDataToPowerPoint` função e substitua-a pela seguinte função:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-199">Find the `writeDataToPowerPoint` function and replace it with the following function:</span></span>

    ```javascript
    function writeDataToPowerPoint(result) {
      let data = [];
      let oneDriveInfo = filterOneDriveInfo(result);

      for (let i = 0; i < oneDriveInfo.length; i++) {
        if (oneDriveInfo[i] !== null) {
          data.push(oneDriveInfo[i]);
        }
      }

      let objectNames = "";
      for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "\n";
      }

      Office.context.document.setSelectedDataAsync(
        objectNames, 
        function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            throw asyncResult.error.message;
          }
      });
    }
    ```

4. <span data-ttu-id="6c2ff-200">Exclua a `writeDataToExcel` função.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-200">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="6c2ff-201">Exclua a `writeDataToOutlook` função.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-201">Delete the `writeDataToOutlook` function.</span></span>

6. <span data-ttu-id="6c2ff-202">Exclua a `writeDataToWord` função.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-202">Delete the `writeDataToWord` function.</span></span>

<span data-ttu-id="6c2ff-203">Depois de fazer essas alterações, pule para a seção [Experimente](#try-it-out) , deste artigo, para experimentar o suplemento atualizado.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-203">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-powerpoint-add-in-typescript"></a><span data-ttu-id="6c2ff-204">Alterações necessárias para um suplemento do PowerPoint (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-204">Changes required for a PowerPoint add-in (TypeScript)</span></span>

<span data-ttu-id="6c2ff-205">Se o suplemento for um suplemento do PowerPoint que foi criado com TypeScript, abra **./src/TaskPane/TaskPane.TS**, localize a `writeDataToOfficeDocument` função e substitua-a pela seguinte função:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-205">If your add-in is a PowerPoint add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

```typescript
export function writeDataToOfficeDocument(result: Object): void {
  let data: string[] = [];

  let itemNames: string[] = [];
  let oneDriveItems = result["value"];
  for (let item of oneDriveItems) {
    itemNames.push(item["name"]);
  };

  for (let i = 0; i < itemNames.length; i++) {
    if (itemNames[i] !== null) {
      data.push(itemNames[i]);
    }
  }

  let objectNames: string = "";
  for (let i = 0; i < data.length; i++) {
    objectNames += data[i] + "\n";
  }

  Office.context.document.setSelectedDataAsync(objectNames, function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}
```

<span data-ttu-id="6c2ff-206">Depois de fazer essas alterações, pule para a seção [Experimente](#try-it-out) , deste artigo, para experimentar o suplemento atualizado.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-206">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-word-add-in-javascript"></a><span data-ttu-id="6c2ff-207">Alterações necessárias para um suplemento do Word (JavaScript)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-207">Changes required for a Word add-in (JavaScript)</span></span>

<span data-ttu-id="6c2ff-208">Se o suplemento for um suplemento do Word que foi criado com JavaScript, faça as seguintes alterações em **./src/helpers/documentHelper.js**:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-208">If your add-in is a Word add-in that was created with JavaScript, make the following changes in **./src/helpers/documentHelper.js**:</span></span>

1. <span data-ttu-id="6c2ff-209">Encontre a `writeDataToOfficeDocument` função e substitua-a pela seguinte função:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-209">Find the `writeDataToOfficeDocument` function and replace it with the following function:</span></span>

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToWord(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. <span data-ttu-id="6c2ff-210">Encontre a `filterUserProfileInfo` função e substitua-a pela seguinte função:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-210">Find the `filterUserProfileInfo` function and replace it with the following function:</span></span>

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

3. <span data-ttu-id="6c2ff-211">Encontre a `writeDataToWord` função e substitua-a pela seguinte função:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-211">Find the `writeDataToWord` function and replace it with the following function:</span></span>

    ```javascript
    function writeDataToWord(result) {
      return Word.run(function (context) {
        let data = [];
        let oneDriveInfo = filterOneDriveInfo(result);

        for (let i = 0; i < oneDriveInfo.length; i++) {
          if (oneDriveInfo[i] !== null) {
            data.push(oneDriveInfo[i]);
          }
        }

        const documentBody = context.document.body;
        for (let i = 0; i < data.length; i++) {
          if (data[i] !== null) {
            documentBody.insertParagraph(data[i], "End");
          }
        }

        return context.sync();
      });
    }
    ```

4. <span data-ttu-id="6c2ff-212">Exclua a `writeDataToExcel` função.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-212">Delete the `writeDataToExcel` function.</span></span>

5. <span data-ttu-id="6c2ff-213">Exclua a `writeDataToOutlook` função.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-213">Delete the `writeDataToOutlook` function.</span></span>

6. <span data-ttu-id="6c2ff-214">Exclua a `writeDataToPowerPoint` função.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-214">Delete the `writeDataToPowerPoint` function.</span></span>

<span data-ttu-id="6c2ff-215">Depois de fazer essas alterações, pule para a seção [Experimente](#try-it-out) , deste artigo, para experimentar o suplemento atualizado.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-215">After you've made these changes, skip ahead to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

### <a name="changes-required-for-a-word-add-in-typescript"></a><span data-ttu-id="6c2ff-216">Alterações necessárias para um suplemento do Word (TypeScript)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-216">Changes required for a Word add-in (TypeScript)</span></span>

<span data-ttu-id="6c2ff-217">Se o suplemento for um suplemento do Word que foi criado com TypeScript, abra **./src/TaskPane/TaskPane.TS**, localize a `writeDataToOfficeDocument` função e substitua-a pela seguinte função:</span><span class="sxs-lookup"><span data-stu-id="6c2ff-217">If your add-in is a Word add-in that was created with TypeScript, open **./src/taskpane/taskpane.ts**, find the `writeDataToOfficeDocument` function, and replace it with the following function:</span></span>

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Word.run(function(context) {
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
      itemNames.push(item["name"]);
    };

    for (let i = 0; i < itemNames.length; i++) {
      if (itemNames[i] !== null) {
        data.push(itemNames[i]);
      }
    }

    const documentBody: Word.Body = context.document.body;
    for (let i = 0; i < data.length; i++) {
      if (data[i] !== null) {
        documentBody.insertParagraph(data[i], "End");
      }
    }
    return context.sync();
  });
}
```

<span data-ttu-id="6c2ff-218">Depois de fazer essas alterações, continue na seção [Experimente](#try-it-out) deste artigo para experimentar o suplemento atualizado em seu site.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-218">After you've made these changes, continue to the [Try it out](#try-it-out) section of this article to try out your updated add-in.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="6c2ff-219">Experimente</span><span class="sxs-lookup"><span data-stu-id="6c2ff-219">Try it out</span></span>

<span data-ttu-id="6c2ff-220">Se o suplemento for um suplemento do Excel, Word ou PowerPoint, conclua as etapas da seção a seguir para experimentá-lo. Se o suplemento for um suplemento do Outlook, conclua as etapas na seção do [Outlook](#outlook) .</span><span class="sxs-lookup"><span data-stu-id="6c2ff-220">If your add-in is an Excel, Word, or PowerPoint add-in, complete the steps in the following section to try it out. If your add-in is an Outlook add-in, complete the steps in the [Outlook](#outlook) section instead.</span></span>

### <a name="excel-word-and-powerpoint"></a><span data-ttu-id="6c2ff-221">Excel, Word e PowerPoint</span><span class="sxs-lookup"><span data-stu-id="6c2ff-221">Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="6c2ff-222">Execute as etapas a seguir para experimentar um suplemento do Excel, do Word ou do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-222">Complete the following steps to try out an Excel, Word, or PowerPoint add-in.</span></span>

1. <span data-ttu-id="6c2ff-223">Na pasta raiz do projeto, execute o seguinte comando para compilar o projeto, inicie o servidor Web local e Sideload seu suplemento no aplicativo cliente do Office selecionado anteriormente.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-223">In the root folder of the project, run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6c2ff-224">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-224">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="6c2ff-225">Se for solicitado a instalação de um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-225">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="6c2ff-226">No aplicativo cliente do Office que é aberto quando você executa o comando anterior (ou seja, Excel, Word ou PowerPoint), certifique-se de que você está conectado com um usuário que seja membro da mesma organização do Microsoft 365 que a conta de administrador do Microsoft 365 que você usou para se conectar ao Azure durante a [configuração do SSO](sso-quickstart.md#configure-sso) para o aplicativo.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-226">In the Office client application that opens when you run the previous command (i.e., Excel, Word or PowerPoint), make sure that you're signed in with a user that's a member of the same Microsoft 365 organization as the Microsoft 365 administrator account that you used to connect to Azure while [configuring SSO](sso-quickstart.md#configure-sso) for the app.</span></span> <span data-ttu-id="6c2ff-227">Isso estabelecerá as condições apropriadas para que o SSO seja bem-sucedido.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-227">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="6c2ff-228">No aplicativo cliente do Office, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-228">In the Office client application, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="6c2ff-229">A imagem a seguir mostra esse botão no Excel.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-229">The following image shows this button in Excel.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="6c2ff-231">Na parte inferior do painel de tarefas, escolha o botão **ler meu onedrive for Business** para iniciar o processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-231">At the bottom of the task pane, choose the **Read my OneDrive for Business** button to initiate the SSO process.</span></span> 

5. <span data-ttu-id="6c2ff-232">Se uma janela de diálogo for exibida solicitando permissões em nome do suplemento, isso significa que não há suporte ao SSO no seu cenário e, em vez disso, o suplemento voltou para um método alternativo de autenticação do usuário.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-232">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="6c2ff-233">Isso pode ocorrer quando o administrador de locatários não concedeu consentimento para o suplemento para acessar o Microsoft Graph ou quando o usuário não está conectado ao Office com uma conta Microsoft válida ou uma conta de treinamento ou de trabalho do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-233">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft account or Microsoft 365 Education or Work account.</span></span> <span data-ttu-id="6c2ff-234">Escolha o botão **Aceitar** na janela de diálogo para continuar.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-234">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Caixa de diálogo Solicitação de permissões](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="6c2ff-236">Após um usuário aceitar a solicitação de permissões, elas não serão solicitadas novamente no futuro.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-236">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

6. <span data-ttu-id="6c2ff-237">O suplemento lê os dados do OneDrive for Business do usuário conectado e grava os nomes dos 10 arquivos e pastas principais no documento.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-237">The add-in reads data from the signed-in user's OneDrive for Business and writes the names of the top 10 files and folders to the document.</span></span> <span data-ttu-id="6c2ff-238">A imagem a seguir mostra um exemplo de nomes de arquivos e pastas gravados em uma planilha do Excel.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-238">The following image shows an example of file and folder names written to an Excel worksheet.</span></span>

    ![Informações sobre o OneDrive for Business na planilha do Excel](../images/sso-onedrive-info-excel.png)

### <a name="outlook"></a><span data-ttu-id="6c2ff-240">Outlook</span><span class="sxs-lookup"><span data-stu-id="6c2ff-240">Outlook</span></span>

<span data-ttu-id="6c2ff-241">Execute as etapas a seguir para experimentar um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-241">Complete the following steps to try out an Outlook add-in.</span></span>

1. <span data-ttu-id="6c2ff-242">Na pasta raiz do projeto, execute o seguinte comando para compilar o projeto e iniciar o servidor Web local.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-242">In the root folder of the project, run the following command to build the project and start the local web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6c2ff-243">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-243">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="6c2ff-244">Se for solicitado a instalação de um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-244">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="6c2ff-245">Siga as instruções [Realizar sideload dos suplementos do Outlook para teste](/outlook/add-ins/sideload-outlook-add-ins-for-testing)para realizar o sideload do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-245">Follow the instructions in [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing) to sideload the add-in in Outlook.</span></span> <span data-ttu-id="6c2ff-246">Certifique-se de que você está conectado ao Outlook com um usuário que é membro da mesma organização do Microsoft 365 que a conta de administrador do Microsoft 365 que você usou para se conectar ao Azure durante a [configuração do SSO](sso-quickstart.md#configure-sso) para o aplicativo.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-246">Make sure that you're signed in to Outlook with a user that's a member of the same Microsoft 365 organization as the Microsoft 365 administrator account that you used to connect to Azure while [configuring SSO](sso-quickstart.md#configure-sso) for the app.</span></span> <span data-ttu-id="6c2ff-247">Isso estabelecerá as condições apropriadas para que o SSO seja bem-sucedido.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-247">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="6c2ff-248">Escreva uma nova mensagem no Outlook.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-248">In Outlook, compose a new message.</span></span>

4. <span data-ttu-id="6c2ff-249">Na janela redigir mensagem, escolha o botão **Exibir painel de tarefas** na faixa de opções para abrir o painel de tarefas de suplemento.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-249">In the message compose window, choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Outlook](../images/outlook-sso-ribbon-button.png)

5. <span data-ttu-id="6c2ff-251">Na parte inferior do painel de tarefas, escolha o botão **ler meu onedrive for Business** para iniciar o processo de SSO.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-251">At the bottom of the task pane, choose the **Read my OneDrive for Business** button to initiate the SSO process.</span></span> 

6. <span data-ttu-id="6c2ff-252">Se uma janela de diálogo for exibida solicitando permissões em nome do suplemento, isso significa que não há suporte ao SSO no seu cenário e, em vez disso, o suplemento voltou para um método alternativo de autenticação do usuário.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-252">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="6c2ff-253">Isso pode ocorrer quando o administrador de locatários não concedeu consentimento para o suplemento para acessar o Microsoft Graph ou quando o usuário não está conectado ao Office com uma conta Microsoft válida ou uma conta de treinamento ou de trabalho do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-253">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft account or Microsoft 365 Education or Work account.</span></span> <span data-ttu-id="6c2ff-254">Escolha o botão **Aceitar** na janela de diálogo para continuar.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-254">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Caixa de diálogo Solicitação de permissões](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="6c2ff-256">Após um usuário aceitar a solicitação de permissões, elas não serão solicitadas novamente no futuro.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-256">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

7. <span data-ttu-id="6c2ff-257">O suplemento lê os dados do OneDrive for Business do usuário conectado e grava os nomes dos 10 arquivos e pastas principais no corpo da mensagem de email.</span><span class="sxs-lookup"><span data-stu-id="6c2ff-257">The add-in reads data from the signed-in user's OneDrive for Business and writes the names of the top 10 files and folders to the body of the email message.</span></span>

    ![Informações sobre o OneDrive for Business na mensagem do Outlook](../images/sso-onedrive-info-outlook.png)

## <a name="next-steps"></a><span data-ttu-id="6c2ff-259">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="6c2ff-259">Next steps</span></span>

<span data-ttu-id="6c2ff-260">Parabéns, você personalizou com êxito a funcionalidade do suplemento habilitado para SSO que você criou com o gerador Yeoman no [início rápido de SSO](sso-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="6c2ff-260">Congratulations, you've successfully customized the functionality of the SSO-enabled add-in that you created with the Yeoman generator in the [SSO quick start](sso-quickstart.md).</span></span> <span data-ttu-id="6c2ff-261">Para saber mais sobre as etapas de configuração do SSO que o gerador Yeoman concluiu automaticamente e o código que facilita o processo de SSO, confira o tutorial [Criar um Suplemento do Office com Node.js que usa logon único](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="6c2ff-261">To learn more about SSO configuration steps that the Yeoman generator completed automatically, and the code that facilitates the SSO process, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="see-also"></a><span data-ttu-id="6c2ff-262">Confira também</span><span class="sxs-lookup"><span data-stu-id="6c2ff-262">See also</span></span>

- [<span data-ttu-id="6c2ff-263">Habilitar o logon único para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6c2ff-263">Enable single sign-on for Office Add-ins</span></span>](../develop/sso-in-office-add-ins.md)
- [<span data-ttu-id="6c2ff-264">Início rápido logon único (SSO).</span><span class="sxs-lookup"><span data-stu-id="6c2ff-264">Single sign-on (SSO) quick start</span></span>](sso-quickstart.md)
- [<span data-ttu-id="6c2ff-265">Criar um Suplemento do Office com Node.js que usa logon único</span><span class="sxs-lookup"><span data-stu-id="6c2ff-265">Create a Node.js Office Add-in that uses single sign-on</span></span>](../develop/create-sso-office-add-ins-nodejs.md)
- [<span data-ttu-id="6c2ff-266">Solucionar problemas de mensagens de erro no logon único (SSO)</span><span class="sxs-lookup"><span data-stu-id="6c2ff-266">Troubleshoot error messages for single sign-on (SSO)</span></span>](../develop/troubleshoot-sso-in-office-add-ins.md)
