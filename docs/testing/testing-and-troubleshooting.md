---
title: Solucionar erros de usuários com suplementos do Office
description: Saiba como solucionar erros do usuário em Office de complementos.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: dc186e942d129d4a7ae1ce2a326d0e5a0e1629e1
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348626"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="e2e23-103">Solucionar erros de usuários com suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e2e23-103">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="e2e23-p101">Às vezes, seus usuários podem encontrar problemas com suplementos do Office desenvolvidos por você. Por exemplo, um suplemento falha ao carregar ou está inacessível. Use as informações neste artigo para ajudar a resolver problemas comuns que os usuários têm com o seu suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="e2e23-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="e2e23-107">Também é possível usar o [Fiddler](https://www.telerik.com/fiddler) para identificar e depurar problemas com os suplementos.</span><span class="sxs-lookup"><span data-stu-id="e2e23-107">You can also use [Fiddler](https://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="e2e23-108">Erros comuns e etapas de solução de problemas</span><span class="sxs-lookup"><span data-stu-id="e2e23-108">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="e2e23-109">A tabela a seguir lista as mensagens de erro comuns que os usuários podem receber e as etapas que os usuários podem seguir para resolver os erros.</span><span class="sxs-lookup"><span data-stu-id="e2e23-109">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="e2e23-110">**Mensagem de erro**</span><span class="sxs-lookup"><span data-stu-id="e2e23-110">**Error message**</span></span>|<span data-ttu-id="e2e23-111">**Resolução**</span><span class="sxs-lookup"><span data-stu-id="e2e23-111">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="e2e23-112">Erro do aplicativo: catálogo não pôde ser alcançado</span><span class="sxs-lookup"><span data-stu-id="e2e23-112">App error: Catalog could not be reached</span></span>|<span data-ttu-id="e2e23-p102">Verifique as configurações do firewall. "Catálogo" refere-se ao AppSource. Essa mensagem indica que o usuário não consegue acessar o AppSource.</span><span class="sxs-lookup"><span data-stu-id="e2e23-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="e2e23-p103">ERRO DO APLICATIVO: este aplicativo não pôde ser iniciado. Feche essa caixa de diálogo para ignorar o problema ou clique em "Reiniciar"para tentar novamente.</span><span class="sxs-lookup"><span data-stu-id="e2e23-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="e2e23-117">Verifique se as atualizações mais recentes do Office foram instaladas ou baixe a [atualização do Office 2013](https://support.microsoft.com/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="e2e23-117">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>|
|<span data-ttu-id="e2e23-118">Erro: objeto não dá suporte à propriedade ou ao método 'defineProperty'</span><span class="sxs-lookup"><span data-stu-id="e2e23-118">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="e2e23-119">Confirme se o Internet Explorer não está sendo executado no modo de compatibilidade.</span><span class="sxs-lookup"><span data-stu-id="e2e23-119">Confirm that Internet Explorer is not running in Compatibility Mode.</span></span> <span data-ttu-id="e2e23-120">Vá para Ferramentas > **Modo de Exibição** de Compatibilidade Configurações .</span><span class="sxs-lookup"><span data-stu-id="e2e23-120">Go to Tools > **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="e2e23-p105">Não foi possível carregar o aplicativo porque não há suporte para sua versão do navegador. Clique aqui para obter uma lista de versões do navegador compatíveis.</span><span class="sxs-lookup"><span data-stu-id="e2e23-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="e2e23-p106">Verifique se o navegador dá suporte a armazenamento local HTML5 ou redefina as configurações do Internet Explorer. Para saber mais sobre os navegadores compatíveis, confira [Requisitos para a execução de Suplementos do Office](../concepts/requirements-for-running-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="e2e23-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a><span data-ttu-id="e2e23-125">Ao instalar um suplemento, você verá “erro ao carregar suplemento” na barra de status</span><span class="sxs-lookup"><span data-stu-id="e2e23-125">When installing an add-in, you see "Error loading add-in" in the status bar</span></span>

1. <span data-ttu-id="e2e23-126">Feche o Office.</span><span class="sxs-lookup"><span data-stu-id="e2e23-126">Close Office.</span></span>
1. <span data-ttu-id="e2e23-127">Verifique se o manifesto é valido</span><span class="sxs-lookup"><span data-stu-id="e2e23-127">Verify that the manifest is valid</span></span>
1. <span data-ttu-id="e2e23-128">Reinicie o suplemento</span><span class="sxs-lookup"><span data-stu-id="e2e23-128">Restart the add-in</span></span>
1. <span data-ttu-id="e2e23-129">Instale o suplemento novamente.</span><span class="sxs-lookup"><span data-stu-id="e2e23-129">Install the add-in again.</span></span>

<span data-ttu-id="e2e23-130">Você também pode enviar comentários: se estiver usando o Excel no Windows ou Mac, envie comentários à equipe de extensibilidade do Office diretamente do Excel.</span><span class="sxs-lookup"><span data-stu-id="e2e23-130">You can also give us feedback: if using Excel on Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="e2e23-131">Para fazer isso, selecione **Arquivo** | **Comentário** | **Enviar um Rosto Triste**.</span><span class="sxs-lookup"><span data-stu-id="e2e23-131">To do this, select **File** | **Feedback** | **Send a Frown**.</span></span> <span data-ttu-id="e2e23-132">Enviando um rosto triste, você fornece os logs necessários para entendermos o problema.</span><span class="sxs-lookup"><span data-stu-id="e2e23-132">Sending a frown provides the necessary logs to understand the issue.</span></span>

## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="e2e23-133">O suplemento do Outlook não funciona corretamente</span><span class="sxs-lookup"><span data-stu-id="e2e23-133">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="e2e23-134">Se um suplemento do Outlook executado no Windows e [usando o Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) não está funcionando corretamente, tente ativar a depuração de scripts no Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="e2e23-134">If an Outlook add-in running on Windows and [using Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="e2e23-135">Vá para Ferramentas > **Opções da Internet**  >  **Avançadas**.</span><span class="sxs-lookup"><span data-stu-id="e2e23-135">Go to Tools > **Internet Options** > **Advanced**.</span></span>

- <span data-ttu-id="e2e23-136">Em **Navegação, desmarque** **Desabilitar a depuração de script (Internet Explorer)** e **Desabilitar a depuração de script (Outros)**.</span><span class="sxs-lookup"><span data-stu-id="e2e23-136">Under **Browsing**, uncheck **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>

<span data-ttu-id="e2e23-137">Recomendamos que você desmarque essas configurações somente para solucionar o problema.</span><span class="sxs-lookup"><span data-stu-id="e2e23-137">We recommend that you uncheck these settings only to troubleshoot the issue.</span></span> <span data-ttu-id="e2e23-138">Se você deixar desmarcado, receberá prompts durante a navegação.</span><span class="sxs-lookup"><span data-stu-id="e2e23-138">If you leave them unchecked, you will get prompts when you browse.</span></span> <span data-ttu-id="e2e23-139">Depois que o problema for resolvido, verifique **Desabilitar a depuração de script (Internet Explorer)** e **Desabilitar a depuração de script (Outros)** novamente.</span><span class="sxs-lookup"><span data-stu-id="e2e23-139">After the issue is resolved, check **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="e2e23-140">O suplemento não é ativado no Office 2013</span><span class="sxs-lookup"><span data-stu-id="e2e23-140">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="e2e23-141">Se o complemento não for ativado quando o usuário executar as etapas a seguir.</span><span class="sxs-lookup"><span data-stu-id="e2e23-141">If the add-in doesn't activate when the user performs the following steps.</span></span>


1. <span data-ttu-id="e2e23-142">Entrar com a conta da Microsoft no Office 2013.</span><span class="sxs-lookup"><span data-stu-id="e2e23-142">Signs in with their Microsoft account in Office 2013.</span></span>

1. <span data-ttu-id="e2e23-143">Habilitar a verificação de duas etapas para a conta da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="e2e23-143">Enables two-step verification for their Microsoft account.</span></span>

1. <span data-ttu-id="e2e23-144">Verificar a identidade ao ser solicitado quando tentar inserir um suplemento.</span><span class="sxs-lookup"><span data-stu-id="e2e23-144">Verifies their identity when prompted when they try to insert an add-in.</span></span>

<span data-ttu-id="e2e23-145">Verifique se as atualizações mais recentes do Office foram instaladas ou baixe a [atualização do Office 2013](https://support.microsoft.com/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="e2e23-145">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>

## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="e2e23-146">Não é possível exibir a caixa de diálogo do suplemento</span><span class="sxs-lookup"><span data-stu-id="e2e23-146">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="e2e23-147">Quando o usuário usa um suplemento do Office, ele é solicitado a permitir a exibição de uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="e2e23-147">When using an Office Add-in, the user is asked to allow a dialog box to be displayed.</span></span> <span data-ttu-id="e2e23-148">O usuário escolhe **Permitir** e ocorre a seguinte mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="e2e23-148">The user chooses **Allow**, and the following error message occurs.</span></span>

<span data-ttu-id="e2e23-p110">"As configurações de segurança do navegador nos impedem de criar uma caixa de diálogo. Tente outro navegador ou configure o navegador para que a 'URL' e o domínio mostrado na barra de endereço estejam na mesma zona de segurança".</span><span class="sxs-lookup"><span data-stu-id="e2e23-p110">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![Captura de tela da mensagem de erro da caixa de diálogo.](../images/dialog-prevented.png)

|<span data-ttu-id="e2e23-152">**Navegadores afetados**</span><span class="sxs-lookup"><span data-stu-id="e2e23-152">**Affected browsers**</span></span>|<span data-ttu-id="e2e23-153">**Plataformas afetadas**</span><span class="sxs-lookup"><span data-stu-id="e2e23-153">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="e2e23-154">Internet Explorer, Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="e2e23-154">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="e2e23-155">Office na Web</span><span class="sxs-lookup"><span data-stu-id="e2e23-155">Office on the web</span></span>|

<span data-ttu-id="e2e23-p111">Para resolver o problema, os administradores ou usuários finais podem adicionar o domínio do suplemento à lista de sites confiáveis no Internet Explorer. Use o mesmo procedimento se estiver trabalhando com o navegador Internet Explorer ou Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="e2e23-p111">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e2e23-158">Caso não confie no suplemento, não adicione a respectiva URL à lista de sites confiáveis.</span><span class="sxs-lookup"><span data-stu-id="e2e23-158">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="e2e23-159">Para adicionar uma URL à lista de sites confiáveis:</span><span class="sxs-lookup"><span data-stu-id="e2e23-159">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="e2e23-160">No **Painel de Controle**, abra **Opções da Internet** > **Security**.</span><span class="sxs-lookup"><span data-stu-id="e2e23-160">In **Control Panel**, go to **Internet options** > **Security**.</span></span>
1. <span data-ttu-id="e2e23-161">Escolha a zona de **Sites confiáveis** e escolha **Sites**.</span><span class="sxs-lookup"><span data-stu-id="e2e23-161">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
1. <span data-ttu-id="e2e23-162">Insira a URL exibida na mensagem de erro e escolha **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="e2e23-162">Enter the URL that appears in the error message, and choose **Add**.</span></span>
1. <span data-ttu-id="e2e23-p112">Tente usar o suplemento novamente. Se o problema persistir, verifique as configurações de outras zonas de segurança e confira se o domínio do suplemento está na mesma zona que a URL exibida na barra de endereços do aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="e2e23-p112">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="e2e23-p113">Esse problema ocorre quando a API da caixa de diálogo é usada no modo pop-up. Para evitar esse problema, use o sinalizador [displayInFrame](/javascript/api/office/office.ui). Isso requer que a página tenha suporte para exibição dentro de um iframe. O exemplo a seguir mostra como usar o sinalizador.</span><span class="sxs-lookup"><span data-stu-id="e2e23-p113">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInIFrame:true}, callback);
```

## <a name="see-also"></a><span data-ttu-id="e2e23-169">Confira também</span><span class="sxs-lookup"><span data-stu-id="e2e23-169">See also</span></span>

- [<span data-ttu-id="e2e23-170">Solucionar erros de desenvolvimento com Office de complementos</span><span class="sxs-lookup"><span data-stu-id="e2e23-170">Troubleshoot development errors with Office Add-ins</span></span>](troubleshoot-development-errors.md)
