---
title: Faça seu suplemento do Office ser compatível com um suplemento COM existente
description: Habilita a compatibilidade entre o seu Add-in do Office e o seu complemento COM equivalente.
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: b5235255987cc6a644491bc548b92701b350a179
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836848"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="64aab-103">Faça seu suplemento do Office ser compatível com um suplemento COM existente</span><span class="sxs-lookup"><span data-stu-id="64aab-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="64aab-104">Se você tiver um add-in COM existente, poderá criar funcionalidade equivalente em seu Add-in do Office, permitindo assim que sua solução seja executado em outras plataformas, como o Office na Web ou mac.</span><span class="sxs-lookup"><span data-stu-id="64aab-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Mac.</span></span> <span data-ttu-id="64aab-105">Em alguns casos, seu Add-in do Office pode não ser capaz de fornecer toda a funcionalidade disponível no complemento COM correspondente.</span><span class="sxs-lookup"><span data-stu-id="64aab-105">In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in.</span></span> <span data-ttu-id="64aab-106">Nessas situações, o seu complemento COM pode oferecer uma experiência de usuário melhor no Windows do que o correspondente do Office Add-in pode fornecer.</span><span class="sxs-lookup"><span data-stu-id="64aab-106">In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.</span></span>

<span data-ttu-id="64aab-107">Você pode configurar seu Add-in do Office para que, quando o complemento COM equivalente já estiver instalado no computador de um usuário, o Office no Windows executa o add-in COM em vez do Office Add-in.</span><span class="sxs-lookup"><span data-stu-id="64aab-107">You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in.</span></span> <span data-ttu-id="64aab-108">O complemento COM é chamado de "equivalente" porque o Office fará a transição perfeita entre o complemento COM e o Complemento do Office de acordo com o qual está instalado o computador de um usuário.</span><span class="sxs-lookup"><span data-stu-id="64aab-108">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed a user's computer.</span></span>

> [!NOTE]
> <span data-ttu-id="64aab-109">Esse recurso é suportado pelas seguintes plataformas, quando conectado a uma assinatura do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="64aab-109">This feature is supported by the following platforms, when connected to a Microsoft 365 subscription.</span></span>
>
> - <span data-ttu-id="64aab-110">Excel, Word e PowerPoint na Web</span><span class="sxs-lookup"><span data-stu-id="64aab-110">Excel, Word, and PowerPoint on the web</span></span>
> - <span data-ttu-id="64aab-111">Excel, Word e PowerPoint no Windows (versão 1904 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="64aab-111">Excel, Word, and PowerPoint on Windows (version 1904 or later)</span></span>
> - <span data-ttu-id="64aab-112">Excel, Word e PowerPoint no Mac (versão 13.329 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="64aab-112">Excel, Word, and PowerPoint on Mac (version 13.329 or later)</span></span>
> - <span data-ttu-id="64aab-113">Outlook no Windows (versão 2102 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="64aab-113">Outlook on Windows (version 2102 or later)</span></span>

## <a name="specify-an-equivalent-com-add-in"></a><span data-ttu-id="64aab-114">Especificar um complemento COM equivalente</span><span class="sxs-lookup"><span data-stu-id="64aab-114">Specify an equivalent COM add-in</span></span>

### <a name="manifest"></a><span data-ttu-id="64aab-115">Manifesto</span><span class="sxs-lookup"><span data-stu-id="64aab-115">Manifest</span></span>

> [!IMPORTANT]
> <span data-ttu-id="64aab-116">Aplica-se ao Excel, PowerPoint e Word.</span><span class="sxs-lookup"><span data-stu-id="64aab-116">Applies to Excel, PowerPoint, and Word.</span></span> <span data-ttu-id="64aab-117">Suporte do Outlook em breve.</span><span class="sxs-lookup"><span data-stu-id="64aab-117">Outlook support coming soon.</span></span>

<span data-ttu-id="64aab-118">Para habilitar a compatibilidade entre o seu add-in do Office e o seu complemento COM, identifique o complemento COM equivalente no [manifesto](add-in-manifests.md) do seu Add-in do Office.</span><span class="sxs-lookup"><span data-stu-id="64aab-118">To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in.</span></span> <span data-ttu-id="64aab-119">Em seguida, o Office no Windows usará o complemento COM em vez do Office Add-in, se ambos estão instalados.</span><span class="sxs-lookup"><span data-stu-id="64aab-119">Then Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.</span></span>

<span data-ttu-id="64aab-120">O exemplo a seguir mostra a parte do manifesto que especifica um complemento COM como um complemento equivalente.</span><span class="sxs-lookup"><span data-stu-id="64aab-120">The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in.</span></span> <span data-ttu-id="64aab-121">O valor do elemento identifica o complemento COM e o `ProgId` [elemento EquivalentAddins](../reference/manifest/equivalentaddins.md) deve ser posicionado imediatamente antes da marca de `VersionOverrides` fechamento.</span><span class="sxs-lookup"><span data-stu-id="64aab-121">The value of the `ProgId` element identifies the COM add-in and the [EquivalentAddins](../reference/manifest/equivalentaddins.md) element must be positioned immediately before the closing `VersionOverrides` tag.</span></span>

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

> [!TIP]
> <span data-ttu-id="64aab-122">Para obter informações sobre o complemento COM e a compatibilidade de UDF XLL, consulte Tornar suas funções personalizadas compatíveis com funções definidas pelo usuário [XLL.](../excel/make-custom-functions-compatible-with-xll-udf.md)</span><span class="sxs-lookup"><span data-stu-id="64aab-122">For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

### <a name="group-policy"></a><span data-ttu-id="64aab-123">Política de grupo</span><span class="sxs-lookup"><span data-stu-id="64aab-123">Group policy</span></span>

> [!IMPORTANT]
> <span data-ttu-id="64aab-124">Aplica-se somente ao Outlook.</span><span class="sxs-lookup"><span data-stu-id="64aab-124">Applies to Outlook only.</span></span>

<span data-ttu-id="64aab-125">Para declarar compatibilidade entre o seu **add-in** da Web do Outlook e o complemento COM/VSTO, identifique o complemento COM equivalente na política de grupo Desative os complementos da Web do Outlook cujo complemento COM ou VSTO equivalente é instalado configurando no computador do usuário.</span><span class="sxs-lookup"><span data-stu-id="64aab-125">To declare compatibility between your Outlook web add-in and COM/VSTO add-in, identify the equivalent COM add-in in the group policy **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed** by configuring on the user's machine.</span></span> <span data-ttu-id="64aab-126">Em seguida, o Outlook no Windows usará o complemento COM em vez do complemento da Web, se ambos estão instalados.</span><span class="sxs-lookup"><span data-stu-id="64aab-126">Then Outlook on Windows will use the COM add-in instead of the web add-in, if they're both installed.</span></span>

1. <span data-ttu-id="64aab-127">Baixe a ferramenta [Modelos Administrativos mais](https://www.microsoft.com/download/details.aspx?id=49030)recentes, preste atenção às Instruções de **Instalação da ferramenta.**</span><span class="sxs-lookup"><span data-stu-id="64aab-127">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030), paying attention to the tool's **Install Instructions**.</span></span>
1. <span data-ttu-id="64aab-128">Abra o Editor de Política de Grupo Local (**gpedit.msc**).</span><span class="sxs-lookup"><span data-stu-id="64aab-128">Open the Local Group Policy Editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="64aab-129">Navegue **até Configuração do** Usuário Modelos  >     >  **Administrativos do Microsoft Outlook 2016**  >  **Diversos**.</span><span class="sxs-lookup"><span data-stu-id="64aab-129">Navigate to **User Configuration** > **Administrative Templates**  > **Microsoft Outlook 2016** > **Miscellaneous**.</span></span>
1. <span data-ttu-id="64aab-130">Selecione a **configuração Desativar os complementos da Web do Outlook cujos complementos COM ou VSTO equivalentes estão instalados**.</span><span class="sxs-lookup"><span data-stu-id="64aab-130">Select the setting **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed**.</span></span>
1. <span data-ttu-id="64aab-131">Abra o link para editar a configuração de política.</span><span class="sxs-lookup"><span data-stu-id="64aab-131">Open the link to edit the policy setting.</span></span>
1. <span data-ttu-id="64aab-132">Na caixa de diálogo **Os complementos da Web do Outlook para desativar**:</span><span class="sxs-lookup"><span data-stu-id="64aab-132">In the dialog **Outlook web add-ins to deactivate**:</span></span>
    1. <span data-ttu-id="64aab-133">Definir **o nome** do valor como o encontrado no manifesto do complemento da `Id` Web.</span><span class="sxs-lookup"><span data-stu-id="64aab-133">Set **Value name** to the `Id` found in the web add-in's manifest.</span></span> <span data-ttu-id="64aab-134">**Importante**: *Não adicione* chaves ao redor da `{}` entrada.</span><span class="sxs-lookup"><span data-stu-id="64aab-134">**Important**: Do *not* add curly braces `{}` around the entry.</span></span>
    1. <span data-ttu-id="64aab-135">Definir **Valor** como `ProgId` o do complemento COM/VSTO equivalente.</span><span class="sxs-lookup"><span data-stu-id="64aab-135">Set **Value** to the `ProgId` of the equivalent COM/VSTO add-in.</span></span>
    1. <span data-ttu-id="64aab-136">Selecione **OK** para colocar a atualização em vigor.</span><span class="sxs-lookup"><span data-stu-id="64aab-136">Select **OK** to put the update into effect.</span></span>
    <span data-ttu-id="64aab-137">![Captura de tela mostrando a caixa de diálogo "Os complementos da Web do Outlook para desativar"](../images/outlook-deactivate-gpo-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="64aab-137">![Screenshot showing the dialog "Outlook web add-ins to deactivate"](../images/outlook-deactivate-gpo-dialog.png)</span></span>

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="64aab-138">Comportamento equivalente para usuários</span><span class="sxs-lookup"><span data-stu-id="64aab-138">Equivalent behavior for users</span></span>

<span data-ttu-id="64aab-139">Quando um [complemento COM](#specify-an-equivalent-com-add-in)equivalente é especificado, o Office no Windows não exibirá a interface de usuário do seu Complemento do Office (UI) se o complemento COM equivalente estiver instalado.</span><span class="sxs-lookup"><span data-stu-id="64aab-139">When an [equivalent COM add-in is specified](#specify-an-equivalent-com-add-in), Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed.</span></span> <span data-ttu-id="64aab-140">O Office oculta apenas os botões de faixa de opções do Office Add-in e não impede a instalação.</span><span class="sxs-lookup"><span data-stu-id="64aab-140">Office only hides the ribbon buttons of the Office Add-in and does not prevent installation.</span></span> <span data-ttu-id="64aab-141">Portanto, seu Complemento do Office ainda aparecerá nos seguintes locais na interface do usuário:</span><span class="sxs-lookup"><span data-stu-id="64aab-141">Therefore your Office Add-in will still appear in the following locations within the UI:</span></span>

- <span data-ttu-id="64aab-142">Em **Meus complementos**</span><span class="sxs-lookup"><span data-stu-id="64aab-142">Under **My add-ins**</span></span>
- <span data-ttu-id="64aab-143">Como entrada no gerenciador de faixa de opções (somente Excel, Word e PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="64aab-143">As an entry in the ribbon manager (Excel, Word, and PowerPoint only)</span></span>

> [!NOTE]
> <span data-ttu-id="64aab-144">A especificação de um complemento COM equivalente no manifesto não tem efeito em outras plataformas, como o Office na Web ou no Mac.</span><span class="sxs-lookup"><span data-stu-id="64aab-144">Specifying an equivalent COM add-in in the manifest has no effect on other platforms like Office on the web or on Mac.</span></span>

<span data-ttu-id="64aab-145">Os cenários a seguir descrevem o que acontece dependendo de como o usuário adquire o Office Add-in.</span><span class="sxs-lookup"><span data-stu-id="64aab-145">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="64aab-146">Aquisição do AppSource de um Add-in do Office</span><span class="sxs-lookup"><span data-stu-id="64aab-146">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="64aab-147">Se um usuário adquirir o Office Add-in do AppSource e o complemento COM equivalente já estiver instalado, o Office:</span><span class="sxs-lookup"><span data-stu-id="64aab-147">If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="64aab-148">Instale o Office Add-in.</span><span class="sxs-lookup"><span data-stu-id="64aab-148">Install the Office Add-in.</span></span>
2. <span data-ttu-id="64aab-149">Ocultar a interface do usuário do Complemento do Office na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="64aab-149">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="64aab-150">Exibe um chamado para o usuário que aponta para o botão de faixa de opções do complemento COM.</span><span class="sxs-lookup"><span data-stu-id="64aab-150">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="64aab-151">Implantação centralizada do Office Add-in</span><span class="sxs-lookup"><span data-stu-id="64aab-151">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="64aab-152">Se um administrador implantar o Add-in do Office em seu locatário usando a implantação centralizada e o complemento COM equivalente já estiver instalado, o usuário deverá reiniciar o Office antes de ver as alterações.</span><span class="sxs-lookup"><span data-stu-id="64aab-152">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes.</span></span> <span data-ttu-id="64aab-153">Depois que o Office reiniciar, ele irá:</span><span class="sxs-lookup"><span data-stu-id="64aab-153">After Office restarts, it will:</span></span>

1. <span data-ttu-id="64aab-154">Instale o Office Add-in.</span><span class="sxs-lookup"><span data-stu-id="64aab-154">Install the Office Add-in.</span></span>
2. <span data-ttu-id="64aab-155">Ocultar a interface do usuário do Complemento do Office na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="64aab-155">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="64aab-156">Exibe um chamado para o usuário que aponta para o botão de faixa de opções do complemento COM.</span><span class="sxs-lookup"><span data-stu-id="64aab-156">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="64aab-157">Documento compartilhado com o Add-in incorporado do Office</span><span class="sxs-lookup"><span data-stu-id="64aab-157">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="64aab-158">Se um usuário tiver o add-in COM instalado e, em seguida, receber um documento compartilhado com o Complemento do Office incorporado, quando ele abrir o documento, o Office:</span><span class="sxs-lookup"><span data-stu-id="64aab-158">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="64aab-159">Solicitar que o usuário confie no Office Add-in.</span><span class="sxs-lookup"><span data-stu-id="64aab-159">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="64aab-160">Se for confiável, o Office Add-in será instalado.</span><span class="sxs-lookup"><span data-stu-id="64aab-160">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="64aab-161">Ocultar a interface do usuário do Complemento do Office na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="64aab-161">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="64aab-162">Outro comportamento de complemento COM</span><span class="sxs-lookup"><span data-stu-id="64aab-162">Other COM add-in behavior</span></span>

### <a name="excel-powerpoint-word"></a><span data-ttu-id="64aab-163">Excel, PowerPoint, Word</span><span class="sxs-lookup"><span data-stu-id="64aab-163">Excel, PowerPoint, Word</span></span>

<span data-ttu-id="64aab-164">Se um usuário desinstalar o complemento COM equivalente, o Office no Windows restaurará a interface do usuário do Office Add-in.</span><span class="sxs-lookup"><span data-stu-id="64aab-164">If a user uninstalls the equivalent COM add-in, then Office on Windows restores the Office Add-in UI.</span></span>

<span data-ttu-id="64aab-165">Depois de especificar um complemento COM equivalente para o seu Complemento do Office, o Office interrompe o processamento de atualizações para o seu Add-in do Office.</span><span class="sxs-lookup"><span data-stu-id="64aab-165">After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="64aab-166">Para adquirir as atualizações mais recentes do Office Add-in, o usuário deve primeiro desinstalar o complemento COM.</span><span class="sxs-lookup"><span data-stu-id="64aab-166">To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.</span></span>

### <a name="outlook"></a><span data-ttu-id="64aab-167">Outlook</span><span class="sxs-lookup"><span data-stu-id="64aab-167">Outlook</span></span>

<span data-ttu-id="64aab-168">O complemento COM/VSTO deve ser conectado quando o Outlook for iniciado para que o complemento da Web correspondente seja desabilitado.</span><span class="sxs-lookup"><span data-stu-id="64aab-168">The COM/VSTO add-in must be connected when Outlook is started in order for the corresponding web add-in to be disabled.</span></span>

<span data-ttu-id="64aab-169">Se o complemento COM/VSTO for desconectado durante uma sessão subsequente do Outlook, o complemento da Web provavelmente permanecerá desabilitado até que o Outlook seja reiniciado.</span><span class="sxs-lookup"><span data-stu-id="64aab-169">If the COM/VSTO add-in is then disconnected during a subsequent Outlook session, the web add-in will likely remain disabled until Outlook is restarted.</span></span>

## <a name="see-also"></a><span data-ttu-id="64aab-170">Confira também</span><span class="sxs-lookup"><span data-stu-id="64aab-170">See also</span></span>

- [<span data-ttu-id="64aab-171">Tornar suas funções personalizadas compatíveis com funções definidas pelo usuário XLL</span><span class="sxs-lookup"><span data-stu-id="64aab-171">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
