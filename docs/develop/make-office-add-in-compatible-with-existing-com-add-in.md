---
title: Faça seu suplemento do Office ser compatível com um suplemento COM existente
description: Habilita a compatibilidade entre seu Office e o seu complemento COM equivalente.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 85e5d8cc06aa599862c92b59a26c744f28ca2d22
ms.sourcegitcommit: 95fc1fc8a0dbe8fc94f0ea647836b51cc7f8601d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/14/2021
ms.locfileid: "53418682"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="44159-103">Faça seu suplemento do Office ser compatível com um suplemento COM existente</span><span class="sxs-lookup"><span data-stu-id="44159-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="44159-104">Se você tiver um complemento COM existente, poderá criar funcionalidade equivalente em seu Office Add-in, permitindo que sua solução seja executado em outras plataformas, como Office na Web ou Mac.</span><span class="sxs-lookup"><span data-stu-id="44159-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Mac.</span></span> <span data-ttu-id="44159-105">Em alguns casos, seu Office de usuário pode não ser capaz de fornecer toda a funcionalidade disponível no complemento COM correspondente.</span><span class="sxs-lookup"><span data-stu-id="44159-105">In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in.</span></span> <span data-ttu-id="44159-106">Nessas situações, o seu add-in COM pode fornecer uma melhor experiência do usuário Windows do que o Office que o Add-in pode fornecer.</span><span class="sxs-lookup"><span data-stu-id="44159-106">In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.</span></span>

<span data-ttu-id="44159-107">Você pode configurar seu Office Add-in para que, quando o complemento COM equivalente já estiver instalado no computador de um usuário, o Office no Windows executa o add-in COM em vez do Office Add-in.</span><span class="sxs-lookup"><span data-stu-id="44159-107">You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in.</span></span> <span data-ttu-id="44159-108">O add-in COM é chamado de "equivalente" porque o Office fará uma transição perfeita entre o complemento COM e o Office De acordo com o qual está instalado o computador de um usuário.</span><span class="sxs-lookup"><span data-stu-id="44159-108">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed a user's computer.</span></span>

> [!NOTE]
> <span data-ttu-id="44159-109">Esse recurso é suportado pela plataforma e aplicativos a seguir, quando conectado a uma assinatura Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="44159-109">This feature is supported by the following platform and applications, when connected to a Microsoft 365 subscription.</span></span> <span data-ttu-id="44159-110">Os complementos COM não podem ser instalados em qualquer outra plataforma, portanto, nessas plataformas, o elemento de manifesto discutido posteriormente neste artigo, , é `EquivalentAddins` ignorado.</span><span class="sxs-lookup"><span data-stu-id="44159-110">COM add-ins cannot be installed on any other platform, so on those platforms the manifest element that is discussed later in this article, `EquivalentAddins`, is ignored.</span></span>
>
> - <span data-ttu-id="44159-111">Excel, Word e PowerPoint no Windows (versão 1904 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="44159-111">Excel, Word, and PowerPoint on Windows (version 1904 or later)</span></span>

## <a name="specify-an-equivalent-com-add-in"></a><span data-ttu-id="44159-112">Especificar um complemento COM equivalente</span><span class="sxs-lookup"><span data-stu-id="44159-112">Specify an equivalent COM add-in</span></span>

### <a name="manifest"></a><span data-ttu-id="44159-113">Manifesto</span><span class="sxs-lookup"><span data-stu-id="44159-113">Manifest</span></span>

> [!IMPORTANT]
> <span data-ttu-id="44159-114">Aplica-se Excel, PowerPoint e Word.</span><span class="sxs-lookup"><span data-stu-id="44159-114">Applies to Excel, PowerPoint, and Word.</span></span> <span data-ttu-id="44159-115">Outlook suporte em breve.</span><span class="sxs-lookup"><span data-stu-id="44159-115">Outlook support coming soon.</span></span>

<span data-ttu-id="44159-116">Para habilitar a compatibilidade entre o seu Office e o complemento COM, [](add-in-manifests.md) identifique o complemento COM equivalente no manifesto do seu Office Add-in.</span><span class="sxs-lookup"><span data-stu-id="44159-116">To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in.</span></span> <span data-ttu-id="44159-117">Em seguida, Office no Windows usará o add-in COM em vez do Office do Office, se ambos estão instalados.</span><span class="sxs-lookup"><span data-stu-id="44159-117">Then Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.</span></span>

<span data-ttu-id="44159-118">O exemplo a seguir mostra a parte do manifesto que especifica um complemento COM como um complemento equivalente.</span><span class="sxs-lookup"><span data-stu-id="44159-118">The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in.</span></span> <span data-ttu-id="44159-119">O valor do elemento identifica o complemento COM e o `ProgId` [elemento EquivalentAddins](../reference/manifest/equivalentaddins.md) deve ser posicionado imediatamente antes da marca de `VersionOverrides` fechamento.</span><span class="sxs-lookup"><span data-stu-id="44159-119">The value of the `ProgId` element identifies the COM add-in and the [EquivalentAddins](../reference/manifest/equivalentaddins.md) element must be positioned immediately before the closing `VersionOverrides` tag.</span></span>

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
> <span data-ttu-id="44159-120">Para obter informações sobre o complemento COM e a compatibilidade de UDF XLL, consulte Tornar suas funções personalizadas compatíveis com funções definidas pelo usuário [XLL.](../excel/make-custom-functions-compatible-with-xll-udf.md)</span><span class="sxs-lookup"><span data-stu-id="44159-120">For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

### <a name="group-policy"></a><span data-ttu-id="44159-121">Política de grupo</span><span class="sxs-lookup"><span data-stu-id="44159-121">Group policy</span></span>

> [!IMPORTANT]
> <span data-ttu-id="44159-122">Aplica-se Outlook somente.</span><span class="sxs-lookup"><span data-stu-id="44159-122">Applies to Outlook only.</span></span>

<span data-ttu-id="44159-123">Para declarar compatibilidade entre o seu **add-in** da Web Outlook e o com/VSTO add-in, identifique o complemento COM equivalente na política de grupo Desative os complementos da Web do Outlook cuja COM equivalente ou um complemento VSTO está instalado configurando-se no computador do usuário.</span><span class="sxs-lookup"><span data-stu-id="44159-123">To declare compatibility between your Outlook web add-in and COM/VSTO add-in, identify the equivalent COM add-in in the group policy **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed** by configuring on the user's machine.</span></span> <span data-ttu-id="44159-124">Em seguida Outlook no Windows usará o add-in COM em vez do complemento da Web, se ambos estão instalados.</span><span class="sxs-lookup"><span data-stu-id="44159-124">Then Outlook on Windows will use the COM add-in instead of the web add-in, if they're both installed.</span></span>

1. <span data-ttu-id="44159-125">Baixe a ferramenta [Modelos Administrativos mais](https://www.microsoft.com/download/details.aspx?id=49030)recentes, preste atenção às Instruções de **Instalação da ferramenta.**</span><span class="sxs-lookup"><span data-stu-id="44159-125">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030), paying attention to the tool's **Install Instructions**.</span></span>
1. <span data-ttu-id="44159-126">Abra o Editor de Política de Grupo Local (**gpedit.msc**).</span><span class="sxs-lookup"><span data-stu-id="44159-126">Open the Local Group Policy Editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="44159-127">Navegue **até Configuração do** Usuário Modelos  >     >  **Administrativos microsoft Outlook 2016**  >  **diversos**.</span><span class="sxs-lookup"><span data-stu-id="44159-127">Navigate to **User Configuration** > **Administrative Templates**  > **Microsoft Outlook 2016** > **Miscellaneous**.</span></span>
1. <span data-ttu-id="44159-128">Selecione a configuração Desativar Outlook da Web, cuja COM ou **VSTO add-in** equivalente está instalado .</span><span class="sxs-lookup"><span data-stu-id="44159-128">Select the setting **Deactivate Outlook web add-ins whose equivalent COM or VSTO add-in is installed**.</span></span>
1. <span data-ttu-id="44159-129">Abra o link para editar a configuração de política.</span><span class="sxs-lookup"><span data-stu-id="44159-129">Open the link to edit the policy setting.</span></span>
1. <span data-ttu-id="44159-130">Na caixa de **diálogo Outlook de web para desativar**:</span><span class="sxs-lookup"><span data-stu-id="44159-130">In the dialog **Outlook web add-ins to deactivate**:</span></span>
    1. <span data-ttu-id="44159-131">Definir **o nome** do valor como o encontrado no manifesto do complemento da `Id` Web.</span><span class="sxs-lookup"><span data-stu-id="44159-131">Set **Value name** to the `Id` found in the web add-in's manifest.</span></span> <span data-ttu-id="44159-132">**Importante**: *Não adicione* chaves ao redor da `{}` entrada.</span><span class="sxs-lookup"><span data-stu-id="44159-132">**Important**: Do *not* add curly braces `{}` around the entry.</span></span>
    1. <span data-ttu-id="44159-133">**Desmarcar** Valor `ProgId` como o do com/VSTO equivalente.</span><span class="sxs-lookup"><span data-stu-id="44159-133">Set **Value** to the `ProgId` of the equivalent COM/VSTO add-in.</span></span>
    1. <span data-ttu-id="44159-134">Selecione **OK** para colocar a atualização em vigor.</span><span class="sxs-lookup"><span data-stu-id="44159-134">Select **OK** to put the update into effect.</span></span>
    <span data-ttu-id="44159-135">![Captura de tela mostrando a caixa de diálogo "Outlook de web para desativar".](../images/outlook-deactivate-gpo-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="44159-135">![Screenshot showing the dialog "Outlook web add-ins to deactivate".](../images/outlook-deactivate-gpo-dialog.png)</span></span>

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="44159-136">Comportamento equivalente para usuários</span><span class="sxs-lookup"><span data-stu-id="44159-136">Equivalent behavior for users</span></span>

<span data-ttu-id="44159-137">Quando um complemento [COM](#specify-an-equivalent-com-add-in)equivalente é especificado, o Office no Windows não exibirá Office interface de usuário do seu complemento (UI) se o complemento COM equivalente estiver instalado.</span><span class="sxs-lookup"><span data-stu-id="44159-137">When an [equivalent COM add-in is specified](#specify-an-equivalent-com-add-in), Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed.</span></span> <span data-ttu-id="44159-138">Office oculta apenas os botões de faixa de opções do Office e não impede a instalação.</span><span class="sxs-lookup"><span data-stu-id="44159-138">Office only hides the ribbon buttons of the Office Add-in and does not prevent installation.</span></span> <span data-ttu-id="44159-139">Portanto, seu Office de usuário ainda aparecerá nos seguintes locais dentro da interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="44159-139">Therefore your Office Add-in will still appear in the following locations within the UI.</span></span>

- <span data-ttu-id="44159-140">Em **Meus complementos**</span><span class="sxs-lookup"><span data-stu-id="44159-140">Under **My add-ins**</span></span>
- <span data-ttu-id="44159-141">Como entrada no gerenciador de faixa de opções (Excel, Word e PowerPoint somente)</span><span class="sxs-lookup"><span data-stu-id="44159-141">As an entry in the ribbon manager (Excel, Word, and PowerPoint only)</span></span>

> [!NOTE]
> <span data-ttu-id="44159-142">A especificação de um complemento COM equivalente no manifesto não tem efeito em outras plataformas, como Office na Web ou no Mac.</span><span class="sxs-lookup"><span data-stu-id="44159-142">Specifying an equivalent COM add-in in the manifest has no effect on other platforms like Office on the web or on Mac.</span></span>

<span data-ttu-id="44159-143">Os cenários a seguir descrevem o que acontece dependendo de como o usuário adquire o Office Add-in.</span><span class="sxs-lookup"><span data-stu-id="44159-143">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="44159-144">Aquisição do AppSource de um Office Add-in</span><span class="sxs-lookup"><span data-stu-id="44159-144">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="44159-145">Se um usuário adquirir o Office do AppSource e o complemento COM equivalente já estiver instalado, Office:</span><span class="sxs-lookup"><span data-stu-id="44159-145">If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="44159-146">Instale o Office de usuário.</span><span class="sxs-lookup"><span data-stu-id="44159-146">Install the Office Add-in.</span></span>
2. <span data-ttu-id="44159-147">Ocultar a Office interface do usuário de complemento na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="44159-147">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="44159-148">Exibe um chamado para o usuário que aponta para o botão de faixa de opções do complemento COM.</span><span class="sxs-lookup"><span data-stu-id="44159-148">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="44159-149">Implantação centralizada do Office Desemporto</span><span class="sxs-lookup"><span data-stu-id="44159-149">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="44159-150">Se um administrador implantar o Office Add-in em seu locatário usando a implantação centralizada e o complemento COM equivalente já estiver instalado, o usuário deverá reiniciar o Office antes de ver quaisquer alterações.</span><span class="sxs-lookup"><span data-stu-id="44159-150">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes.</span></span> <span data-ttu-id="44159-151">Depois Office reiniciar, ele irá:</span><span class="sxs-lookup"><span data-stu-id="44159-151">After Office restarts, it will:</span></span>

1. <span data-ttu-id="44159-152">Instale o Office de usuário.</span><span class="sxs-lookup"><span data-stu-id="44159-152">Install the Office Add-in.</span></span>
2. <span data-ttu-id="44159-153">Ocultar a Office interface do usuário de complemento na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="44159-153">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="44159-154">Exibe um chamado para o usuário que aponta para o botão de faixa de opções do complemento COM.</span><span class="sxs-lookup"><span data-stu-id="44159-154">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="44159-155">Documento compartilhado com o Office Incorporado</span><span class="sxs-lookup"><span data-stu-id="44159-155">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="44159-156">Se um usuário tiver o complemento COM instalado e, em seguida, receber um documento compartilhado com o Office Add-in incorporado, quando abrir o documento, Office irá:</span><span class="sxs-lookup"><span data-stu-id="44159-156">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="44159-157">Solicitar que o usuário confie no Office Add-in.</span><span class="sxs-lookup"><span data-stu-id="44159-157">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="44159-158">Se for confiável, o Office de usuário será instalado.</span><span class="sxs-lookup"><span data-stu-id="44159-158">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="44159-159">Ocultar a Office interface do usuário de complemento na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="44159-159">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="44159-160">Outro comportamento de complemento COM</span><span class="sxs-lookup"><span data-stu-id="44159-160">Other COM add-in behavior</span></span>

### <a name="excel-powerpoint-word"></a><span data-ttu-id="44159-161">Excel, PowerPoint, Word</span><span class="sxs-lookup"><span data-stu-id="44159-161">Excel, PowerPoint, Word</span></span>

<span data-ttu-id="44159-162">Se um usuário desinstalar o add-in COM equivalente, Office em Windows restaurará a interface do usuário do Office Desinstalar.</span><span class="sxs-lookup"><span data-stu-id="44159-162">If a user uninstalls the equivalent COM add-in, then Office on Windows restores the Office Add-in UI.</span></span>

<span data-ttu-id="44159-163">Depois de especificar um add-in COM equivalente para seu Office de Office, o Office interrompe o processamento de atualizações para seu Office Add-in.</span><span class="sxs-lookup"><span data-stu-id="44159-163">After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="44159-164">Para adquirir as atualizações mais recentes para o Office, o usuário deve primeiro desinstalar o complemento COM.</span><span class="sxs-lookup"><span data-stu-id="44159-164">To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.</span></span>

### <a name="outlook"></a><span data-ttu-id="44159-165">Outlook</span><span class="sxs-lookup"><span data-stu-id="44159-165">Outlook</span></span>

<span data-ttu-id="44159-166">O complemento COM/VSTO deve ser conectado quando o Outlook for iniciado para que o complemento da Web correspondente seja desabilitado.</span><span class="sxs-lookup"><span data-stu-id="44159-166">The COM/VSTO add-in must be connected when Outlook is started in order for the corresponding web add-in to be disabled.</span></span>

<span data-ttu-id="44159-167">Se o complemento COM/VSTO for desconectado durante uma sessão de Outlook subsequente, o complemento da Web provavelmente permanecerá desabilitado até que Outlook seja reiniciado.</span><span class="sxs-lookup"><span data-stu-id="44159-167">If the COM/VSTO add-in is then disconnected during a subsequent Outlook session, the web add-in will likely remain disabled until Outlook is restarted.</span></span>

## <a name="see-also"></a><span data-ttu-id="44159-168">Confira também</span><span class="sxs-lookup"><span data-stu-id="44159-168">See also</span></span>

- [<span data-ttu-id="44159-169">Tornar suas funções personalizadas compatíveis com funções definidas pelo usuário XLL</span><span class="sxs-lookup"><span data-stu-id="44159-169">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
