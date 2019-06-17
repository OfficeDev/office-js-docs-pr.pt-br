---
title: Torne o seu suplemento do Office compatível com um suplemento COM existente
description: Habilitar a compatibilidade entre o suplemento do Office e o suplemento COM equivalente
ms.date: 06/13/2019
localization_priority: Normal
ms.openlocfilehash: 1dd6de5e07d835cc017f95cd1a992a5f5d188ef1
ms.sourcegitcommit: ee5b4935b5ee1db567a13627b2f87471ee8b8165
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/13/2019
ms.locfileid: "34933755"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in-preview"></a><span data-ttu-id="b89cd-103">Tornar o suplemento do Office compatível com um suplemento de COM existente (visualização)</span><span class="sxs-lookup"><span data-stu-id="b89cd-103">Make your Office Add-in compatible with an existing COM add-in (preview)</span></span>

<span data-ttu-id="b89cd-104">Se você tiver um suplemento COM existente, poderá criar uma funcionalidade equivalente no suplemento do Office, permitindo assim que sua solução seja executada em outras plataformas, como o Office na Web ou o Office no Mac.</span><span class="sxs-lookup"><span data-stu-id="b89cd-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Office on Mac.</span></span> <span data-ttu-id="b89cd-105">Em alguns casos, o suplemento do Office pode não ser capaz de fornecer toda a funcionalidade que está disponível no suplemento COM correspondente.</span><span class="sxs-lookup"><span data-stu-id="b89cd-105">In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in.</span></span> <span data-ttu-id="b89cd-106">Nessas situações, o suplemento COM pode fornecer uma experiência de usuário melhor no Windows do que o suplemento do Office correspondente pode fornecer.</span><span class="sxs-lookup"><span data-stu-id="b89cd-106">In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.</span></span>

<span data-ttu-id="b89cd-107">Você pode configurar seu suplemento do Office para que, quando o suplemento COM equivalente já estiver instalado no computador de um usuário, o Office no Windows execute o suplemento COM em vez do suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="b89cd-107">You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in.</span></span> <span data-ttu-id="b89cd-108">O suplemento de COM é chamado de "equivalente" porque o Office faz uma transição transparente entre o suplemento de COM e o suplemento do Office de acordo com o qual está instalado o computador de um usuário.</span><span class="sxs-lookup"><span data-stu-id="b89cd-108">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed a user's computer.</span></span>

> [!NOTE]
> <span data-ttu-id="b89cd-109">Este recurso está atualmente em versão prévia e não tem suporte para uso em ambientes de produção.</span><span class="sxs-lookup"><span data-stu-id="b89cd-109">This feature is currently in preview and not supported for use in production environments.</span></span> <span data-ttu-id="b89cd-110">Ele está disponível no Excel, Word e PowerPoint versão 16.0.11629.20214 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="b89cd-110">It's available in Excel, Word, and PowerPoint version 16.0.11629.20214 or later.</span></span> <span data-ttu-id="b89cd-111">Para acessar essa compilação, você deve ter uma assinatura do Office 365 e participar do programa [Office](https://products.office.com/office-insider) Insider no nível do insider. \*\*\*\*</span><span class="sxs-lookup"><span data-stu-id="b89cd-111">To access this build, you must have an Office 365 subscription and join the [Office Insider](https://products.office.com/office-insider) program at the **Insider** level.</span></span>

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a><span data-ttu-id="b89cd-112">Especificar um suplemento COM equivalente no manifesto</span><span class="sxs-lookup"><span data-stu-id="b89cd-112">Specify an equivalent COM add-in in the manifest</span></span>

<span data-ttu-id="b89cd-113">Para habilitar a compatibilidade entre o suplemento do Office e o suplemento de COM, identifique o suplemento COM equivalente no [manifesto](add-in-manifests.md) do suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="b89cd-113">To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in.</span></span> <span data-ttu-id="b89cd-114">O Office no Windows usará o suplemento COM em vez do suplemento do Office, se eles estiverem instalados.</span><span class="sxs-lookup"><span data-stu-id="b89cd-114">Then Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.</span></span>

<span data-ttu-id="b89cd-115">O exemplo a seguir mostra a parte do manifesto que especifica um suplemento de COM como um suplemento equivalente.</span><span class="sxs-lookup"><span data-stu-id="b89cd-115">The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in.</span></span> <span data-ttu-id="b89cd-116">O valor do `ProgID` elemento identifica o suplemento de com.</span><span class="sxs-lookup"><span data-stu-id="b89cd-116">The value of the `ProgID` element identifies the COM add-in.</span></span>

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgID>ContosoCOMAddin</ProgID>
      <Type>COM</Type>
    </EquivalentAddin>
  <EquivalentAddins>
  ...
</VersionOverrides>
```

> [!TIP]
> <span data-ttu-id="b89cd-117">Para obter informações sobre o suplemento de COM e a compatibilidade do XLL UDF, confira [tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário do XLL](../excel/make-custom-functions-compatible-with-xll-udf.md).</span><span class="sxs-lookup"><span data-stu-id="b89cd-117">For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="b89cd-118">Comportamento equivalente para usuários</span><span class="sxs-lookup"><span data-stu-id="b89cd-118">Equivalent behavior for users</span></span>

<span data-ttu-id="b89cd-119">Quando um suplemento COM equivalente é especificado no manifesto do suplemento do Office, o Office no Windows não exibirá a interface do usuário do suplemento do Office se o suplemento COM equivalente estiver instalado.</span><span class="sxs-lookup"><span data-stu-id="b89cd-119">When an equivalent COM add-in is specified in the Office Add-in manifest, Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed.</span></span> <span data-ttu-id="b89cd-120">O Office só oculta os botões da faixa de opções do suplemento do Office e não impede a instalação.</span><span class="sxs-lookup"><span data-stu-id="b89cd-120">Office only hides the ribbon buttons of the Office Add-in and does not prevent installation.</span></span> <span data-ttu-id="b89cd-121">Portanto, o suplemento do Office ainda aparecerá nos seguintes locais dentro da interface do usuário:</span><span class="sxs-lookup"><span data-stu-id="b89cd-121">Therefore your Office Add-in will still appear in the following locations within the UI:</span></span>

- <span data-ttu-id="b89cd-122">Em **meus suplementos**</span><span class="sxs-lookup"><span data-stu-id="b89cd-122">Under **My add-ins**</span></span>
- <span data-ttu-id="b89cd-123">Como uma entrada no Gerenciador de faixa de opções</span><span class="sxs-lookup"><span data-stu-id="b89cd-123">As an entry in the ribbon manager</span></span>

> [!NOTE]
> <span data-ttu-id="b89cd-124">A especificação de um suplemento COM equivalente no manifesto não tem efeito sobre outras plataformas como o Office na Web ou o Office para Mac.</span><span class="sxs-lookup"><span data-stu-id="b89cd-124">Specifying an equivalent COM add-in in the manifest has no effect on other platforms like Office on the web or Office for Mac.</span></span>

<span data-ttu-id="b89cd-125">Os cenários a seguir descrevem o que acontece dependendo de como o usuário adquire o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="b89cd-125">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="b89cd-126">Aquisição do AppSource de um suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="b89cd-126">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="b89cd-127">Se um usuário adquire o suplemento do Office do AppSource e o suplemento COM equivalente já estiver instalado, o Office irá:</span><span class="sxs-lookup"><span data-stu-id="b89cd-127">If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="b89cd-128">Instalar o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="b89cd-128">Install the Office Add-in.</span></span>
2. <span data-ttu-id="b89cd-129">Ocultar a interface do usuário do suplemento do Office na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="b89cd-129">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="b89cd-130">Exibe uma chamada para o usuário que aponta o botão da faixa de opções suplemento de COM.</span><span class="sxs-lookup"><span data-stu-id="b89cd-130">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="b89cd-131">Implantação centralizada do suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="b89cd-131">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="b89cd-132">Se um administrador implantar o suplemento do Office em seu locatário usando a implantação centralizada e o suplemento COM equivalente já estiver instalado, o usuário deverá reiniciar o Office antes de ver as alterações.</span><span class="sxs-lookup"><span data-stu-id="b89cd-132">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes.</span></span> <span data-ttu-id="b89cd-133">Após a reinicialização do Office, ela irá:</span><span class="sxs-lookup"><span data-stu-id="b89cd-133">After Office restarts, it will:</span></span>

1. <span data-ttu-id="b89cd-134">Instalar o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="b89cd-134">Install the Office Add-in.</span></span>
2. <span data-ttu-id="b89cd-135">Ocultar a interface do usuário do suplemento do Office na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="b89cd-135">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="b89cd-136">Exibe uma chamada para o usuário que aponta o botão da faixa de opções suplemento de COM.</span><span class="sxs-lookup"><span data-stu-id="b89cd-136">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="b89cd-137">Documento compartilhado com o suplemento incorporado do Office</span><span class="sxs-lookup"><span data-stu-id="b89cd-137">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="b89cd-138">Se um usuário tiver o suplemento COM instalado e, em seguida, receber um documento compartilhado com o suplemento do Office incorporado, quando abrir o documento, o Office irá:</span><span class="sxs-lookup"><span data-stu-id="b89cd-138">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="b89cd-139">Solicitar que o usuário confie no suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="b89cd-139">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="b89cd-140">Se confiável, o suplemento do Office será instalado.</span><span class="sxs-lookup"><span data-stu-id="b89cd-140">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="b89cd-141">Ocultar a interface do usuário do suplemento do Office na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="b89cd-141">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="b89cd-142">Outro comportamento de suplemento de COM</span><span class="sxs-lookup"><span data-stu-id="b89cd-142">Other COM add-in behavior</span></span>

<span data-ttu-id="b89cd-143">Se um usuário desinstalar o suplemento COM equivalente, o Office no Windows restaura a interface do usuário do suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="b89cd-143">If a user uninstalls the equivalent COM add-in, then Office on Windows restores the Office Add-in UI.</span></span>

<span data-ttu-id="b89cd-144">Depois de especificar um suplemento COM equivalente para o suplemento do Office, o Office interrompe o processamento de atualizações para seu suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="b89cd-144">After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="b89cd-145">Para adquirir as atualizações mais recentes para o suplemento do Office, o usuário deve primeiro desinstalar o suplemento de COM.</span><span class="sxs-lookup"><span data-stu-id="b89cd-145">To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="b89cd-146">Confira também</span><span class="sxs-lookup"><span data-stu-id="b89cd-146">See also</span></span>

- [<span data-ttu-id="b89cd-147">Tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário XLL</span><span class="sxs-lookup"><span data-stu-id="b89cd-147">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
