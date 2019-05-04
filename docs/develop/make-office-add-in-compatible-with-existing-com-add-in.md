---
title: Tornar o suplemento do Office compatível com um suplemento de COM existente
description: Habilitar a compatibilidade com um suplemento COM equivalente que tenha a mesma funcionalidade do suplemento do Office
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 8f3780814163cc4dd21311b362d1d821a14b3e80
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356831"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="ad054-103">Tornar o suplemento do Office compatível com um suplemento de COM existente</span><span class="sxs-lookup"><span data-stu-id="ad054-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="ad054-104">Se você tiver um suplemento COM existente, poderá criar uma funcionalidade equivalente no suplemento do Office para estender seus recursos de solução para outras plataformas, como online ou macOS.</span><span class="sxs-lookup"><span data-stu-id="ad054-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="ad054-105">No enTanto, os suplementos do Office não possuem todas as funcionalidades disponíveis em suplementos de COM. O suplemento de COM pode fornecer uma experiência melhor do que o suplemento do Office no Windows no Excel, Word e PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="ad054-105">However, Office Add-ins don't have all of the functionality available in COM add-ins. Your COM add-in may provide a better experience than the Office Add-in on Windows in Excel, Word, and PowerPoint.</span></span>

<span data-ttu-id="ad054-106">Você pode configurar seu suplemento do Office para que, quando um suplemento COM equivalente já estiver instalado no computador do usuário, o Office execute o suplemento COM em vez do suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="ad054-106">You can configure your Office Add-in so that when an equivalent COM add-in is already installed on the user's computer, Office runs the COM add-in instead of your Office Add-in.</span></span> <span data-ttu-id="ad054-107">O suplemento de COM é chamado de "equivalente", pois o Office faz uma transição transparente entre o suplemento de COM e o suplemento do Office, dependendo do que está instalado no Windows.</span><span class="sxs-lookup"><span data-stu-id="ad054-107">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in depending on which is installed on Windows.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a><span data-ttu-id="ad054-108">Especificar um suplemento COM equivalente no manifesto</span><span class="sxs-lookup"><span data-stu-id="ad054-108">Specify an equivalent COM add-in in the manifest</span></span>

<span data-ttu-id="ad054-109">Para habilitar a compatibilidade com um suplemento de COM existente, identifique o suplemento COM equivalente no manifesto do suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="ad054-109">To enable compatibility with an existing COM add-in, identify the equivalent COM add-in in the manifest of your Office Add-in.</span></span> <span data-ttu-id="ad054-110">O Office usará o suplemento COM em vez do suplemento do Office ao ser executado no Windows.</span><span class="sxs-lookup"><span data-stu-id="ad054-110">Then Office will use the COM add-in instead of your Office Add-in when running on Windows.</span></span>

<span data-ttu-id="ad054-111">Especifique o `ProgID` do suplemento com equivalente.</span><span class="sxs-lookup"><span data-stu-id="ad054-111">Specify the `ProgID` of the equivalent COM add-in.</span></span> <span data-ttu-id="ad054-112">O Office usará a interface de usuário do suplemento COM em vez da interface do usuário do suplemento do Office quando o suplemento de COM estiver instalado.</span><span class="sxs-lookup"><span data-stu-id="ad054-112">Office will then use the COM add-in UI instead of your Office Add-in's UI when the COM add-in is installed.</span></span>

<span data-ttu-id="ad054-113">O exemplo a seguir mostra como especificar um suplemento de COM e um XLL como equivalente.</span><span class="sxs-lookup"><span data-stu-id="ad054-113">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="ad054-114">Em geral, você especifica tanto tanto quanto à integridade este exemplo mostra tanto no contexto.</span><span class="sxs-lookup"><span data-stu-id="ad054-114">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="ad054-115">Eles são identificados por seus `ProgID` e `FileName` , respectivamente.</span><span class="sxs-lookup"><span data-stu-id="ad054-115">They are identified by their `ProgID` and `FileName` respectively.</span></span> <span data-ttu-id="ad054-116">Para obter mais informações sobre a compatibilidade XLL, consulte [tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário do XLL](../excel/make-custom-functions-compatible-with-xll-udf.md).</span><span class="sxs-lookup"><span data-stu-id="ad054-116">For more information on XLL compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

```xml
<VersionOverrides>
...
<EquivalentAddins>
  <EquivalentAddin>
    <ProgID>ContosoCOMAddin</ProgID>
    <Type>COM</Type>
  </EquivalentAddin>

  <EquivalentAddin>
    <FileName>contosofunctions.xll</FileName>
    <Type>XLL</Type>
  </EquivalentAddin>
<EquivalentAddins>
...
</VersionOverrides>
```

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="ad054-117">Comportamento equivalente para usuários</span><span class="sxs-lookup"><span data-stu-id="ad054-117">Equivalent behavior for users</span></span>

<span data-ttu-id="ad054-118">Quando um suplemento COM equivalente é especificado no manifesto do suplemento do Office, o Office suprime a interface do usuário do suplemento do Office no Windows quando o suplemento COM equivalente está instalado.</span><span class="sxs-lookup"><span data-stu-id="ad054-118">When an equivalent COM add-in is specified in the Office Add-in manifest, Office suppresses your Office Add-in's UI on Windows when the equivalent COM add-in is installed.</span></span> <span data-ttu-id="ad054-119">Isso não afeta a interface do usuário do suplemento do Office em outras plataformas como online ou macOS.</span><span class="sxs-lookup"><span data-stu-id="ad054-119">This does not affect your Office Add-in's UI on other platforms like online or macOS.</span></span> <span data-ttu-id="ad054-120">O Office só oculta os botões da faixa de opções e não impede a instalação.</span><span class="sxs-lookup"><span data-stu-id="ad054-120">Office only hides the ribbon buttons and does not prevent installation.</span></span> <span data-ttu-id="ad054-121">Portanto, o suplemento do Office ainda aparecerá nos seguintes locais de interface do usuário:</span><span class="sxs-lookup"><span data-stu-id="ad054-121">Therefore your Office Add-in will still appear in the following UI locations:</span></span>

- <span data-ttu-id="ad054-122">Em **meus suplementos** , pois ele é tecnicamente instalado.</span><span class="sxs-lookup"><span data-stu-id="ad054-122">Under **My add-ins** because it is technically installed.</span></span>
- <span data-ttu-id="ad054-123">Como uma entrada no Gerenciador de faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="ad054-123">As an entry in the ribbon manager.</span></span>

<span data-ttu-id="ad054-124">Os cenários a seguir descrevem o que acontece dependendo de como o usuário adquire o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="ad054-124">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="ad054-125">Aquisição do AppSource de um suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="ad054-125">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="ad054-126">Se um usuário baixar o suplemento do Office do AppSource e o suplemento COM equivalente já estiver instalado, o Office irá:</span><span class="sxs-lookup"><span data-stu-id="ad054-126">If a user downloads the Office Add-in from AppSource, and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="ad054-127">Instalar o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="ad054-127">Install the Office Add-in.</span></span>
2. <span data-ttu-id="ad054-128">Ocultar a interface do usuário do suplemento do Office na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="ad054-128">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="ad054-129">Exibe uma chamada para o usuário que aponta o botão da faixa de opções suplemento de COM.</span><span class="sxs-lookup"><span data-stu-id="ad054-129">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="ad054-130">Implantação centralizada do suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="ad054-130">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="ad054-131">Se um administrador implantar o suplemento do Office em seu locatário usando a implantação centralizada e o suplemento COM equivalente já estiver instalado, o usuário precisará reiniciar o Office para que ele possa ver as alterações.</span><span class="sxs-lookup"><span data-stu-id="ad054-131">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user needs to restart Office before they will see any changes.</span></span> <span data-ttu-id="ad054-132">Após a reinicialização do Office, ela irá:</span><span class="sxs-lookup"><span data-stu-id="ad054-132">After Office restarts, it will:</span></span>

1. <span data-ttu-id="ad054-133">Instalar o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="ad054-133">Install the Office Add-in.</span></span>
2. <span data-ttu-id="ad054-134">Ocultar a interface do usuário do suplemento do Office na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="ad054-134">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="ad054-135">Exibe uma chamada para o usuário que aponta o botão da faixa de opções suplemento de COM.</span><span class="sxs-lookup"><span data-stu-id="ad054-135">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="ad054-136">Documento compartilhado com o suplemento incorporado do Office</span><span class="sxs-lookup"><span data-stu-id="ad054-136">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="ad054-137">Se um usuário tiver o suplemento COM instalado e, em seguida, receber um documento compartilhado com o suplemento do Office incorporado, quando abrir o documento, o Office irá:</span><span class="sxs-lookup"><span data-stu-id="ad054-137">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="ad054-138">Solicitar que o usuário confie no suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="ad054-138">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="ad054-139">Se confiável, o suplemento do Office será instalado.</span><span class="sxs-lookup"><span data-stu-id="ad054-139">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="ad054-140">Ocultar a interface do usuário do suplemento do Office na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="ad054-140">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="ad054-141">Outro comportamento de suplemento de COM</span><span class="sxs-lookup"><span data-stu-id="ad054-141">Other COM add-in behavior</span></span>

<span data-ttu-id="ad054-142">Se um usuário desinstala o suplemento de COM, o Office restaura a interface do usuário do suplemento do Office no Windows para o suplemento do Office instalado equivalente.</span><span class="sxs-lookup"><span data-stu-id="ad054-142">If a user uninstalls the COM add-in, then Office restores the Office Add-in UI on Windows for the equivalent installed Office Add-in.</span></span>

<span data-ttu-id="ad054-143">Após especificar um suplemento COM equivalente para o suplemento do Office, o Office interromperá o processamento de atualizações para seu suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="ad054-143">Once you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="ad054-144">O usuário deve desinstalar o suplemento de COM para obter as atualizações mais recentes para o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="ad054-144">The user must uninstall the COM add-in order to get the latest updates for the Office Add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="ad054-145">Confira também</span><span class="sxs-lookup"><span data-stu-id="ad054-145">See also</span></span>

- [<span data-ttu-id="ad054-146">Tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário XLL</span><span class="sxs-lookup"><span data-stu-id="ad054-146">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)