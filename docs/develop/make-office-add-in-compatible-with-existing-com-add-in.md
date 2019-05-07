---
title: Tornar seu suplemento do Excel compatível com um suplemento de COM existente
description: Habilitar a compatibilidade com um suplemento COM equivalente que tenha a mesma funcionalidade do seu suplemento do Excel
ms.date: 05/06/2019
localization_priority: Normal
ms.openlocfilehash: 0890e14466a2cd8f5aff2d1bcf307a43cff28127
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628169"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in-preview"></a><span data-ttu-id="3dc92-103">Tornar o suplemento do Office compatível com um suplemento de COM existente (visualização)</span><span class="sxs-lookup"><span data-stu-id="3dc92-103">Make your Office Add-in compatible with an existing COM add-in (preview)</span></span>

<span data-ttu-id="3dc92-104">Se você tiver um suplemento COM existente, poderá criar uma funcionalidade equivalente no suplemento do Excel para estender seus recursos de solução para outras plataformas, como online ou macOS.</span><span class="sxs-lookup"><span data-stu-id="3dc92-104">If you have an existing COM add-in, you can build equivalent functionality in your Excel add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="3dc92-105">No entanto, os suplementos do Excel não possuem todas as funcionalidades disponíveis em suplementos de COM. O suplemento de COM pode fornecer uma experiência melhor do que o suplemento do Excel no Windows.</span><span class="sxs-lookup"><span data-stu-id="3dc92-105">However, Excel add-ins don't have all of the functionality available in COM add-ins. Your COM add-in may provide a better experience than the Excel add-in on Windows.</span></span>

<span data-ttu-id="3dc92-106">Você pode configurar seu suplemento do Excel para que, quando um suplemento COM equivalente já estiver instalado no computador do usuário, o Office execute o suplemento COM em vez do suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="3dc92-106">You can configure your Excel add-in so that when an equivalent COM add-in is already installed on the user's computer, Office runs the COM add-in instead of your Excel add-in.</span></span> <span data-ttu-id="3dc92-107">O suplemento de COM é chamado de "equivalente", pois o Office faz uma transição transparente entre o suplemento de COM e o suplemento do Excel, dependendo do que está instalado no Windows.</span><span class="sxs-lookup"><span data-stu-id="3dc92-107">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Excel add-in depending on which is installed on Windows.</span></span>

[!include[COM add-in and XLL UDF compatibility requirements note](../includes/xll-compatibility-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a><span data-ttu-id="3dc92-108">Especificar um suplemento COM equivalente no manifesto</span><span class="sxs-lookup"><span data-stu-id="3dc92-108">Specify an equivalent COM add-in in the manifest</span></span>

<span data-ttu-id="3dc92-109">Para habilitar a compatibilidade com um suplemento de COM existente, identifique o suplemento COM equivalente no manifesto do suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="3dc92-109">To enable compatibility with an existing COM add-in, identify the equivalent COM add-in in the manifest of your Excel add-in.</span></span> <span data-ttu-id="3dc92-110">Em seguida, o Office usará o suplemento COM em vez do seu suplemento do Excel ao executar o Windows.</span><span class="sxs-lookup"><span data-stu-id="3dc92-110">Then Office will use the COM add-in instead of your Excel add-in when running on Windows.</span></span>

<span data-ttu-id="3dc92-111">Especifique o `ProgID` do suplemento com equivalente.</span><span class="sxs-lookup"><span data-stu-id="3dc92-111">Specify the `ProgID` of the equivalent COM add-in.</span></span> <span data-ttu-id="3dc92-112">O Office usará a interface de usuário do suplemento COM em vez da interface do usuário do suplemento do Excel quando o suplemento de COM estiver instalado.</span><span class="sxs-lookup"><span data-stu-id="3dc92-112">Office will then use the COM add-in UI instead of your Excel add-in's UI when the COM add-in is installed.</span></span>

<span data-ttu-id="3dc92-113">O exemplo a seguir mostra como especificar um suplemento de COM e um XLL como equivalente.</span><span class="sxs-lookup"><span data-stu-id="3dc92-113">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="3dc92-114">Em geral, você especifica tanto tanto quanto à integridade este exemplo mostra tanto no contexto.</span><span class="sxs-lookup"><span data-stu-id="3dc92-114">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="3dc92-115">Eles são identificados por seus `ProgID` e `FileName` , respectivamente.</span><span class="sxs-lookup"><span data-stu-id="3dc92-115">They are identified by their `ProgID` and `FileName` respectively.</span></span> <span data-ttu-id="3dc92-116">Para obter mais informações sobre a compatibilidade XLL, consulte [tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário do XLL](../excel/make-custom-functions-compatible-with-xll-udf.md).</span><span class="sxs-lookup"><span data-stu-id="3dc92-116">For more information on XLL compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

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

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="3dc92-117">Comportamento equivalente para usuários</span><span class="sxs-lookup"><span data-stu-id="3dc92-117">Equivalent behavior for users</span></span>

<span data-ttu-id="3dc92-118">Quando um suplemento COM equivalente é especificado no manifesto de suplemento do Excel, o Office suprime sua interface do usuário do suplemento do Excel no Windows quando o suplemento COM equivalente está instalado.</span><span class="sxs-lookup"><span data-stu-id="3dc92-118">When an equivalent COM add-in is specified in the Excel add-in manifest, Office suppresses your Excel add-in's UI on Windows when the equivalent COM add-in is installed.</span></span> <span data-ttu-id="3dc92-119">Isso não afeta a interface do usuário do seu suplemento do Excel em outras plataformas como online ou macOS.</span><span class="sxs-lookup"><span data-stu-id="3dc92-119">This does not affect your Excel add-in's UI on other platforms like online or macOS.</span></span> <span data-ttu-id="3dc92-120">O Office só oculta os botões da faixa de opções e não impede a instalação.</span><span class="sxs-lookup"><span data-stu-id="3dc92-120">Office only hides the ribbon buttons and does not prevent installation.</span></span> <span data-ttu-id="3dc92-121">Portanto, o suplemento do Excel ainda aparecerá nos seguintes locais de interface do usuário:</span><span class="sxs-lookup"><span data-stu-id="3dc92-121">Therefore your Excel add-in will still appear in the following UI locations:</span></span>

- <span data-ttu-id="3dc92-122">Em **meus suplementos** , pois ele é tecnicamente instalado.</span><span class="sxs-lookup"><span data-stu-id="3dc92-122">Under **My add-ins** because it is technically installed.</span></span>
- <span data-ttu-id="3dc92-123">Como uma entrada no Gerenciador de faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="3dc92-123">As an entry in the ribbon manager.</span></span>

<span data-ttu-id="3dc92-124">Os cenários a seguir descrevem o que acontece dependendo de como o usuário adquire o suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="3dc92-124">The following scenarios describe what happens depending on how the user acquires the Excel add-in.</span></span>

### <a name="appsource-acquisition-of-an-excel-add-in"></a><span data-ttu-id="3dc92-125">AppSource aquisição de um suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="3dc92-125">AppSource acquisition of an Excel add-in</span></span>

<span data-ttu-id="3dc92-126">Se um usuário baixar o suplemento do Excel do AppSource e o suplemento COM equivalente já estiver instalado, o Office irá:</span><span class="sxs-lookup"><span data-stu-id="3dc92-126">If a user downloads the Excel add-in from AppSource, and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="3dc92-127">Instalar o suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="3dc92-127">Install the Excel add-in.</span></span>
2. <span data-ttu-id="3dc92-128">Ocultar a interface do usuário do suplemento do Excel na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="3dc92-128">Hide the Excel add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="3dc92-129">Exibe uma chamada para o usuário que aponta o botão da faixa de opções suplemento de COM.</span><span class="sxs-lookup"><span data-stu-id="3dc92-129">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-excel-add-in"></a><span data-ttu-id="3dc92-130">Implantação centralizada do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="3dc92-130">Centralized deployment of Excel add-in</span></span>

<span data-ttu-id="3dc92-131">Se um administrador implantar o suplemento do Excel em seu locatário usando a implantação centralizada e o suplemento COM equivalente já estiver instalado, o usuário precisará reiniciar o Office para que ele possa ver as alterações.</span><span class="sxs-lookup"><span data-stu-id="3dc92-131">If an admin deploys the Excel add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user needs to restart Office before they will see any changes.</span></span> <span data-ttu-id="3dc92-132">Após a reinicialização do Office, ela irá:</span><span class="sxs-lookup"><span data-stu-id="3dc92-132">After Office restarts, it will:</span></span>

1. <span data-ttu-id="3dc92-133">Instalar o suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="3dc92-133">Install the Excel add-in.</span></span>
2. <span data-ttu-id="3dc92-134">Ocultar a interface do usuário do suplemento do Excel na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="3dc92-134">Hide the Excel add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="3dc92-135">Exibe uma chamada para o usuário que aponta o botão da faixa de opções suplemento de COM.</span><span class="sxs-lookup"><span data-stu-id="3dc92-135">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-excel-add-in"></a><span data-ttu-id="3dc92-136">Documento compartilhado com o suplemento incorporado do Excel</span><span class="sxs-lookup"><span data-stu-id="3dc92-136">Document shared with embedded Excel add-in</span></span>

<span data-ttu-id="3dc92-137">Se um usuário tiver o suplemento COM instalado e, em seguida, receber um documento compartilhado com o suplemento do Excel incorporado, quando abrir o documento, o Office irá:</span><span class="sxs-lookup"><span data-stu-id="3dc92-137">If a user has the COM add-in installed, and then gets a shared document with the embedded Excel add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="3dc92-138">Solicitar que o usuário confie no suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="3dc92-138">Prompt the user to trust the Excel add-in.</span></span>
2. <span data-ttu-id="3dc92-139">Se confiável, o suplemento do Excel será instalado.</span><span class="sxs-lookup"><span data-stu-id="3dc92-139">If trusted, the Excel add-in will install.</span></span>
3. <span data-ttu-id="3dc92-140">Ocultar a interface do usuário do suplemento do Excel na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="3dc92-140">Hide the Excel add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="3dc92-141">Outro comportamento de suplemento de COM</span><span class="sxs-lookup"><span data-stu-id="3dc92-141">Other COM add-in behavior</span></span>

<span data-ttu-id="3dc92-142">Se um usuário desinstala o suplemento de COM, o Office restaura a interface de usuário do suplemento do Excel no Windows para o suplemento do Excel instalado equivalente.</span><span class="sxs-lookup"><span data-stu-id="3dc92-142">If a user uninstalls the COM add-in, then Office restores the Excel add-in UI on Windows for the equivalent installed Excel add-in.</span></span>

<span data-ttu-id="3dc92-143">Depois de especificar um suplemento de COM equivalente para seu suplemento do Excel, o Office interrompe o processamento de atualizações para seu suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="3dc92-143">Once you specify an equivalent COM add-in for your Excel add-in, Office stops processing updates for your Excel add-in.</span></span> <span data-ttu-id="3dc92-144">O usuário deve desinstalar o suplemento de COM para obter as atualizações mais recentes para o suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="3dc92-144">The user must uninstall the COM add-in order to get the latest updates for the Excel add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="3dc92-145">Confira também</span><span class="sxs-lookup"><span data-stu-id="3dc92-145">See also</span></span>

- [<span data-ttu-id="3dc92-146">Tornar suas funções personalizadas compatíveis com as funções definidas pelo usuário XLL</span><span class="sxs-lookup"><span data-stu-id="3dc92-146">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
