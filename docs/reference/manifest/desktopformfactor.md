---
title: Elemento DesktopFormFactor no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: dea632f7f8afa5d9b69f257798022e9e520e9394
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433737"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="3f063-102">Elemento DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="3f063-102">DesktopFormFactor element</span></span>

<span data-ttu-id="3f063-p101">Especifica as configurações de um suplemento para o fator forma da área de trabalho. O fator de forma da área de trabalho inclui o Office para Windows, Office para Mac e Office Online. Ele contém todas as informações do suplemento para o fator forma da área de trabalho, exceto para o nó **Resources**.</span><span class="sxs-lookup"><span data-stu-id="3f063-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="3f063-p102">Cada definição de DesktopFormFactor contém o elemento **FunctionFile** e um ou mais elementos **ExtensionPoint**. Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="3f063-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="3f063-108">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3f063-108">Child elements</span></span>

| <span data-ttu-id="3f063-109">Elemento</span><span class="sxs-lookup"><span data-stu-id="3f063-109">Element</span></span>                               | <span data-ttu-id="3f063-110">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3f063-110">Required</span></span> | <span data-ttu-id="3f063-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="3f063-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="3f063-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="3f063-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="3f063-113">Sim</span><span class="sxs-lookup"><span data-stu-id="3f063-113">Yes</span></span>      | <span data-ttu-id="3f063-114">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="3f063-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="3f063-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="3f063-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="3f063-116">Sim</span><span class="sxs-lookup"><span data-stu-id="3f063-116">Yes</span></span>      | <span data-ttu-id="3f063-117">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3f063-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="3f063-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="3f063-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="3f063-119">Não</span><span class="sxs-lookup"><span data-stu-id="3f063-119">No</span></span>       | <span data-ttu-id="3f063-120">Define o texto explicativo que aparece ao instalar o suplemento em hosts do Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="3f063-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="3f063-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="3f063-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="3f063-122">Não</span><span class="sxs-lookup"><span data-stu-id="3f063-122">No</span></span> | <span data-ttu-id="3f063-123">Define se o suplemento do Outlook está disponível em cenários de representante e é definido como *false* por padrão.</span><span class="sxs-lookup"><span data-stu-id="3f063-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span><br><br><span data-ttu-id="3f063-124">**Importante**: esse elemento só está disponível no conjunto de requisitos de visualização de suplementos do Outlook em comparação com o Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="3f063-124">**Important**: This element is only available in the Outlook add-ins Preview requirement set against Exchange Online.</span></span> <span data-ttu-id="3f063-125">Os suplementos que usam esse elemento não podem ser publicados no AppSource ou implantados por meio da implantação centralizada.</span><span class="sxs-lookup"><span data-stu-id="3f063-125">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="3f063-126">Exemplo de DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="3f063-126">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
