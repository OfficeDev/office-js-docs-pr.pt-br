---
title: Elemento DesktopFormFactor no arquivo de manifesto
description: Especifica as configurações de um suplemento para o fator forma da área de trabalho.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 18828e6b61a45ae2dc1528b3f7a54e664af09519
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292311"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="97478-103">Elemento DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="97478-103">DesktopFormFactor element</span></span>

<span data-ttu-id="97478-104">Especifica as configurações de um suplemento para o fator forma da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="97478-104">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="97478-105">O fator de forma da área de trabalho inclui o Office na Web, Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="97478-105">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="97478-106">Ele contém todas as informações do suplemento para o fator de forma da área de trabalho, exceto para o nó de **recursos** .</span><span class="sxs-lookup"><span data-stu-id="97478-106">It contains all the add-in information for the desktop form factor except for the **Resources** node.</span></span>

<span data-ttu-id="97478-107">Cada definição de DesktopFormFactor contém o elemento **functionfile** e um ou mais elementos **ExtensionPoint** .</span><span class="sxs-lookup"><span data-stu-id="97478-107">Each DesktopFormFactor definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="97478-108">Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="97478-108">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="97478-109">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="97478-109">Child elements</span></span>

| <span data-ttu-id="97478-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="97478-110">Element</span></span>                               | <span data-ttu-id="97478-111">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="97478-111">Required</span></span> | <span data-ttu-id="97478-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="97478-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="97478-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="97478-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="97478-114">Sim</span><span class="sxs-lookup"><span data-stu-id="97478-114">Yes</span></span>      | <span data-ttu-id="97478-115">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="97478-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="97478-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="97478-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="97478-117">Sim</span><span class="sxs-lookup"><span data-stu-id="97478-117">Yes</span></span>      | <span data-ttu-id="97478-118">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="97478-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="97478-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="97478-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="97478-120">Não</span><span class="sxs-lookup"><span data-stu-id="97478-120">No</span></span>       | <span data-ttu-id="97478-121">Define o texto explicativo que aparece ao instalar o suplemento no Word, no Excel ou no PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="97478-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint.</span></span> |
| [<span data-ttu-id="97478-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="97478-122">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="97478-123">Não</span><span class="sxs-lookup"><span data-stu-id="97478-123">No</span></span> | <span data-ttu-id="97478-124">Define se o suplemento do Outlook está disponível nos cenários de representante.</span><span class="sxs-lookup"><span data-stu-id="97478-124">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="97478-125">Definido como *false* por padrão.</span><span class="sxs-lookup"><span data-stu-id="97478-125">Set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="97478-126">Exemplo de DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="97478-126">DesktopFormFactor example</span></span>

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
        <!-- Information on this extension point. -->
      </ExtensionPoint>
      <!-- Possibly more ExtensionPoint elements. -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
