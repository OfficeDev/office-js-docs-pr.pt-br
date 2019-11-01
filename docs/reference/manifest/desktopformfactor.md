---
title: Elemento DesktopFormFactor no arquivo de manifesto
description: ''
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: bada3cd4cff7973517aedb83235a224ef6c273eb
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37901959"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="35030-102">Elemento DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="35030-102">DesktopFormFactor element</span></span>

<span data-ttu-id="35030-103">Especifica as configurações de um suplemento para o fator forma da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="35030-103">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="35030-104">O fator de forma da área de trabalho inclui o Office na Web, Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="35030-104">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="35030-105">Ele contém todas as informações do suplemento para o fator forma da área de trabalho, exceto para o nó **Resources**.</span><span class="sxs-lookup"><span data-stu-id="35030-105">It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="35030-p102">Cada definição de DesktopFormFactor contém o elemento **FunctionFile** e um ou mais elementos **ExtensionPoint**. Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="35030-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="35030-108">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="35030-108">Child elements</span></span>

| <span data-ttu-id="35030-109">Elemento</span><span class="sxs-lookup"><span data-stu-id="35030-109">Element</span></span>                               | <span data-ttu-id="35030-110">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="35030-110">Required</span></span> | <span data-ttu-id="35030-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="35030-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="35030-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="35030-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="35030-113">Sim</span><span class="sxs-lookup"><span data-stu-id="35030-113">Yes</span></span>      | <span data-ttu-id="35030-114">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="35030-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="35030-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="35030-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="35030-116">Sim</span><span class="sxs-lookup"><span data-stu-id="35030-116">Yes</span></span>      | <span data-ttu-id="35030-117">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="35030-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="35030-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="35030-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="35030-119">Não</span><span class="sxs-lookup"><span data-stu-id="35030-119">No</span></span>       | <span data-ttu-id="35030-120">Define o texto explicativo que aparece ao instalar o suplemento em hosts do Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="35030-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="35030-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="35030-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="35030-122">Não</span><span class="sxs-lookup"><span data-stu-id="35030-122">No</span></span> | <span data-ttu-id="35030-123">Define se o suplemento do Outlook está disponível em cenários de representante e é definido como *false* por padrão.</span><span class="sxs-lookup"><span data-stu-id="35030-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="35030-124">Exemplo de DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="35030-124">DesktopFormFactor example</span></span>

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
