---
title: Elemento DesktopFormFactor no arquivo de manifesto
description: Especifica as configurações de um suplemento para o fator forma da área de trabalho.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: bfea6900e6b07d8dc7ad5b5256703d873242d88c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718360"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="f0f07-103">Elemento DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="f0f07-103">DesktopFormFactor element</span></span>

<span data-ttu-id="f0f07-104">Especifica as configurações de um suplemento para o fator forma da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="f0f07-104">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="f0f07-105">O fator de forma da área de trabalho inclui o Office na Web, Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="f0f07-105">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="f0f07-106">Ele contém todas as informações do suplemento para o fator de forma da área de trabalho, exceto para o nó de **recursos** .</span><span class="sxs-lookup"><span data-stu-id="f0f07-106">It contains all the add-in information for the desktop form factor except for the **Resources** node.</span></span>

<span data-ttu-id="f0f07-107">Cada definição de DesktopFormFactor contém o elemento **functionfile** e um ou mais elementos **ExtensionPoint** .</span><span class="sxs-lookup"><span data-stu-id="f0f07-107">Each DesktopFormFactor definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="f0f07-108">Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="f0f07-108">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="f0f07-109">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="f0f07-109">Child elements</span></span>

| <span data-ttu-id="f0f07-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="f0f07-110">Element</span></span>                               | <span data-ttu-id="f0f07-111">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="f0f07-111">Required</span></span> | <span data-ttu-id="f0f07-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="f0f07-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="f0f07-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="f0f07-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="f0f07-114">Sim</span><span class="sxs-lookup"><span data-stu-id="f0f07-114">Yes</span></span>      | <span data-ttu-id="f0f07-115">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="f0f07-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="f0f07-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="f0f07-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="f0f07-117">Sim</span><span class="sxs-lookup"><span data-stu-id="f0f07-117">Yes</span></span>      | <span data-ttu-id="f0f07-118">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f0f07-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="f0f07-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="f0f07-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="f0f07-120">Não</span><span class="sxs-lookup"><span data-stu-id="f0f07-120">No</span></span>       | <span data-ttu-id="f0f07-121">Define o texto explicativo que aparece ao instalar o suplemento em hosts do Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="f0f07-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="f0f07-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="f0f07-122">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="f0f07-123">Não</span><span class="sxs-lookup"><span data-stu-id="f0f07-123">No</span></span> | <span data-ttu-id="f0f07-124">Define se o suplemento do Outlook está disponível em cenários de representante e é definido como *false* por padrão.</span><span class="sxs-lookup"><span data-stu-id="f0f07-124">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="f0f07-125">Exemplo de DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="f0f07-125">DesktopFormFactor example</span></span>

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
