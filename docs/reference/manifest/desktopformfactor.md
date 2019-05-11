---
title: Elemento DesktopFormFactor no arquivo de manifesto
description: ''
ms.date: 05/08/2019
localization_priority: Normal
ms.openlocfilehash: b46536886d59692d03976083412a8b8d2e6ae859
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952387"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="f7d06-102">Elemento DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="f7d06-102">DesktopFormFactor element</span></span>

<span data-ttu-id="f7d06-103">Especifica as configurações de um suplemento para o fator forma da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="f7d06-103">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="f7d06-104">O fator de forma da área de trabalho inclui o Office no Windows, o Office para Mac e o Office Online.</span><span class="sxs-lookup"><span data-stu-id="f7d06-104">The desktop form factor includes Office on Windows, Office for Mac, and Office Online.</span></span> <span data-ttu-id="f7d06-105">Ele contém todas as informações do suplemento para o fator forma da área de trabalho, exceto para o nó **Resources**.</span><span class="sxs-lookup"><span data-stu-id="f7d06-105">It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="f7d06-p102">Cada definição de DesktopFormFactor contém o elemento **FunctionFile** e um ou mais elementos **ExtensionPoint**. Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="f7d06-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="f7d06-108">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="f7d06-108">Child elements</span></span>

| <span data-ttu-id="f7d06-109">Elemento</span><span class="sxs-lookup"><span data-stu-id="f7d06-109">Element</span></span>                               | <span data-ttu-id="f7d06-110">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="f7d06-110">Required</span></span> | <span data-ttu-id="f7d06-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="f7d06-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="f7d06-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="f7d06-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="f7d06-113">Sim</span><span class="sxs-lookup"><span data-stu-id="f7d06-113">Yes</span></span>      | <span data-ttu-id="f7d06-114">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="f7d06-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="f7d06-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="f7d06-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="f7d06-116">Sim</span><span class="sxs-lookup"><span data-stu-id="f7d06-116">Yes</span></span>      | <span data-ttu-id="f7d06-117">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f7d06-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="f7d06-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="f7d06-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="f7d06-119">Não</span><span class="sxs-lookup"><span data-stu-id="f7d06-119">No</span></span>       | <span data-ttu-id="f7d06-120">Define o texto explicativo que aparece ao instalar o suplemento em hosts do Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="f7d06-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="f7d06-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="f7d06-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="f7d06-122">Não</span><span class="sxs-lookup"><span data-stu-id="f7d06-122">No</span></span> | <span data-ttu-id="f7d06-123">Define se o suplemento do Outlook está disponível em cenários de representante e é definido como *false* por padrão.</span><span class="sxs-lookup"><span data-stu-id="f7d06-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span><br><br><span data-ttu-id="f7d06-124">**Importante**: como o acesso de representante para suplementos do Outlook está atualmente em versão prévia, os suplementos que usam `SupportSharedFolders` o elemento não podem ser publicados no AppSource ou implantados por meio da implantação centralizada.</span><span class="sxs-lookup"><span data-stu-id="f7d06-124">**Important**: Because delegate access for Outlook add-ins is currently in preview, add-ins that use the `SupportSharedFolders` element cannot be published to AppSource or deployed via centralized deployment.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="f7d06-125">Exemplo de DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="f7d06-125">DesktopFormFactor example</span></span>

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
