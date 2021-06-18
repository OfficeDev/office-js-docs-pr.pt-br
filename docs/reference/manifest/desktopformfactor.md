---
title: Elemento DesktopFormFactor no arquivo de manifesto
description: Especifica as configurações de um suplemento para o fator forma da área de trabalho.
ms.date: 06/15/2021
localization_priority: Normal
ms.openlocfilehash: 66673d83fd8608a1ec10492d7a944b0515de61c0
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007786"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="cd202-103">Elemento DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="cd202-103">DesktopFormFactor element</span></span>

<span data-ttu-id="cd202-104">Especifica as configurações de um suplemento para o fator forma da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="cd202-104">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="cd202-105">O fator de formulário da área de trabalho inclui Office na Web, Windows e Mac.</span><span class="sxs-lookup"><span data-stu-id="cd202-105">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="cd202-106">Ele contém todas as informações de complemento para o fator de formulário da área de trabalho, exceto para o **nó Recursos.**</span><span class="sxs-lookup"><span data-stu-id="cd202-106">It contains all the add-in information for the desktop form factor except for the **Resources** node.</span></span>

<span data-ttu-id="cd202-107">Cada definição desktopFormFactor contém o **elemento FunctionFile** e um ou mais **elementos ExtensionPoint.**</span><span class="sxs-lookup"><span data-stu-id="cd202-107">Each DesktopFormFactor definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="cd202-108">Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="cd202-108">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="cd202-109">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="cd202-109">Child elements</span></span>

| <span data-ttu-id="cd202-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="cd202-110">Element</span></span>                               | <span data-ttu-id="cd202-111">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="cd202-111">Required</span></span> | <span data-ttu-id="cd202-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="cd202-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="cd202-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="cd202-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="cd202-114">Sim</span><span class="sxs-lookup"><span data-stu-id="cd202-114">Yes</span></span>      | <span data-ttu-id="cd202-115">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="cd202-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="cd202-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="cd202-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="cd202-117">Sim</span><span class="sxs-lookup"><span data-stu-id="cd202-117">Yes</span></span>      | <span data-ttu-id="cd202-118">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="cd202-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="cd202-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="cd202-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="cd202-120">Não</span><span class="sxs-lookup"><span data-stu-id="cd202-120">No</span></span>       | <span data-ttu-id="cd202-121">Define o texto explicante que aparece ao instalar o complemento no Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="cd202-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint.</span></span> |
| [<span data-ttu-id="cd202-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="cd202-122">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="cd202-123">Não</span><span class="sxs-lookup"><span data-stu-id="cd202-123">No</span></span> | <span data-ttu-id="cd202-124">Define se o Outlook está disponível em cenários de caixa de correio compartilhada (agora em visualização) e pastas compartilhadas (ou seja, acesso de representante).</span><span class="sxs-lookup"><span data-stu-id="cd202-124">Defines whether the Outlook add-in is available in shared mailbox (now in preview) and shared folders (that is, delegate access) scenarios.</span></span> <span data-ttu-id="cd202-125">Definir como *false* por padrão.</span><span class="sxs-lookup"><span data-stu-id="cd202-125">Set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="cd202-126">Exemplo de DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="cd202-126">DesktopFormFactor example</span></span>

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
