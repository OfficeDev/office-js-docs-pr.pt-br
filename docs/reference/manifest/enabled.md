---
title: Elemento Enabled no arquivo de manifesto
description: Saiba como especificar se um comando de suplemento está desabilitado quando o suplemento é iniciado.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: a47ab97ff5a159c73bea52f130ce0c16efe2b6b6
ms.sourcegitcommit: 0e7ed44019d6564c79113639af831ea512fa0a13
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/09/2020
ms.locfileid: "42566181"
---
# <a name="enabled-element"></a><span data-ttu-id="a27d0-103">Elemento Enabled</span><span class="sxs-lookup"><span data-stu-id="a27d0-103">Enabled element</span></span>

<span data-ttu-id="a27d0-104">Especifica se um controle de [botão](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls) está habilitado quando o suplemento é iniciado.</span><span class="sxs-lookup"><span data-stu-id="a27d0-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="a27d0-105">O elemento **Enabled** é um elemento filho do [controle](control.md).</span><span class="sxs-lookup"><span data-stu-id="a27d0-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="a27d0-106">Se for omitido, o padrão será `true`.</span><span class="sxs-lookup"><span data-stu-id="a27d0-106">If it is omitted, the default is `true`.</span></span> 

<span data-ttu-id="a27d0-107">O controle pai também pode ser habilitado e desabilitado programaticamente.</span><span class="sxs-lookup"><span data-stu-id="a27d0-107">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="a27d0-108">Para obter mais informações, consulte [habilitar e desabilitar comandos de suplemento](/office/dev/add-ins/design/disable-add-in-commands).</span><span class="sxs-lookup"><span data-stu-id="a27d0-108">For more information, see [Enable and Disable Add-in Commands](/office/dev/add-ins/design/disable-add-in-commands).</span></span>

## <a name="example"></a><span data-ttu-id="a27d0-109">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a27d0-109">Example</span></span>

```xml
<Enabled>false</Enabled>
```

