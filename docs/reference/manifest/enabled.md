---
title: Elemento Enabled no arquivo de manifesto
description: Saiba como especificar se um comando de suplemento está desabilitado quando o suplemento é iniciado.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 4c2c013c8e55966ba2678755536ce04ae3014ed0
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596897"
---
# <a name="enabled-element"></a><span data-ttu-id="dac5d-103">Elemento Enabled</span><span class="sxs-lookup"><span data-stu-id="dac5d-103">Enabled element</span></span>

<span data-ttu-id="dac5d-104">Especifica se um controle de [botão](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls) está habilitado quando o suplemento é iniciado.</span><span class="sxs-lookup"><span data-stu-id="dac5d-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="dac5d-105">O elemento **Enabled** é um elemento filho do [controle](control.md).</span><span class="sxs-lookup"><span data-stu-id="dac5d-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="dac5d-106">Se for omitido, o padrão será `true`.</span><span class="sxs-lookup"><span data-stu-id="dac5d-106">If it is omitted, the default is `true`.</span></span>

<span data-ttu-id="dac5d-107">O controle pai também pode ser habilitado e desabilitado programaticamente.</span><span class="sxs-lookup"><span data-stu-id="dac5d-107">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="dac5d-108">Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](../../design/disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="dac5d-108">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

## <a name="example"></a><span data-ttu-id="dac5d-109">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dac5d-109">Example</span></span>

```xml
<Enabled>false</Enabled>
```
