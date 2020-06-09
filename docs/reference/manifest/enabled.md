---
title: Elemento Enabled no arquivo de manifesto
description: Saiba como especificar se um comando de suplemento está desabilitado quando o suplemento é iniciado.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 2849689fec99190c3a9b039c6c04069bc8194ee1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611565"
---
# <a name="enabled-element"></a><span data-ttu-id="fcafe-103">Elemento Enabled</span><span class="sxs-lookup"><span data-stu-id="fcafe-103">Enabled element</span></span>

<span data-ttu-id="fcafe-104">Especifica se um controle de [botão](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls) está habilitado quando o suplemento é iniciado.</span><span class="sxs-lookup"><span data-stu-id="fcafe-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="fcafe-105">O elemento **Enabled** é um elemento filho do [controle](control.md).</span><span class="sxs-lookup"><span data-stu-id="fcafe-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="fcafe-106">Se for omitido, o padrão será `true` .</span><span class="sxs-lookup"><span data-stu-id="fcafe-106">If it is omitted, the default is `true`.</span></span>

<span data-ttu-id="fcafe-107">O controle pai também pode ser habilitado e desabilitado programaticamente.</span><span class="sxs-lookup"><span data-stu-id="fcafe-107">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="fcafe-108">Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](../../design/disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="fcafe-108">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

## <a name="example"></a><span data-ttu-id="fcafe-109">Exemplo</span><span class="sxs-lookup"><span data-stu-id="fcafe-109">Example</span></span>

```xml
<Enabled>false</Enabled>
```
