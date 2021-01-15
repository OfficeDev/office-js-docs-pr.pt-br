---
title: Elemento Enabled no arquivo de manifesto
description: Saiba como especificar se um comando de suplemento está desabilitado quando o suplemento é iniciado.
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: be18767638af6f2be6352cea46739f6a01b7dd45
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771386"
---
# <a name="enabled-element"></a><span data-ttu-id="bfa33-103">Elemento Enabled</span><span class="sxs-lookup"><span data-stu-id="bfa33-103">Enabled element</span></span>

<span data-ttu-id="bfa33-104">Especifica se um controle de [botão](control.md#button-control) ou de [menu](control.md#menu-dropdown-button-controls) está habilitado quando o suplemento é iniciado.</span><span class="sxs-lookup"><span data-stu-id="bfa33-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="bfa33-105">O elemento **Enabled** é um elemento filho do [controle](control.md).</span><span class="sxs-lookup"><span data-stu-id="bfa33-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="bfa33-106">Se for omitido, o padrão será `true` .</span><span class="sxs-lookup"><span data-stu-id="bfa33-106">If it is omitted, the default is `true`.</span></span>

<span data-ttu-id="bfa33-107">Este elemento só é válido no Excel; ou seja, quando o `Name` atributo do elemento [host](host.md) é "Workbook".</span><span class="sxs-lookup"><span data-stu-id="bfa33-107">This element is only valid in Excel; that is, when the `Name` attribute of the [Host](host.md) element is "Workbook".</span></span>

<span data-ttu-id="bfa33-108">O controle pai também pode ser habilitado e desabilitado programaticamente.</span><span class="sxs-lookup"><span data-stu-id="bfa33-108">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="bfa33-109">Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](../../design/disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="bfa33-109">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

## <a name="example"></a><span data-ttu-id="bfa33-110">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bfa33-110">Example</span></span>

```xml
<Enabled>false</Enabled>
```
