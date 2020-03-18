---
title: Elemento Permissions no arquivo de manifesto
description: O elemento Permissions especifica o nível de acesso da API para o suplemento do Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 91e024a2f13ea7605941c8c17a642f325cbcd61d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717996"
---
# <a name="permissions-element"></a><span data-ttu-id="56e32-103">Elemento Permissions</span><span class="sxs-lookup"><span data-stu-id="56e32-103">Permissions element</span></span>

<span data-ttu-id="56e32-104">Especifica o nível de acesso da API para seu Suplemento do Office. Você deve solicitar permissões com base no princípio do privilégio mínimo.</span><span class="sxs-lookup"><span data-stu-id="56e32-104">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="56e32-105">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="56e32-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="56e32-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="56e32-106">Syntax</span></span>

<span data-ttu-id="56e32-107">Para suplementos de conteúdo e de painel de tarefas:</span><span class="sxs-lookup"><span data-stu-id="56e32-107">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="56e32-108">Para suplementos de email</span><span class="sxs-lookup"><span data-stu-id="56e32-108">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="56e32-109">Contido em</span><span class="sxs-lookup"><span data-stu-id="56e32-109">Contained in</span></span>

[<span data-ttu-id="56e32-110">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="56e32-110">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="56e32-111">Comentários</span><span class="sxs-lookup"><span data-stu-id="56e32-111">Remarks</span></span>

<span data-ttu-id="56e32-112">Para obter mais detalhes, consulte [solicitando permissões para uso da API em suplementos](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) e [noções básicas sobre permissões de suplementos do Outlook](../../outlook/understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="56e32-112">For more detail, see [Requesting permissions for API use in add-ins](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md).</span></span>
