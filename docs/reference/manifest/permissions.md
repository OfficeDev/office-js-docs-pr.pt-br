---
title: Elemento Permissions no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: a70d72e454273873c6a30ffd82c3a2a5194f55e0
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851303"
---
# <a name="permissions-element"></a><span data-ttu-id="b677b-102">Elemento Permissions</span><span class="sxs-lookup"><span data-stu-id="b677b-102">Permissions element</span></span>

<span data-ttu-id="b677b-103">Especifica o nível de acesso da API para seu Suplemento do Office. Você deve solicitar permissões com base no princípio do privilégio mínimo.</span><span class="sxs-lookup"><span data-stu-id="b677b-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="b677b-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="b677b-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b677b-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="b677b-105">Syntax</span></span>

<span data-ttu-id="b677b-106">Para suplementos de conteúdo e de painel de tarefas:</span><span class="sxs-lookup"><span data-stu-id="b677b-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="b677b-107">Para suplementos de email</span><span class="sxs-lookup"><span data-stu-id="b677b-107">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="b677b-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="b677b-108">Contained in</span></span>

[<span data-ttu-id="b677b-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b677b-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="b677b-110">Comentários</span><span class="sxs-lookup"><span data-stu-id="b677b-110">Remarks</span></span>

<span data-ttu-id="b677b-111">Para obter mais detalhes, consulte [solicitando permissões para uso da API em suplementos](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) e [noções básicas sobre permissões de suplementos do Outlook](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="b677b-111">For more detail, see [Requesting permissions for API use in add-ins](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
