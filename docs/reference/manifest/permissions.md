---
title: Elemento Permissions no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 95cb45f89e2a5b92edc29bf32d0b47fcb2dbf8ce
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165542"
---
# <a name="permissions-element"></a><span data-ttu-id="feb40-102">Elemento Permissions</span><span class="sxs-lookup"><span data-stu-id="feb40-102">Permissions element</span></span>

<span data-ttu-id="feb40-103">Especifica o nível de acesso da API para seu Suplemento do Office. Você deve solicitar permissões com base no princípio do privilégio mínimo.</span><span class="sxs-lookup"><span data-stu-id="feb40-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="feb40-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="feb40-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="feb40-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="feb40-105">Syntax</span></span>

<span data-ttu-id="feb40-106">Para suplementos de conteúdo e de painel de tarefas:</span><span class="sxs-lookup"><span data-stu-id="feb40-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="feb40-107">Para suplementos de email</span><span class="sxs-lookup"><span data-stu-id="feb40-107">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="feb40-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="feb40-108">Contained in</span></span>

[<span data-ttu-id="feb40-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="feb40-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="feb40-110">Comentários</span><span class="sxs-lookup"><span data-stu-id="feb40-110">Remarks</span></span>

<span data-ttu-id="feb40-111">Para obter mais detalhes, consulte [solicitando permissões para uso da API em suplementos](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) e [noções básicas sobre permissões de suplementos do Outlook](../../outlook/understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="feb40-111">For more detail, see [Requesting permissions for API use in add-ins](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md).</span></span>
