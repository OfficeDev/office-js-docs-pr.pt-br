---
title: Elemento Permissions no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3442a8e0caee442ce1b38c5ff39cfd1ef5088fb7
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450657"
---
# <a name="permissions-element"></a><span data-ttu-id="f240d-102">Elemento Permissions</span><span class="sxs-lookup"><span data-stu-id="f240d-102">Permissions element</span></span>

<span data-ttu-id="f240d-103">Especifica o nível de acesso da API para seu Suplemento do Office. Você deve solicitar permissões com base no princípio do privilégio mínimo.</span><span class="sxs-lookup"><span data-stu-id="f240d-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="f240d-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="f240d-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f240d-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="f240d-105">Syntax</span></span>

<span data-ttu-id="f240d-106">Para suplementos de conteúdo e de painel de tarefas:</span><span class="sxs-lookup"><span data-stu-id="f240d-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="f240d-107">Para suplementos de email</span><span class="sxs-lookup"><span data-stu-id="f240d-107">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="f240d-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="f240d-108">Contained in</span></span>

[<span data-ttu-id="f240d-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="f240d-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="f240d-110">Comentários</span><span class="sxs-lookup"><span data-stu-id="f240d-110">Remarks</span></span>

<span data-ttu-id="f240d-111">Para saber mais, confira [Solicitando permissões para API usadas em suplementos de conteúdo e de painel de tarefas](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) e [Entendendo as permissões de suplemento do Outlook](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="f240d-111">For more detail, see [Requesting permissions for API use in content and task pane add-ins](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
