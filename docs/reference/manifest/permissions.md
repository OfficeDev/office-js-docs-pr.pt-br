---
title: Elemento Permissions no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 9193651ec0c795cdb55eb3fc6576dbacd59e0fb2
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432351"
---
# <a name="permissions-element"></a><span data-ttu-id="2a9ff-102">Elemento Permissions</span><span class="sxs-lookup"><span data-stu-id="2a9ff-102">Permissions element</span></span>

<span data-ttu-id="2a9ff-103">Especifica o nível de acesso da API para seu Suplemento do Office. Você deve solicitar permissões com base no princípio do privilégio mínimo.</span><span class="sxs-lookup"><span data-stu-id="2a9ff-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="2a9ff-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="2a9ff-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="2a9ff-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="2a9ff-105">Syntax</span></span>

<span data-ttu-id="2a9ff-106">Para suplementos de conteúdo e de painel de tarefas:</span><span class="sxs-lookup"><span data-stu-id="2a9ff-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="2a9ff-107">Para suplementos de email</span><span class="sxs-lookup"><span data-stu-id="2a9ff-107">For mail add-ins:</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="2a9ff-108">Contido em</span><span class="sxs-lookup"><span data-stu-id="2a9ff-108">Contained in</span></span>

[<span data-ttu-id="2a9ff-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="2a9ff-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="2a9ff-110">Comentários</span><span class="sxs-lookup"><span data-stu-id="2a9ff-110">Remarks</span></span>

<span data-ttu-id="2a9ff-111">Para saber mais, confira [Solicitando permissões para API usadas em suplementos de conteúdo e de painel de tarefas](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) e [Entendendo as permissões de suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="2a9ff-111">For more detail, see [Requesting permissions for API use in content and task pane add-ins](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
