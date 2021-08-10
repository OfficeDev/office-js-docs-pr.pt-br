---
title: Elemento Permissions no arquivo de manifesto
description: O elemento Permissions especifica o nível de acesso à API para seu Office Add-in.
ms.date: 06/26/2020
localization_priority: Normal
ms.openlocfilehash: 2f2ccb4f6ec691b19cadea76a06520a9bad7a0b6c0e51699f2c8db67a3030de0
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089007"
---
# <a name="permissions-element"></a>Elemento Permissions

Especifica o nível de acesso da API para seu Suplemento do Office. Você deve solicitar permissões com base no princípio do privilégio mínimo.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

Para suplementos de conteúdo e de painel de tarefas:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

Para suplementos de email:

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Comentários

Para obter mais detalhes, consulte Solicitando permissões para uso da API em conteúdo e [complementos](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) do painel de tarefas e [Noções básicas Outlook](../../outlook/understanding-outlook-add-in-permissions.md)permissões de Outlook de complemento.
