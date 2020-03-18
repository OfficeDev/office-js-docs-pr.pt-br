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
# <a name="permissions-element"></a>Elemento Permissions

Especifica o nível de acesso da API para seu Suplemento do Office. Você deve solicitar permissões com base no princípio do privilégio mínimo.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

Para suplementos de conteúdo e de painel de tarefas:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

Para suplementos de email

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Comentários

Para obter mais detalhes, consulte [solicitando permissões para uso da API em suplementos](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) e [noções básicas sobre permissões de suplementos do Outlook](../../outlook/understanding-outlook-add-in-permissions.md).
