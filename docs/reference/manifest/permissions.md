---
title: Elemento Permissions no arquivo de manifesto
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3442a8e0caee442ce1b38c5ff39cfd1ef5088fb7
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872057"
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

Para saber mais, confira [Solicitando permissões para API usadas em suplementos de conteúdo e de painel de tarefas](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) e [Entendendo as permissões de suplemento do Outlook](/outlook/add-ins/understanding-outlook-add-in-permissions).
