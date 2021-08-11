---
title: Elemento DefaultSettings no arquivo de manifesto
description: Especifica a localização de origem padrão e outras configurações padrão para o suplemento de conteúdo ou de painel de tarefas.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 11e398d86a702f4e45a5cea7b63e0380ce65d1749d0660789e96477744d73079
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095891"
---
# <a name="defaultsettings-element"></a>Elemento DefaultSettings

Especifica a localização de origem padrão e outras configurações padrão para o suplemento de conteúdo ou de painel de tarefas.

**Tipo de suplemento:** Conteúdo, Painel de tarefas

## <a name="syntax"></a>Sintaxe

```XML
<DefaultSettings>
  ...
</DefaultSettings>
```

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Pode conter

|Elemento|Conteúdo|Email|TaskPane|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>Comentários

O local de origem e outras configurações no **elemento DefaultSettings** aplicam-se somente a complementos de conteúdo e painel de tarefas. Para os complementos de email, especifique os locais padrão para arquivos de origem e outras configurações padrão no [elemento FormSettings.](formsettings.md)
