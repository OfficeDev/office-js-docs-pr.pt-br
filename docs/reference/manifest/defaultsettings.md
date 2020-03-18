---
title: Elemento DefaultSettings no arquivo de manifesto
description: Especifica a localização de origem padrão e outras configurações padrão para o suplemento de conteúdo ou de painel de tarefas.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: b97f692a1fd39e4b1f55080f6ed77e623be0000c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718367"
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

|**Elemento**|**Content**|**Email**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>Comentários

O local de origem e outras configurações no elemento **DefaultSettings** só se aplicam a suplementos de conteúdo e de painel de tarefas. Para suplementos de email, você especifica os locais padrão para arquivos de origem e outras configurações padrão no elemento [FormSettings](formsettings.md) .

