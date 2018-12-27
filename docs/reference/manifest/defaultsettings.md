---
title: Elemento DefaultSettings no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 0c109d5d893cf9d3502f1cbf1724007f01e623e6
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433751"
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

|**Element**|**Conteúdo**|**Email**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](sourcelocation.md)|x||x|
|[RequestedWidth](requestedwidth.md)|x|||
|[RequestedHeight](requestedheight.md)|x|||

## <a name="remarks"></a>Comentários

O local de origem e outras configurações no elemento **DefaultSettings** se aplicam apenas a suplementos de conteúdo e de painel de tarefas. Para suplementos de email, você especifica os locais padrão para os arquivos de origem e outras configurações padrão no elemento [FormSettings](formsettings.md).

