---
title: Elemento Requirements no arquivo de manifesto
description: O elemento Requirements especifica o conjunto de requisitos mínimo e os métodos que seu Office de complemento precisa para ativar.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3020037b48e3f759acf6a7e2758bb8c1fd2dd36429e0b21613e22fca33cacc1a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098100"
---
# <a name="requirements-element"></a>Elemento Requirements

Especifica o conjunto mínimo de Office da API JavaScript[(](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) conjuntos de requisitos e/ou métodos) que seu Office Add-in precisa ativar.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Pode conter

|Elemento|Conteúdo|Email|TaskPane|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[Métodos](methods.md)|x||x|

## <a name="remarks"></a>Comentários

Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).
