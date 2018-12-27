---
title: Elemento Requirements no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 2544e9b01b2d4d3ddc0a0c6238b4a5b0e6c4f832
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432701"
---
# <a name="requirements-element"></a>Elemento Requirements

Especifica o conjunto mínimo de requisitos da API JavaScript para Office ([conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) e/ou métodos) que o Suplemento do Office precisa ativar.

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

|**Element**|**Conteúdo**|**Email**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[Métodos](methods.md)|x||x|

## <a name="remarks"></a>Comentários

Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

