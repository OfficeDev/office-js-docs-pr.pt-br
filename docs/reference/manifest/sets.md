---
title: Elemento Sets no arquivo de manifesto
description: O elemento Sets especifica o conjunto mínimo de Office API JavaScript que seu Office Desempio exige para ativar.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: bd8f8311bb06a8e9e98fc408aece6395ab5643b1
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936548"
---
# <a name="sets-element"></a>Elemento Sets

Especifica o subconjunto mínimo da API JavaScript Office que o seu Office Descrição requer para ativar.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a>Contido em

[Requisitos](requirements.md)

## <a name="can-contain"></a>Pode conter

[Set](set.md)

## <a name="attributes"></a>Atributos

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|cadeia de caracteres|opcional|Especifica o valor padrão do atributo **MinVersion** para todos os elementos [Set](set.md) filho. O valor padrão é "1.1".|

## <a name="remarks"></a>Comentários

Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Para obter mais informações sobre o atributo **MinVersion** do elemento **Set** e o atributo **DefaultMinVersion** do elemento **Sets,** consulte Definir o elemento [Requirements no manifesto](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).

