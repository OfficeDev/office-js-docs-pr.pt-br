---
title: Elemento Sets no arquivo de manifesto
description: O elemento sets especifica o conjunto mínimo de API JavaScript do Office que o suplemento do Office exige para ativar.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 8c1c97bfc2934ecf3cc20b472b29a03805603729
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608730"
---
# <a name="sets-element"></a>Elemento Sets

Especifica o subconjunto mínimo da API JavaScript do Office que o suplemento do Office exige para ativar.

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

|**Atributo**|**Tipo**|**Obrigatório**|**Descrição**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|cadeia de caracteres|opcional|Especifica o valor do atributo **MinVersion** padrão para todos os elementos do [conjunto](set.md) filho. O valor padrão é "1.1".|

## <a name="remarks"></a>Comentários

Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Para obter mais informações sobre o atributo **MinVersion** do elemento **set** e o atributo **DefaultMinVersion** do elemento **sets** , confira [definir o elemento requirements no manifesto](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).

