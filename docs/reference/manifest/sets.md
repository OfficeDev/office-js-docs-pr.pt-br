---
title: Elemento Sets no arquivo de manifesto
description: O elemento Sets especifica o conjunto mínimo de Office API JavaScript que seu Office Desempio exige para ativar.
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 38707ec78a79e9104dd21f9fa5ceab8c6fbd2c79
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151996"
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

