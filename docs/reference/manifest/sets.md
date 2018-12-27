---
title: Elemento Sets no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: b7e78ae05f8409f38c885a1d6a328347d00d0df1
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433652"
---
# <a name="sets-element"></a>Elemento Sets

Especifica o subconjunto mínimo da API do JavaScript para Office que o Suplemento do Office exige para ativar.

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
|DefaultMinVersion|cadeia de caracteres|opcional|Especifica o valor padrão do atributo  **MinVersion** para todos os elementos [Set](set.md) filho. O valor padrão é "1.1".|

## <a name="remarks"></a>Comentários

Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Para saber mais sobre o atributo **MinVersion** do elemento **Set** e o atributo **DefaultMinVersion** do elemento **Sets**, confira [Definir o elemento Requirements no manifesto](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).

