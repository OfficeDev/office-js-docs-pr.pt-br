---
title: Elemento Method no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: fded84344182bb45597b00a794f18defaa44d3b3
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432820"
---
# <a name="method-element"></a>Elemento Method

Especifica um método individual a partir da API do JavaScript para Office que o Suplemento do Office exige para ativar.

**Tipo de suplemento:** Conteúdo, Painel de tarefas

## <a name="syntax"></a>Sintaxe

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>Contido em

[Methods](methods.md)

## <a name="attributes"></a>Atributos

|**Atributo**|**Tipo**|**Obrigatório**|**Descrição**|
|:-----|:-----|:-----|:-----|
|Nome|cadeia de caracteres|obrigatório|Especifica o nome do método necessário qualificado com seu objeto pai. Por exemplo, para especificar o método **getSelectedDataAsync**, você deve especificar `"Document.getSelectedDataAsync"`.|

## <a name="remarks"></a>Comentários

Os elementos **Method** e **Methods** não têm suporte nos suplementos de email. Para saber mais sobre conjuntos de requisitos, consulte [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

> [!IMPORTANT] 
> Como não há forma de especificar o requisito de versão mínimo de métodos individuais, para verificar se um método está disponível no tempo de execução, você também deve usar uma instrução **if** ao chamar esse método no script do suplemento. Para saber mais sobre como fazer isso, consulte [Noções básicas da API JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

