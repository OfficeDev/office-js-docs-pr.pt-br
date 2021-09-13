---
title: Elemento Method no arquivo de manifesto
description: O elemento Method especifica um método individual da API javaScript Office que os seus Office Desempresos exigem para ativar.
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 037446f5027a97214d2b1be6ee99c8f6822b33b9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59148931"
---
# <a name="method-element"></a>Elemento Method

Especifica um método individual da API javaScript Office que seu Office Descrição requer para ativar.

**Tipo de suplemento:** Conteúdo, Painel de tarefas

## <a name="syntax"></a>Sintaxe

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a>Contido em

[Methods](methods.md)

## <a name="attributes"></a>Atributos

|Atributo|Tipo|Obrigatório|Descrição|
|:-----|:-----|:-----|:-----|
|Nome|cadeia de caracteres|obrigatório|Especifica o nome do método necessário qualificado com seu objeto pai. Por exemplo, para especificar o `getSelectedDataAsync` método, você deve especificar `"Document.getSelectedDataAsync"` .|

## <a name="remarks"></a>Comentários

Os `Methods` elementos e não são `Method` suportados por complementos de email. Para obter mais informações sobre conjuntos de requisitos, [consulte Office versões e conjuntos de requisitos.](../../develop/office-versions-and-requirement-sets.md)

> [!IMPORTANT]
> Como não há forma de especificar o requisito de versão mínimo de métodos individuais, para verificar se um método está disponível no tempo de execução, você também deve usar uma instrução **if** ao chamar esse método no script do suplemento. Para obter mais informações sobre como fazer isso, consulte [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).
