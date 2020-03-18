---
title: Tempos de execução no arquivo de manifesto (versão prévia)
description: O elemento de Runtime especifica o tempo de execução do seu suplemento.
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 5797aa78ae3667461de48de481ff44f14c307ced
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720418"
---
# <a name="runtimes-element-preview"></a>Elemento de runtimes (visualização)

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

Especifica o tempo de execução do suplemento e permite funções personalizadas, botões da faixa de opções e o painel de tarefas para usar o mesmo tempo de execução do JavaScript. Filho do `<Host>` elemento no seu arquivo de manifesto. Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

**Tipo de suplemento:** Painel de tarefas

> [!IMPORTANT]
> O tempo de execução compartilhado está atualmente em versão prévia e só está disponível no Excel no Windows. Para experimentar os recursos de visualização, você precisará ingressar no [Office Insider](https://insider.office.com/).

## <a name="syntax"></a>Sintaxe

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Contido em 
[Host](./host.md)

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Runtime**     | Sim |  O tempo de execução do suplemento.

## <a name="see-also"></a>Também confira

- [Runtime](runtime.md)
