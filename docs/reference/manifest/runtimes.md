---
title: Tempos de execução no arquivo de manifesto (versão prévia)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 17e53b53d55ea9547cdfc5c4f89f8f4c3a7ab75e
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/26/2020
ms.locfileid: "42283868"
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

## <a name="see-also"></a>Confira também

- [Runtime](runtime.md)
