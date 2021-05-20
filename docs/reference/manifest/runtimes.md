---
title: Tempos de execução no arquivo de manifesto
description: O elemento Runtimes especifica o tempo de execução do seu complemento.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 80336674c6d954bb9e0c6892feb41cb2f03c5859
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555294"
---
# <a name="runtimes-element"></a>Elemento runtimes

Especifica o tempo de execução do seu complemento. Filho do [`<Host>`](host.md) elemento.

> [!NOTE]
> Ao ser executado em Office em Windows, um complemento que tem um `<Runtimes>` elemento em seu manifesto não é necessariamente executado no mesmo controle de webview que de outra forma seria. Para obter mais informações sobre como as versões de Windows e Office determinar qual controle do webview é normalmente usado, consulte [Navegadores usados por Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). Se as condições descritas lá para o uso Microsoft Edge com o WebView2 (baseado em Chromium) forem atendidas, o complemento usará esse navegador, quer ele tenha ou não um `<Runtimes>` elemento. No entanto, quando essas condições não são atendidas, um complemento com um `<Runtimes>` elemento sempre usa o Internet Explorer 11, independentemente da Windows ou Microsoft 365 versão.

**Tipo de complemento:** Painel de tarefas, Correio

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a>Sintaxe

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Contido em

[Host](host.md)

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
| [Tempo de execução](runtime.md) | Sim |  O tempo de execução para o seu complemento. **Importante**: No momento, você só pode definir um `<Runtime>` elemento. |

## <a name="see-also"></a>Confira também

- [Tempo de execução](runtime.md)
- [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configure seu Outlook complemento para ativação baseada em eventos](../../outlook/autolaunch.md)
