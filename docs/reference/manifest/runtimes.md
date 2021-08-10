---
title: Tempos de execução no arquivo de manifesto
description: O elemento Runtimes especifica o tempo de execução do seu complemento.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: e5ec70449d3984671d507131ac8d4fc0f7617cdcda1ad8f99b4f4bf52773aded
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57091623"
---
# <a name="runtimes-element"></a>Elemento Runtimes

Especifica o tempo de execução do seu complemento. Filho do [`<Host>`](host.md) elemento.

> [!NOTE]
> Ao executar no Office no Windows, um complemento que tenha um elemento em seu manifesto não necessariamente é executado no mesmo controle de webview como faria `<Runtimes>` de outra forma. Para obter mais informações sobre como as versões de Windows e Office determinam qual controle webview normalmente é usado, consulte [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). Se as condições descritas lá para o uso de Microsoft Edge com WebView2 (baseadas em Chromium) são atendidas, o complemento usa esse navegador se ele tem ou não um `<Runtimes>` elemento. No entanto, quando essas condições não são atendidas, um complemento com um elemento sempre usa o Internet Explorer 11, independentemente da versão Windows `<Runtimes>` ou Microsoft 365.

**Tipo de complemento:** Painel de tarefas, Email

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
| [Tempo de execução](runtime.md) | Sim |  O tempo de execução do seu complemento. **Importante**: No momento, você só pode definir um `<Runtime>` elemento. |

## <a name="see-also"></a>Confira também

- [Tempo de execução](runtime.md)
- [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configurar seu Outlook para ativação baseada em eventos](../../outlook/autolaunch.md)
