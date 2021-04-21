---
title: Tempos de execução no arquivo de manifesto
description: O elemento Runtimes especifica o tempo de execução do seu complemento.
ms.date: 04/16/2021
localization_priority: Normal
ms.openlocfilehash: 8f4a602c05b9af7bde9f644ef40b61a214e66cd5
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917083"
---
# <a name="runtimes-element"></a>Elemento Runtimes

Especifica o tempo de execução do seu complemento. Filho do [`<Host>`](host.md) elemento.

> [!NOTE]
> Ao executar no Office no Windows, um add-in que tenha um elemento em seu manifesto não necessariamente é executado no mesmo controle `<Runtimes>` de webview como faria. Para obter mais informações sobre como as versões do Windows e do Office determinam qual controle webview normalmente é usado, consulte [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). Se as condições descritas lá para o uso do Microsoft Edge com WebView2 (baseado em Chromium) são atendidas, o complemento usa esse navegador se ele tem ou não um `<Runtimes>` elemento. No entanto, quando essas condições não são atendidas, um complemento com um elemento sempre usa o Internet Explorer 11, independentemente da versão do Windows ou `<Runtimes>` do Microsoft 365.

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
| [Tempo de execução](runtime.md) | Sim |  O tempo de execução do seu complemento. |

## <a name="see-also"></a>Confira também

- [Tempo de execução](runtime.md)
- [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configurar seu complemento do Outlook para ativação baseada em eventos](../../outlook/autolaunch.md)
