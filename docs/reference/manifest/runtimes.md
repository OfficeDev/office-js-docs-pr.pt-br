---
title: Tempos de execução no arquivo de manifesto
description: O elemento Runtimes especifica o tempo de execução do seu complemento.
ms.date: 09/28/2021
ms.localizationpriority: medium
ms.openlocfilehash: 758bb7b830009d6691190a0279440a52da724624
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138601"
---
# <a name="runtimes-element"></a>Elemento Runtimes

Especifica o tempo de execução do seu complemento. Filho do [`<Host>`](host.md) elemento.

> [!NOTE]
> Ao executar no Office no Windows, um complemento que tenha um elemento em seu manifesto não necessariamente é executado no mesmo controle de webview como faria `<Runtimes>` de outra forma. Para obter mais informações sobre como as versões de Windows e Office determinam qual controle webview normalmente é usado, consulte [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). Se as condições descritas lá para o uso de Microsoft Edge com WebView2 (baseadas em Chromium) são atendidas, o complemento usa esse navegador se ele tem ou não um `<Runtimes>` elemento. No entanto, quando essas condições não são atendidas, um complemento com um elemento sempre usa o Internet Explorer 11, independentemente da versão Windows `<Runtimes>` ou Microsoft 365.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nestes esquemas VersionOverrides:**

 - Painel de tarefas 1.0
 - Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos:**

- [SharedRuntime 1.1](../requirement-sets/shared-runtime-requirement-sets.md) (Somente quando usado em um complemento do painel de tarefas.)

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
| [Runtime](runtime.md) | Sim |  O tempo de execução do seu complemento. **Importante**: No momento, você só pode definir um `<Runtime>` elemento. |

## <a name="see-also"></a>Confira também

- [Runtime](runtime.md)
- [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configurar seu Outlook para ativação baseada em eventos](../../outlook/autolaunch.md)
