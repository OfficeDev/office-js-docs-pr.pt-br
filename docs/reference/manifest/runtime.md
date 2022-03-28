---
title: Tempo de execução no arquivo de manifesto
description: O elemento Runtime configura seu complemento para usar um tempo de execução JavaScript compartilhado para seus vários componentes, por exemplo, faixa de opções, painel de tarefas, funções personalizadas.
ms.date: 03/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 38920dc43349be8da629785167d03252578f2a42
ms.sourcegitcommit: 64942cdd79d7976a0291c75463d01cb33a8327d8
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/25/2022
ms.locfileid: "64404671"
---
# <a name="runtime-element"></a>Elemento Runtime

Configura seu complemento para usar um tempo de execução javaScript compartilhado para que vários componentes sejam executados no mesmo tempo de execução. Filho do [`<Runtimes>`](runtimes.md) elemento.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

 - Painel de tarefas 1.0
 - Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [SharedRuntime 1.1](../requirement-sets/shared-runtime-requirement-sets.md) (Somente quando usado em um complemento do painel de tarefas.)

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a>Sintaxe

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Contido em

- [Tempos de execução](runtimes.md)

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
| [Override](override.md) | Não | **Outlook**: especifica o local da URL do arquivo JavaScript que Outlook Desktop requer para manipuladores de ponto de extensão [LaunchEvent](../../reference/manifest/extensionpoint.md#launchevent). **Importante**: no momento, você só pode definir um `<Override>` elemento e ele deve ser do tipo `javascript`.|

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **resid**  |  Sim  | Especifica o local da URL da página HTML do seu complemento. O `resid` pode ter não mais de 32 caracteres e deve corresponder a `id` um atributo de um `Url` elemento no `Resources` elemento. |
|  [lifetime](#lifetime-attribute)  |  Não  | O valor padrão para `lifetime` é `short` e não precisa ser especificado. Outlook de ativação baseada em evento usam apenas o `short` valor. Se você quiser usar um tempo de execução compartilhado em um Excel de Excel, de definir explicitamente o valor como `long`. |

### <a name="lifetime-attribute"></a>atributo lifetime

Opcional. Representa o período de tempo em que o add-in tem permissão para ser executado.

**Valores disponíveis**

`short`: Padrão. Usado apenas para Outlook de ativação baseada em eventos. Depois que o add-in for ativado, ele será executado por um período máximo de tempo, conforme especificado pela plataforma. Atualmente, isso é cerca de 5 minutos. Esse é o único valor suportado pelo Outlook.

`long`: Usado somente ao configurar um [tempo de execução JavaScript compartilhado](../../develop/configure-your-add-in-to-use-a-shared-runtime.md). O complemento pode iniciar no documento aberto e executado indefinidamente. Por exemplo, o código do painel de tarefas continuará sendo executado mesmo quando o usuário fechar o painel de tarefas. Esse é o único valor suportado pelo tempo de execução compartilhado.

## <a name="see-also"></a>Confira também

- [Tempos de execução](runtimes.md)
- [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configurar seu Outlook para ativação baseada em eventos](../../outlook/autolaunch.md)
