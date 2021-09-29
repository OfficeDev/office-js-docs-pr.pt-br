---
title: Elemento MobileFormFactor no arquivo de manifesto
description: O elemento MobileFormFactor especifica as configurações do fator de forma móvel para um complemento.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: c0b701064f3a8c0244dc62226899c4317b9075c4
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990674"
---
# <a name="mobileformfactor-element"></a>Elemento MobileFormFactor

Especifica as configurações de um suplemento para um fator forma móvel. Ele contém todas as informações do suplemento para o fator forma móvel, exceto para o nó **Resources**.

Cada **definição MobileFormFactor** contém o **elemento FunctionFile** e um ou mais **elementos ExtensionPoint.** Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).

O elemento **MobileFormFactor** é definido no esquema VersionOverrides 1.1. O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.

**Tipo de suplemento:** Email

## <a name="child-elements"></a>Elementos filho

| Elemento                             | Obrigatório | Descrição  |
|:------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md) | Sim      | Define onde um suplemento expõe a funcionalidade. |
| [FunctionFile](functionfile.md)     | Sim      | Uma URL para um arquivo que contém funções JavaScript.|

## <a name="mobileformfactor-example"></a>Exemplo de MobileFormFactor

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
