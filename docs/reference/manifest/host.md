---
title: Elemento Host no arquivo de manifesto
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: debb4d59f75ce974ffb21d853c6b65a579c4e685
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127566"
---
# <a name="host-element"></a>Elemento Host

Especifica um tipo de aplicativo individual do Office em que o suplemento deve ser ativado.

> [!IMPORTANT] 
> A sintaxe do elemento **Host** varia de acordo com a definição do elemento, se dentro do [manifesto básico](#basic-manifest) ou dentro do nó [VersionOverrides](#versionoverrides-node). No entanto, a funcionalidade é a mesma.  

## <a name="basic-manifest"></a>Manifesto básico

Quando definido no manifesto básico (em [OfficeApp](officeapp.md)), o tipo de host é determinado pelo atributo `Name`.

### <a name="attributes"></a>Atributos

| Atributo     | Tipo   | Obrigatório | Descrição                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [Nome](#name) | cadeia de caracteres | obrigatório | O nome do tipo de aplicativo host do Office. |

### <a name="name"></a>Name
Especifica o tipo de Host destinado por esse suplemento. O valor deve ser uma das seguintes opções:

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

### <a name="example"></a>Exemplo
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a>Nó VersionOverrides
Quando definido em [VersionOverrides](versionoverrides.md), o tipo de host é determinado pelo atributo `xsi:type`. 

### <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Sim  | Descreve o host do Office a que essas configurações se aplicam.|

### <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [DesktopFormFactor](desktopformfactor.md)    |  Sim   |  Define as configurações do fator forma da área de trabalho. |
|  [MobileFormFactor](mobileformfactor.md)    |  Não   |  Define as configurações do fator forma móvel. **Observação:** Esse elemento só é suportado no Outlook no iOS. |
|  [AllFormFactors](allformfactors.md)    |  Não   |  Define as configurações de todos os fatores forma. Usado somente pelas funções personalizadas no Excel. |

### <a name="xsitype"></a>xsi:type

Controla a qual host do Office (Word, Excel, PowerPoint, Outlook, OneNote) as configurações contidas se aplicam. O valor deve ser uma das seguintes opções:

- `Document` (Word)
- `MailHost` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Workbook` (Excel)

## <a name="host-example"></a>Exemplo de host 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
