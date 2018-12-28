---
title: Elemento Host no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 37b772261ad82b4f899e73314a08ffd1dd03b442
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432806"
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
|  [MobileFormFactor](mobileformfactor.md)    |  Não   |  Define as configurações do fator forma móvel. **Observação:** esse elemento só tem suporte no Outlook para iOS. |
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
