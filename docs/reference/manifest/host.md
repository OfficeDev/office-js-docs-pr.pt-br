---
title: Elemento Host no arquivo de manifesto
description: Especifica um tipo de aplicativo individual do Office em que o suplemento deve ser ativado.
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 5db9df97c4ba558d54756b983a26cb7b71e049d5
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611810"
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

Especifica o tipo de Host destinado por esse suplemento. O valor deve ser um dos seguintes.

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

> [!IMPORTANT]
> Não recomendamos mais criar e usar aplicativos Web do Access e bancos de dados no SharePoint. Como alternativa, use o [Microsoft PowerApps](https://powerapps.microsoft.com/) para criar soluções de negócios sem código para dispositivos móveis e Web.

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
|  [MobileFormFactor](mobileformfactor.md)    |  Não   |  Define as configurações do fator forma móvel. **Observação:** Esse elemento só é suportado no Outlook no iOS e no Android. |
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
