---
title: Elemento Host no arquivo de manifesto
description: Especifica um tipo de aplicativo individual do Office em que o suplemento deve ser ativado.
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 45d4ed42946038699be235ff3912c071a92ff226
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348325"
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
| [Nome](#name) | cadeia de caracteres | obrigatório | O nome do tipo de aplicativo Office cliente. |

### <a name="name"></a>Name

Especifica o tipo de Host destinado por esse suplemento. O valor deve ser uma das seguintes opções:

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
|  [xsi:type](#xsitype)  |  Sim  | Descreve o aplicativo Office em que essas configurações se aplicam.|

### <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [DesktopFormFactor](desktopformfactor.md)    |  Sim   |  Define as configurações do fator forma da área de trabalho. |
|  [MobileFormFactor](mobileformfactor.md)    |  Não   |  Define as configurações do fator forma móvel. **Observação:** Esse elemento só tem suporte em Outlook iOS e Android. |
|  [AllFormFactors](allformfactors.md)    |  Não   |  Define as configurações de todos os fatores forma. Usado somente pelas funções personalizadas no Excel. |

### <a name="xsitype"></a>xsi:type

Controla qual Office aplicativo (Word, Excel, PowerPoint, Outlook, OneNote) onde as configurações contidas se aplicam. O valor deve ser uma das seguintes opções:

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
