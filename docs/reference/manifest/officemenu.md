---
title: Elemento OfficeMenu no arquivo de manifesto
description: O elemento OfficeMenu define uma coleção de controles a serem adicionados ao menu Office de contexto.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 741e5cabacbc28bf0ebcfdaab8068b280a356dd473d0a83a2378d253df977232
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57094175"
---
# <a name="officemenu-element"></a>Elemento OfficeMenu

Define um conjunto de controles que serão adicionados ao menu de contexto do Office. Aplica-se aos suplementos do Word, do Excel, do PowerPoint e do OneNote.

## <a name="attributes"></a>Atributos

| Atributo            | Obrigatório | Descrição                          |
|:---------------------|:--------:|:-------------------------------------|
| [xsi:type](#xsitype) | Sim      | O tipo de OfficeMenu está sendo definido.|

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Control](#control)    | Sim |  Conjunto de um ou mais objetos Control.  |

## <a name="xsitype"></a>xsi:type

Especifica um menu interno do aplicativo cliente do Office no qual você deseja adicionar esse suplemento do Office.

- `ContextMenuText` -  Exibe o item no menu de contexto quando o texto for selecionado e o usuário abre o menu de contexto (clica com o botão direito do mouse) no texto selecionado. Aplica-se a Word, Excel, PowerPoint e OneNote.
- `ContextMenuCell` -  Exibe o item no menu de contexto quando o usuário abre o menu de contexto (clica com o botão direito do mouse) em uma célula na planilha. Aplica-se ao Excel.

## <a name="control"></a>Control

Cada elemento **OfficeMenu** requer um ou mais controles de [menu](control.md#menu-dropdown-button-controls). 

## <a name="example"></a>Exemplo

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />
          </Action>
        </Item>
      </Items>
    </Control>
</OfficeMenu>
```
