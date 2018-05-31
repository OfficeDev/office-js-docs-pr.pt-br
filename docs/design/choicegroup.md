---
title: Componente ChoiceGroup no Office UI Fabric
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 78da2fae781039663bfe2bac159bfbe50192c023
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437210"
---
# <a name="choicegroup-component-in-office-ui-fabric"></a>Componente ChoiceGroup no Office UI Fabric

O componente ChoiceGroup, também conhecido como um botão de opção, apresenta aos usuários duas ou mais opções mutuamente exclusivas. Os usuários podem selecionar apenas um botão do ChoiceGroup em um grupo. Cada opção é representada por um botão do ChoiceGroup. 
  
#### <a name="example-choicegroup-in-a-task-pane"></a>Exemplo: ChoiceGroup em um painel de tarefas

 ![Imagem mostrando um ChoiceGroup](../images/overview-with-app-choicegroup.png)

## <a name="best-practices"></a>Práticas recomendadas

|**Faça**|**Não faça**|
|:------------|:--------------|
|Mantenha as opções de ChoiceGroup no mesmo nível.<br/><br/>![Exemplo do que fazer com ChoiceGroup](../images/choice-do.png)<br/>|Não utilize ChoiceGroups ou caixas de seleção aninhados.<br/><br/>![Exemplo do que não fazer com ChoiceGroup](../images/choice-dont.png)<br/>|
|Use ChoiceGroups com duas a sete opções, verificando se há espaço suficiente na tela para mostrar todas as opções. Caso contrário, use uma caixa de seleção ou lista suspensa.|Não use quando as opções forem números com uma gradação fixa, por exemplo, 10, 20, 30 e assim por diante. Em vez disso, use um componente de controle deslizante.|
|Se os usuários não puderem escolher nenhuma das opções, considere incluir uma opção como **Nenhum** ou **Não se aplica**.|Não use dois botões de ChoiceGroup para uma única opção binária.|
|Se possível, alinhe os botões de ChoiceGroup verticalmente em vez de horizontalmente. O alinhamento horizontal é mais difícil de ler e localizar.||
|Liste as opções em ordem lógica, por exemplo, da opção mais provável a ser selecionada até a menos, da operação mais simples até a mais complexa ou do menor risco para o maior risco. |Não use ordenação alfabética porque é dependente do idioma.|

## <a name="variants"></a>Variantes

|**Variação**|**Descrição**|**Exemplo**|
|:------------|:--------------|:----------|
|**ChoiceGroups**|Use quando não forem necessárias imagens para fazer uma escolha.|![Imagem da variante de ChoiceGroup](../images/radio.png)<br/>|
|**ChoiceGroups usando imagens**|Use quando forem necessárias imagens para fazer uma escolha.|![Variante de ChoiceGroup com imagem](../images/radio-image.png)<br/>|

## <a name="implementation"></a>Implementação

Para saber mais, confira [ChoiceGroup](https://dev.office.com/fabric#/components/choicegroup) e [Primeiros passos com exemplo de código do Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## <a name="see-also"></a>Veja também

- [Padrões de design da experiência do usuário](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Office UI Fabric em Suplementos do Office](office-ui-fabric.md)
