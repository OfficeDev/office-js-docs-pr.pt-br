---
title: Caixas de diálogo em Suplementos do Office
description: Conheça as práticas recomendadas para o design visual das caixas de diálogo em suplementos do Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2f3b25fac7f12494e6b5a1e0a32e72baa345e978
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717191"
---
# <a name="dialog-boxes-in-office-add-ins"></a>Caixas de diálogo em Suplementos do Office
 
Caixas de diálogo são superfícies que flutuam acima da janela do aplicativo do Office ativo. Você pode usar caixas de diálogo para fornecer espaço adicional na tela para tarefas como páginas de entrada que não podem ser abertas diretamente em um painel de tarefas ou solicitações para confirmar uma ação executada por um usuário ou mostrar vídeos que podem ser muito pequenos se confinados a um painel de tarefas.

*Figura 1. Layout típico de uma caixa de diálogo*

![Uma imagem de exemplo que exibe um layout típico de uma caixa de diálogo](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a>Práticas recomendadas

|**Faça**|**Não faça**|
|:-----|:--------|
|<ul><li>Inclua um título descritivo com o nome de suplemento, juntamente com a tarefa atual.</li></ul>|<ul><li>Não adicione o nome da sua empresa ao título.</li></ul>|
||<ul><li>Não abra uma caixa de diálogo, a menos que o cenário exija isso.</li></ul>|

## <a name="implementation"></a>Implementação

Confira um exemplo que implementa uma caixa de diálogo em [Exemplo de API de caixa de diálogo de suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) no GitHub.

## <a name="see-also"></a>Confira também

- [Objeto Dialog](/javascript/api/office/office.dialog)
- [Padrões de design da experiência do usuário para suplementos do Office](../design/ux-design-pattern-templates.md)
