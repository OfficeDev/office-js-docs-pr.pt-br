---
title: Caixas de diálogo em Suplementos do Office
description: ''
ms.date: 02/28/2019
localization_priority: Priority
ms.openlocfilehash: 3638006c30515a1fcae93ccfdbd43e0e92005c37
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691101"
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

- [Objeto Dialog](https://docs.microsoft.com/javascript/api/office/office.dialog)
- [Padrões de design da experiência do usuário para suplementos do Office](../design/ux-design-pattern-templates.md)


