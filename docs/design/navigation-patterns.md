---
title: Padrões de navegação para Suplementos do Office
description: Saiba mais sobre as práticas recomendadas para usar barras de comandos, barras de guias e botões voltar para projetar a navegação de um suplemento do Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 6fb025a897cfc820117a0b6153acc92c2aeb837e
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718752"
---
# <a name="navigation-patterns"></a>Padrões de navegação

Os principais recursos de um suplemento são acessados por meio de tipos de comandos específicos e área de tela limitada. É importante que a navegação seja intuitiva, forneça contexto e permita que o usuário se mova facilmente por todo o suplemento.

## <a name="best-practices"></a>Práticas recomendadas

| Fazer    | Não fazer |
| :---- | :---- |
| Certifique-se de que o usuário tenha uma opção de navegação claramente visível. | Não complique o processo de navegação usando a interface de usuário não padrão.
| Utilize os seguintes componentes, conforme aplicável, para permitir que os usuários naveguem pelo suplemento. | Não dificulte para o usuário entender seu local ou contexto atual dentro do suplemento.



## <a name="command-bar"></a>Barra de comandos

A Barra de comandos é uma superfície que abriga comandos que operam no conteúdo da janela, painel ou região pai sobre o qual ela reside. Recursos opcionais incluem um ponto de acesso de menu vertical suspenso, pesquisa e comandos laterais.

![Comandos: especificações para o painel de tarefas da área de trabalho](../images/add-in-command-bar.png)



## <a name="tab-bar"></a>Barra de guias

Mostra a navegação usando botões com texto empilhado na vertical e ícones. Use a barra de guias para proporcionar uma navegação em guias com títulos curtos e descritivos.

![Barra de guias: especificações para o painel de tarefas da área de trabalho](../images/add-in-tab-bar.png)


## <a name="back-button"></a>Botão Voltar

O botão Voltar permite que os usuários se recuperem de uma ação de navegação detalhada. Esse padrão ajuda a garantir que os usuários sigam uma série de etapas ordenadas.  

![Botão Voltar: especificações para o painel de tarefas da área de trabalho](../images/add-in-back-button.png)
