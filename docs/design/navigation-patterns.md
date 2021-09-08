---
title: Padrões de navegação para Suplementos do Office
description: Saiba as práticas recomendadas para usar barras de comando, barras de tabulação e botões de fundo, para projetar a navegação de um Office Add-in.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 3bb350ede78bef684899f26e4818eba440677541
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58939114"
---
# <a name="navigation-patterns"></a>Padrões de navegação

Os principais recursos de um suplemento são acessados por meio de tipos de comandos específicos e área de tela limitada. É importante que a navegação seja intuitiva, forneça contexto e permita que o usuário se mova facilmente por todo o suplemento.

## <a name="best-practices"></a>Práticas recomendadas

| Fazer    | Não fazer |
| :---- | :---- |
| Certifique-se de que o usuário tenha uma opção de navegação claramente visível. | Não complique o processo de navegação usando a interface de usuário não padrão.
| Utilize os seguintes componentes, conforme aplicável, para permitir que os usuários naveguem pelo suplemento. | Não dificulte para o usuário entender seu local ou contexto atual dentro do suplemento.

## <a name="command-bar"></a>Barra de comandos

CommandBar é uma superfície no painel de tarefas que abriga comandos que operam no conteúdo da janela, do painel ou da região pai que reside acima. Recursos opcionais incluem um ponto de acesso de menu vertical suspenso, pesquisa e comandos laterais.

![Ilustração mostrando uma barra de comandos dentro Office painel de tarefas do aplicativo da área de trabalho. Este exemplo mostra uma barra de comandos imediatamente abaixo do nome do complemento que inclui um menu de hambúrguer e uma pesquisa.](../images/add-in-command-bar.png)

## <a name="tab-bar"></a>Barra de guias

A barra de guias mostra a navegação usando botões com texto e ícones empilhados verticalmente. Use a barra de guias para fornecer a navegação usando guias com títulos curtos e descritivos.

![Ilustração mostrando uma barra de guias em um Office de tarefas do aplicativo de área de trabalho. Este exemplo mostra uma barra de guias imediatamente abaixo do nome do add-in com as guias "Home", "Configurações", "Favorites" e "Account".](../images/add-in-tab-bar.png)

## <a name="back-button"></a>Botão Voltar

O botão voltar permite que os usuários se recuperem de uma ação de navegação de detalhamento. Esse padrão ajuda a garantir que os usuários sigam uma série de etapas ordenadas.

![Ilustração mostrando um botão voltar em um painel Office de tarefas do aplicativo da área de trabalho. Este exemplo mostra um botão voltar imediatamente abaixo do nome do complemento, na parte superior esquerda.](../images/add-in-back-button.png)
