---
title: Padrões de navegação para Suplementos do Office
description: Saiba mais sobre as práticas recomendadas para usar barras de comandos, barras de guias e botões voltar para projetar a navegação de um suplemento do Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 3bb350ede78bef684899f26e4818eba440677541
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132029"
---
# <a name="navigation-patterns"></a>Padrões de navegação

Os principais recursos de um suplemento são acessados por meio de tipos de comandos específicos e área de tela limitada. É importante que a navegação seja intuitiva, forneça contexto e permita que o usuário se mova facilmente por todo o suplemento.

## <a name="best-practices"></a>Práticas recomendadas

| Fazer    | Não fazer |
| :---- | :---- |
| Certifique-se de que o usuário tenha uma opção de navegação claramente visível. | Não complique o processo de navegação usando a interface de usuário não padrão.
| Utilize os seguintes componentes, conforme aplicável, para permitir que os usuários naveguem pelo suplemento. | Não dificulte para o usuário entender seu local ou contexto atual dentro do suplemento.

## <a name="command-bar"></a>Barra de comandos

O CommandBar é uma superfície dentro do painel de tarefas que abriga comandos que operam no conteúdo da janela, painel ou região pai que residem acima. Recursos opcionais incluem um ponto de acesso de menu vertical suspenso, pesquisa e comandos laterais.

![Ilustração mostrando uma barra de comandos dentro de um painel de tarefas de aplicativo da área de trabalho do Office. Este exemplo mostra uma barra de comandos imediatamente abaixo do nome do suplemento que inclui um menu e uma pesquisa.](../images/add-in-command-bar.png)

## <a name="tab-bar"></a>Barra de guias

A barra de guias mostra a navegação usando botões com texto empilhado verticalmente e ícones. Use a barra de guias para fornecer a navegação usando guias com títulos curtos e descritivos.

![Ilustração mostrando uma barra de guias dentro de um painel de tarefas de aplicativo da área de trabalho do Office. Este exemplo mostra uma barra de guias imediatamente abaixo do nome do suplemento com as guias "Home", "Settings", "Favorites" e "Account".](../images/add-in-tab-bar.png)

## <a name="back-button"></a>Botão Voltar

O botão voltar permite que os usuários se recuperem de uma ação de navegação de busca detalhada. Esse padrão ajuda a garantir que os usuários sigam uma série de etapas ordenadas.

![Ilustração mostrando um botão voltar dentro de um painel de tarefas de aplicativo da área de trabalho do Office. Este exemplo mostra um botão voltar imediatamente abaixo do nome do suplemento, no canto superior esquerdo.](../images/add-in-back-button.png)
