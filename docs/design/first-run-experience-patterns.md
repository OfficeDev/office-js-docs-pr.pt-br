---
title: Padrões de tela de apresentação para suplemento dos Office
description: Saiba as práticas recomendadas para projetar experiências de primeira Office de complementos.
ms.date: 07/08/2018
localization_priority: Normal
ms.openlocfilehash: cd268e227f6d4c6cc5aae5c954a39e0c19315330
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/05/2021
ms.locfileid: "53774004"
---
# <a name="first-run-experience-patterns"></a>Padrões de tela de apresentação

Uma tela de apresentação (FRE) é a introdução de um usuário para o suplemento. Um FRE é exibida quando um usuário abre um suplemento pela primeira vez e fornece informações sobre as funções, recursos e/ou os benefícios do suplemento. Essa experiência ajuda a moldar a impressão do usuário de um suplemento e pode influenciar fortemente sua probabilidade de voltar e continuar usando o suplemento.

## <a name="best-practices"></a>Práticas recomendadas

Siga estas práticas recomendadas ao criar sua experiência de primeira etapa.

|Fazer|Não fazer|
|:------|:------|
|Forneça uma simples e breve introdução para as principais ações do suplemento. | Não inclua informações e legendas que não sejam relevantes ao guia de introdução.
|Forneça aos usuários a oportunidade de concluir uma ação que impactará positivamente o uso do add-in. | Não espere que os usuários aprendam tudo ao mesmo tempo. Concentre-se na ação que fornece o maior valor.
|Crie uma experiência envolvente que os usuários desejem concluir. | Não force os usuários a clicar na experiência da tela de apresentação. Forneça aos usuários uma opção para ignorar a tela de apresentação. |

Considere se mostrar aos usuários a tela de apresentação uma vez ou periodicamente é importante para seu cenário. Por exemplo, se o suplemento for usado apenas periodicamente, os usuários poderão ficar menos familiarizados com seu suplemento e poderão se beneficiar de outra interação com a tela de apresentação.

Aplique os seguintes padrões, conforme aplicável, para criar ou aprimorar a tela de apresentação do seu suplemento.

## <a name="carousel"></a>Carrossel

O carrossel apresenta aos usuários uma série de recursos ou página de informações antes que eles comecem a usar o suplemento.

*Figura 1. Permitir que os usuários avancem ou pulem as páginas in início do fluxo de carrossel*

![Ilustração mostrando a etapa 1 de um carrossel na primeira experiência de Office de tarefas do aplicativo de área de trabalho. Neste exemplo, uma ação "Ignorar" é incluída na parte superior direita do painel de tarefas.](../images/add-in-FRE-step-1.png)

*Figura 2. Minimizar o número de telas de carrossel apenas para o que é necessário para comunicar efetivamente sua mensagem*

![Ilustração mostrando a etapa 2 de um carrossel na primeira experiência de Office de tarefas do aplicativo de área de trabalho. Neste exemplo, há três telas de carrossel no painel de tarefas.](../images/add-in-FRE-step-2.png)

*Figura 3. Fornecer uma chamada clara para a ação para sair da primeira experiência de executar*

![Ilustração mostrando a etapa 3 de um carrossel na primeira experiência de Office de tarefas do aplicativo de área de trabalho. Neste exemplo, a terceira e última tela do painel de tarefas mostra um botão para começar.](../images/add-in-FRE-step-3.png)

## <a name="value-placemat"></a>Roteiro de valor

O posicionamento do valor informa a proposta de valor do seu suplemento com posicionamento do logotipo, uma proposta de valor claramente definida, destaques ou resumo do recurso e uma chamada para ação.

*Figura 4. Um placemat de valor com logotipo, proposta de valor desmarcada, resumo de recursos e chamada para ação*

![Ilustração mostrando um placemat de valor na primeira experiência de Office de aplicativos de área de trabalho. Neste exemplo, o painel de tarefas exibe o logotipo do complemento, uma descrição do complemento e um botão para começar.](../images/add-in-FRE-value.png)

### <a name="video-placemat"></a>Roteiro de vídeo

O roteiro de vídeo mostra aos usuários um vídeo antes que eles comecem a usar o suplemento.

*Figura 5. Primeiro executar o placemat de vídeo - A tela contém uma imagem de still do vídeo com um botão de reprodução e um botão de chamada para ação des clara*

![Ilustração mostrando um placemat de vídeo na primeira experiência de Office de aplicativos de área de trabalho.](../images/add-in-FRE-video.png)

*Figura 6. Player de vídeo - Usuários apresentados com um vídeo dentro de uma janela de diálogo*

![Ilustração mostrando um vídeo em uma janela de diálogo com um Office de área de trabalho e painel de tarefas do complemento em segundo plano.](../images/add-in-FRE-video-dialog.png)
