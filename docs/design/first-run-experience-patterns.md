---
title: Padrões de tela de apresentação para suplemento dos Office
description: Saiba as práticas recomendadas para projetar experiências de tela de apresentação em suplementos do Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 00785df2cfd2f41b41917ea720c154e24b72f779
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132064"
---
# <a name="first-run-experience-patterns"></a>Padrões de tela de apresentação

Uma tela de apresentação (FRE) é a introdução de um usuário para o suplemento. Um FRE é exibida quando um usuário abre um suplemento pela primeira vez e fornece informações sobre as funções, recursos e/ou os benefícios do suplemento. Essa experiência ajuda a moldar a impressão do usuário de um suplemento e pode influenciar fortemente sua probabilidade de voltar e continuar usando o suplemento.

## <a name="best-practices"></a>Práticas recomendadas

Siga estas práticas recomendadas ao criar sua tela de apresentação:

|Fazer|Não fazer|
|:------|:------|
|Forneça uma simples e breve introdução para as principais ações do suplemento. | Não inclua informações e legendas que não sejam relevantes ao guia de introdução.
|Forneça aos usuários a oportunidade de concluir uma ação que impactará positivamente o uso do add-in. | Não espere que os usuários aprendam tudo ao mesmo tempo. Concentre-se na ação que fornece o maior valor.
|Crie uma experiência envolvente que os usuários desejem concluir. | Não force os usuários a clicar na experiência da tela de apresentação. Forneça aos usuários uma opção para ignorar a tela de apresentação. |

Considere se mostrar aos usuários a tela de apresentação uma vez ou periodicamente é importante para seu cenário. Por exemplo, se o suplemento for usado apenas periodicamente, os usuários poderão ficar menos familiarizados com seu suplemento e poderão se beneficiar de outra interação com a tela de apresentação.

Aplique os seguintes padrões, conforme aplicável, para criar ou aprimorar a tela de apresentação do seu suplemento.

## <a name="carousel"></a>Carrossel

O carrossel apresenta aos usuários uma série de recursos ou página de informações antes que eles comecem a usar o suplemento.

*Figura 1. Permitir que os usuários avancem ou ignorem as páginas iniciais do fluxo de carrossel*

![Ilustração mostrando a etapa 1 de um carrossel na experiência de primeira execução de um painel de tarefas de aplicativo da área de trabalho do Office. Neste exemplo, uma ação "ignorar" é incluída no canto superior direito do painel de tarefas.](../images/add-in-FRE-step-1.png)

*Figura 2. Minimizar o número de telas de carrossel apenas para o que é necessário para comunicar efetivamente sua mensagem*

![Ilustração mostrando a etapa 2 de um carrossel na experiência de primeira execução de um painel de tarefas de aplicativo da área de trabalho do Office. Neste exemplo, há 3 telas de carrossel no painel de tarefas.](../images/add-in-FRE-step-2.png)

*Figura 3. Fornecer um plano de ação claro para sair da experiência de primeira execução*

![Ilustração mostrando a etapa 3 de um carrossel na experiência de primeira execução de um painel de tarefas de aplicativo da área de trabalho do Office. Neste exemplo, a terceira e última tela do painel de tarefas mostra um botão para começar.](../images/add-in-FRE-step-3.png)

## <a name="value-placemat"></a>Roteiro de valor

O posicionamento do valor informa a proposta de valor do seu suplemento com posicionamento do logotipo, uma proposta de valor claramente definida, destaques ou resumo do recurso e uma chamada para ação.

*Figura 4. Um valor roteiro com logotipo, proposta de valor claro, Resumo de recursos e plano de ação*

![Ilustração mostrando um valor roteiro na primeira experiência de execução de um painel de tarefas de aplicativo da área de trabalho do Office. Neste exemplo, o painel de tarefas exibe o logotipo do suplemento, uma descrição do suplemento e um botão para começar.](../images/add-in-FRE-value.png)

### <a name="video-placemat"></a>Roteiro de vídeo

O roteiro de vídeo mostra aos usuários um vídeo antes que eles comecem a usar o suplemento.

*Figura 5. Tela de apresentação de roteiro-a tela contém uma imagem estática do vídeo com um botão Play e um botão limpar chamada para ação*

![Ilustração mostrando um roteiro de vídeo na experiência de primeira execução de um painel de tarefas de aplicativo da área de trabalho do Office](../images/add-in-FRE-video.png)

*Figura 6. Player de vídeo-os usuários apresentados com um vídeo em uma janela de diálogo*

![Ilustração mostrando um vídeo em uma janela de diálogo com um aplicativo de área de trabalho e um painel de tarefas de suplemento do Office em segundo plano](../images/add-in-FRE-video-dialog.png)
