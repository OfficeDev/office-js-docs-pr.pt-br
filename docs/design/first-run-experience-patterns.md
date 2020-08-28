---
title: Padrões de tela de apresentação para suplemento dos Office
description: Saiba as práticas recomendadas para projetar experiências de tela de apresentação em suplementos do Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: c0528e869dd8ee7fe779785fb1a9b6d347deab75
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292950"
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

*Figura 1: permitir que os usuários avancem ou ignorem as páginas iniciais do fluxo de carrossel.* 
 ![ Primeira execução-carrossel etapa 1-especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-step-1.png)



*Figura 2: minimize o número de telas de carrossel que você apresenta ao usuário apenas para o que é necessário para comunicar efetivamente sua mensagem.* 
 ![ Primeira execução-carrossel etapa 2-especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-step-2.png)


*Figura 3: forneça um plano de ação claro para sair da experiência de primeira execução.* 
 ![ Primeira execução-carrossel etapa 3-especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-step-3.png)



## <a name="value-placemat"></a>Roteiro de valor

O posicionamento do valor informa a proposta de valor do seu suplemento com posicionamento do logotipo, uma proposta de valor claramente definida, destaques ou resumo do recurso e uma chamada para ação.



![Roteiro de primeiro valor de execução-especificações do painel de tarefas da área de trabalho ](../images/add-in-FRE-value.png)
 *um valor roteiro com logotipo, proposta de valor clara, Resumo de recursos e plano de ação.*


### <a name="video-placemat"></a>Roteiro de vídeo

O roteiro de vídeo mostra aos usuários um vídeo antes que eles comecem a usar o suplemento.


*Figura 1: primeira execução roteiro-a tela contém uma imagem estática do vídeo com um botão Play e um botão limpar chamada para ação.* 
 ![ Roteiro de vídeo – especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-video.png)



*Figura 2: player de vídeo-os usuários são apresentados com um vídeo em uma janela de diálogo.* 
 ![ Vídeo roteiro-diálogo-especificações do painel de tarefas da área de trabalho](../images/add-in-FRE-video-dialog.png)
