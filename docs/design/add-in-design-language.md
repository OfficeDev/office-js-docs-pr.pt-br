---
title: Idioma de design de suplemento do Office
description: ''
ms.date: 12/04/2017
localization_priority: Priority
ms.openlocfilehash: abeeca02687c44f6a3cc513867ff3eb637c07348
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388693"
---
# <a name="office-add-in-design-language"></a><span data-ttu-id="d7b93-102">Idioma de design de suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="d7b93-102">Office Add-in design language</span></span>

<span data-ttu-id="d7b93-p101">A linguagem de design do Office é um sistema visual claro e simples que garante a consistência nas experiências. Ela contém um conjunto de elementos visuais que definem as interfaces do Office, incluindo:</span><span class="sxs-lookup"><span data-stu-id="d7b93-p101">The Office design language is a clean and simple visual system that ensures consistency across experiences. It contains a set of visual elements that define Office interfaces, including:</span></span>

- <span data-ttu-id="d7b93-105">Um tipo de fonte padrão</span><span class="sxs-lookup"><span data-stu-id="d7b93-105">A standard typeface</span></span>
- <span data-ttu-id="d7b93-106">Uma paleta de cores comuns</span><span class="sxs-lookup"><span data-stu-id="d7b93-106">A common color palette</span></span>
- <span data-ttu-id="d7b93-107">Um conjunto de pesos e tamanhos tipográficos</span><span class="sxs-lookup"><span data-stu-id="d7b93-107">A set of typographic sizes and weights</span></span>
- <span data-ttu-id="d7b93-108">Diretrizes de ícones</span><span class="sxs-lookup"><span data-stu-id="d7b93-108">Icon guidelines</span></span>
- <span data-ttu-id="d7b93-109">Ativos de ícones compartilhados</span><span class="sxs-lookup"><span data-stu-id="d7b93-109">Shared icon assets</span></span>
- <span data-ttu-id="d7b93-110">Definições de animação</span><span class="sxs-lookup"><span data-stu-id="d7b93-110">Animation definitions</span></span>
- <span data-ttu-id="d7b93-111">Componentes comuns</span><span class="sxs-lookup"><span data-stu-id="d7b93-111">Common components</span></span>

<span data-ttu-id="d7b93-p102">O [Office UI Fabric](https://developer.microsoft.com/fabric) é a estrutura de front-end oficial para criação com a linguagem de design do Office. O uso do Fabric é opcional, mas é a maneira mais rápida de garantir que os suplementos sejam como uma extensão natural do Office. Tire proveito do Fabric para projetar e criar suplementos que complementam o Office.</span><span class="sxs-lookup"><span data-stu-id="d7b93-p102">[Office UI Fabric](https://developer.microsoft.com/fabric) is the official front-end framework for building with the Office design language. Using Fabric is optional, but it is the fastest way to ensure that your add-ins feel like a natural extension of Office. Take advantage of Fabric to design and build add-ins that complement Office.</span></span>

<span data-ttu-id="d7b93-p103">Vários suplementos do Office estão associados a uma marca pré-existente. Você pode manter uma marca forte e sua linguagem visual ou de componente no suplemento. Procure oportunidades para manter sua própria linguagem visual durante a integração ao Office. Considere maneiras de substituir cores, tipografia, ícones ou outros elementos estilísticos pelos elementos de sua própria marca do Office. Considere maneiras de seguir layouts comuns de suplemento ou padrões de design da experiência do usuário durante a inserção de controles e componentes que são familiares para seus clientes.</span><span class="sxs-lookup"><span data-stu-id="d7b93-p103">Many Office Add-ins are associated with a preexisting brand. You can retain a strong brand and its visual or component language in your add-in. Look for opportunities to retain your own visual language while integrating with Office. Consider ways to swap out Office colors, typography, icons, or other stylistic elements with elements of your own brand. Consider ways to follow common add-in layouts or UX design patterns while inserting controls and components that are familiar to your customers.</span></span>

<span data-ttu-id="d7b93-p104">Inserir uma interface do usuário baseada em HTML com uma forte identidade visual no Office pode criar dissonâncias para os clientes. Encontre um equilíbrio que se ajuste perfeitamente ao Office, mas também se alinhe claramente à sua marca pai ou serviço. Quando um suplemento não se ajusta ao Office, normalmente é porque elementos estilísticos estão em conflito. Por exemplo, a tipografia é muito grande e está fora da grade, as cores são contrastantes ou particularmente fortes ou as animações são supérfluas e se comportam de maneira diferente do Office. A aparência e o comportamento de controles ou componentes se desviam demasiadamente dos padrões do Office.</span><span class="sxs-lookup"><span data-stu-id="d7b93-p104">Inserting a heavily branded HTML-based UI inside of Office can create dissonance for customers. Find a balance that fits seamlessly in Office but also clearly aligns with your service or parent brand. When an add-in does not fit with Office, it's often because stylistic elements conflict. For example, typography is too large and off grid, colors are contrasting or particularly loud, or animations are superfluous and behave differently than Office. The appearance and behavior of controls or components veer too far from Office standards.</span></span>
