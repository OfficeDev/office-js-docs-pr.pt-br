---
title: Idioma de design de suplemento do Office
description: Saiba como tornar seu Office de complemento visualmente compatível com Office.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 4374eaad1e660728a347d19a012d345b642755af
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330056"
---
# <a name="office-add-in-design-language"></a><span data-ttu-id="ae96a-103">Idioma de design de suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="ae96a-103">Office Add-in design language</span></span>

<span data-ttu-id="ae96a-p101">A linguagem de design do Office é um sistema visual claro e simples que garante a consistência nas experiências. Ela contém um conjunto de elementos visuais que definem as interfaces do Office, incluindo:</span><span class="sxs-lookup"><span data-stu-id="ae96a-p101">The Office design language is a clean and simple visual system that ensures consistency across experiences. It contains a set of visual elements that define Office interfaces, including:</span></span>

- <span data-ttu-id="ae96a-106">Um tipo de fonte padrão</span><span class="sxs-lookup"><span data-stu-id="ae96a-106">A standard typeface</span></span>
- <span data-ttu-id="ae96a-107">Uma paleta de cores comuns</span><span class="sxs-lookup"><span data-stu-id="ae96a-107">A common color palette</span></span>
- <span data-ttu-id="ae96a-108">Um conjunto de pesos e tamanhos tipográficos</span><span class="sxs-lookup"><span data-stu-id="ae96a-108">A set of typographic sizes and weights</span></span>
- <span data-ttu-id="ae96a-109">Diretrizes de ícones</span><span class="sxs-lookup"><span data-stu-id="ae96a-109">Icon guidelines</span></span>
- <span data-ttu-id="ae96a-110">Ativos de ícones compartilhados</span><span class="sxs-lookup"><span data-stu-id="ae96a-110">Shared icon assets</span></span>
- <span data-ttu-id="ae96a-111">Definições de animação</span><span class="sxs-lookup"><span data-stu-id="ae96a-111">Animation definitions</span></span>
- <span data-ttu-id="ae96a-112">Componentes comuns</span><span class="sxs-lookup"><span data-stu-id="ae96a-112">Common components</span></span>

<span data-ttu-id="ae96a-113">[A interface do usuário fluente](../design/add-in-design.md) é a estrutura front-end oficial para a criação com o Office de design.</span><span class="sxs-lookup"><span data-stu-id="ae96a-113">[Fluent UI](../design/add-in-design.md) is the official front-end framework for building with the Office design language.</span></span> <span data-ttu-id="ae96a-114">Usar a interface do usuário fluente é opcional, mas é a maneira mais rápida de garantir que seus complementos se sintam como uma extensão natural de Office.</span><span class="sxs-lookup"><span data-stu-id="ae96a-114">Using Fluent UI is optional, but it is the fastest way to ensure that your add-ins feel like a natural extension of Office.</span></span> <span data-ttu-id="ae96a-115">Aproveite a interface do usuário fluente para projetar e criar os complementos que complementam Office.</span><span class="sxs-lookup"><span data-stu-id="ae96a-115">Take advantage of Fluent UI to design and build add-ins that complement Office.</span></span>

<span data-ttu-id="ae96a-p103">Vários suplementos do Office estão associados a uma marca pré-existente. Você pode manter uma marca forte e sua linguagem visual ou de componente no suplemento. Procure oportunidades para manter sua própria linguagem visual durante a integração ao Office. Considere maneiras de substituir cores, tipografia, ícones ou outros elementos estilísticos pelos elementos de sua própria marca do Office. Considere maneiras de seguir layouts comuns de suplemento ou padrões de design da experiência do usuário durante a inserção de controles e componentes que são familiares para seus clientes.</span><span class="sxs-lookup"><span data-stu-id="ae96a-p103">Many Office Add-ins are associated with a preexisting brand. You can retain a strong brand and its visual or component language in your add-in. Look for opportunities to retain your own visual language while integrating with Office. Consider ways to swap out Office colors, typography, icons, or other stylistic elements with elements of your own brand. Consider ways to follow common add-in layouts or UX design patterns while inserting controls and components that are familiar to your customers.</span></span>

<span data-ttu-id="ae96a-p104">Inserir uma interface do usuário baseada em HTML com uma forte identidade visual no Office pode criar dissonâncias para os clientes. Encontre um equilíbrio que se ajuste perfeitamente ao Office, mas também se alinhe claramente à sua marca pai ou serviço. Quando um suplemento não se ajusta ao Office, normalmente é porque elementos estilísticos estão em conflito. Por exemplo, a tipografia é muito grande e está fora da grade, as cores são contrastantes ou particularmente fortes ou as animações são supérfluas e se comportam de maneira diferente do Office. A aparência e o comportamento de controles ou componentes se desviam demasiadamente dos padrões do Office.</span><span class="sxs-lookup"><span data-stu-id="ae96a-p104">Inserting a heavily branded HTML-based UI inside of Office can create dissonance for customers. Find a balance that fits seamlessly in Office but also clearly aligns with your service or parent brand. When an add-in does not fit with Office, it's often because stylistic elements conflict. For example, typography is too large and off grid, colors are contrasting or particularly loud, or animations are superfluous and behave differently than Office. The appearance and behavior of controls or components veer too far from Office standards.</span></span>
