---
title: Diretrizes de acessibilidade para suplementos do Office
description: Saiba como tornar o suplemento do Office acessível a todos os usuários.
ms.date: 09/24/2018
localization_priority: Normal
ms.openlocfilehash: 61028c86e9ff79271b67d217e2dc93df300af006
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718619"
---
# <a name="accessibility-guidelines"></a><span data-ttu-id="60155-103">Diretrizes de acessibilidade</span><span class="sxs-lookup"><span data-stu-id="60155-103">Accessibility guidelines</span></span>

<span data-ttu-id="60155-p101">À medida que você projeta e desenvolve seus suplementos do Office, convém verificar se todos os usuários e clientes potenciais são capazes de usar seu suplemento com êxito. Aplique as seguintes diretrizes para garantir que sua solução seja acessível a todos os públicos.</span><span class="sxs-lookup"><span data-stu-id="60155-p101">As you design and develop your Office Add-ins, you'll want to ensure that all potential users and customers are able to use your add-in successfully. Apply the following guidelines to ensure that your solution is accessible to all audiences.</span></span>

## <a name="design-for-multiple-input-methods"></a><span data-ttu-id="60155-106">Projetar para vários métodos de entrada</span><span class="sxs-lookup"><span data-stu-id="60155-106">Design for multiple input methods</span></span>

- <span data-ttu-id="60155-p102">Certifique-se de que os usuários possam realizar operações usando apenas o teclado. Os usuários devem conseguir se mover para todos os elementos acionáveis da página usando uma combinação das teclas Tab e de setas.</span><span class="sxs-lookup"><span data-stu-id="60155-p102">Ensure that users can perform operations by using only the keyboard. Users should be able to move to all actionable elements on the page by using a combination of the Tab and arrow keys.</span></span>
- <span data-ttu-id="60155-109">Em um dispositivo móvel, quando os usuários operam um controle por toque, o dispositivo deve fornecer um feedback sonoro útil.</span><span class="sxs-lookup"><span data-stu-id="60155-109">On a mobile device, when users operate a control by touch, the device should provide useful audio feedback.</span></span>
- <span data-ttu-id="60155-110">Forneça rótulos úteis para todos os controles interativos.</span><span class="sxs-lookup"><span data-stu-id="60155-110">Provide helpful labels for all interactive controls.</span></span> 

## <a name="make-your-add-in-easy-to-use"></a><span data-ttu-id="60155-111">Tornar seu suplemento fácil de usar</span><span class="sxs-lookup"><span data-stu-id="60155-111">Make your add-in easy to use</span></span>

- <span data-ttu-id="60155-112">Não dependa de um único atributo, como cor, tamanho, forma, local, orientação ou som, para atribuir significados na sua interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="60155-112">Don't rely on a single attribute, such as color, size, shape, location, orientation, or sound, to convey meaning in your UI.</span></span>
- <span data-ttu-id="60155-113">Evite alterações inesperadas de contexto, como mover o foco para outro elemento da interface do usuário sem uma ação do usuário.</span><span class="sxs-lookup"><span data-stu-id="60155-113">Avoid unexpected changes of context, such as moving the focus to a different UI element without user action.</span></span>
- <span data-ttu-id="60155-114">Ofereça uma maneira de verificar, confirmar ou reverter todas as ações de associação.</span><span class="sxs-lookup"><span data-stu-id="60155-114">Provide a way to verify, confirm, or reverse all binding actions.</span></span>
- <span data-ttu-id="60155-115">Forneça uma maneira de pausar ou parar mídias, como áudio e vídeo.</span><span class="sxs-lookup"><span data-stu-id="60155-115">Provide a way to pause or stop media, such as audio and video.</span></span>
- <span data-ttu-id="60155-116">Não estabeleça um limite de tempo para uma ação do usuário.</span><span class="sxs-lookup"><span data-stu-id="60155-116">Do not impose a time limit for user action.</span></span>

## <a name="make-your-add-in-easy-to-see"></a><span data-ttu-id="60155-117">Deixar seu suplemento fácil de ver</span><span class="sxs-lookup"><span data-stu-id="60155-117">Make your add-in easy to see</span></span>

- <span data-ttu-id="60155-118">Evite mudanças de cor inesperadas.</span><span class="sxs-lookup"><span data-stu-id="60155-118">Avoid unexpected color changes.</span></span>
- <span data-ttu-id="60155-p103">Forneça informações significativas e em tempo hábil para descrever elementos de interface do usuário, títulos e cabeçalhos, entradas e erros. Verifique se os nomes dos controles descrevem adequadamente o objetivo do controle.</span><span class="sxs-lookup"><span data-stu-id="60155-p103">Provide meaningful and timely information to describe UI elements, titles and headings, inputs, and errors. Ensure that names of controls adequately describe the intent of the control.</span></span>
- <span data-ttu-id="60155-121">Siga as [diretrizes padrão](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) de contraste de cor.</span><span class="sxs-lookup"><span data-stu-id="60155-121">Follow [standard guidelines](https://www.w3.org/TR/UNDERSTANDING-WCAG20/visual-audio-contrast-contrast.html) for color contrast.</span></span>

## <a name="account-for-assistive-technologies"></a><span data-ttu-id="60155-122">Incluir tecnologias adaptativas</span><span class="sxs-lookup"><span data-stu-id="60155-122">Account for assistive technologies</span></span>

- <span data-ttu-id="60155-123">Evite usar recursos que interfiram em tecnologias adaptativas, incluindo em interações visuais, auditivas ou outras.</span><span class="sxs-lookup"><span data-stu-id="60155-123">Avoid using features that interfere with assistive technologies, including visual, audio, or other interactions.</span></span>
- <span data-ttu-id="60155-p104">Não forneça o texto em um formato de imagem. Os leitores de tela não podem ler o texto em imagens.</span><span class="sxs-lookup"><span data-stu-id="60155-p104">Do not provide text in an image format. Screen readers cannot read text within images.</span></span>
- <span data-ttu-id="60155-126">Forneça uma maneira para os usuários ajustarem ou desativarem todas as fontes de áudio.</span><span class="sxs-lookup"><span data-stu-id="60155-126">Provide a way for users to adjust or mute all audio sources.</span></span>
- <span data-ttu-id="60155-127">Forneça uma maneira para os usuários ativarem legendas ou descrições de áudio com fontes de áudio.</span><span class="sxs-lookup"><span data-stu-id="60155-127">Provide a way for users to turn on captions or audio description with audio sources.</span></span>
- <span data-ttu-id="60155-128">Forneça alternativas para o som como um meio para alertar os usuários, como indicações visuais ou vibrações.</span><span class="sxs-lookup"><span data-stu-id="60155-128">Provide alternatives to sound as a means to alert users, such as visual cues or vibrations.</span></span>

## <a name="see-also"></a><span data-ttu-id="60155-129">Confira também</span><span class="sxs-lookup"><span data-stu-id="60155-129">See also</span></span>

- [<span data-ttu-id="60155-130">Diretrizes de Acessibilidade para Conteúdo da Web (WCAG) 2.0</span><span class="sxs-lookup"><span data-stu-id="60155-130">Web Content Accessibility Guidelines (WCAG) 2.0</span></span>](https://www.w3.org/TR/wcag2ict/#REF-WCAG20)
- [<span data-ttu-id="60155-131">Orientações sobre a Aplicação das WCAG 2.0 para Tecnologias de Comunicação e Informação que não Sejam da Web (WCAG2ICT)</span><span class="sxs-lookup"><span data-stu-id="60155-131">Guidance on Applying WCAG 2.0 to Non-Web Information and Communications Technologies (WCAG2ICT)</span></span>](https://www.w3.org/TR/wcag2ict/)
- [<span data-ttu-id="60155-132">Padrão Europeu para requisitos de acessibilidade para Tecnologias de Comunicação e Informação (ICT)</span><span class="sxs-lookup"><span data-stu-id="60155-132">European Standard on accessibility requirements for Information and Communication Technologies (ICT)</span></span>](https://www.etsi.org/deliver/etsi_en/301500_301599/301549/01.00.00_20/en_301549v010000c.pdf) 
