---
title: Design de suplementos do Outlook
description: Diretrizes para ajudar a projetar e construir um suplemento atraente, que oferece o melhor do seu aplicativo diretamente para o Outlook – no Windows, na Web, no iOS, no Mac e no Android.
ms.date: 06/24/2019
localization_priority: Priority
ms.openlocfilehash: a669d2cf0a98ffa0ca7b7dfc3fcc5b71d291a0e0
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077131"
---
# <a name="outlook-add-in-design-guidelines"></a><span data-ttu-id="e42d4-103">Diretrizes de design de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="e42d4-103">Outlook add-in design guidelines</span></span>

<span data-ttu-id="e42d4-p101">Os suplementos são uma ótima maneira de os parceiros estenderem a funcionalidade do Outlook para além do conjunto de recursos base. Os suplementos permitem que os usuários acessem experiências, tarefas e conteúdo de terceiros sem precisar sair da caixa de entrada. Uma vez instalados, os suplementos do Outlook estão disponíveis em todos os dispositivos e plataformas.</span><span class="sxs-lookup"><span data-stu-id="e42d4-p101">Add-ins are a great way for partners to extend the functionality of Outlook beyond our core feature set. Add-ins enable users to access third-party experiences, tasks, and content without needing to leave their inbox. Once installed, Outlook add-ins are available on every platform and device.</span></span>  

<span data-ttu-id="e42d4-107">As seguintes diretrizes gerais o ajudarão a projetar e construir um suplemento atraente, que oferece o melhor do seu aplicativo diretamente para o Outlook – no Windows, na Web, no iOS, no Mac e no Android.</span><span class="sxs-lookup"><span data-stu-id="e42d4-107">The following high-level guidelines will help you design and build a compelling add-in, which brings the best of your app right into Outlook&mdash;on Windows, Web, iOS, Mac, and Android.</span></span>

## <a name="principles"></a><span data-ttu-id="e42d4-108">Princípios</span><span class="sxs-lookup"><span data-stu-id="e42d4-108">Principles</span></span>

1. <span data-ttu-id="e42d4-109">**Se concentrar em algumas tarefas importantes. Realizá-las bem**</span><span class="sxs-lookup"><span data-stu-id="e42d4-109">**Focus on a few key tasks; do them well**</span></span>

   <span data-ttu-id="e42d4-p102">Os suplementos melhor projetados são fáceis de usar, concentrados e agregam valor real para os usuários. Como seu suplemento será executado dentro do Outlook, há ênfase adicional colocada nesse princípio. O Outlook é um aplicativo de produtividade – é onde as pessoas vão para realizar tarefas.</span><span class="sxs-lookup"><span data-stu-id="e42d4-p102">The best designed add-ins are simple to use, focused, and provide real value to users. Because your add-in will run inside of Outlook, there is additional emphasis placed on this principle. Outlook is a productivity app&mdash;it's where people go to get things done.</span></span>

   <span data-ttu-id="e42d4-p103">Você será uma extensão de nossa experiência e é importante para garantir que os cenários habilitados pareçam naturais dentro do Outlook. Considere cuidadosamente os casos de uso comuns que se beneficiarão mais de ter chamadas para eles dentro das experiências de email e calendário.</span><span class="sxs-lookup"><span data-stu-id="e42d4-p103">You will be an extension of our experience and it is important to make sure the scenarios you enable feel like a natural fit inside of Outlook. Think carefully about which of your common use cases will benefit the most from having hooks to them from within our email and calendaring experiences.</span></span>

   <span data-ttu-id="e42d4-p104">Um suplemento não deve tentar fazer tudo o que o seu aplicativo faz. O foco deve ser nas ações usadas com mais frequência e apropriadas, no contexto do conteúdo do Outlook. Pense em seu plano de chamada à ação e esclareça o que o usuário deve fazer quando abre o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="e42d4-p104">An add-in should not attempt to do everything your app does. The focus should be on the most frequently used, and appropriate, actions in the context of Outlook content. Think about your call to action and make it clear what the user should do when your task pane opens.</span></span>

2. <span data-ttu-id="e42d4-118">**Faça com que ele fique o mais nativo possível**</span><span class="sxs-lookup"><span data-stu-id="e42d4-118">**Make it feel as native as possible**</span></span>

   <span data-ttu-id="e42d4-p105">O suplemento deve ser projetado usando padrões nativos da plataforma na qual o Outlook estiver em execução. Para fazer isso, respeite e implemente as diretrizes de interação e visuais estabelecidas por cada plataforma. O Outlook tem suas próprias diretrizes e elas também é importante considerá-las. Um suplemento bem projetado será uma mistura apropriada de sua experiência, da plataforma e do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e42d4-p105">Your add-in should be designed using patterns native to the platform that Outlook is running on. To achieve this, be sure to respect and implement the interaction and visual guidelines set forth by each platform. Outlook has its own guidelines and those are also important to consider. A well-designed add-in will be an appropriate blend of your experience, the platform, and Outlook.</span></span>

   <span data-ttu-id="e42d4-p106">Isso significa que o suplemento será visualmente diferente quando for executado no Outlook no iOS em comparação com o Outlook no Android. Recomendamos que você confira o [Framework7](https://framework7.io/) como uma opção para ajudá-lo com o estilo.</span><span class="sxs-lookup"><span data-stu-id="e42d4-p106">This does mean that your add-in will have to visually be different when it runs in Outlook on iOS versus Android. We recommend taking a look at [Framework7](https://framework7.io/) as one option to help you with styling.</span></span>

3. <span data-ttu-id="e42d4-125">**Torne-o agradável de usar e acerte nos detalhes**</span><span class="sxs-lookup"><span data-stu-id="e42d4-125">**Make it enjoyable to use and get the details right**</span></span>

   <span data-ttu-id="e42d4-p107">As pessoas gostam de usar produtos funcionais e visualmente atraentes. Você pode ajudar a garantir o sucesso de seu suplemento ao criar uma experiência na qual você considerou cuidadosamente cada interação e detalhe visual. As etapas necessárias para concluir uma tarefa devem ser claras e relevantes. O ideal é que nenhuma ação leve mais do que um ou dois cliques.</span><span class="sxs-lookup"><span data-stu-id="e42d4-p107">People enjoy using products that are both functionally and visually appealing. You can help ensure the success of your add-in by crafting an experience where you've carefully considered every interaction and visual detail. The necessary steps to complete a task must be clear and relevant. Ideally, no action should be further than a click or two away.</span></span> 
   
   <span data-ttu-id="e42d4-130">Tente não tirar um usuário do contexto para concluir uma ação.</span><span class="sxs-lookup"><span data-stu-id="e42d4-130">Try not to take a user out of context to complete an action.</span></span> <span data-ttu-id="e42d4-131">Um usuário deve ter facilidade para entrar e sair de seu suplemento e voltar para o que estava fazendo antes.</span><span class="sxs-lookup"><span data-stu-id="e42d4-131">A user should easily be able to get in and out of your add-in and back to whatever she was doing before.</span></span> <span data-ttu-id="e42d4-132">Um suplemento não deve ser um destino onde se gaste muito tempo; ele deve ser um aprimoramento de nossa funcionalidade principal.</span><span class="sxs-lookup"><span data-stu-id="e42d4-132">An add-in is not meant to be a destination to spend a lot of time in&mdash;it is an enhancement to our core functionality.</span></span> <span data-ttu-id="e42d4-133">Se feito corretamente, seu suplemento nos ajudará a cumprir a meta de tornar as pessoas mais produtivas.</span><span class="sxs-lookup"><span data-stu-id="e42d4-133">If done properly, your add-in will help us deliver on the goal of making people more productive.</span></span>

4. <span data-ttu-id="e42d4-134">**Use sua marca de forma sensata**</span><span class="sxs-lookup"><span data-stu-id="e42d4-134">**Brand wisely**</span></span>

   <span data-ttu-id="e42d4-p109">Valorizamos uma ótima identidade visual, e sabemos que é importante proporcionar aos usuários sua experiência única. Mas sentimos que a melhor maneira de garantir o sucesso de seu suplemento é construir uma experiência intuitiva que incorpora sutilmente elementos de sua marca versus a exibição de elementos de marca persistentes ou intrusivos que apenas distraem um usuário de se mover através de seu sistema sem restrições.</span><span class="sxs-lookup"><span data-stu-id="e42d4-p109">We value great branding, and we know it is important to provide users with your unique experience. But we feel the best way to ensure your add-in's success is to build an intuitive experience that subtly incorporates elements of your brand versus displaying persistent or obtrusive brand elements that only distract a user from moving through your system in an unencumbered manner.</span></span> 
    
   <span data-ttu-id="e42d4-137">Uma boa maneira de incorporar sua marca de forma significativa é utilizar as cores, os ícones e a voz de sua marca, presumindo que esses itens não entrem em conflito com os padrões da plataforma de sua preferência ou os requisitos de acessibilidade.</span><span class="sxs-lookup"><span data-stu-id="e42d4-137">A good way to incorporate your brand in a meaningful way is through the use of your brand colors, icons, and voice&mdash;assuming these don't conflict with the preferred platform patterns or accessibility requirements.</span></span> <span data-ttu-id="e42d4-138">Tente manter o foco no conteúdo e na conclusão de tarefas, não na marca.</span><span class="sxs-lookup"><span data-stu-id="e42d4-138">Strive to keep the focus on content and task completion, not brand attention.</span></span> 
    
   > [!NOTE]
   >  <span data-ttu-id="e42d4-139">Os anúncios não devem ser mostrados em suplementos no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="e42d4-139">Ads should not be shown within add-ins on iOS or Android.</span></span>

## <a name="design-patterns"></a><span data-ttu-id="e42d4-140">Padrões de design</span><span class="sxs-lookup"><span data-stu-id="e42d4-140">Design patterns</span></span>

> [!NOTE]
> <span data-ttu-id="e42d4-141">Ainda que os princípios acima se apliquem a todos os pontos de extremidade/plataformas, os seguintes padrões e exemplos são específicos para suplementos de dispositivos móveis na plataforma iOS.</span><span class="sxs-lookup"><span data-stu-id="e42d4-141">While the above principles apply to all endpoints/platforms, the following patterns and examples are specific to mobile add-ins on the iOS platform.</span></span>

<span data-ttu-id="e42d4-p111">Para ajudar você a criar um suplemento bem projetado, temos [modelos](../design/ux-design-pattern-templates.md) que contêm padrões iOS de dispositivos móveis que funcionam no ambiente do Outlook Mobile. Aproveitar esses padrões específicos ajudará a garantir que seu suplemento pareça nativo na plataforma iOS e no Outlook Mobile. Esses padrões também estão detalhados abaixo. Embora não seja completa, esse é o início de uma biblioteca que continuaremos a criar conforme descobrimos novos parceiros de paradigmas que gostaríamos de incluir em seus suplementos.</span><span class="sxs-lookup"><span data-stu-id="e42d4-p111">To help you create a well-designed add-in, we have [templates](../design/ux-design-pattern-templates.md) that contain iOS mobile patterns that work within the Outlook Mobile environment. Leveraging these specific patterns will help ensure your add-in feels native to both the iOS platform and Outlook Mobile. These patterns are also detailed below. While not exhaustive, this is the start of a library that we will continue to build upon as we uncover additional paradigms partners wish to include in their add-ins.</span></span>  

### <a name="overview"></a><span data-ttu-id="e42d4-146">Visão geral</span><span class="sxs-lookup"><span data-stu-id="e42d4-146">Overview</span></span>

<span data-ttu-id="e42d4-147">Um suplemento típico é composto pelos seguintes componentes.</span><span class="sxs-lookup"><span data-stu-id="e42d4-147">A typical add-in is made up of the following components.</span></span>

![Diagrama de padrões UX básicos para um painel de tarefas no iOS.](../images/outlook-mobile-design-overview.png)

![Diagrama de padrões UX básicos para um painel de tarefas no Android.](../images/outlook-mobile-design-overview-android.jpg)

### <a name="loading"></a><span data-ttu-id="e42d4-150">Carregando</span><span class="sxs-lookup"><span data-stu-id="e42d4-150">Loading</span></span>

<span data-ttu-id="e42d4-p112">Quando um usuário toca no seu suplemento, a UX deverá ser exibida o mais rapidamente possível. Se houver qualquer atraso, use uma barra de progresso ou um indicador de atividade. Uma barra de progresso deve ser usada quando o período é determinável e um indicador de atividade deve ser usado quando o período não pode ser determinado.</span><span class="sxs-lookup"><span data-stu-id="e42d4-p112">When a user taps on your add-in, the UX should display as quickly as possible. If there is any delay, use a progress bar or activity indicator. A progress bar should be used when the amount of time is determinable and an activity indicator should be used when the amount of time is indeterminable.</span></span>

<span data-ttu-id="e42d4-154">**Um exemplo de páginas de carregamento no iOS**</span><span class="sxs-lookup"><span data-stu-id="e42d4-154">**An example of loading pages on iOS**</span></span>

![Exemplos de uma barra de progresso e um indicador de atividade no iOS.](../images/outlook-mobile-design-loading.png)

<span data-ttu-id="e42d4-156">**Um exemplo de páginas de carregamento no Android**</span><span class="sxs-lookup"><span data-stu-id="e42d4-156">**An example of loading pages on Android**</span></span>

![Exemplos de uma barra de progresso e um indicador de atividade no Android.](../images/outlook-mobile-design-loading-android.jpg)


### <a name="sign-insign-up"></a><span data-ttu-id="e42d4-158">Entrar/Inscrever-se</span><span class="sxs-lookup"><span data-stu-id="e42d4-158">Sign in/Sign up</span></span>

<span data-ttu-id="e42d4-159">Torne seu fluxo de entrada (e inscrição) simples e fácil de usar.</span><span class="sxs-lookup"><span data-stu-id="e42d4-159">Make your sign in (and sign up) flow straightforward and simple to use.</span></span>

<span data-ttu-id="e42d4-160">**Uma página de exemplo para entrar e se inscrever no iOS**</span><span class="sxs-lookup"><span data-stu-id="e42d4-160">**An example page to sign in and sign up on iOS**</span></span>

![Exemplos de páginas de entrada e inscrição no iOS.](../images/outlook-mobile-design-signin.png)

<span data-ttu-id="e42d4-162">**Uma página de entrada de exemplo no Android**</span><span class="sxs-lookup"><span data-stu-id="e42d4-162">**An example sign in page on Android**</span></span>

![Exemplos de página de entrada no Android.](../images/outlook-mobile-design-signin-android.png)

### <a name="brand-bar"></a><span data-ttu-id="e42d4-164">Barra da marca</span><span class="sxs-lookup"><span data-stu-id="e42d4-164">Brand bar</span></span>

<span data-ttu-id="e42d4-p113">A primeira tela do seu suplemento deve incluir o elemento de identidade visual. Projetada para reconhecimento, a barra de marca também ajuda a definir o contexto para o usuário. Como a barra de navegação contém o nome da sua empresa/marca, não é necessário repetir a barra de marca nas páginas seguintes.</span><span class="sxs-lookup"><span data-stu-id="e42d4-p113">The first screen of your add-in should include your branding element. Designed for recognition, the brand bar also helps set context for the user. Because the navigation bar contains the name of your company/brand, it's unnecessary to repeat the brand bar on subsequent pages.</span></span>

<span data-ttu-id="e42d4-168">**Um exemplo de identidade visual no iOS**</span><span class="sxs-lookup"><span data-stu-id="e42d4-168">**An example of branding on iOS**</span></span>

![Exemplos de barras de marca no iOS.](../images/outlook-mobile-design-branding.png)

<span data-ttu-id="e42d4-170">**Um exemplo de identidade visual no Android**</span><span class="sxs-lookup"><span data-stu-id="e42d4-170">**An example of branding on Android**</span></span>

![Exemplos de barras de marca no Android.](../images/outlook-mobile-design-branding-android.png)

### <a name="margins"></a><span data-ttu-id="e42d4-172">Margens</span><span class="sxs-lookup"><span data-stu-id="e42d4-172">Margins</span></span>

<span data-ttu-id="e42d4-173">Margens móveis devem ser definidas para 15px (8% da tela) de cada lado, para alinhar ao iOS do Outlook e 16px de cada lado para alinhar ao Android do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e42d4-173">Mobile margins should be set to 15px (8% of screen) for each side, to align with Outlook iOS and 16px for each side to align with Outlook Android.</span></span>

![Exemplos de margens no iOS.](../images/outlook-mobile-design-margins.png)

### <a name="typography"></a><span data-ttu-id="e42d4-175">Tipografia</span><span class="sxs-lookup"><span data-stu-id="e42d4-175">Typography</span></span>

<span data-ttu-id="e42d4-176">Uso da tipografia está alinhado ao Outlook iOS e é mantido simples para facilitar a análise.</span><span class="sxs-lookup"><span data-stu-id="e42d4-176">Typography usage is aligned to Outlook iOS and is kept simple for scannability.</span></span>

<span data-ttu-id="e42d4-177">**Tipografia no iOS**</span><span class="sxs-lookup"><span data-stu-id="e42d4-177">**Typography on iOS**</span></span>

![Exemplos de tipografia para iOS.](../images/outlook-mobile-design-typography.png)

<span data-ttu-id="e42d4-179">**Tipografia no Android**</span><span class="sxs-lookup"><span data-stu-id="e42d4-179">**Typography on Android**</span></span>

![Exemplos de tipografia para Android.](../images/outlook-mobile-design-typography-android.png)

### <a name="color-palette"></a><span data-ttu-id="e42d4-181">Paleta de cores</span><span class="sxs-lookup"><span data-stu-id="e42d4-181">Color palette</span></span>

<span data-ttu-id="e42d4-p114">O uso da cor é sutil no Outlook iOS.  Para alinhar, pedimos que o uso da cor esteja localizado nas ações e nos estados de erro, com apenas a marca de barra usando uma cor exclusiva.</span><span class="sxs-lookup"><span data-stu-id="e42d4-p114">Color usage is subtle in Outlook iOS.  To align, we ask that usage of color is localized to actions and error states, with only the brand bar using a unique color.</span></span>

![Paleta de cores para iOS.](../images/outlook-mobile-design-color-palette.png)

### <a name="cells"></a><span data-ttu-id="e42d4-185">Células</span><span class="sxs-lookup"><span data-stu-id="e42d4-185">Cells</span></span>

<span data-ttu-id="e42d4-186">Como a barra de navegação não pode ser usada para rotular uma página, use títulos de seção em páginas de etiquetas.</span><span class="sxs-lookup"><span data-stu-id="e42d4-186">Since the navigation bar cannot be used to label a page, use section titles to label pages.</span></span>

<span data-ttu-id="e42d4-187">**Exemplos de células no iOS**</span><span class="sxs-lookup"><span data-stu-id="e42d4-187">**Examples of cells on iOS**</span></span>

![Tipos de célula para iOS.](../images/outlook-mobile-design-cell-types.png)
* * *
![Ações recomendadas para células para iOS.](../images/outlook-mobile-design-cell-dos.png)
* * *
![Não recomendado para iOS.](../images/outlook-mobile-design-cell-donts.png)
* * *
![Células e entradas para iOS.](../images/outlook-mobile-design-cell-input.png)

<span data-ttu-id="e42d4-192">**Exemplos de células no Android**</span><span class="sxs-lookup"><span data-stu-id="e42d4-192">**Examples of cells on Android**</span></span>

![Tipos de célula para Android.](../images/outlook-mobile-design-cell-type-android.png)
* * *
![Ações recomendadas para células para Android.](../images/outlook-mobile-design-cell-dos-android.png)
* * *
![Ações não recomendadas para células para Android.](../images/outlook-mobile-design-cell-donts-android.png)
* * *
![Células e entradas para Android parte 1.](../images/outlook-mobile-design-cell-input-1-android.png)

![Células e entradas para Android parte 2.](../images/outlook-mobile-design-cell-input-2-android.png)

### <a name="actions"></a><span data-ttu-id="e42d4-198">Ações</span><span class="sxs-lookup"><span data-stu-id="e42d4-198">Actions</span></span>

<span data-ttu-id="e42d4-199">Mesmo que o aplicativo manipule uma infinidade de ações, considere as mais importantes que deseja que o suplemento execute e concentre-se nelas.</span><span class="sxs-lookup"><span data-stu-id="e42d4-199">Even if your app handles a multitude of actions, think about the most important ones you want your add-in to perform, and concentrate on those.</span></span>

<span data-ttu-id="e42d4-200">**Exemplos de ações no iOS**</span><span class="sxs-lookup"><span data-stu-id="e42d4-200">**Examples of actions on iOS**</span></span>

![Ações e células no iOS.](../images/outlook-mobile-design-action-cells.png)
* * *
![Ações recomendadas para iOS.](../images/outlook-mobile-design-action-dos.png)

<span data-ttu-id="e42d4-203">**Exemplos de ações no Android**</span><span class="sxs-lookup"><span data-stu-id="e42d4-203">**Examples of actions on Android**</span></span>

![Ações e células no Android.](../images/outlook-mobile-design-action-cells-android.png)
* * *
![Ações recomendadas para Android.](../images/outlook-mobile-design-action-dos-android.png)

### <a name="buttons"></a><span data-ttu-id="e42d4-206">Botões</span><span class="sxs-lookup"><span data-stu-id="e42d4-206">Buttons</span></span>

<span data-ttu-id="e42d4-207">Botões são usados quando existem outros elementos UX abaixo (versus ações, onde a ação é o último elemento na tela).</span><span class="sxs-lookup"><span data-stu-id="e42d4-207">Buttons are used when there are other UX elements below (vs. actions, where the action is the last element on the screen).</span></span>

<span data-ttu-id="e42d4-208">**Exemplos de botões no iOS**</span><span class="sxs-lookup"><span data-stu-id="e42d4-208">**Examples of buttons on iOS**</span></span>

![Exemplos de botões para iOS.](../images/outlook-mobile-design-buttons.png)

<span data-ttu-id="e42d4-210">**Exemplos de botões no Android**</span><span class="sxs-lookup"><span data-stu-id="e42d4-210">**Examples of buttons on Android**</span></span>

![Exemplos de botões para Android.](../images/outlook-mobile-design-buttons-android.png)

### <a name="tabs"></a><span data-ttu-id="e42d4-212">Guias</span><span class="sxs-lookup"><span data-stu-id="e42d4-212">Tabs</span></span>

<span data-ttu-id="e42d4-213">Guias podem auxiliar na organização do conteúdo.</span><span class="sxs-lookup"><span data-stu-id="e42d4-213">Tabs can aid in content organization.</span></span>

<span data-ttu-id="e42d4-214">**Exemplos de guias no iOS**</span><span class="sxs-lookup"><span data-stu-id="e42d4-214">**Examples of tabs on iOS**</span></span>

![Exemplos de guias para iOS.](../images/outlook-mobile-design-tabs.png)

<span data-ttu-id="e42d4-216">**Exemplos de guias no Android**</span><span class="sxs-lookup"><span data-stu-id="e42d4-216">**Examples of tabs on Android**</span></span>

![Exemplos de guias para Android.](../images/outlook-mobile-design-tabs-android.png)

### <a name="icons"></a><span data-ttu-id="e42d4-218">Ícones</span><span class="sxs-lookup"><span data-stu-id="e42d4-218">Icons</span></span>

<span data-ttu-id="e42d4-p115">Os ícones devem seguir o design atual do Outlook para iOS quando possível. Use nosso padrão tamanho e cor.</span><span class="sxs-lookup"><span data-stu-id="e42d4-p115">Icons should follow the current Outlook iOS design when possible. Use our standard size and color.</span></span>

<span data-ttu-id="e42d4-221">**Exemplos de ícones no iOS**</span><span class="sxs-lookup"><span data-stu-id="e42d4-221">**Examples of icons on iOS**</span></span>

![Exemplos de ícones para iOS.](../images/outlook-mobile-design-icons.png)

<span data-ttu-id="e42d4-223">**Exemplos de ícones no Android**</span><span class="sxs-lookup"><span data-stu-id="e42d4-223">**Examples of icons on Android**</span></span>

![Exemplos de ícones para Android.](../images/outlook-mobile-design-icons-android.jpg)

## <a name="end-to-end-examples"></a><span data-ttu-id="e42d4-225">Exemplos de ponta a ponta</span><span class="sxs-lookup"><span data-stu-id="e42d4-225">End-to-end examples</span></span>

<span data-ttu-id="e42d4-226">Para o lançamento de nossos suplementos do Outlook Mobile v1, trabalhamos junto a nossos parceiros que estavam criando suplementos. Como uma maneira de mostrar o potencial de seus suplementos no Outlook Mobile, nosso designer reuniu fluxos de ponta a ponta de cada suplemento, aproveitando nossas diretrizes e padrões.</span><span class="sxs-lookup"><span data-stu-id="e42d4-226">For our v1 Outlook Mobile Add-ins launch, we worked closely with our partners who were building add-ins. As a way to showcase the potential of their add-ins on Outlook Mobile, our designer put together end-to-end flows for each add-in, leveraging our guidelines and patterns.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e42d4-227">Estes exemplos destinam-se a realçar o modo ideal de abordar a interação e o design visual de um suplemento e podem não corresponder exatamente aos conjuntos de recursos nas versões enviadas dos suplementos.</span><span class="sxs-lookup"><span data-stu-id="e42d4-227">These examples are meant to highlight the ideal way to approach both the interaction and visual design of an add-in and may not match the exact feature sets in the shipped versions of the add-ins.</span></span> 

### <a name="giphy"></a><span data-ttu-id="e42d4-228">GIPHY</span><span class="sxs-lookup"><span data-stu-id="e42d4-228">GIPHY</span></span>

<span data-ttu-id="e42d4-229">**Um exemplo do GIPHY no iOS**</span><span class="sxs-lookup"><span data-stu-id="e42d4-229">**An example of GIPHY on iOS**</span></span>

![Design de ponta a ponta para o suplemento GIPHY no iOS.](../images/outlook-mobile-design-giphy.png)

<span data-ttu-id="e42d4-231">**Um exemplo do GIPHY no Android**</span><span class="sxs-lookup"><span data-stu-id="e42d4-231">**An example of GIPHY on Android**</span></span>

![Design de ponta a ponta para o suplemento GIPHY no Android.](../images/outlook-mobile-design-giphy-android.png)

### <a name="nimble"></a><span data-ttu-id="e42d4-233">Nimble</span><span class="sxs-lookup"><span data-stu-id="e42d4-233">Nimble</span></span>

<span data-ttu-id="e42d4-234">**Um exemplo do Nimble no iOS**</span><span class="sxs-lookup"><span data-stu-id="e42d4-234">**An example of Nimble on iOS**</span></span>

![Design de ponta a ponta para o suplemento Nimble no iOS.](../images/outlook-mobile-design-nimble.png)

<span data-ttu-id="e42d4-236">**Um exemplo do Nimble no Android**</span><span class="sxs-lookup"><span data-stu-id="e42d4-236">**An example of Nimble on Android**</span></span>

![Design de ponta a ponta para o suplemento Nimble no Android.](../images/outlook-mobile-design-nimble-android.png)

### <a name="trello"></a><span data-ttu-id="e42d4-238">Trello</span><span class="sxs-lookup"><span data-stu-id="e42d4-238">Trello</span></span>

<span data-ttu-id="e42d4-239">**Um exemplo do Trello no iOS**</span><span class="sxs-lookup"><span data-stu-id="e42d4-239">**An example of Trello on iOS**</span></span>

![Design de ponta a ponta para o suplemento Trello, parte 1, no iOS.](../images/outlook-mobile-design-trello-1.png)
* * *
![Design de ponta a ponta para o suplemento Trello, parte 2, no iOS.](../images/outlook-mobile-design-trello-2.png)
* * *
![Design de ponta a ponta para o suplemento Trello, parte 3, no iOS.](../images/outlook-mobile-design-trello-3.png)

<span data-ttu-id="e42d4-243">**Um exemplo do Trello no Android**</span><span class="sxs-lookup"><span data-stu-id="e42d4-243">**An example of Trello on Android**</span></span>

![Design de ponta a ponta para o suplemento Trello, parte 1, no Android.](../images/outlook-mobile-design-trello-1-android.png)
* * *
![Design de ponta a ponta para o suplemento Trello, parte 2, no Android.](../images/outlook-mobile-design-trello-2-android.png)

### <a name="dynamics-crm"></a><span data-ttu-id="e42d4-246">Dynamics CRM</span><span class="sxs-lookup"><span data-stu-id="e42d4-246">Dynamics CRM</span></span>

<span data-ttu-id="e42d4-247">**Um exemplo do Dynamics CRM no iOS**</span><span class="sxs-lookup"><span data-stu-id="e42d4-247">**An example of Dynamics CRM on iOS**</span></span>

![Design de ponta a ponta para o suplemento Dynamics CRM no iOS.](../images/outlook-mobile-design-crm.png)

<span data-ttu-id="e42d4-249">**Um exemplo do Dynamics CRM no Android**</span><span class="sxs-lookup"><span data-stu-id="e42d4-249">**An example of Dynamics CRM on Android**</span></span>

![Design de ponta a ponta para o suplemento Dynamics CRM no Android.](../images/outlook-mobile-design-crm-android.png)
