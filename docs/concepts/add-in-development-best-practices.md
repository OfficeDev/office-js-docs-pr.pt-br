---
title: Práticas recomendadas para o desenvolvimento de suplementos do Office
description: Aplique as práticas recomendadas ao desenvolver para criar Office de complementos.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 06b7f74692edbba1bc0ecdde723c4a661e830970
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330077"
---
# <a name="best-practices-for-developing-office-add-ins"></a><span data-ttu-id="ec46f-103">Práticas recomendadas para o desenvolvimento de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ec46f-103">Best practices for developing Office Add-ins</span></span>

<span data-ttu-id="ec46f-p101">Os suplementos eficazes oferecem uma funcionalidade exclusiva e fascinante que estende os aplicativos do Office de uma maneira visualmente atraente. Para criar um excelente suplemento, ofereça uma primeira experiência envolvente para seus usuários, desenvolva uma experiência de interface de usuário de alto nível e otimize o desempenho do seu suplemento. Aplique as práticas recomendadas descritas neste artigo para criar suplementos que ajudem os usuários a concluir suas tarefas de forma rápida e eficiente.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p101">Effective add-ins offer unique and compelling functionality that extends Office applications in a visually appealing way. To create a great add-in, provide an engaging first-time experience for your users, design a first-class UI experience, and optimize your add-in's performance. Apply the best practices described in this article to create add-ins that help your users complete their tasks quickly and efficiently.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="provide-clear-value"></a><span data-ttu-id="ec46f-107">Fornecer um valor claro</span><span class="sxs-lookup"><span data-stu-id="ec46f-107">Provide clear value</span></span>

- <span data-ttu-id="ec46f-p102">Crie suplementos que ajudem os usuários a concluir tarefas de forma rápida e eficiente. Concentre-se nos cenários que fazem sentido para aplicativos do Office. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="ec46f-p102">Create add-ins that help users complete tasks quickly and efficiently. Focus on scenarios that make sense for Office applications. For example:</span></span>
  - <span data-ttu-id="ec46f-111">Torne as principais tarefas de criação mais rápidas e fáceis, com menos interrupções.</span><span class="sxs-lookup"><span data-stu-id="ec46f-111">Make core authoring tasks faster and easier, with fewer interruptions.</span></span>
  - <span data-ttu-id="ec46f-112">Habilite novos cenários no Office.</span><span class="sxs-lookup"><span data-stu-id="ec46f-112">Enable new scenarios within Office.</span></span>
  - <span data-ttu-id="ec46f-113">Incorporar serviços complementares em Office aplicativos.</span><span class="sxs-lookup"><span data-stu-id="ec46f-113">Embed complementary services within Office applications.</span></span>
  - <span data-ttu-id="ec46f-114">Melhore a experiência do Office para aumentar a produtividade.</span><span class="sxs-lookup"><span data-stu-id="ec46f-114">Improve the Office experience to enhance productivity.</span></span>
- <span data-ttu-id="ec46f-115">Certifique-se de que o valor do seu suplemento seja claro para os usuários desde o princípio, [criando uma experiência envolvente na primeira execução](#create-an-engaging-first-run-experience).</span><span class="sxs-lookup"><span data-stu-id="ec46f-115">Make sure that the value of your add-in is clear to users right away by [creating an engaging first run experience](#create-an-engaging-first-run-experience).</span></span>
- <span data-ttu-id="ec46f-p103">Crie uma [listagem eficaz do AppSource](/office/dev/store/create-effective-office-store-listings). Deixe claro quais são os benefícios do seu suplemento no título e na descrição. Não dependa da sua marca para dizer o que seu suplemento faz.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p103">Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings). Make the benefits of your add-in clear in your title and description. Don't rely on your brand to communicate what your add-in does.</span></span>

## <a name="create-an-engaging-first-run-experience"></a><span data-ttu-id="ec46f-119">Criar uma experiência envolvente na primeira execução</span><span class="sxs-lookup"><span data-stu-id="ec46f-119">Create an engaging first-run experience</span></span>

- <span data-ttu-id="ec46f-p104">Envolva os novos usuários com uma primeira experiência altamente útil e intuitiva. Observe que, mesmo depois de baixar o suplemento da loja, os usuários ainda estão decidindo se vão utilizá-lo.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p104">Engage new users with a highly usable and intuitive first experience. Note that users are still deciding whether to use or abandon an add-in after they download it from the store.</span></span>

- <span data-ttu-id="ec46f-p105">Deixe claro quais são as etapas que usuário terá que seguir para se envolver com seu suplemento. Use vídeos, diagramas, painéis de paginação ou outros recursos para atrair usuários.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p105">Make the steps that the user needs to take to engage with your add-in clear. Use videos, placemats, paging panels, or other resources to entice users.</span></span>

- <span data-ttu-id="ec46f-124">Reforce a proposta de valor do seu suplemento no início, em vez de apenas pedir que seus usuários entrem.</span><span class="sxs-lookup"><span data-stu-id="ec46f-124">Reinforce the value proposition of your add-in on launch, rather than just asking users to sign in.</span></span>

- <span data-ttu-id="ec46f-125">Forneça uma interface do usuário informativa e torne sua interface do usuário pessoal.</span><span class="sxs-lookup"><span data-stu-id="ec46f-125">Provide teaching UI to guide users and make your UI personal.</span></span>

  ![Captura de tela mostrando uma comparação "Do" vs. "Don't".](../images/contoso-part-catalog-do-dont.png)

- <span data-ttu-id="ec46f-129">Se seu suplemento de conteúdo estiver vinculado a dados no documento do usuário, inclua exemplos de dados ou um modelo para mostrar aos usuários o formato de dados a ser usado.</span><span class="sxs-lookup"><span data-stu-id="ec46f-129">If your content add-in binds to data in the user's document, include sample data or a template to show users the data format to use.</span></span>

  ![Captura de tela mostrando uma comparação "Do" vs. "Don't".](../images/add-in-title.png)

- <span data-ttu-id="ec46f-p108">Ofereça [avaliações gratuitas](/office/dev/store/decide-on-a-pricing-model). Caso o suplemento exija uma assinatura, disponibilize algumas funcionalidades sem a necessidade da assinatura.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p108">Offer [free trials](/office/dev/store/decide-on-a-pricing-model). If your add-in requires a subscription, make some functionality available without a subscription.</span></span>

- <span data-ttu-id="ec46f-p109">Simplifique o processo de inscrição. Preencha automaticamente as informações (email, nome de exibição) e ignore as verificações de email.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p109">Make signup simple. Prefill information (email, display name) and skip email verifications.</span></span>

- <span data-ttu-id="ec46f-p110">Evite os pop-ups. Se você tiver de usá-los, oriente o usuário sobre como habilitar o seu pop-up.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p110">Avoid pop ups. If you have to use them, guide the user to enable your pop up.</span></span>

<span data-ttu-id="ec46f-139">Para padrões que podem ser aplicados ao desenvolver sua experiência de primeira execução, consulte [Padrões de design da experiência do usuário para suplementos do Office](../design/first-run-experience-patterns.md).</span><span class="sxs-lookup"><span data-stu-id="ec46f-139">For patterns that you can apply as you develop your first-run experience, see [UX design patterns for Office Add-ins](../design/first-run-experience-patterns.md).</span></span>

## <a name="use-add-in-commands"></a><span data-ttu-id="ec46f-140">Usar comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="ec46f-140">Use add-in commands</span></span>

- <span data-ttu-id="ec46f-p111">Fornece ao suplemento pontos de entrada relevantes da interface do usuário usando os comandos do suplemento. Confira mais detalhes, inclusive as práticas recomendadas de design, nos [comandos de suplemento](../design/add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="ec46f-p111">Provide relevant UI entry points for your add-in by using add-in commands. For details, including design best practices, see [add-in commands](../design/add-in-commands.md).</span></span>

## <a name="apply-ux-design-principles"></a><span data-ttu-id="ec46f-143">Aplicar os princípios de design de UX</span><span class="sxs-lookup"><span data-stu-id="ec46f-143">Apply UX design principles</span></span>

- <span data-ttu-id="ec46f-144">Assegure-se de que a aparência e a funcionalidade de seus suplementos complementam a experiência do Office.</span><span class="sxs-lookup"><span data-stu-id="ec46f-144">Ensure that the look and feel and functionality of your add-in complements the Office experience.</span></span> <span data-ttu-id="ec46f-145">Consulte [Design the UI of Office Add-ins](../design/add-in-design.md).</span><span class="sxs-lookup"><span data-stu-id="ec46f-145">See [Design the UI of Office Add-ins](../design/add-in-design.md).</span></span>

- <span data-ttu-id="ec46f-p113">Favoreça o conteúdo através do Chrome. Evite elementos de interface do usuário supérfluos que não agregam valor à experiência do usuário.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p113">Favor content over chrome. Avoid superfluous UI elements that don't add value to the user experience.</span></span>

- <span data-ttu-id="ec46f-p114">Mantenha os usuários no controle. Verifique se os usuários compreenderam as decisões importantes e podem reverter facilmente as ações realizadas pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p114">Keep users in control. Ensure that users understand important decisions, and can easily reverse actions the add-in performs.</span></span>

- <span data-ttu-id="ec46f-p115">Use uma identidade visual para inspirar confiança e orientar os usuários. Não use o recurso de identidade visual para sobrecarregar ou enviar anúncios aos usuários.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p115">Use branding to inspire trust and orient users. Do not use branding to overwhelm or advertise to users.</span></span>

- <span data-ttu-id="ec46f-p116">Evite a necessidade de rolagem. Otimize para a resolução 1366 x 768.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p116">Avoid scrolling. Optimize for 1366 x 768 resolution.</span></span>

- <span data-ttu-id="ec46f-154">Não inclua imagens não licenciadas.</span><span class="sxs-lookup"><span data-stu-id="ec46f-154">Do not include unlicensed images.</span></span>

- <span data-ttu-id="ec46f-155">Use uma [linguagem clara e simples](../design/voice-guidelines.md) no seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="ec46f-155">Use [clear and simple language](../design/voice-guidelines.md) in your add-in.</span></span>

- <span data-ttu-id="ec46f-156">Preocupe-se com a acessibilidade: facilite a interação dos usuários com o seu suplemento e inclua tecnologias adaptativas, como leitores de tela.</span><span class="sxs-lookup"><span data-stu-id="ec46f-156">Account for accessibility - make your add-in easy for all users to interact with, and accommodate assistive technologies such as screen readers.</span></span>

- <span data-ttu-id="ec46f-p117">Desenvolva para todas as plataformas e métodos de entrada, incluindo teclado/mouse e [toque](#optimize-for-touch). Certifique-se de que sua interface do usuário responda a diferentes fatores forma.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p117">Design for all platforms and input methods, including mouse/keyboard and [touch](#optimize-for-touch). Ensure that your UI is responsive to different form factors.</span></span>

### <a name="optimize-for-touch"></a><span data-ttu-id="ec46f-159">Otimizar para toque</span><span class="sxs-lookup"><span data-stu-id="ec46f-159">Optimize for touch</span></span>

- <span data-ttu-id="ec46f-160">Use a [propriedade Context.touchEnabled](/javascript/api/office/office.context#touchenabled) para detectar se o aplicativo Office que o seu complemento executa está habilitado para toque.</span><span class="sxs-lookup"><span data-stu-id="ec46f-160">Use the [Context.touchEnabled](/javascript/api/office/office.context#touchenabled) property to detect whether the Office application that your add-in runs on is touch enabled.</span></span>

  > [!NOTE]
  > <span data-ttu-id="ec46f-161">Essa propriedade não tem suporte no Outlook.</span><span class="sxs-lookup"><span data-stu-id="ec46f-161">This property is not supported in Outlook.</span></span>

- <span data-ttu-id="ec46f-p118">Verifique se todos os controles são dimensionados adequadamente para interação por toque. Por exemplo, se os botões têm destinos de toque adequados e se as caixas de entrada têm a dimensão correta para que os usuários insiram entradas.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p118">Ensure that all controls are appropriately sized for touch interaction. For example, buttons have adequate touch targets, and input boxes are large enough for users to enter input.</span></span>

- <span data-ttu-id="ec46f-164">Não confie nos métodos de entrada sem toque, como passar o cursor ou clicar com o botão direito do mouse.</span><span class="sxs-lookup"><span data-stu-id="ec46f-164">Do not rely on non-touch input methods like hover or right-click.</span></span>

- <span data-ttu-id="ec46f-p119">Verifique se o suplemento funciona nos modos retrato e paisagem. Observe que em dispositivos de toque, parte do suplemento pode ficar oculta pelo teclado virtual.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p119">Ensure that your add-in works in both portrait and landscape modes. Be aware that on touch devices, part of your add-in might be hidden by the soft keyboard.</span></span>

- <span data-ttu-id="ec46f-167">Teste seu suplemento em um dispositivo real usando o [sideload](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="ec46f-167">Test your add-in on a real device by using [sideloading](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).</span></span>

> [!NOTE]
> <span data-ttu-id="ec46f-168">Se você estiver usando a [interface do usuário do Fluent React](../design/using-office-ui-fabric-react.md) seus elementos de design, muitos desses elementos são integrados ao sistema de design.</span><span class="sxs-lookup"><span data-stu-id="ec46f-168">If you're using [Fluent UI React](../design/using-office-ui-fabric-react.md) for your design elements, many of these elements are built into the design system.</span></span>


## <a name="optimize-and-monitor-add-in-performance"></a><span data-ttu-id="ec46f-169">Otimizar e monitorar o desempenho do suplemento</span><span class="sxs-lookup"><span data-stu-id="ec46f-169">Optimize and monitor add-in performance</span></span>

- <span data-ttu-id="ec46f-p120">Crie a percepção de respostas rápidas da interface do usuário. Seu suplemento deverá ser carregado em 500 ms ou menos.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p120">Create the perception of fast UI responses. Your add-in should load in 500 ms or less.</span></span>

- <span data-ttu-id="ec46f-172">Certifique-se de que todas as interações do usuário respondam em menos de um segundo.</span><span class="sxs-lookup"><span data-stu-id="ec46f-172">Ensure that all user interactions respond in under one second.</span></span>

- <span data-ttu-id="ec46f-173">Forneça indicadores de carregamento para operações com longa execução.</span><span class="sxs-lookup"><span data-stu-id="ec46f-173">Provide loading indicators for long-running operations.</span></span>

- <span data-ttu-id="ec46f-p121">Use uma CDN para hospedar imagens, recursos e bibliotecas comuns. Carregue o máximo possível de um só lugar.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p121">Use a CDN to host images, resources, and common libraries. Load as much as you can from one place.</span></span>

- <span data-ttu-id="ec46f-p122">Siga as práticas da Web padrão para otimizar a página. Use apenas versões reduzidas das bibliotecas na produção. Carregue somente os recursos que você precisar e otimize como os recursos são carregados.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p122">Follow standard web practices to optimize your web page. In production, use only minified versions of libraries. Only load resources that you need, and optimize how resources are loaded.</span></span>

- <span data-ttu-id="ec46f-p123">Se o tempo de execução das operações demorar, forneça feedback aos usuários. Observe os limites relacionados na tabela a seguir. Saiba mais em [Limites de recurso e otimização de desempenho para Suplementos do Office](../concepts/resource-limits-and-performance-optimization.md).</span><span class="sxs-lookup"><span data-stu-id="ec46f-p123">If operations take time to execute, provide feedback to users. Note the thresholds listed in the following table. For additional information, see [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md).</span></span>

  |<span data-ttu-id="ec46f-182">Classe de interação</span><span class="sxs-lookup"><span data-stu-id="ec46f-182">Interaction class</span></span>|<span data-ttu-id="ec46f-183">Target</span><span class="sxs-lookup"><span data-stu-id="ec46f-183">Target</span></span>|<span data-ttu-id="ec46f-184">Limite superior</span><span class="sxs-lookup"><span data-stu-id="ec46f-184">Upper bound</span></span>|<span data-ttu-id="ec46f-185">Percepção humana</span><span class="sxs-lookup"><span data-stu-id="ec46f-185">Human perception</span></span>|
  |:-----|:-----|:-----|:-----|
  |<span data-ttu-id="ec46f-186">Instantâneo</span><span class="sxs-lookup"><span data-stu-id="ec46f-186">Instant</span></span>|<span data-ttu-id="ec46f-187"><=50 ms</span><span class="sxs-lookup"><span data-stu-id="ec46f-187"><=50 ms</span></span>|<span data-ttu-id="ec46f-188">100 ms</span><span class="sxs-lookup"><span data-stu-id="ec46f-188">100 ms</span></span>|<span data-ttu-id="ec46f-189">Nenhum atraso considerável.</span><span class="sxs-lookup"><span data-stu-id="ec46f-189">No noticeable delay.</span></span>|
  |<span data-ttu-id="ec46f-190">Rápida</span><span class="sxs-lookup"><span data-stu-id="ec46f-190">Fast</span></span>|<span data-ttu-id="ec46f-191">50 – 100 ms.</span><span class="sxs-lookup"><span data-stu-id="ec46f-191">50-100 ms</span></span>|<span data-ttu-id="ec46f-192">200 ms</span><span class="sxs-lookup"><span data-stu-id="ec46f-192">200 ms</span></span>|<span data-ttu-id="ec46f-p124">Atraso mínimo considerável. Não são necessários comentários.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p124">Minimally noticeable delay. No feedback necessary.</span></span>|
  |<span data-ttu-id="ec46f-195">Típico</span><span class="sxs-lookup"><span data-stu-id="ec46f-195">Typical</span></span>|<span data-ttu-id="ec46f-196">100 – 300 ms</span><span class="sxs-lookup"><span data-stu-id="ec46f-196">100-300 ms</span></span>|<span data-ttu-id="ec46f-197">500 ms</span><span class="sxs-lookup"><span data-stu-id="ec46f-197">500 ms</span></span>|<span data-ttu-id="ec46f-p125">Rápido, mas não o suficiente para ser descrito como rápido. Não são necessários comentários.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p125">Quick, but too slow to be described as fast. No feedback necessary.</span></span>|
  |<span data-ttu-id="ec46f-200">Dinâmico</span><span class="sxs-lookup"><span data-stu-id="ec46f-200">Responsive</span></span>|<span data-ttu-id="ec46f-201">300 – 500 ms.</span><span class="sxs-lookup"><span data-stu-id="ec46f-201">300-500 ms</span></span>|<span data-ttu-id="ec46f-202">1 segundo</span><span class="sxs-lookup"><span data-stu-id="ec46f-202">1 second</span></span>|<span data-ttu-id="ec46f-p126">Não muito rápido, embora pareça ser dinâmico. Não são necessários comentários.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p126">Not fast, but still feels responsive. No feedback necessary.</span></span>|
  |<span data-ttu-id="ec46f-205">Contínuo</span><span class="sxs-lookup"><span data-stu-id="ec46f-205">Continuous</span></span>|<span data-ttu-id="ec46f-206">>500 ms</span><span class="sxs-lookup"><span data-stu-id="ec46f-206">>500 ms</span></span>|<span data-ttu-id="ec46f-207">5 segundos</span><span class="sxs-lookup"><span data-stu-id="ec46f-207">5 seconds</span></span>|<span data-ttu-id="ec46f-p127">Tempo de espera médio, já não parece ser dinâmico. Podem ser necessários comentários.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p127">Medium wait, no longer feels responsive. Might need feedback.</span></span>|
  |<span data-ttu-id="ec46f-210">Cativo</span><span class="sxs-lookup"><span data-stu-id="ec46f-210">Captive</span></span>|<span data-ttu-id="ec46f-211">>500 ms</span><span class="sxs-lookup"><span data-stu-id="ec46f-211">>500 ms</span></span>|<span data-ttu-id="ec46f-212">10 segundos</span><span class="sxs-lookup"><span data-stu-id="ec46f-212">10 seconds</span></span>|<span data-ttu-id="ec46f-p128">Longo, mas não o suficiente para fazer executar outra ação. Podem ser necessários comentários.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p128">Long, but not long enough to do something else. Might need feedback.</span></span>|
  |<span data-ttu-id="ec46f-215">Estendida</span><span class="sxs-lookup"><span data-stu-id="ec46f-215">Extended</span></span>|<span data-ttu-id="ec46f-216">>500 ms</span><span class="sxs-lookup"><span data-stu-id="ec46f-216">>500 ms</span></span>|<span data-ttu-id="ec46f-217">>10 segundos</span><span class="sxs-lookup"><span data-stu-id="ec46f-217">>10 seconds</span></span>|<span data-ttu-id="ec46f-p129">Longo o suficiente para realizar outra ação durante o tempo de espera. Podem ser necessários comentários.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p129">Long enough to do something else while waiting. Might need feedback.</span></span>|
  |<span data-ttu-id="ec46f-220">Longa execução</span><span class="sxs-lookup"><span data-stu-id="ec46f-220">Long running</span></span>|<span data-ttu-id="ec46f-221">> 5 segundos</span><span class="sxs-lookup"><span data-stu-id="ec46f-221">>5 seconds</span></span>|<span data-ttu-id="ec46f-222">> 1 minuto</span><span class="sxs-lookup"><span data-stu-id="ec46f-222">>1 minute</span></span>|<span data-ttu-id="ec46f-223">Os usuários certamente farão algo mais.</span><span class="sxs-lookup"><span data-stu-id="ec46f-223">Users will certainly do something else.</span></span>|

- <span data-ttu-id="ec46f-224">Monitore a integridade do serviço e use a telemetria para monitorar o sucesso do usuário.</span><span class="sxs-lookup"><span data-stu-id="ec46f-224">Monitor your service health, and use telemetry to monitor user success.</span></span>

- <span data-ttu-id="ec46f-225">Minimize as trocas de dados entre o add-in e o Office documento.</span><span class="sxs-lookup"><span data-stu-id="ec46f-225">Minimize data exchanges between the add-in and the Office document.</span></span> <span data-ttu-id="ec46f-226">Para obter mais informações, [consulte Evite usar o método context.sync em loops](correlated-objects-pattern.md).</span><span class="sxs-lookup"><span data-stu-id="ec46f-226">For more information, see [Avoid using the context.sync method in loops](correlated-objects-pattern.md).</span></span>

## <a name="market-your-add-in"></a><span data-ttu-id="ec46f-227">Comercializar seu suplemento</span><span class="sxs-lookup"><span data-stu-id="ec46f-227">Market your add-in</span></span>

- <span data-ttu-id="ec46f-p131">Publique seu suplemento no [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) e [promova-o](/office/dev/store/promote-your-office-store-solution) pelo seu site. Crie uma [listagem eficaz do AppSource](/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="ec46f-p131">Publish your add-in to [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) and [promote it](/office/dev/store/promote-your-office-store-solution) from your website. Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings).</span></span>

- <span data-ttu-id="ec46f-p132">Use títulos sucintos e descritivos para o suplemento. Inclua no máximo 128 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p132">Use succinct and descriptive add-in titles. Include no more than 128 characters.</span></span>

- <span data-ttu-id="ec46f-p133">Escreva descrições curtas e atraentes para o seu suplemento. Responda a pergunta "Qual problema este suplemento resolve?".</span><span class="sxs-lookup"><span data-stu-id="ec46f-p133">Write short, compelling descriptions of your add-in. Answer the question "What problem does this add-in solve?".</span></span>

- <span data-ttu-id="ec46f-p134">Transmita a proposta de valor do seu suplemento em seu título e descrição. Não confie apenas em sua marca.</span><span class="sxs-lookup"><span data-stu-id="ec46f-p134">Convey the value proposition of your add-in in your title and description. Don't rely on your brand.</span></span>

- <span data-ttu-id="ec46f-236">Crie um site para ajudar os usuários a encontrar e utilizar seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="ec46f-236">Create a website to help users find and use your add-in.</span></span>

## <a name="use-javascript-that-supports-internet-explorer"></a><span data-ttu-id="ec46f-237">Usar JavaScript compatível com o Internet Explorer</span><span class="sxs-lookup"><span data-stu-id="ec46f-237">Use JavaScript that supports Internet Explorer</span></span>

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="see-also"></a><span data-ttu-id="ec46f-238">Confira também</span><span class="sxs-lookup"><span data-stu-id="ec46f-238">See also</span></span>

- [<span data-ttu-id="ec46f-239">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ec46f-239">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="ec46f-240">Saiba mais sobre o Programa para Desenvolvedores do Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="ec46f-240">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
