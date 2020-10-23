---
title: Práticas recomendadas para o desenvolvimento de suplementos do Office
description: Aplique as práticas recomendadas ao desenvolver para criar suplementos do Office.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 8ce0482e108e7b8774442a2b0669a0e76bb401f9
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740858"
---
# <a name="best-practices-for-developing-office-add-ins"></a><span data-ttu-id="1e813-103">Práticas recomendadas para o desenvolvimento de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="1e813-103">Best practices for developing Office Add-ins</span></span>

<span data-ttu-id="1e813-p101">Os suplementos eficazes oferecem uma funcionalidade exclusiva e fascinante que estende os aplicativos do Office de uma maneira visualmente atraente. Para criar um excelente suplemento, ofereça uma primeira experiência envolvente para seus usuários, desenvolva uma experiência de interface de usuário de alto nível e otimize o desempenho do seu suplemento. Aplique as práticas recomendadas descritas neste artigo para criar suplementos que ajudem os usuários a concluir suas tarefas de forma rápida e eficiente.</span><span class="sxs-lookup"><span data-stu-id="1e813-p101">Effective add-ins offer unique and compelling functionality that extends Office applications in a visually appealing way. To create a great add-in, provide an engaging first-time experience for your users, design a first-class UI experience, and optimize your add-in's performance. Apply the best practices described in this article to create add-ins that help your users complete their tasks quickly and efficiently.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="provide-clear-value"></a><span data-ttu-id="1e813-107">Fornecer um valor claro</span><span class="sxs-lookup"><span data-stu-id="1e813-107">Provide clear value</span></span>

- <span data-ttu-id="1e813-p102">Crie suplementos que ajudem os usuários a concluir tarefas de forma rápida e eficiente. Concentre-se nos cenários que fazem sentido para aplicativos do Office. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="1e813-p102">Create add-ins that help users complete tasks quickly and efficiently. Focus on scenarios that make sense for Office applications. For example:</span></span>
 - <span data-ttu-id="1e813-111">Torne as principais tarefas de criação mais rápidas e fáceis, com menos interrupções.</span><span class="sxs-lookup"><span data-stu-id="1e813-111">Make core authoring tasks faster and easier, with fewer interruptions.</span></span>
 - <span data-ttu-id="1e813-112">Habilite novos cenários no Office.</span><span class="sxs-lookup"><span data-stu-id="1e813-112">Enable new scenarios within Office.</span></span>
 - <span data-ttu-id="1e813-113">Inserir serviços complementares nos aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="1e813-113">Embed complementary services within Office applications.</span></span>
 - <span data-ttu-id="1e813-114">Melhore a experiência do Office para aumentar a produtividade.</span><span class="sxs-lookup"><span data-stu-id="1e813-114">Improve the Office experience to enhance productivity.</span></span>
- <span data-ttu-id="1e813-115">Certifique-se de que o valor do seu suplemento seja claro para os usuários desde o princípio, [criando uma experiência envolvente na primeira execução](#create-an-engaging-first-run-experience).</span><span class="sxs-lookup"><span data-stu-id="1e813-115">Make sure that the value of your add-in is clear to users right away by [creating an engaging first run experience](#create-an-engaging-first-run-experience).</span></span>
- <span data-ttu-id="1e813-p103">Crie uma [listagem eficaz do AppSource](/office/dev/store/create-effective-office-store-listings). Deixe claro quais são os benefícios do seu suplemento no título e na descrição. Não dependa da sua marca para dizer o que seu suplemento faz.</span><span class="sxs-lookup"><span data-stu-id="1e813-p103">Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings). Make the benefits of your add-in clear in your title and description. Don't rely on your brand to communicate what your add-in does.</span></span>


## <a name="create-an-engaging-first-run-experience"></a><span data-ttu-id="1e813-119">Criar uma experiência envolvente na primeira execução</span><span class="sxs-lookup"><span data-stu-id="1e813-119">Create an engaging first-run experience</span></span>

- <span data-ttu-id="1e813-p104">Envolva os novos usuários com uma primeira experiência altamente útil e intuitiva. Observe que, mesmo depois de baixar o suplemento da loja, os usuários ainda estão decidindo se vão utilizá-lo.</span><span class="sxs-lookup"><span data-stu-id="1e813-p104">Engage new users with a highly usable and intuitive first experience. Note that users are still deciding whether to use or abandon an add-in after they download it from the store.</span></span>

- <span data-ttu-id="1e813-p105">Deixe claro quais são as etapas que usuário terá que seguir para se envolver com seu suplemento. Use vídeos, diagramas, painéis de paginação ou outros recursos para atrair usuários.</span><span class="sxs-lookup"><span data-stu-id="1e813-p105">Make the steps that the user needs to take to engage with your add-in clear. Use videos, placemats, paging panels, or other resources to entice users.</span></span>

- <span data-ttu-id="1e813-124">Reforce a proposta de valor do seu suplemento no início, em vez de apenas pedir que seus usuários entrem.</span><span class="sxs-lookup"><span data-stu-id="1e813-124">Reinforce the value proposition of your add-in on launch, rather than just asking users to sign in.</span></span>

- <span data-ttu-id="1e813-125">Forneça uma interface do usuário informativa e torne sua interface do usuário pessoal.</span><span class="sxs-lookup"><span data-stu-id="1e813-125">Provide teaching UI to guide users and make your UI personal.</span></span>

   ![Uma captura de tela que mostra um painel de tarefas de suplemento com etapas de introdução ao lado de um suplemento sem etapas de introdução](../images/contoso-part-catalog-do-dont.png)

- <span data-ttu-id="1e813-127">Se seu suplemento de conteúdo estiver vinculado a dados no documento do usuário, inclua exemplos de dados ou um modelo para mostrar aos usuários o formato de dados a ser usado.</span><span class="sxs-lookup"><span data-stu-id="1e813-127">If your content add-in binds to data in the user's document, include sample data or a template to show users the data format to use.</span></span>

   ![Uma captura de tela que mostra um suplemento de conteúdo com dados ao lado de um suplemento de conteúdo sem dados](../images/add-in-title.png)

- <span data-ttu-id="1e813-p106">Ofereça [avaliações gratuitas](/office/dev/store/decide-on-a-pricing-model). Caso o suplemento exija uma assinatura, disponibilize algumas funcionalidades sem a necessidade da assinatura.</span><span class="sxs-lookup"><span data-stu-id="1e813-p106">Offer [free trials](/office/dev/store/decide-on-a-pricing-model). If your add-in requires a subscription, make some functionality available without a subscription.</span></span>

- <span data-ttu-id="1e813-p107">Simplifique o processo de inscrição. Preencha automaticamente as informações (email, nome de exibição) e ignore as verificações de email.</span><span class="sxs-lookup"><span data-stu-id="1e813-p107">Make signup simple. Prefill information (email, display name) and skip email verifications.</span></span>

- <span data-ttu-id="1e813-p108">Evite os pop-ups. Se você tiver de usá-los, oriente o usuário sobre como habilitar o seu pop-up.</span><span class="sxs-lookup"><span data-stu-id="1e813-p108">Avoid pop ups. If you have to use them, guide the user to enable your pop up.</span></span>

<span data-ttu-id="1e813-135">Para padrões que podem ser aplicados ao desenvolver sua experiência de primeira execução, consulte [Padrões de design da experiência do usuário para suplementos do Office](../design/first-run-experience-patterns.md).</span><span class="sxs-lookup"><span data-stu-id="1e813-135">For patterns that you can apply as you develop your first-run experience, see [UX design patterns for Office Add-ins](../design/first-run-experience-patterns.md).</span></span>

## <a name="use-add-in-commands"></a><span data-ttu-id="1e813-136">Usar comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="1e813-136">Use add-in commands</span></span>

- <span data-ttu-id="1e813-p109">Fornece ao suplemento pontos de entrada relevantes da interface do usuário usando os comandos do suplemento. Confira mais detalhes, inclusive as práticas recomendadas de design, nos [comandos de suplemento](../design/add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="1e813-p109">Provide relevant UI entry points for your add-in by using add-in commands. For details, including design best practices, see [add-in commands](../design/add-in-commands.md).</span></span>

## <a name="apply-ux-design-principles"></a><span data-ttu-id="1e813-139">Aplicar os princípios de design de UX</span><span class="sxs-lookup"><span data-stu-id="1e813-139">Apply UX design principles</span></span>

- <span data-ttu-id="1e813-p110">Assegure-se de que a aparência e a funcionalidade de seus suplementos complementam a experiência do Office. Use o [Office UI Fabric](https://developer.microsoft.com/fabric).</span><span class="sxs-lookup"><span data-stu-id="1e813-p110">Ensure that the look and feel and functionality of your add-in complements the Office experience. Use [Office UI Fabric](https://developer.microsoft.com/fabric).</span></span>

- <span data-ttu-id="1e813-p111">Favoreça o conteúdo através do Chrome. Evite elementos de interface do usuário supérfluos que não agregam valor à experiência do usuário.</span><span class="sxs-lookup"><span data-stu-id="1e813-p111">Favor content over chrome. Avoid superfluous UI elements that don't add value to the user experience.</span></span>

- <span data-ttu-id="1e813-p112">Mantenha os usuários no controle. Verifique se os usuários compreenderam as decisões importantes e podem reverter facilmente as ações realizadas pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="1e813-p112">Keep users in control. Ensure that users understand important decisions, and can easily reverse actions the add-in performs.</span></span>

- <span data-ttu-id="1e813-p113">Use uma identidade visual para inspirar confiança e orientar os usuários. Não use o recurso de identidade visual para sobrecarregar ou enviar anúncios aos usuários.</span><span class="sxs-lookup"><span data-stu-id="1e813-p113">Use branding to inspire trust and orient users. Do not use branding to overwhelm or advertise to users.</span></span>

- <span data-ttu-id="1e813-p114">Evite a necessidade de rolagem. Otimize para a resolução 1366 x 768.</span><span class="sxs-lookup"><span data-stu-id="1e813-p114">Avoid scrolling. Optimize for 1366 x 768 resolution.</span></span>

- <span data-ttu-id="1e813-150">Não inclua imagens não licenciadas.</span><span class="sxs-lookup"><span data-stu-id="1e813-150">Do not include unlicensed images.</span></span>

- <span data-ttu-id="1e813-151">Use uma [linguagem clara e simples](../design/voice-guidelines.md) no seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="1e813-151">Use [clear and simple language](../design/voice-guidelines.md) in your add-in.</span></span>

- <span data-ttu-id="1e813-152">Preocupe-se com a acessibilidade: facilite a interação dos usuários com o seu suplemento e inclua tecnologias adaptativas, como leitores de tela.</span><span class="sxs-lookup"><span data-stu-id="1e813-152">Account for accessibility - make your add-in easy for all users to interact with, and accommodate assistive technologies such as screen readers.</span></span>

- <span data-ttu-id="1e813-p115">Desenvolva para todas as plataformas e métodos de entrada, incluindo teclado/mouse e [toque](#optimize-for-touch). Certifique-se de que sua interface do usuário responda a diferentes fatores forma.</span><span class="sxs-lookup"><span data-stu-id="1e813-p115">Design for all platforms and input methods, including mouse/keyboard and [touch](#optimize-for-touch). Ensure that your UI is responsive to different form factors.</span></span>

### <a name="optimize-for-touch"></a><span data-ttu-id="1e813-155">Otimizar para toque</span><span class="sxs-lookup"><span data-stu-id="1e813-155">Optimize for touch</span></span>

- <span data-ttu-id="1e813-156">Use a propriedade [Context. touchEnabled](/javascript/api/office/office.context#touchenabled) para detectar se o aplicativo do Office no qual o suplemento é executado está habilitado para toque.</span><span class="sxs-lookup"><span data-stu-id="1e813-156">Use the [Context.touchEnabled](/javascript/api/office/office.context#touchenabled) property to detect whether the Office application that your add-in runs on is touch enabled.</span></span>

  > [!NOTE]
  > <span data-ttu-id="1e813-157">Essa propriedade não tem suporte no Outlook.</span><span class="sxs-lookup"><span data-stu-id="1e813-157">This property is not supported in Outlook.</span></span>

- <span data-ttu-id="1e813-p116">Verifique se todos os controles são dimensionados adequadamente para interação por toque. Por exemplo, se os botões têm destinos de toque adequados e se as caixas de entrada têm a dimensão correta para que os usuários insiram entradas.</span><span class="sxs-lookup"><span data-stu-id="1e813-p116">Ensure that all controls are appropriately sized for touch interaction. For example, buttons have adequate touch targets, and input boxes are large enough for users to enter input.</span></span>

- <span data-ttu-id="1e813-160">Não confie nos métodos de entrada sem toque, como passar o cursor ou clicar com o botão direito do mouse.</span><span class="sxs-lookup"><span data-stu-id="1e813-160">Do not rely on non-touch input methods like hover or right-click.</span></span>

- <span data-ttu-id="1e813-p117">Verifique se o suplemento funciona nos modos retrato e paisagem. Observe que em dispositivos de toque, parte do suplemento pode ficar oculta pelo teclado virtual.</span><span class="sxs-lookup"><span data-stu-id="1e813-p117">Ensure that your add-in works in both portrait and landscape modes. Be aware that on touch devices, part of your add-in might be hidden by the soft keyboard.</span></span>

- <span data-ttu-id="1e813-163">Teste seu suplemento em um dispositivo real usando o [sideload](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="1e813-163">Test your add-in on a real device by using [sideloading](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).</span></span>

> [!NOTE]
> <span data-ttu-id="1e813-164">Se você está usando o [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) nos seus elementos de design, muitos desses elementos já foram tratados.</span><span class="sxs-lookup"><span data-stu-id="1e813-164">If you're using [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) for your design elements, many of these elements are taken care of.</span></span>


## <a name="optimize-and-monitor-add-in-performance"></a><span data-ttu-id="1e813-165">Otimizar e monitorar o desempenho do suplemento</span><span class="sxs-lookup"><span data-stu-id="1e813-165">Optimize and monitor add-in performance</span></span>

- <span data-ttu-id="1e813-p118">Crie a percepção de respostas rápidas da interface do usuário. Seu suplemento deverá ser carregado em 500 ms ou menos.</span><span class="sxs-lookup"><span data-stu-id="1e813-p118">Create the perception of fast UI responses. Your add-in should load in 500 ms or less.</span></span>

- <span data-ttu-id="1e813-168">Certifique-se de que todas as interações do usuário respondam em menos de um segundo.</span><span class="sxs-lookup"><span data-stu-id="1e813-168">Ensure that all user interactions respond in under one second.</span></span>

-  <span data-ttu-id="1e813-169">Forneça indicadores de carregamento para operações com longa execução.</span><span class="sxs-lookup"><span data-stu-id="1e813-169">Provide loading indicators for long-running operations.</span></span>

- <span data-ttu-id="1e813-p119">Use uma CDN para hospedar imagens, recursos e bibliotecas comuns. Carregue o máximo possível de um só lugar.</span><span class="sxs-lookup"><span data-stu-id="1e813-p119">Use a CDN to host images, resources, and common libraries. Load as much as you can from one place.</span></span>

- <span data-ttu-id="1e813-p120">Siga as práticas da Web padrão para otimizar a página. Use apenas versões reduzidas das bibliotecas na produção. Carregue somente os recursos que você precisar e otimize como os recursos são carregados.</span><span class="sxs-lookup"><span data-stu-id="1e813-p120">Follow standard web practices to optimize your web page. In production, use only minified versions of libraries. Only load resources that you need, and optimize how resources are loaded.</span></span>

- <span data-ttu-id="1e813-p121">Se o tempo de execução das operações demorar, forneça feedback aos usuários. Observe os limites relacionados na tabela a seguir. Saiba mais em [Limites de recurso e otimização de desempenho para Suplementos do Office](../concepts/resource-limits-and-performance-optimization.md).</span><span class="sxs-lookup"><span data-stu-id="1e813-p121">If operations take time to execute, provide feedback to users. Note the thresholds listed in the following table. For additional information, see [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md).</span></span>

  |<span data-ttu-id="1e813-178">**Classe de interação**</span><span class="sxs-lookup"><span data-stu-id="1e813-178">**Interaction class**</span></span>|<span data-ttu-id="1e813-179">**Destino**</span><span class="sxs-lookup"><span data-stu-id="1e813-179">**Target**</span></span>|<span data-ttu-id="1e813-180">**Limite superior**</span><span class="sxs-lookup"><span data-stu-id="1e813-180">**Upper bound**</span></span>|<span data-ttu-id="1e813-181">**Percepção humana**</span><span class="sxs-lookup"><span data-stu-id="1e813-181">**Human perception**</span></span>|
  |:-----|:-----|:-----|:-----|
  |<span data-ttu-id="1e813-182">Instantâneo</span><span class="sxs-lookup"><span data-stu-id="1e813-182">Instant</span></span>|<span data-ttu-id="1e813-183"><=50 ms</span><span class="sxs-lookup"><span data-stu-id="1e813-183"><=50 ms</span></span>|<span data-ttu-id="1e813-184">100 ms</span><span class="sxs-lookup"><span data-stu-id="1e813-184">100 ms</span></span>|<span data-ttu-id="1e813-185">Nenhum atraso considerável.</span><span class="sxs-lookup"><span data-stu-id="1e813-185">No noticeable delay.</span></span>|
  |<span data-ttu-id="1e813-186">Rápida</span><span class="sxs-lookup"><span data-stu-id="1e813-186">Fast</span></span>|<span data-ttu-id="1e813-187">50 – 100 ms.</span><span class="sxs-lookup"><span data-stu-id="1e813-187">50-100 ms</span></span>|<span data-ttu-id="1e813-188">200 ms</span><span class="sxs-lookup"><span data-stu-id="1e813-188">200 ms</span></span>|<span data-ttu-id="1e813-p122">Atraso mínimo considerável. Não são necessários comentários.</span><span class="sxs-lookup"><span data-stu-id="1e813-p122">Minimally noticeable delay. No feedback necessary.</span></span>|
  |<span data-ttu-id="1e813-191">Típico</span><span class="sxs-lookup"><span data-stu-id="1e813-191">Typical</span></span>|<span data-ttu-id="1e813-192">100 – 300 ms</span><span class="sxs-lookup"><span data-stu-id="1e813-192">100-300 ms</span></span>|<span data-ttu-id="1e813-193">500 ms</span><span class="sxs-lookup"><span data-stu-id="1e813-193">500 ms</span></span>|<span data-ttu-id="1e813-p123">Rápido, mas não o suficiente para ser descrito como rápido. Não são necessários comentários.</span><span class="sxs-lookup"><span data-stu-id="1e813-p123">Quick, but too slow to be described as fast. No feedback necessary.</span></span>|
  |<span data-ttu-id="1e813-196">Dinâmico</span><span class="sxs-lookup"><span data-stu-id="1e813-196">Responsive</span></span>|<span data-ttu-id="1e813-197">300 – 500 ms.</span><span class="sxs-lookup"><span data-stu-id="1e813-197">300-500 ms</span></span>|<span data-ttu-id="1e813-198">1 segundo</span><span class="sxs-lookup"><span data-stu-id="1e813-198">1 second</span></span>|<span data-ttu-id="1e813-p124">Não muito rápido, embora pareça ser dinâmico. Não são necessários comentários.</span><span class="sxs-lookup"><span data-stu-id="1e813-p124">Not fast, but still feels responsive. No feedback necessary.</span></span>|
  |<span data-ttu-id="1e813-201">Contínuo</span><span class="sxs-lookup"><span data-stu-id="1e813-201">Continuous</span></span>|<span data-ttu-id="1e813-202">>500 ms</span><span class="sxs-lookup"><span data-stu-id="1e813-202">>500 ms</span></span>|<span data-ttu-id="1e813-203">5 segundos</span><span class="sxs-lookup"><span data-stu-id="1e813-203">5 seconds</span></span>|<span data-ttu-id="1e813-p125">Tempo de espera médio, já não parece ser dinâmico. Podem ser necessários comentários.</span><span class="sxs-lookup"><span data-stu-id="1e813-p125">Medium wait, no longer feels responsive. Might need feedback.</span></span>|
  |<span data-ttu-id="1e813-206">Cativo</span><span class="sxs-lookup"><span data-stu-id="1e813-206">Captive</span></span>|<span data-ttu-id="1e813-207">>500 ms</span><span class="sxs-lookup"><span data-stu-id="1e813-207">>500 ms</span></span>|<span data-ttu-id="1e813-208">10 segundos</span><span class="sxs-lookup"><span data-stu-id="1e813-208">10 seconds</span></span>|<span data-ttu-id="1e813-p126">Longo, mas não o suficiente para fazer executar outra ação. Podem ser necessários comentários.</span><span class="sxs-lookup"><span data-stu-id="1e813-p126">Long, but not long enough to do something else. Might need feedback.</span></span>|
  |<span data-ttu-id="1e813-211">Estendida</span><span class="sxs-lookup"><span data-stu-id="1e813-211">Extended</span></span>|<span data-ttu-id="1e813-212">>500 ms</span><span class="sxs-lookup"><span data-stu-id="1e813-212">>500 ms</span></span>|<span data-ttu-id="1e813-213">>10 segundos</span><span class="sxs-lookup"><span data-stu-id="1e813-213">>10 seconds</span></span>|<span data-ttu-id="1e813-p127">Longo o suficiente para realizar outra ação durante o tempo de espera. Podem ser necessários comentários.</span><span class="sxs-lookup"><span data-stu-id="1e813-p127">Long enough to do something else while waiting. Might need feedback.</span></span>|
  |<span data-ttu-id="1e813-216">Longa execução</span><span class="sxs-lookup"><span data-stu-id="1e813-216">Long running</span></span>|<span data-ttu-id="1e813-217">> 5 segundos</span><span class="sxs-lookup"><span data-stu-id="1e813-217">>5 seconds</span></span>|<span data-ttu-id="1e813-218">> 1 minuto</span><span class="sxs-lookup"><span data-stu-id="1e813-218">>1 minute</span></span>|<span data-ttu-id="1e813-219">Os usuários certamente farão algo mais.</span><span class="sxs-lookup"><span data-stu-id="1e813-219">Users will certainly do something else.</span></span>|

- <span data-ttu-id="1e813-220">Monitore a integridade do serviço e use a telemetria para monitorar o sucesso do usuário.</span><span class="sxs-lookup"><span data-stu-id="1e813-220">Monitor your service health, and use telemetry to monitor user success.</span></span>

- <span data-ttu-id="1e813-221">Minimize as trocas de dados entre o suplemento e o documento do Office.</span><span class="sxs-lookup"><span data-stu-id="1e813-221">Minimize data exchanges between the add-in and the Office document.</span></span> <span data-ttu-id="1e813-222">Para obter mais informações, consulte [Evite usar o método Context. Sync em loops](correlated-objects-pattern.md).</span><span class="sxs-lookup"><span data-stu-id="1e813-222">For more information, see [Avoid using the context.sync method in loops](correlated-objects-pattern.md).</span></span>

## <a name="market-your-add-in"></a><span data-ttu-id="1e813-223">Comercializar seu suplemento</span><span class="sxs-lookup"><span data-stu-id="1e813-223">Market your add-in</span></span>

- <span data-ttu-id="1e813-p129">Publique seu suplemento no [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) e [promova-o](/office/dev/store/promote-your-office-store-solution) pelo seu site. Crie uma [listagem eficaz do AppSource](/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="1e813-p129">Publish your add-in to [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) and [promote it](/office/dev/store/promote-your-office-store-solution) from your website. Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings).</span></span>

- <span data-ttu-id="1e813-p130">Use títulos sucintos e descritivos para o suplemento. Inclua no máximo 128 caracteres.</span><span class="sxs-lookup"><span data-stu-id="1e813-p130">Use succinct and descriptive add-in titles. Include no more than 128 characters.</span></span>

- <span data-ttu-id="1e813-p131">Escreva descrições curtas e atraentes para o seu suplemento. Responda a pergunta "Qual problema este suplemento resolve?".</span><span class="sxs-lookup"><span data-stu-id="1e813-p131">Write short, compelling descriptions of your add-in. Answer the question "What problem does this add-in solve?".</span></span>

- <span data-ttu-id="1e813-p132">Transmita a proposta de valor do seu suplemento em seu título e descrição. Não confie apenas em sua marca.</span><span class="sxs-lookup"><span data-stu-id="1e813-p132">Convey the value proposition of your add-in in your title and description. Don't rely on your brand.</span></span>

- <span data-ttu-id="1e813-232">Crie um site para ajudar os usuários a encontrar e utilizar seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="1e813-232">Create a website to help users find and use your add-in.</span></span>

## <a name="use-javascript-that-supports-internet-explorer"></a><span data-ttu-id="1e813-233">Usar JavaScript que suporte o Internet Explorer</span><span class="sxs-lookup"><span data-stu-id="1e813-233">Use JavaScript that supports Internet Explorer</span></span>

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="see-also"></a><span data-ttu-id="1e813-234">Confira também</span><span class="sxs-lookup"><span data-stu-id="1e813-234">See also</span></span>

- [<span data-ttu-id="1e813-235">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="1e813-235">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="1e813-236">Saiba mais sobre o programa de desenvolvedor 365 da Microsoft</span><span class="sxs-lookup"><span data-stu-id="1e813-236">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
