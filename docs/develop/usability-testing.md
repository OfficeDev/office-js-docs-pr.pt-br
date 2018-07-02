---
title: Teste de usabilidade de Suplementos do Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 410b8d7ede22cf222ee2df794e438c7f5f8881dd
ms.sourcegitcommit: 4e4f7c095e8f33b06bd8a02534ee901125eb1d17
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/28/2018
ms.locfileid: "20085263"
---
# <a name="usability-testing-for-office-add-ins"></a><span data-ttu-id="55da4-102">Teste de usabilidade de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="55da4-102">Usability testing for Office Add-ins</span></span>

<span data-ttu-id="55da4-p101">Um excelente design de suplemento considera os comportamentos do usuário. Como seus próprios conceitos prévios influenciam suas decisões de design, é importante testar designs com usuários reais para garantir que seus suplementos funcionem bem para seus clientes.</span><span class="sxs-lookup"><span data-stu-id="55da4-p101">A great add-in design takes user behaviors into account. Because your own preconceptions influence your design decisions, it’s important to test designs with real users to make sure that your add-ins work well for your customers.</span></span> 

<span data-ttu-id="55da4-p102">É possível executar testes de usabilidade de maneiras diferentes. Para muitos desenvolvedores de suplementos, estudos de usabilidade remota não moderada são os que economizam mais tempo e dinheiro. Vários serviços de testes populares facilitam isso. Veja alguns exemplos:</span><span class="sxs-lookup"><span data-stu-id="55da4-p102">You can run usability tests in different ways. For many add-in developers, remote, unmoderated usability studies are the most time and cost effective. Several popular testing services make this easy; the following are some examples:</span></span> 

 - [<span data-ttu-id="55da4-108">UserTesting.com</span><span class="sxs-lookup"><span data-stu-id="55da4-108">UserTesting.com</span></span>](https://www.UserTesting.com)
 - [<span data-ttu-id="55da4-109">Optimalworkshop.com</span><span class="sxs-lookup"><span data-stu-id="55da4-109">Optimalworkshop.com</span></span>](https://www.Optimalworkshop.com)
 - [<span data-ttu-id="55da4-110">Userzoom.com</span><span class="sxs-lookup"><span data-stu-id="55da4-110">Userzoom.com</span></span>](https://www.Userzoom.com)

<span data-ttu-id="55da4-111">Esses serviços de teste o ajudam a simplificar a criação do plano de teste e remover a necessidade de buscar participantes ou moderar os testes.</span><span class="sxs-lookup"><span data-stu-id="55da4-111">These testing services help you to streamline test plan creation and remove the need to seek out participants or moderate the tests.</span></span> 

<span data-ttu-id="55da4-p103">Você precisa de apenas cinco participantes para descobrir a maioria dos problemas de usabilidade no seu design. Incorpore testes pequenos regularmente durante o ciclo de desenvolvimento para garantir que seu produto seja centralizado no usuário.</span><span class="sxs-lookup"><span data-stu-id="55da4-p103">You need only five participants to uncover most usability issues in your design. Incorporate small tests regularly throughout your development cycle to ensure that your product is user-centered.</span></span>

> [!NOTE]
> <span data-ttu-id="55da4-p104">Recomendamos que você teste a usabilidade do seu suplemento em várias plataformas. Para [publicar](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store) seu suplemento no AppSource, ele deve funcionar em todas as [plataformas compatíveis com os métodos que você definir](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="55da4-p104">We recommend that you test the usability of your add-in across multiple platforms. To [publish your add-in to AppSource](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store), it must work on all [platforms that support the methods that you define](../overview/office-add-in-availability.md).</span></span>

## <a name="1---sign-up-for-a-testing-service"></a><span data-ttu-id="55da4-116">1.   Inscreva-se em um serviço de teste</span><span class="sxs-lookup"><span data-stu-id="55da4-116">1.   Sign up for a testing service</span></span>

<span data-ttu-id="55da4-117">Saiba mais em [Seleção de uma ferramenta online para o teste de usuário remoto não moderado.](https://www.nngroup.com/articles/unmoderated-user-testing-tools/)</span><span class="sxs-lookup"><span data-stu-id="55da4-117">For more information, see [Selecting an Online Tool for Unmoderated Remote User Testing.](https://www.nngroup.com/articles/unmoderated-user-testing-tools/)</span></span>

## <a name="2-develop-your-research-questions"></a><span data-ttu-id="55da4-118">2. Desenvolva as perguntas da sua pesquisa</span><span class="sxs-lookup"><span data-stu-id="55da4-118">2. Develop your research questions</span></span>
 
<span data-ttu-id="55da4-p105">As perguntas da pesquisa definem os objetivos de sua pesquisa e guiam seu plano de teste. Suas perguntas o ajudarão a identificar os participantes para recrutar e as tarefas que eles executarão. Certifique-se de que suas perguntas de pesquisa sejam o mais específicas possível. Você também pode procurar responder perguntas mais amplas.</span><span class="sxs-lookup"><span data-stu-id="55da4-p105">Research questions define the objectives of your research and guide your test plan. Your questions will help you identify participants to recruit and the tasks they will perform. Make your research questions as specific as you can. You can also seek to answer broader questions.</span></span>
 
<span data-ttu-id="55da4-123">A seguir, alguns exemplos de perguntas de pesquisa:</span><span class="sxs-lookup"><span data-stu-id="55da4-123">The following are some examples of research questions:</span></span>
  
<span data-ttu-id="55da4-124">**Específicas**</span><span class="sxs-lookup"><span data-stu-id="55da4-124">**Specific**</span></span>  

 - <span data-ttu-id="55da4-125">Os usuários percebem o link "avaliação gratuita" na página inicial?</span><span class="sxs-lookup"><span data-stu-id="55da4-125">Do users notice the "free trial" link on the landing page?</span></span>
 - <span data-ttu-id="55da4-126">Quando os usuários inserem conteúdo do suplemento em seu documento eles entendem onde é inserido no documento?</span><span class="sxs-lookup"><span data-stu-id="55da4-126">When users insert content from the add-in to their document, do they understand where in the document it is inserted?</span></span>

<span data-ttu-id="55da4-127">**Amplas**</span><span class="sxs-lookup"><span data-stu-id="55da4-127">**Broad**</span></span>  

 - <span data-ttu-id="55da4-128">Quais são os pontos mais problemáticos para usuário em nosso suplemento?</span><span class="sxs-lookup"><span data-stu-id="55da4-128">What are the biggest pain points for the user in our add-in?</span></span>
 - <span data-ttu-id="55da4-129">Os usuários entendem o significado dos ícones na barra de comandos, antes de clicar neles?</span><span class="sxs-lookup"><span data-stu-id="55da4-129">Do users understand the meaning of the icons in our command bar, before they click on them?</span></span>
 - <span data-ttu-id="55da4-130">Os usuários localizam o menu configurações com facilidade?</span><span class="sxs-lookup"><span data-stu-id="55da4-130">Can users easily find the settings menu?</span></span>

<span data-ttu-id="55da4-p106">É importante obter dados de toda a jornada do usuário – da descoberta do suplemento à instalação e utilização dele. Considere perguntas de pesquisa que abordem os seguintes aspectos da experiência do usuário no suplemento:</span><span class="sxs-lookup"><span data-stu-id="55da4-p106">It’s important to get data on the entire user journey – from discovering your add-in, to installing and using it. Consider research questions that address the following aspects of the add-in user experience:</span></span>
 
 - <span data-ttu-id="55da4-133">Localização do suplemento na Loja</span><span class="sxs-lookup"><span data-stu-id="55da4-133">Finding your add-in in AppSource</span></span>
 - <span data-ttu-id="55da4-134">Escolha da instalação do suplemento</span><span class="sxs-lookup"><span data-stu-id="55da4-134">Choosing to install your add-in</span></span>
 - <span data-ttu-id="55da4-135">Experiência de primeira execução</span><span class="sxs-lookup"><span data-stu-id="55da4-135">First run experience</span></span>
 - <span data-ttu-id="55da4-136">Comandos da faixa de opções</span><span class="sxs-lookup"><span data-stu-id="55da4-136">Ribbon commands</span></span>
 - <span data-ttu-id="55da4-137">Interface do Usuário do Suplemento</span><span class="sxs-lookup"><span data-stu-id="55da4-137">Add-in UI</span></span>
 - <span data-ttu-id="55da4-138">Como o suplemento interage com o espaço do documento do aplicativo do Office</span><span class="sxs-lookup"><span data-stu-id="55da4-138">How the add-in interacts with the document space of the Office application</span></span>
 - <span data-ttu-id="55da4-139">Qual o nível de controle que o usuário tem nos fluxos de inserção de conteúdo</span><span class="sxs-lookup"><span data-stu-id="55da4-139">How much control the user has over any content insertion flows</span></span>

<span data-ttu-id="55da4-140">Para saber mais, veja [Escrever perguntas eficazes.](http://help.usertesting.com/customer/en/portal/articles/2077663-writing-effective-questions)</span><span class="sxs-lookup"><span data-stu-id="55da4-140">For more information, see [Writing Effective Questions.](http://help.usertesting.com/customer/en/portal/articles/2077663-writing-effective-questions)</span></span>
 
## <a name="3-identify-participants-to-target"></a><span data-ttu-id="55da4-141">3. Identifique os participantes que serão o alvo</span><span class="sxs-lookup"><span data-stu-id="55da4-141">3. Identify participants to target</span></span>
 
<span data-ttu-id="55da4-p107">O teste remoto de serviços pode oferecer a você o controle de várias características dos participantes do teste. Pense cuidadosamente sobre que tipos de usuários você deseja buscar. Nos seus estágios iniciais de coleta de dados, talvez seja melhor recrutar uma ampla variedade de participantes para identificar problemas de usabilidade mais óbvios. Posteriormente, você pode optar por grupos segmentados como usuários avançados do Office, ocupações específicas ou faixas etárias específicas.</span><span class="sxs-lookup"><span data-stu-id="55da4-p107">Remote testing services can give you control over many characteristics of your test participants. Think carefully about what kinds of users you want to target. In your early stages of data collection, it might be better to recruit a wide variety of participants to identify more obvious usability issues. Later, you might choose to target groups like advanced Office users, particular occupations, or specific age ranges.</span></span>
 
## <a name="4-create-the-participant-screener"></a><span data-ttu-id="55da4-146">4. Crie o verificador de participantes</span><span class="sxs-lookup"><span data-stu-id="55da4-146">4. Create the participant screener</span></span>
 
<span data-ttu-id="55da4-p108">O verificador é o conjunto de perguntas e requisitos que você apresentará aos participantes do teste em potencial para verificá-los para o teste. Tenha em mente que os participantes de serviços como UserTesting.com têm interesse financeiro em se qualificar para seu teste. É uma boa ideia incluir perguntas difíceis em sua verificação se desejar excluir determinados usuários do teste.</span><span class="sxs-lookup"><span data-stu-id="55da4-p108">The screener is the set of questions and requirements you will present to prospective test participants to screen them for your test. Keep in mind that participants for services like UserTesting.com have a financial interest in qualifying for your test. It's a good idea to include trick questions in your screener if you want to  exclude certain users from the test.</span></span> 
 
<span data-ttu-id="55da4-150">Por exemplo, se deseja encontrar participantes que estão familiarizados com o GitHub, para filtrar os usuários que possam se mostrar incorretamente, inclua respostas falsas na lista de possíveis respostas.</span><span class="sxs-lookup"><span data-stu-id="55da4-150">For example, if you want to find participants who are familiar with GitHub, to filter out users who might misrepresent themselves, include fakes in the list of possible answers.</span></span>

<span data-ttu-id="55da4-151">**Com quais dos seguintes repositórios de código fonte você tem familiaridade?**</span><span class="sxs-lookup"><span data-stu-id="55da4-151">**Which of the following source code repositories are you familiar with?**</span></span>  
 <span data-ttu-id="55da4-p109">a. SourceShelf [*Rejeitar*]</span><span class="sxs-lookup"><span data-stu-id="55da4-p109">a. SourceShelf  [*Reject*]</span></span>  
 <span data-ttu-id="55da4-p110">b. CodeContainer [*Rejeitar*]</span><span class="sxs-lookup"><span data-stu-id="55da4-p110">b. CodeContainer  [*Reject*]</span></span>  
 <span data-ttu-id="55da4-p111">c. GitHub [*Deve selecionar*]</span><span class="sxs-lookup"><span data-stu-id="55da4-p111">c. GitHub  [*Must select*]</span></span>  
 <span data-ttu-id="55da4-p112">d. BitBucket [*Pode selecionar*]</span><span class="sxs-lookup"><span data-stu-id="55da4-p112">d. BitBucket  [*May select*]</span></span>  
 <span data-ttu-id="55da4-p113">e. CloudForge [*Pode selecionar*]</span><span class="sxs-lookup"><span data-stu-id="55da4-p113">e. CloudForge  [*May select*]</span></span>  

<span data-ttu-id="55da4-162">Se estiver planejando testar uma compilação em funcionamento do suplemento, as perguntas a seguir podem verificar os usuários que conseguirão fazer isso.</span><span class="sxs-lookup"><span data-stu-id="55da4-162">If you are planning to test a live build of your add-in, the following questions can screen for users who will be able to do this.</span></span> 

<span data-ttu-id="55da4-163">**Este teste requer que você tenha o Microsoft PowerPoint 2016. Você tem o PowerPoint 2016?**</span><span class="sxs-lookup"><span data-stu-id="55da4-163">**This test requires you to have Microsoft PowerPoint 2016. Do you have PowerPoint 2016?**</span></span>  
 <span data-ttu-id="55da4-p114">a. Sim [*Deve selecionar*]</span><span class="sxs-lookup"><span data-stu-id="55da4-p114">a. Yes [*Must select*]</span></span>  
 <span data-ttu-id="55da4-p115">b. Não [*Rejeitar*]</span><span class="sxs-lookup"><span data-stu-id="55da4-p115">b. No [*Reject*]</span></span>  
 <span data-ttu-id="55da4-p116">c. Não sei [*Rejeitar*]</span><span class="sxs-lookup"><span data-stu-id="55da4-p116">c. I don’t know [*Reject*]</span></span>  

<span data-ttu-id="55da4-170">**Este teste requer que você instale um suplemento gratuito para PowerPoint 2016 e crie uma conta gratuita para usá-lo. Você está disposto a instalar um suplemento e criar uma conta gratuita?**</span><span class="sxs-lookup"><span data-stu-id="55da4-170">**This test requires you to install a free add-in for PowerPoint 2016, and create a free account to use it. Are you willing to install an add-in and create a free account?**</span></span>  
 <span data-ttu-id="55da4-p117">a. Sim [*Deve selecionar*]</span><span class="sxs-lookup"><span data-stu-id="55da4-p117">a. Yes [*Must select*]</span></span>  
 <span data-ttu-id="55da4-p118">b. Não [*Rejeitar*]</span><span class="sxs-lookup"><span data-stu-id="55da4-p118">b. No [*Reject*]</span></span>  

<span data-ttu-id="55da4-175">Para saber mais, veja [Práticas recomendadas do verificador de perguntas.](http://help.usertesting.com/customer/en/portal/articles/2077835-screener-question-best-practices)</span><span class="sxs-lookup"><span data-stu-id="55da4-175">For more information, see [Screener Questions Best Practices.](http://help.usertesting.com/customer/en/portal/articles/2077835-screener-question-best-practices)</span></span>
 
## <a name="5-create-tasks-and-questions-for-participants"></a><span data-ttu-id="55da4-176">5. Crie tarefas e perguntas para os participantes</span><span class="sxs-lookup"><span data-stu-id="55da4-176">5. Create tasks and questions for participants</span></span>
 
<span data-ttu-id="55da4-p119">Tente priorizar o que você quer testar para que seja possível limitar o número de tarefas e perguntas do participante. Alguns serviços pagam os participantes apenas para um determinado período para que você certifique-se de não excedê-lo.</span><span class="sxs-lookup"><span data-stu-id="55da4-p119">Try to prioritize what you want tested so that you can limit the number of tasks and questions for the participant. Some services pay participants only for a set amount of time, so you want to make sure not to go over.</span></span>

<span data-ttu-id="55da4-p120">Tente observar como os participantes se comportam em vez de perguntar sobre eles sempre que possível. Se você precisar perguntar sobre comportamentos, pergunte o que os participantes fizeram no passado, em vez do que o que eles esperariam fazer em uma situação. Isso tende a fornecer resultados mais confiáveis.</span><span class="sxs-lookup"><span data-stu-id="55da4-p120">Try to observe participant behaviors instead of asking about them, whenever possible. If you need to ask about behaviors, ask about what participants have done in the past, rather than what they would expect to do in a situation. This tends to give more reliable results.</span></span>
 
<span data-ttu-id="55da4-p121">O principal desafio no teste não moderado é garantir que seus participantes entendam suas tarefas e cenários. Suas orientações devem ser *claras e concisas*. Inevitavelmente, se houver potencial para confusão, alguém ficará confuso.</span><span class="sxs-lookup"><span data-stu-id="55da4-p121">The main challenge in unmoderated testing is making sure your participants understand your tasks and scenarios. Your directions should be *clear and concise*. Inevitably, if there is potential for confusion, someone will be confused.</span></span> 

<span data-ttu-id="55da4-p122">Não pense que o usuário estará na tela que deve estar em um determinado momento durante o teste. Considere informar a tela em que eles precisam estar para iniciar a próxima tarefa.</span><span class="sxs-lookup"><span data-stu-id="55da4-p122">Don't assume that your user will be on the screen they’re supposed to be on at any given point during the test. Consider telling them what screen they need to be on to start the next task.</span></span> 

<span data-ttu-id="55da4-187">Saiba mais em [Como escrever tarefas excelentes.](http://help.usertesting.com/customer/en/portal/articles/2077824-writing-great-tasks)</span><span class="sxs-lookup"><span data-stu-id="55da4-187">For more information, see [Writing Great Tasks.](http://help.usertesting.com/customer/en/portal/articles/2077824-writing-great-tasks)</span></span>

## <a name="6-create-a-prototype-to-match-the-tasks-and-questions"></a><span data-ttu-id="55da4-188">6. Crie um protótipo para corresponder às tarefas e perguntas</span><span class="sxs-lookup"><span data-stu-id="55da4-188">6. Create a prototype to match the tasks and questions</span></span>
 
<span data-ttu-id="55da4-p123">Você também pode testar o suplemento em funcionamento ou pode testar um protótipo. Tenha em mente que, se você deseja testar o suplemento em funcionamento, é necessário buscar participantes que tenham o Office 2016, estejam dispostos a instalar o suplemento e criar uma conta (a menos que você tenha as credenciais de logon para fornecer). Você precisará certificar-se de que o suplemento será instalado com êxito.</span><span class="sxs-lookup"><span data-stu-id="55da4-p123">You can either test your live add-in, or you can test a prototype. Keep in mind that if you want to test the live add-in, you need to screen for participants that have Office 2016, are willing to install the add-in, and are willing to sign up for an account (unless you have logon credentials to provide them.) You'll then need to make sure that they successfully install your add-in.</span></span> 

<span data-ttu-id="55da4-p124">Em média, são necessários cerca de cinco minutos para orientar os usuários sobre como instalar um suplemento. A seguir, um exemplo de etapas de instalação claras e concisas. Ajuste as etapas com base nas condições específicas do teste.</span><span class="sxs-lookup"><span data-stu-id="55da4-p124">On average, it takes about 5 minutes to walk users through how to install an add-in. The following is an example of clear, concise installation steps. Adjust the steps based on the specifics of your test.</span></span>

<span data-ttu-id="55da4-194">**Instale o suplemento (insira o nome do suplemento aqui) para o PowerPoint 2016 usando as seguintes instruções:**</span><span class="sxs-lookup"><span data-stu-id="55da4-194">**Please install the (insert your add-in name here) add-in for PowerPoint 2016, using the following instructions:**</span></span> 

1. <span data-ttu-id="55da4-195">Abra o Microsoft PowerPoint 2016.</span><span class="sxs-lookup"><span data-stu-id="55da4-195">Open Microsoft PowerPoint 2016.</span></span>
2. <span data-ttu-id="55da4-196">Selecione **Apresentação em Branco.**</span><span class="sxs-lookup"><span data-stu-id="55da4-196">Select **Blank Presentation.**</span></span>
3. <span data-ttu-id="55da4-197">Vá para **Inserir > Meus Suplementos.**</span><span class="sxs-lookup"><span data-stu-id="55da4-197">Go to **Insert > My Add-ins.**</span></span>
5. <span data-ttu-id="55da4-198">Na janela pop-up, escolha **Loja.**</span><span class="sxs-lookup"><span data-stu-id="55da4-198">In the popup window, choose **Store.**</span></span>
6. <span data-ttu-id="55da4-199">Digite (Nome do suplemento) na caixa de pesquisa.</span><span class="sxs-lookup"><span data-stu-id="55da4-199">Type (Add-in name) in the search box.</span></span>
7. <span data-ttu-id="55da4-200">Escolha (Nome do suplemento).</span><span class="sxs-lookup"><span data-stu-id="55da4-200">Choose (Add-in name).</span></span>
8. <span data-ttu-id="55da4-201">Tire um momento para observar a página da Loja de forma a se familiarizar com o suplemento.</span><span class="sxs-lookup"><span data-stu-id="55da4-201">Take a moment to look at the Store page to familiarize yourself with the add-in.</span></span>
9. <span data-ttu-id="55da4-202">Escolha **Adicionar** para instalar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="55da4-202">Choose **Add** to install the add-in.</span></span>

<span data-ttu-id="55da4-p125">Você pode testar um protótipo em qualquer nível de interação e fidelidade visual. Para vinculação e interatividade mais complexas, considere uma ferramenta de criação de protótipo como a [InVision](https://www.invisionapp.com). Se você deseja testar telas estáticas, é possível hospedar imagens online e enviar a URL correspondente para os participantes ou fornecer um link para uma apresentação online do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="55da4-p125">You can test a prototype at any level of interaction and visual fidelity. For more complex linking and interactivity, consider a prototyping tool like [InVision](https://www.invisionapp.com). If you just want to test static screens, you can host images online and send participants the corresponding URL, or give them a link to an online PowerPoint presentation.</span></span> 

## <a name="7-run-a-pilot-test"></a><span data-ttu-id="55da4-206">7. Execute um teste piloto</span><span class="sxs-lookup"><span data-stu-id="55da4-206">7. Run a pilot test</span></span>

<span data-ttu-id="55da4-p126">Pode ser difícil acertar no protótipo e na lista de tarefas/perguntas. Os usuários podem ficar confusos com as tarefas ou podem se perder em seu protótipo. Você deve fazer um teste piloto 1 a 3 usuários para trabalhar corrigir os inevitáveis problemas com o formato do teste. Isso ajudará a garantir que suas perguntas sejam claras, que o protótipo esteja configurado corretamente e que você esteja capturando o tipo de dados que está procurando.</span><span class="sxs-lookup"><span data-stu-id="55da4-p126">It can be tricky to get the prototype and your task/question list right. Users might be confused by tasks, or might get lost in your prototype. You should run a pilot test with 1-3 users to work out the inevitable issues with the test format. This will help to ensure that your questions are clear, that the prototype is set up correctly, and that you’re capturing the type of data you’re looking for.</span></span>

## <a name="8-run-the-test"></a><span data-ttu-id="55da4-211">8. Execute o teste</span><span class="sxs-lookup"><span data-stu-id="55da4-211">8. Run the test</span></span>

<span data-ttu-id="55da4-p127">Depois que você solicitar o teste, receberá notificações por email quando os participantes o concluírem. A menos que tenha direcionado para um grupo específico de participantes, os testes normalmente são concluídos dentro de algumas horas.</span><span class="sxs-lookup"><span data-stu-id="55da4-p127">After you order your test, you will get email notifications when participants complete it. Unless you’ve targeted a specific group of participants, the tests are usually completed within a few hours.</span></span>

## <a name="9-analyze-results"></a><span data-ttu-id="55da4-214">9. Analise os resultados</span><span class="sxs-lookup"><span data-stu-id="55da4-214">9. Analyze results</span></span>

<span data-ttu-id="55da4-p128">Essa é a parte em que você tenta fazer com que os dados coletados façam sentido. Ao assistir os vídeos de teste, anote os problemas e os êxitos do usuário. Evite tentar interpretar o significado dos dados até que tenha exibido todos os resultados.</span><span class="sxs-lookup"><span data-stu-id="55da4-p128">This is the part where you try to make sense of the data you’ve collected. While watching the test videos, record notes about problems and successes the user has. Avoid trying to interpret the meaning of the data until you have viewed all the results.</span></span> 

<span data-ttu-id="55da4-p129">Um único participante com um problema de usabilidade não é suficiente para gerar uma alteração no design. Dois ou mais participantes que encontram o mesmo problema sugere que outros usuários no geral também encontrarão esse problema.</span><span class="sxs-lookup"><span data-stu-id="55da4-p129">A single participant having a usability issue is not enough to warrant making a change to the design. Two or more participants encountering the same issue suggests that other users in the general population will also encounter that issue.</span></span>

<span data-ttu-id="55da4-p130">Em geral, tome cuidado com como você usa seus dados para tirar conclusões. Não caia na armadilha de tentar fazer com que os dados se ajustem a uma determinada narrativa. Seja honesto sobre o que os dados realmente comprovam, refutam ou apenas falham em oferecer informações. Mantenha a mente aberta. O comportamento do usuário com frequência desafia as expectativas do designer.</span><span class="sxs-lookup"><span data-stu-id="55da4-p130">In general, be careful about how you use your data to draw conclusions. Don’t fall into the trap of trying to make the data fit a certain narrative; be honest about what the data actually proves, disproves, or simply fails to provide any insight about. Keep an open mind; user behavior frequently defies designer’s expectations.</span></span>
 

## <a name="see-also"></a><span data-ttu-id="55da4-223">Veja também</span><span class="sxs-lookup"><span data-stu-id="55da4-223">See also</span></span>
 
 - [<span data-ttu-id="55da4-224">Como conduzir testes de usabilidade</span><span class="sxs-lookup"><span data-stu-id="55da4-224">How to Conduct Usability Testing</span></span>](http://whatpixel.com/howto-conduct-usability-testing/)  
 - [<span data-ttu-id="55da4-225">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="55da4-225">Best Practices</span></span>](http://help.usertesting.com/customer/en/portal/articles/1680726-best-practices)  
 - [<span data-ttu-id="55da4-226">Minimizar desvio</span><span class="sxs-lookup"><span data-stu-id="55da4-226">Minimizing Bias</span></span>](http://downloads.usertesting.com/white_papers/TipSheet_MinimizingBias.pdf)  
