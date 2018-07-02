# <a name="first-run-experience-patterns"></a><span data-ttu-id="d5f9a-101">Padrões de telas de apresentação</span><span class="sxs-lookup"><span data-stu-id="d5f9a-101">First-run experience patterns</span></span>

<span data-ttu-id="d5f9a-102">Uma tela de apresentação (FRE) é a introdução do usuário ao seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-102">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="d5f9a-103">Uma FRE é apresentada quando um usuário abre um suplemento pela primeira vez. Ela fornece informações sobre as funções, recursos e/ou benefícios do suplemento.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-103">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="d5f9a-104">Essa experiência ajuda a moldar a impressão do usuário sobre um suplemento e pode aumentar a probabilidade de que ele retorne e continue a usá-lo.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-104">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="d5f9a-105">Melhores práticas</span><span class="sxs-lookup"><span data-stu-id="d5f9a-105">Best practices</span></span>


<span data-ttu-id="d5f9a-106">Siga estas práticas recomendadas ao criar sua tela de apresentação:</span><span class="sxs-lookup"><span data-stu-id="d5f9a-106">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="d5f9a-107">Faça</span><span class="sxs-lookup"><span data-stu-id="d5f9a-107">Do</span></span>|<span data-ttu-id="d5f9a-108">Não faça</span><span class="sxs-lookup"><span data-stu-id="d5f9a-108">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="d5f9a-109">Forneça uma breve introdução para as principais ações no suplemento.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-109">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="d5f9a-110">Não inclua informações e textos explicativos que não são relevantes para a introdução.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-110">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="d5f9a-111">Dê aos usuários a oportunidade de concluir uma ação que impactará positivamente o uso do suplemento.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-111">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="d5f9a-112">Não espere que os usuários aprendam tudo de uma vez.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-112">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="d5f9a-113">Concentre-se na ação que fornece o maior valor.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-113">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="d5f9a-114">Crie uma experiência envolvente que os usuários queiram concluir.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-114">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="d5f9a-115">Não force os usuários a seguir pela tela de apresentação.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-115">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="d5f9a-116">Dê aos usuários uma opção para ignorar a tela de apresentação.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-116">Give users an option to bypass the first-run experience.</span></span> |



<span data-ttu-id="d5f9a-117">Considere se é importante para o seu cenário mostrar a tela de apresentação aos usuários apenas uma vez ou periodicamente.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-117">Consider whether showing users the first-run experience once or many times is important to your scenario.</span></span> <span data-ttu-id="d5f9a-118">Por exemplo, se o suplemento não for usado com muita frequência, os usuários poderão se beneficiar de um novo contato com a tela de apresentação.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-118">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>



<span data-ttu-id="d5f9a-119">Utilize os seguintes padrões, quando aplicáveis, para criar ou aprimorar a tela de apresentação do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-119">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>



## <a name="carousel"></a><span data-ttu-id="d5f9a-120">Carrossel</span><span class="sxs-lookup"><span data-stu-id="d5f9a-120">Carousel</span></span>


<span data-ttu-id="d5f9a-121">O carrossel leva os usuários por uma série de recursos ou páginas informativas antes de começarem a usar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-121">Walkthrough takes users through a series of features or information before they start using the add-in. (PDF, code)</span></span>

<span data-ttu-id="d5f9a-122">*Figura 1: Permita que os usuários avancem ou pulem as páginas iniciais do fluxo do carrossel.*
![Tela de apresentação - Carrossel - Especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-step-1.png)</span><span class="sxs-lookup"><span data-stu-id="d5f9a-122">*Figure 1: Allow users to advance or skip the beginning pages of the carousel flow.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-1.png)</span></span>



<span data-ttu-id="d5f9a-123">*Figura 2: Reduza o número de telas do carrossel apresentadas ao usuário para apenas aquelas essenciais para comunicar efetivamente sua mensagem*
![Tela de apresentação - Carrossel - Especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-step-2.png)</span><span class="sxs-lookup"><span data-stu-id="d5f9a-123">*Figure 2: Minimize the number of carousel screens you present to the user to only what is needed to effectively communicate your message*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-2.png)</span></span>


<span data-ttu-id="d5f9a-124">*Figura 3: Forneça uma clara chamada para ação para sair da tela de apresentação.*
![Tela de apresentação - Carrossel - Especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-step-3.png)</span><span class="sxs-lookup"><span data-stu-id="d5f9a-124">*Figure 3: Provide a clear call to action to exit the first-run-experience.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-3.png)</span></span>



## <a name="value-placemat"></a><span data-ttu-id="d5f9a-125">Placemat de valor</span><span class="sxs-lookup"><span data-stu-id="d5f9a-125">Value Placemat</span></span>

<span data-ttu-id="d5f9a-126">O placemat de valor apresenta seu suplemento utilizando o logotipo, uma proposta de valor claramente definida, destaques ou resumo dos recursos e uma frase de chamada para ação.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-126">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>



<span data-ttu-id="d5f9a-127">![Tela de apresentação - Placemat de valor - Especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-value.png)
*Um placemat de valor com logotipo, clara proposta de valor, resumo de recursos e chamada para ação.*</span><span class="sxs-lookup"><span data-stu-id="d5f9a-127">![First Run - Value Placemat - Specifications for desktop task pane](../images/add-in-FRE-value.png)
*A value placemat with logo, clear value proposition, feature summary, and call to action.*</span></span>


### <a name="video-placemat"></a><span data-ttu-id="d5f9a-128">Placemat de vídeo</span><span class="sxs-lookup"><span data-stu-id="d5f9a-128">Video Placemat</span></span>

<span data-ttu-id="d5f9a-129">O placemat de vídeo mostra aos usuários um vídeo antes que eles comecem a usar o seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="d5f9a-129">Video shows users a video before they start using your add-in. (spec, code)</span></span>


<span data-ttu-id="d5f9a-130">*Figura 1: Placemat da tela de apresentação. Contém uma imagem estática do vídeo com um botão de reprodução e claro botão de chamada para a ação.*![Placemat de vídeo - Especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-video.png)</span><span class="sxs-lookup"><span data-stu-id="d5f9a-130">*Figure 1: First Run Placemat - The screen contains a still image from the video with a play button and clear call to action button.*![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video.png)</span></span>



<span data-ttu-id="d5f9a-131">*Figura 2: Player de vídeo - Um vídeo dentro de uma janela de diálogo é apresentado para os usuários.*
![Placemat de vídeo - Especificações para o painel de tarefas da área de trabalho](../images/add-in-FRE-video-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="d5f9a-131">*Figure 2: Video Player - Users are presented with a video within a dialog window.*
![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video-dialog.png)</span></span>
