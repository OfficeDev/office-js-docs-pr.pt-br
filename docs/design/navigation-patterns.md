# <a name="navigation-patterns"></a><span data-ttu-id="78e89-101">Padrões de navegação</span><span class="sxs-lookup"><span data-stu-id="78e89-101">Navigation patterns</span></span>

<span data-ttu-id="78e89-102">Os principais recursos de um suplemento são acessados ​​por meio de tipos de comandos específicos e área de tela limitada.</span><span class="sxs-lookup"><span data-stu-id="78e89-102">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="78e89-103">É importante que a navegação seja intuitiva, forneça contexto e permita que o usuário se mova facilmente por todo o suplemento.</span><span class="sxs-lookup"><span data-stu-id="78e89-103">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="78e89-104">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="78e89-104">Best practices</span></span>

| <span data-ttu-id="78e89-105">Fazer</span><span class="sxs-lookup"><span data-stu-id="78e89-105">Do</span></span>    | <span data-ttu-id="78e89-106">Não fazer</span><span class="sxs-lookup"><span data-stu-id="78e89-106">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="78e89-107">Certifique-se de que o usuário tenha uma opção de navegação claramente visível.</span><span class="sxs-lookup"><span data-stu-id="78e89-107">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="78e89-108">Não complique o processo de navegação usando a interface do usuário não padrão.</span><span class="sxs-lookup"><span data-stu-id="78e89-108">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="78e89-109">Utilize os seguintes componentes, conforme aplicável, para permitir que os usuários naveguem pelo seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="78e89-109">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="78e89-110">Não dificulte o usuário a entender seu local ou contexto atual dentro do suplemento</span><span class="sxs-lookup"><span data-stu-id="78e89-110">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>



## <a name="command-bar"></a><span data-ttu-id="78e89-111">Barra de Comandos</span><span class="sxs-lookup"><span data-stu-id="78e89-111">command bar</span></span>

<span data-ttu-id="78e89-112">A Barra de Comandos é uma superfície que abriga os comandos que operam no conteúdo da janela, painel ou região pai que reside acima.</span><span class="sxs-lookup"><span data-stu-id="78e89-112">CommandBar is a surface that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="78e89-113">Os recursos opcionais incluem um ponto de acesso do menu de hambúrguer, pesquisa e comandos laterais.</span><span class="sxs-lookup"><span data-stu-id="78e89-113">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![Comandos - Especificações para o painel de tarefas da área de trabalho](../images/add-in-command-bar.png)



## <a name="tab-bar"></a><span data-ttu-id="78e89-115">Barra de Guias</span><span class="sxs-lookup"><span data-stu-id="78e89-115">Tab bar</span></span>

<span data-ttu-id="78e89-116">Mostra a navegação usando botões com texto e ícones empilhados na vertical.</span><span class="sxs-lookup"><span data-stu-id="78e89-116">Tab bar - Shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="78e89-117">Use a barra de guias para fornecer a navegação usando guias com títulos curtos e descritivos.</span><span class="sxs-lookup"><span data-stu-id="78e89-117">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![Barra de Guias - Especificações para o painel de tarefas da área de trabalho](../images/add-in-tab-bar.png)


## <a name="back-button"></a><span data-ttu-id="78e89-119">Botão voltar</span><span class="sxs-lookup"><span data-stu-id="78e89-119">Back button</span></span>

<span data-ttu-id="78e89-120">O botão voltar permite que os usuários se recuperem de uma ação de navegação de busca detalhada.</span><span class="sxs-lookup"><span data-stu-id="78e89-120">The back button allows users to recover from a drill down navigational action.</span></span> <span data-ttu-id="78e89-121">Esse padrão ajuda a garantir que os usuários sigam uma série de etapas ordenadas.</span><span class="sxs-lookup"><span data-stu-id="78e89-121">Use this pattern to ensure users follow an ordered series of steps.</span></span>  

![Botão Voltar - Especificações para o painel de tarefas da área de trabalho](../images/add-in-back-button.png)
