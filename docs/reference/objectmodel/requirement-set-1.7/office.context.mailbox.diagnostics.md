
# <a name="diagnostics"></a><span data-ttu-id="7eb81-101">diagnóstico</span><span class="sxs-lookup"><span data-stu-id="7eb81-101">diagnostics</span></span>

### <span data-ttu-id="7eb81-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span><span class="sxs-lookup"><span data-stu-id="7eb81-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="7eb81-104">Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="7eb81-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7eb81-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7eb81-105">Requirements</span></span>

|<span data-ttu-id="7eb81-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="7eb81-106">Requirement</span></span>| <span data-ttu-id="7eb81-107">Valor</span><span class="sxs-lookup"><span data-stu-id="7eb81-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="7eb81-108">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7eb81-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7eb81-109">1.0</span><span class="sxs-lookup"><span data-stu-id="7eb81-109">1.0</span></span>|
|[<span data-ttu-id="7eb81-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7eb81-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7eb81-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7eb81-111">ReadItem</span></span>|
|[<span data-ttu-id="7eb81-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7eb81-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7eb81-113">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="7eb81-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="7eb81-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="7eb81-114">Members and methods</span></span>

| <span data-ttu-id="7eb81-115">Membro</span><span class="sxs-lookup"><span data-stu-id="7eb81-115">Member</span></span> | <span data-ttu-id="7eb81-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="7eb81-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="7eb81-117">hostname</span><span class="sxs-lookup"><span data-stu-id="7eb81-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="7eb81-118">Membro</span><span class="sxs-lookup"><span data-stu-id="7eb81-118">Member</span></span> |
| [<span data-ttu-id="7eb81-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="7eb81-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="7eb81-120">Membro</span><span class="sxs-lookup"><span data-stu-id="7eb81-120">Member</span></span> |
| [<span data-ttu-id="7eb81-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="7eb81-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="7eb81-122">Membro</span><span class="sxs-lookup"><span data-stu-id="7eb81-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="7eb81-123">Membros</span><span class="sxs-lookup"><span data-stu-id="7eb81-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="7eb81-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="7eb81-124">hostName :String</span></span>

<span data-ttu-id="7eb81-125">Obtém uma sequência de caracteres que representa o nome do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="7eb81-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="7eb81-126">Uma sequência de caracteres que pode ser um dos valores a seguir: `Outlook`, `Mac Outlook`, `OutlookIOS` ou `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="7eb81-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="7eb81-127">Tipo:</span><span class="sxs-lookup"><span data-stu-id="7eb81-127">Type:</span></span>

*   <span data-ttu-id="7eb81-128">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7eb81-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7eb81-129">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7eb81-129">Requirements</span></span>

|<span data-ttu-id="7eb81-130">Requisito</span><span class="sxs-lookup"><span data-stu-id="7eb81-130">Requirement</span></span>| <span data-ttu-id="7eb81-131">Valor</span><span class="sxs-lookup"><span data-stu-id="7eb81-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="7eb81-132">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7eb81-132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7eb81-133">1.0</span><span class="sxs-lookup"><span data-stu-id="7eb81-133">1.0</span></span>|
|[<span data-ttu-id="7eb81-134">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7eb81-134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7eb81-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7eb81-135">ReadItem</span></span>|
|[<span data-ttu-id="7eb81-136">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7eb81-136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7eb81-137">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="7eb81-137">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="7eb81-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="7eb81-138">hostVersion :String</span></span>

<span data-ttu-id="7eb81-139">Obtém uma cadeia de caracteres que representa a versão do aplicativo host ou do Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="7eb81-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="7eb81-p102">Se o suplemento de email estiver em execução no cliente do Outlook para área de trabalho ou Outlook para iOS, a propriedade `hostVersion` retornará a versão do aplicativo host, o Outlook. No Outlook Web App, a propriedade retorna a versão do Exchange Server. Um exemplo é a cadeia de caracteres `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="7eb81-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="7eb81-143">Tipo:</span><span class="sxs-lookup"><span data-stu-id="7eb81-143">Type:</span></span>

*   <span data-ttu-id="7eb81-144">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="7eb81-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7eb81-145">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7eb81-145">Requirements</span></span>

|<span data-ttu-id="7eb81-146">Requisito</span><span class="sxs-lookup"><span data-stu-id="7eb81-146">Requirement</span></span>| <span data-ttu-id="7eb81-147">Valor</span><span class="sxs-lookup"><span data-stu-id="7eb81-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="7eb81-148">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7eb81-148">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7eb81-149">1.0</span><span class="sxs-lookup"><span data-stu-id="7eb81-149">1.0</span></span>|
|[<span data-ttu-id="7eb81-150">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7eb81-150">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7eb81-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7eb81-151">ReadItem</span></span>|
|[<span data-ttu-id="7eb81-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7eb81-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7eb81-153">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="7eb81-153">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="7eb81-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="7eb81-154">OWAView :String</span></span>

<span data-ttu-id="7eb81-155">Obtém uma cadeia de caracteres que representa o modo de exibição atual do Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="7eb81-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="7eb81-156">A cadeia de caracteres retornada pode ser um dos valores a seguir: `OneColumn`, `TwoColumns` ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="7eb81-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="7eb81-157">Se o aplicativo host não for Outlook Web App, acessar essa propriedade resultará em `undefined`.</span><span class="sxs-lookup"><span data-stu-id="7eb81-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="7eb81-158">O Outlook Web App tem três modos de exibição que correspondem à largura da tela e da janela, e à quantidade de colunas que pode ser exibida:</span><span class="sxs-lookup"><span data-stu-id="7eb81-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="7eb81-p103">`OneColumn`, que é exibido quando a tela é estreita. O Outlook Web App usa esse layout de coluna única em toda a tela de um smartphone.</span><span class="sxs-lookup"><span data-stu-id="7eb81-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="7eb81-p104">`TwoColumns`, que é exibido quando a tela é mais larga. O Outlook Web App usa esse modo de exibição na maioria dos tablets.</span><span class="sxs-lookup"><span data-stu-id="7eb81-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="7eb81-p105">`ThreeColumns`, que é exibido quando a tela é ainda mais larga. Por exemplo, o Outlook Web App usa esse modo de exibição em um modo de tela cheia em um desktop.</span><span class="sxs-lookup"><span data-stu-id="7eb81-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="7eb81-165">Tipo:</span><span class="sxs-lookup"><span data-stu-id="7eb81-165">Type:</span></span>

*   <span data-ttu-id="7eb81-166">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7eb81-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7eb81-167">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7eb81-167">Requirements</span></span>

|<span data-ttu-id="7eb81-168">Requisito</span><span class="sxs-lookup"><span data-stu-id="7eb81-168">Requirement</span></span>| <span data-ttu-id="7eb81-169">Valor</span><span class="sxs-lookup"><span data-stu-id="7eb81-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="7eb81-170">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7eb81-170">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7eb81-171">1.0</span><span class="sxs-lookup"><span data-stu-id="7eb81-171">1.0</span></span>|
|[<span data-ttu-id="7eb81-172">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7eb81-172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7eb81-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7eb81-173">ReadItem</span></span>|
|[<span data-ttu-id="7eb81-174">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7eb81-174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7eb81-175">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="7eb81-175">Compose or read</span></span>|