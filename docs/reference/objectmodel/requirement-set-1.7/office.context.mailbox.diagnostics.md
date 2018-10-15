
# <a name="diagnostics"></a><span data-ttu-id="8885d-101">diagnostics</span><span class="sxs-lookup"><span data-stu-id="8885d-101">diagnostics</span></span>

### <span data-ttu-id="8885d-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span><span class="sxs-lookup"><span data-stu-id="8885d-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="8885d-104">Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="8885d-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8885d-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8885d-105">Requirements</span></span>

|<span data-ttu-id="8885d-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="8885d-106">Requirement</span></span>| <span data-ttu-id="8885d-107">Valor</span><span class="sxs-lookup"><span data-stu-id="8885d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8885d-108">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8885d-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8885d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="8885d-109">1.0</span></span>|
|[<span data-ttu-id="8885d-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8885d-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8885d-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8885d-111">ReadItem</span></span>|
|[<span data-ttu-id="8885d-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8885d-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8885d-113">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="8885d-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8885d-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="8885d-114">Members and methods</span></span>

| <span data-ttu-id="8885d-115">Membro</span><span class="sxs-lookup"><span data-stu-id="8885d-115">Member</span></span> | <span data-ttu-id="8885d-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="8885d-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8885d-117">hostname</span><span class="sxs-lookup"><span data-stu-id="8885d-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="8885d-118">Membro</span><span class="sxs-lookup"><span data-stu-id="8885d-118">Member</span></span> |
| [<span data-ttu-id="8885d-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="8885d-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="8885d-120">Membro</span><span class="sxs-lookup"><span data-stu-id="8885d-120">Member</span></span> |
| [<span data-ttu-id="8885d-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="8885d-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="8885d-122">Membro</span><span class="sxs-lookup"><span data-stu-id="8885d-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="8885d-123">Membros</span><span class="sxs-lookup"><span data-stu-id="8885d-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="8885d-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="8885d-124">hostName :String</span></span>

<span data-ttu-id="8885d-125">Obtém uma sequência de caracteres que representa o nome do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="8885d-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="8885d-126">Uma sequência de caracteres que pode ser um dos valores a seguir: `Outlook`, `Mac Outlook`, `OutlookIOS` ou `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="8885d-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="8885d-127">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8885d-127">Type:</span></span>

*   <span data-ttu-id="8885d-128">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="8885d-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8885d-129">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8885d-129">Requirements</span></span>

|<span data-ttu-id="8885d-130">Requisito</span><span class="sxs-lookup"><span data-stu-id="8885d-130">Requirement</span></span>| <span data-ttu-id="8885d-131">Valor</span><span class="sxs-lookup"><span data-stu-id="8885d-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="8885d-132">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8885d-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8885d-133">1.0</span><span class="sxs-lookup"><span data-stu-id="8885d-133">1.0</span></span>|
|[<span data-ttu-id="8885d-134">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8885d-134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8885d-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8885d-135">ReadItem</span></span>|
|[<span data-ttu-id="8885d-136">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8885d-136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8885d-137">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="8885d-137">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="8885d-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="8885d-138">hostVersion :String</span></span>

<span data-ttu-id="8885d-139">Obtém uma sequência de caracteres que representa a versão do aplicativo host ou do Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="8885d-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="8885d-p102">Se o suplemento de email estiver em execução no cliente da área de trabalho do Outlook ou no Outlook para iOS, a propriedade `hostVersion` retornará a versão do aplicativo host, o Outlook. No Outlook Web App, a propriedade retorna a versão do Exchange Server. Um exemplo é a sequência de caracteres `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="8885d-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="8885d-143">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8885d-143">Type:</span></span>

*   <span data-ttu-id="8885d-144">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="8885d-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8885d-145">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8885d-145">Requirements</span></span>

|<span data-ttu-id="8885d-146">Requisito</span><span class="sxs-lookup"><span data-stu-id="8885d-146">Requirement</span></span>| <span data-ttu-id="8885d-147">Valor</span><span class="sxs-lookup"><span data-stu-id="8885d-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="8885d-148">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8885d-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8885d-149">1.0</span><span class="sxs-lookup"><span data-stu-id="8885d-149">1.0</span></span>|
|[<span data-ttu-id="8885d-150">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8885d-150">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8885d-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8885d-151">ReadItem</span></span>|
|[<span data-ttu-id="8885d-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8885d-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8885d-153">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="8885d-153">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="8885d-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="8885d-154">OWAView :String</span></span>

<span data-ttu-id="8885d-155">Obtém uma sequência de caracteres que representa o modo de exibição atual do Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="8885d-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="8885d-156">A sequência de caracteres retornada pode ser um dos valores a seguir: `OneColumn`, `TwoColumns` ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="8885d-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="8885d-157">Se o aplicativo host não for o Outlook Web App, o acesso a essa propriedade resultará em `undefined`.</span><span class="sxs-lookup"><span data-stu-id="8885d-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="8885d-158">O Outlook Web App tem três modos de exibição que correspondem à largura da tela e da janela, e ao número de colunas que pode ser exibido:</span><span class="sxs-lookup"><span data-stu-id="8885d-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="8885d-p103">`OneColumn`, que é exibido quando a tela é estreita. O Outlook Web App usa esse layout de coluna única em toda a tela de um smartphone.</span><span class="sxs-lookup"><span data-stu-id="8885d-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="8885d-p104">`TwoColumns`, que é exibido quando a tela é mais larga. O Outlook Web App usa esse modo de exibição na maioria dos tablets.</span><span class="sxs-lookup"><span data-stu-id="8885d-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="8885d-p105">`ThreeColumns`, que é exibido quando a tela é larga. Por exemplo, o Outlook Web App usa esse modo de exibição em uma janela de tela inteira em um computador.</span><span class="sxs-lookup"><span data-stu-id="8885d-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="8885d-165">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8885d-165">Type:</span></span>

*   <span data-ttu-id="8885d-166">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="8885d-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8885d-167">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8885d-167">Requirements</span></span>

|<span data-ttu-id="8885d-168">Requisito</span><span class="sxs-lookup"><span data-stu-id="8885d-168">Requirement</span></span>| <span data-ttu-id="8885d-169">Valor</span><span class="sxs-lookup"><span data-stu-id="8885d-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="8885d-170">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8885d-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8885d-171">1.0</span><span class="sxs-lookup"><span data-stu-id="8885d-171">1.0</span></span>|
|[<span data-ttu-id="8885d-172">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8885d-172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8885d-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8885d-173">ReadItem</span></span>|
|[<span data-ttu-id="8885d-174">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8885d-174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8885d-175">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="8885d-175">Compose or read</span></span>|