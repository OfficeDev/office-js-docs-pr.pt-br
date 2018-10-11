
# <a name="diagnostics"></a><span data-ttu-id="a35ac-101">diagnostics</span><span class="sxs-lookup"><span data-stu-id="a35ac-101">diagnostics</span></span>

### <span data-ttu-id="a35ac-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span><span class="sxs-lookup"><span data-stu-id="a35ac-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="a35ac-104">Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="a35ac-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a35ac-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a35ac-105">Requirements</span></span>

|<span data-ttu-id="a35ac-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="a35ac-106">Requirement</span></span>| <span data-ttu-id="a35ac-107">Valor</span><span class="sxs-lookup"><span data-stu-id="a35ac-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a35ac-108">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a35ac-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a35ac-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a35ac-109">1.0</span></span>|
|[<span data-ttu-id="a35ac-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a35ac-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a35ac-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a35ac-111">ReadItem</span></span>|
|[<span data-ttu-id="a35ac-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a35ac-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a35ac-113">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="a35ac-113">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="a35ac-114">Membros</span><span class="sxs-lookup"><span data-stu-id="a35ac-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="a35ac-115">hostName :String</span><span class="sxs-lookup"><span data-stu-id="a35ac-115">hostName :String</span></span>

<span data-ttu-id="a35ac-116">Obtém uma sequência de caracteres que representa o nome do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="a35ac-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="a35ac-117">Uma sequência de caracteres que pode ser um dos valores a seguir: `Outlook`, `OutlookIOS` ou `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="a35ac-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, , or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="a35ac-118">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a35ac-118">Type:</span></span>

*   <span data-ttu-id="a35ac-119">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="a35ac-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a35ac-120">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a35ac-120">Requirements</span></span>

|<span data-ttu-id="a35ac-121">Requisito</span><span class="sxs-lookup"><span data-stu-id="a35ac-121">Requirement</span></span>| <span data-ttu-id="a35ac-122">Valor</span><span class="sxs-lookup"><span data-stu-id="a35ac-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="a35ac-123">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a35ac-123">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a35ac-124">1.0</span><span class="sxs-lookup"><span data-stu-id="a35ac-124">1.0</span></span>|
|[<span data-ttu-id="a35ac-125">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a35ac-125">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a35ac-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a35ac-126">ReadItem</span></span>|
|[<span data-ttu-id="a35ac-127">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="a35ac-127">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a35ac-128">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="a35ac-128">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="a35ac-129">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="a35ac-129">hostVersion :String</span></span>

<span data-ttu-id="a35ac-130">Obtém uma sequência de caracteres que representa a versão do aplicativo host ou do Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="a35ac-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="a35ac-p102">Se o suplemento de email estiver em execução no cliente da área de trabalho do Outlook ou no Outlook para iOS, a propriedade `hostVersion` retornará a versão do aplicativo host, o Outlook. No Outlook Web App, a propriedade retorna a versão do Exchange Server. Um exemplo é a sequência de caracteres `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="a35ac-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="a35ac-134">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a35ac-134">Type:</span></span>

*   <span data-ttu-id="a35ac-135">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="a35ac-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a35ac-136">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a35ac-136">Requirements</span></span>

|<span data-ttu-id="a35ac-137">Requisito</span><span class="sxs-lookup"><span data-stu-id="a35ac-137">Requirement</span></span>| <span data-ttu-id="a35ac-138">Valor</span><span class="sxs-lookup"><span data-stu-id="a35ac-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="a35ac-139">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a35ac-139">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a35ac-140">1.0</span><span class="sxs-lookup"><span data-stu-id="a35ac-140">1.0</span></span>|
|[<span data-ttu-id="a35ac-141">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a35ac-141">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a35ac-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a35ac-142">ReadItem</span></span>|
|[<span data-ttu-id="a35ac-143">Modo aplicável ao Outlook</span><span class="sxs-lookup"><span data-stu-id="a35ac-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a35ac-144">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="a35ac-144">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="a35ac-145">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="a35ac-145">OWAView :String</span></span>

<span data-ttu-id="a35ac-146">Obtém uma sequência de caracteres que representa o modo de exibição atual do Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="a35ac-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="a35ac-147">A sequência de caracteres retornada pode ser um dos valores a seguir: `OneColumn`, `TwoColumns` ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="a35ac-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="a35ac-148">Se o aplicativo host não for o Outlook Web App, acessar essa propriedade resultará em `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a35ac-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="a35ac-149">O Outlook Web App tem três modos de exibição que correspondem à largura da tela e da janela e à quantidade de colunas que pode ser exibida:</span><span class="sxs-lookup"><span data-stu-id="a35ac-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="a35ac-p103">`OneColumn`, que é exibido quando a tela é estreita. O Outlook Web App usa esse layout de coluna única na tela inteira de um smartphone.</span><span class="sxs-lookup"><span data-stu-id="a35ac-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="a35ac-p104">`TwoColumns`, que é exibido quando a tela é mais larga. O Outlook Web App usa esse modo de exibição na maioria dos tablets.</span><span class="sxs-lookup"><span data-stu-id="a35ac-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="a35ac-p105">`ThreeColumns`, que é exibido quando a tela é larga. Por exemplo, o Outlook Web App usa esse modo de exibição em uma janela de tela inteira em um computador.</span><span class="sxs-lookup"><span data-stu-id="a35ac-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="a35ac-156">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a35ac-156">Type:</span></span>

*   <span data-ttu-id="a35ac-157">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="a35ac-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a35ac-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a35ac-158">Requirements</span></span>

|<span data-ttu-id="a35ac-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="a35ac-159">Requirement</span></span>| <span data-ttu-id="a35ac-160">Valor</span><span class="sxs-lookup"><span data-stu-id="a35ac-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="a35ac-161">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a35ac-161">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a35ac-162">1.0</span><span class="sxs-lookup"><span data-stu-id="a35ac-162">1.0</span></span>|
|[<span data-ttu-id="a35ac-163">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a35ac-163">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a35ac-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a35ac-164">ReadItem</span></span>|
|[<span data-ttu-id="a35ac-165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a35ac-165">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a35ac-166">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="a35ac-166">Compose or read</span></span>|