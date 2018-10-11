
# <a name="diagnostics"></a><span data-ttu-id="886d6-101">diagnostics</span><span class="sxs-lookup"><span data-stu-id="886d6-101">diagnostics</span></span>

### <span data-ttu-id="886d6-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span><span class="sxs-lookup"><span data-stu-id="886d6-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="886d6-104">Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="886d6-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="886d6-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="886d6-105">Requirements</span></span>

|<span data-ttu-id="886d6-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="886d6-106">Requirement</span></span>| <span data-ttu-id="886d6-107">Valor</span><span class="sxs-lookup"><span data-stu-id="886d6-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="886d6-108">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="886d6-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="886d6-109">1.0</span><span class="sxs-lookup"><span data-stu-id="886d6-109">1.0</span></span>|
|[<span data-ttu-id="886d6-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="886d6-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="886d6-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="886d6-111">ReadItem</span></span>|
|[<span data-ttu-id="886d6-112">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="886d6-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="886d6-113">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="886d6-113">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="886d6-114">Membros</span><span class="sxs-lookup"><span data-stu-id="886d6-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="886d6-115">hostName :String</span><span class="sxs-lookup"><span data-stu-id="886d6-115">hostName :String</span></span>

<span data-ttu-id="886d6-116">Obtém uma sequência de caracteres que representa o nome do aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="886d6-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="886d6-117">Uma sequência de caracteres que pode ser um dos valores a seguir: `Outlook`, `OutlookIOS` ou `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="886d6-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, , or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="886d6-118">Tipo:</span><span class="sxs-lookup"><span data-stu-id="886d6-118">Type:</span></span>

*   <span data-ttu-id="886d6-119">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="886d6-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="886d6-120">Requisitos</span><span class="sxs-lookup"><span data-stu-id="886d6-120">Requirements</span></span>

|<span data-ttu-id="886d6-121">Requisito</span><span class="sxs-lookup"><span data-stu-id="886d6-121">Requirement</span></span>| <span data-ttu-id="886d6-122">Valor</span><span class="sxs-lookup"><span data-stu-id="886d6-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="886d6-123">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="886d6-123">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="886d6-124">1.0</span><span class="sxs-lookup"><span data-stu-id="886d6-124">1.0</span></span>|
|[<span data-ttu-id="886d6-125">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="886d6-125">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="886d6-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="886d6-126">ReadItem</span></span>|
|[<span data-ttu-id="886d6-127">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="886d6-127">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="886d6-128">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="886d6-128">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="886d6-129">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="886d6-129">hostVersion :String</span></span>

<span data-ttu-id="886d6-130">Obtém uma sequência de caracteres que representa a versão do aplicativo host ou do Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="886d6-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="886d6-p102">Se o suplemento de e-mail estiver em execução no cliente da área de trabalho do Outlook ou no Outlook para iOS, a propriedade `hostVersion` retornará a versão do aplicativo host, o Outlook. No Outlook Web App, a propriedade retorna a versão do Exchange Server. Um exemplo é a sequência de caracteres `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="886d6-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="886d6-134">Tipo:</span><span class="sxs-lookup"><span data-stu-id="886d6-134">Type:</span></span>

*   <span data-ttu-id="886d6-135">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="886d6-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="886d6-136">Requisitos</span><span class="sxs-lookup"><span data-stu-id="886d6-136">Requirements</span></span>

|<span data-ttu-id="886d6-137">Requisito</span><span class="sxs-lookup"><span data-stu-id="886d6-137">Requirement</span></span>| <span data-ttu-id="886d6-138">Valor</span><span class="sxs-lookup"><span data-stu-id="886d6-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="886d6-139">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="886d6-139">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="886d6-140">1.0</span><span class="sxs-lookup"><span data-stu-id="886d6-140">1.0</span></span>|
|[<span data-ttu-id="886d6-141">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="886d6-141">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="886d6-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="886d6-142">ReadItem</span></span>|
|[<span data-ttu-id="886d6-143">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="886d6-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="886d6-144">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="886d6-144">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="886d6-145">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="886d6-145">OWAView :String</span></span>

<span data-ttu-id="886d6-146">Obtém uma sequência de caracteres que representa o modo de exibição atual do Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="886d6-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="886d6-147">A sequência de caracteres retornada pode ser um dos valores a seguir: `OneColumn`, `TwoColumns` ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="886d6-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="886d6-148">Se o aplicativo host não for o Outlook Web App, o acesso a essa propriedade resultará em `undefined`.</span><span class="sxs-lookup"><span data-stu-id="886d6-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="886d6-149">O Outlook Web App tem três modos de exibição que correspondem à largura da tela e da janela, e ao número de colunas que pode ser exibido:</span><span class="sxs-lookup"><span data-stu-id="886d6-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="886d6-p103">`OneColumn`, que é exibido quando a tela é estreita. O Outlook Web App usa esse layout de coluna única em toda a tela de um smartphone.</span><span class="sxs-lookup"><span data-stu-id="886d6-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="886d6-p104">`TwoColumns`, que é exibido quando a tela é mais larga. O Outlook Web App usa esse modo de exibição na maioria dos tablets.</span><span class="sxs-lookup"><span data-stu-id="886d6-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="886d6-p105">`ThreeColumns`, que é exibido quando a tela é larga. Por exemplo, o Outlook Web App usa esse modo de exibição em uma janela de tela inteira em um computador.</span><span class="sxs-lookup"><span data-stu-id="886d6-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="886d6-156">Tipo:</span><span class="sxs-lookup"><span data-stu-id="886d6-156">Type:</span></span>

*   <span data-ttu-id="886d6-157">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="886d6-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="886d6-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="886d6-158">Requirements</span></span>

|<span data-ttu-id="886d6-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="886d6-159">Requirement</span></span>| <span data-ttu-id="886d6-160">Valor</span><span class="sxs-lookup"><span data-stu-id="886d6-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="886d6-161">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="886d6-161">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="886d6-162">1.0</span><span class="sxs-lookup"><span data-stu-id="886d6-162">1.0</span></span>|
|[<span data-ttu-id="886d6-163">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="886d6-163">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="886d6-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="886d6-164">ReadItem</span></span>|
|[<span data-ttu-id="886d6-165">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="886d6-165">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="886d6-166">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="886d6-166">Compose or read</span></span>|