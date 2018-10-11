
# <a name="userprofile"></a><span data-ttu-id="532c5-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="532c5-101">userProfile</span></span>

### <span data-ttu-id="532c5-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="532c5-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="532c5-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="532c5-104">Requirements</span></span>

|<span data-ttu-id="532c5-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="532c5-105">Requirement</span></span>| <span data-ttu-id="532c5-106">Valor</span><span class="sxs-lookup"><span data-stu-id="532c5-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="532c5-107">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="532c5-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="532c5-108">1.0</span><span class="sxs-lookup"><span data-stu-id="532c5-108">1.0</span></span>|
|[<span data-ttu-id="532c5-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="532c5-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="532c5-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="532c5-110">ReadItem</span></span>|
|[<span data-ttu-id="532c5-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="532c5-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="532c5-112">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="532c5-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="532c5-113">Membros</span><span class="sxs-lookup"><span data-stu-id="532c5-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="532c5-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="532c5-114">displayName :String</span></span>

<span data-ttu-id="532c5-115">Obtém o nome para exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="532c5-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="532c5-116">Tipo:</span><span class="sxs-lookup"><span data-stu-id="532c5-116">Type:</span></span>

*   <span data-ttu-id="532c5-117">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="532c5-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="532c5-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="532c5-118">Requirements</span></span>

|<span data-ttu-id="532c5-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="532c5-119">Requirement</span></span>| <span data-ttu-id="532c5-120">Valor</span><span class="sxs-lookup"><span data-stu-id="532c5-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="532c5-121">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="532c5-121">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="532c5-122">1.0</span><span class="sxs-lookup"><span data-stu-id="532c5-122">1.0</span></span>|
|[<span data-ttu-id="532c5-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="532c5-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="532c5-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="532c5-124">ReadItem</span></span>|
|[<span data-ttu-id="532c5-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="532c5-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="532c5-126">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="532c5-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="532c5-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="532c5-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="532c5-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="532c5-128">emailAddress :String</span></span>

<span data-ttu-id="532c5-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="532c5-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="532c5-130">Tipo:</span><span class="sxs-lookup"><span data-stu-id="532c5-130">Type:</span></span>

*   <span data-ttu-id="532c5-131">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="532c5-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="532c5-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="532c5-132">Requirements</span></span>

|<span data-ttu-id="532c5-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="532c5-133">Requirement</span></span>| <span data-ttu-id="532c5-134">Valor</span><span class="sxs-lookup"><span data-stu-id="532c5-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="532c5-135">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="532c5-135">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="532c5-136">1.0</span><span class="sxs-lookup"><span data-stu-id="532c5-136">1.0</span></span>|
|[<span data-ttu-id="532c5-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="532c5-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="532c5-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="532c5-138">ReadItem</span></span>|
|[<span data-ttu-id="532c5-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="532c5-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="532c5-140">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="532c5-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="532c5-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="532c5-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="532c5-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="532c5-142">timeZone :String</span></span>

<span data-ttu-id="532c5-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="532c5-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="532c5-144">Tipo:</span><span class="sxs-lookup"><span data-stu-id="532c5-144">Type:</span></span>

*   <span data-ttu-id="532c5-145">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="532c5-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="532c5-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="532c5-146">Requirements</span></span>

|<span data-ttu-id="532c5-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="532c5-147">Requirement</span></span>| <span data-ttu-id="532c5-148">Valor</span><span class="sxs-lookup"><span data-stu-id="532c5-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="532c5-149">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="532c5-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="532c5-150">1.0</span><span class="sxs-lookup"><span data-stu-id="532c5-150">1.0</span></span>|
|[<span data-ttu-id="532c5-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="532c5-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="532c5-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="532c5-152">ReadItem</span></span>|
|[<span data-ttu-id="532c5-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="532c5-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="532c5-154">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="532c5-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="532c5-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="532c5-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```