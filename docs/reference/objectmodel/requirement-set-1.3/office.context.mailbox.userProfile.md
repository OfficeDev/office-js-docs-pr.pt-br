
# <a name="userprofile"></a><span data-ttu-id="a548e-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="a548e-101">userProfile</span></span>

### <span data-ttu-id="a548e-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="a548e-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="a548e-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a548e-104">Requirements</span></span>

|<span data-ttu-id="a548e-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="a548e-105">Requirement</span></span>| <span data-ttu-id="a548e-106">Valor</span><span class="sxs-lookup"><span data-stu-id="a548e-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a548e-107">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a548e-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a548e-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a548e-108">1.0</span></span>|
|[<span data-ttu-id="a548e-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a548e-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a548e-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a548e-110">ReadItem</span></span>|
|[<span data-ttu-id="a548e-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a548e-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a548e-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a548e-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="a548e-113">Membros</span><span class="sxs-lookup"><span data-stu-id="a548e-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="a548e-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="a548e-114">displayName :String</span></span>

<span data-ttu-id="a548e-115">Obtém o nome para exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="a548e-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="a548e-116">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a548e-116">Type:</span></span>

*   <span data-ttu-id="a548e-117">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="a548e-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a548e-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a548e-118">Requirements</span></span>

|<span data-ttu-id="a548e-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="a548e-119">Requirement</span></span>| <span data-ttu-id="a548e-120">Valor</span><span class="sxs-lookup"><span data-stu-id="a548e-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="a548e-121">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a548e-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a548e-122">1.0</span><span class="sxs-lookup"><span data-stu-id="a548e-122">1.0</span></span>|
|[<span data-ttu-id="a548e-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a548e-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a548e-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a548e-124">ReadItem</span></span>|
|[<span data-ttu-id="a548e-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a548e-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a548e-126">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a548e-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a548e-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a548e-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="a548e-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="a548e-128">emailAddress :String</span></span>

<span data-ttu-id="a548e-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="a548e-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="a548e-130">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a548e-130">Type:</span></span>

*   <span data-ttu-id="a548e-131">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="a548e-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a548e-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a548e-132">Requirements</span></span>

|<span data-ttu-id="a548e-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="a548e-133">Requirement</span></span>| <span data-ttu-id="a548e-134">Valor</span><span class="sxs-lookup"><span data-stu-id="a548e-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="a548e-135">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a548e-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a548e-136">1.0</span><span class="sxs-lookup"><span data-stu-id="a548e-136">1.0</span></span>|
|[<span data-ttu-id="a548e-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a548e-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a548e-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a548e-138">ReadItem</span></span>|
|[<span data-ttu-id="a548e-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a548e-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a548e-140">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a548e-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a548e-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a548e-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="a548e-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="a548e-142">timeZone :String</span></span>

<span data-ttu-id="a548e-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="a548e-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="a548e-144">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a548e-144">Type:</span></span>

*   <span data-ttu-id="a548e-145">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="a548e-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a548e-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a548e-146">Requirements</span></span>

|<span data-ttu-id="a548e-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="a548e-147">Requirement</span></span>| <span data-ttu-id="a548e-148">Valor</span><span class="sxs-lookup"><span data-stu-id="a548e-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="a548e-149">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a548e-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a548e-150">1.0</span><span class="sxs-lookup"><span data-stu-id="a548e-150">1.0</span></span>|
|[<span data-ttu-id="a548e-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a548e-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a548e-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a548e-152">ReadItem</span></span>|
|[<span data-ttu-id="a548e-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a548e-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a548e-154">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a548e-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a548e-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a548e-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```