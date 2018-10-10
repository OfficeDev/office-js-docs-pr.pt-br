
# <a name="userprofile"></a><span data-ttu-id="3bec3-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="3bec3-101">userProfile</span></span>

### <span data-ttu-id="3bec3-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="3bec3-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="3bec3-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3bec3-104">Requirements</span></span>

|<span data-ttu-id="3bec3-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="3bec3-105">Requirement</span></span>| <span data-ttu-id="3bec3-106">Valor</span><span class="sxs-lookup"><span data-stu-id="3bec3-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="3bec3-107">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3bec3-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3bec3-108">1.0</span><span class="sxs-lookup"><span data-stu-id="3bec3-108">1.0</span></span>|
|[<span data-ttu-id="3bec3-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="3bec3-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3bec3-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3bec3-110">ReadItem</span></span>|
|[<span data-ttu-id="3bec3-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3bec3-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3bec3-112">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="3bec3-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="3bec3-113">Membros</span><span class="sxs-lookup"><span data-stu-id="3bec3-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="3bec3-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="3bec3-114">displayName :String</span></span>

<span data-ttu-id="3bec3-115">Obtém o nome para exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="3bec3-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="3bec3-116">Type:</span><span class="sxs-lookup"><span data-stu-id="3bec3-116">Type:</span></span>

*   <span data-ttu-id="3bec3-117">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="3bec3-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3bec3-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3bec3-118">Requirements</span></span>

|<span data-ttu-id="3bec3-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="3bec3-119">Requirement</span></span>| <span data-ttu-id="3bec3-120">Valor</span><span class="sxs-lookup"><span data-stu-id="3bec3-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="3bec3-121">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3bec3-121">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3bec3-122">1.0</span><span class="sxs-lookup"><span data-stu-id="3bec3-122">1.0</span></span>|
|[<span data-ttu-id="3bec3-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="3bec3-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3bec3-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3bec3-124">ReadItem</span></span>|
|[<span data-ttu-id="3bec3-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3bec3-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3bec3-126">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="3bec3-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3bec3-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3bec3-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="3bec3-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="3bec3-128">emailAddress :String</span></span>

<span data-ttu-id="3bec3-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="3bec3-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="3bec3-130">Type:</span><span class="sxs-lookup"><span data-stu-id="3bec3-130">Type:</span></span>

*   <span data-ttu-id="3bec3-131">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="3bec3-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3bec3-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3bec3-132">Requirements</span></span>

|<span data-ttu-id="3bec3-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="3bec3-133">Requirement</span></span>| <span data-ttu-id="3bec3-134">Valor</span><span class="sxs-lookup"><span data-stu-id="3bec3-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="3bec3-135">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3bec3-135">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3bec3-136">1.0</span><span class="sxs-lookup"><span data-stu-id="3bec3-136">1.0</span></span>|
|[<span data-ttu-id="3bec3-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="3bec3-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3bec3-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3bec3-138">ReadItem</span></span>|
|[<span data-ttu-id="3bec3-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3bec3-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3bec3-140">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="3bec3-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3bec3-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3bec3-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="3bec3-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="3bec3-142">timeZone :String</span></span>

<span data-ttu-id="3bec3-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="3bec3-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="3bec3-144">Type:</span><span class="sxs-lookup"><span data-stu-id="3bec3-144">Type:</span></span>

*   <span data-ttu-id="3bec3-145">sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="3bec3-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3bec3-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3bec3-146">Requirements</span></span>

|<span data-ttu-id="3bec3-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="3bec3-147">Requirement</span></span>| <span data-ttu-id="3bec3-148">Valor</span><span class="sxs-lookup"><span data-stu-id="3bec3-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="3bec3-149">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3bec3-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3bec3-150">1.0</span><span class="sxs-lookup"><span data-stu-id="3bec3-150">1.0</span></span>|
|[<span data-ttu-id="3bec3-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="3bec3-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3bec3-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3bec3-152">ReadItem</span></span>|
|[<span data-ttu-id="3bec3-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3bec3-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3bec3-154">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="3bec3-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3bec3-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3bec3-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```