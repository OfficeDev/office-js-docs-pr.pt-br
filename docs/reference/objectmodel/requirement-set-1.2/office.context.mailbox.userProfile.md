
# <a name="userprofile"></a><span data-ttu-id="67d7c-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="67d7c-101">userProfile</span></span>

### <span data-ttu-id="67d7c-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="67d7c-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="67d7c-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="67d7c-104">Requirements</span></span>

|<span data-ttu-id="67d7c-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="67d7c-105">Requirement</span></span>| <span data-ttu-id="67d7c-106">Valor</span><span class="sxs-lookup"><span data-stu-id="67d7c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="67d7c-107">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="67d7c-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="67d7c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="67d7c-108">1.0</span></span>|
|[<span data-ttu-id="67d7c-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="67d7c-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="67d7c-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67d7c-110">ReadItem</span></span>|
|[<span data-ttu-id="67d7c-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="67d7c-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="67d7c-112">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="67d7c-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="67d7c-113">Membros</span><span class="sxs-lookup"><span data-stu-id="67d7c-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="67d7c-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="67d7c-114">displayName :String</span></span>

<span data-ttu-id="67d7c-115">Obtém o nome para exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="67d7c-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="67d7c-116">Tipo:</span><span class="sxs-lookup"><span data-stu-id="67d7c-116">Type:</span></span>

*   <span data-ttu-id="67d7c-117">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="67d7c-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="67d7c-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="67d7c-118">Requirements</span></span>

|<span data-ttu-id="67d7c-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="67d7c-119">Requirement</span></span>| <span data-ttu-id="67d7c-120">Valor</span><span class="sxs-lookup"><span data-stu-id="67d7c-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="67d7c-121">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="67d7c-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="67d7c-122">1.0</span><span class="sxs-lookup"><span data-stu-id="67d7c-122">1.0</span></span>|
|[<span data-ttu-id="67d7c-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="67d7c-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="67d7c-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67d7c-124">ReadItem</span></span>|
|[<span data-ttu-id="67d7c-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="67d7c-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="67d7c-126">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="67d7c-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="67d7c-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="67d7c-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="67d7c-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="67d7c-128">emailAddress :String</span></span>

<span data-ttu-id="67d7c-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="67d7c-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="67d7c-130">Tipo:</span><span class="sxs-lookup"><span data-stu-id="67d7c-130">Type:</span></span>

*   <span data-ttu-id="67d7c-131">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="67d7c-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="67d7c-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="67d7c-132">Requirements</span></span>

|<span data-ttu-id="67d7c-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="67d7c-133">Requirement</span></span>| <span data-ttu-id="67d7c-134">Valor</span><span class="sxs-lookup"><span data-stu-id="67d7c-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="67d7c-135">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="67d7c-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="67d7c-136">1.0</span><span class="sxs-lookup"><span data-stu-id="67d7c-136">1.0</span></span>|
|[<span data-ttu-id="67d7c-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="67d7c-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="67d7c-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67d7c-138">ReadItem</span></span>|
|[<span data-ttu-id="67d7c-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="67d7c-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="67d7c-140">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="67d7c-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="67d7c-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="67d7c-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="67d7c-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="67d7c-142">timeZone :String</span></span>

<span data-ttu-id="67d7c-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="67d7c-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="67d7c-144">Tipo:</span><span class="sxs-lookup"><span data-stu-id="67d7c-144">Type:</span></span>

*   <span data-ttu-id="67d7c-145">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="67d7c-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="67d7c-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="67d7c-146">Requirements</span></span>

|<span data-ttu-id="67d7c-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="67d7c-147">Requirement</span></span>| <span data-ttu-id="67d7c-148">Valor</span><span class="sxs-lookup"><span data-stu-id="67d7c-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="67d7c-149">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="67d7c-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="67d7c-150">1.0</span><span class="sxs-lookup"><span data-stu-id="67d7c-150">1.0</span></span>|
|[<span data-ttu-id="67d7c-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="67d7c-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="67d7c-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67d7c-152">ReadItem</span></span>|
|[<span data-ttu-id="67d7c-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="67d7c-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="67d7c-154">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="67d7c-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="67d7c-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="67d7c-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```