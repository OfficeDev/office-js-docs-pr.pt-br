
# <a name="userprofile"></a><span data-ttu-id="c1efd-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="c1efd-101">userProfile</span></span>

### <span data-ttu-id="c1efd-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="c1efd-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1efd-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c1efd-104">Requirements</span></span>

|<span data-ttu-id="c1efd-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="c1efd-105">Requirement</span></span>| <span data-ttu-id="c1efd-106">Valor</span><span class="sxs-lookup"><span data-stu-id="c1efd-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1efd-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c1efd-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1efd-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c1efd-108">1.0</span></span>|
|[<span data-ttu-id="c1efd-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c1efd-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1efd-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1efd-110">ReadItem</span></span>|
|[<span data-ttu-id="c1efd-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c1efd-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1efd-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="c1efd-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="c1efd-113">Membros</span><span class="sxs-lookup"><span data-stu-id="c1efd-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="c1efd-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="c1efd-114">displayName :String</span></span>

<span data-ttu-id="c1efd-115">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="c1efd-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c1efd-116">Tipo:</span><span class="sxs-lookup"><span data-stu-id="c1efd-116">Type:</span></span>

*   <span data-ttu-id="c1efd-117">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c1efd-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1efd-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c1efd-118">Requirements</span></span>

|<span data-ttu-id="c1efd-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="c1efd-119">Requirement</span></span>| <span data-ttu-id="c1efd-120">Valor</span><span class="sxs-lookup"><span data-stu-id="c1efd-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1efd-121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c1efd-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1efd-122">1.0</span><span class="sxs-lookup"><span data-stu-id="c1efd-122">1.0</span></span>|
|[<span data-ttu-id="c1efd-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c1efd-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1efd-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1efd-124">ReadItem</span></span>|
|[<span data-ttu-id="c1efd-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c1efd-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1efd-126">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="c1efd-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1efd-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c1efd-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="c1efd-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="c1efd-128">emailAddress :String</span></span>

<span data-ttu-id="c1efd-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="c1efd-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c1efd-130">Tipo:</span><span class="sxs-lookup"><span data-stu-id="c1efd-130">Type:</span></span>

*   <span data-ttu-id="c1efd-131">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c1efd-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1efd-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c1efd-132">Requirements</span></span>

|<span data-ttu-id="c1efd-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="c1efd-133">Requirement</span></span>| <span data-ttu-id="c1efd-134">Valor</span><span class="sxs-lookup"><span data-stu-id="c1efd-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1efd-135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c1efd-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1efd-136">1.0</span><span class="sxs-lookup"><span data-stu-id="c1efd-136">1.0</span></span>|
|[<span data-ttu-id="c1efd-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c1efd-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1efd-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1efd-138">ReadItem</span></span>|
|[<span data-ttu-id="c1efd-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c1efd-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1efd-140">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="c1efd-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1efd-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c1efd-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="c1efd-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="c1efd-142">timeZone :String</span></span>

<span data-ttu-id="c1efd-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="c1efd-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c1efd-144">Tipo:</span><span class="sxs-lookup"><span data-stu-id="c1efd-144">Type:</span></span>

*   <span data-ttu-id="c1efd-145">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c1efd-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1efd-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c1efd-146">Requirements</span></span>

|<span data-ttu-id="c1efd-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="c1efd-147">Requirement</span></span>| <span data-ttu-id="c1efd-148">Valor</span><span class="sxs-lookup"><span data-stu-id="c1efd-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1efd-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c1efd-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1efd-150">1.0</span><span class="sxs-lookup"><span data-stu-id="c1efd-150">1.0</span></span>|
|[<span data-ttu-id="c1efd-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c1efd-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1efd-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1efd-152">ReadItem</span></span>|
|[<span data-ttu-id="c1efd-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c1efd-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1efd-154">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="c1efd-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1efd-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c1efd-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```