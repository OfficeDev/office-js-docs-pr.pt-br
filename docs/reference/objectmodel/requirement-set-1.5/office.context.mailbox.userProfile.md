# <a name="userprofile"></a><span data-ttu-id="b117e-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="b117e-101">userProfile</span></span>

### <span data-ttu-id="b117e-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="b117e-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="b117e-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b117e-104">Requirements</span></span>

|<span data-ttu-id="b117e-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="b117e-105">Requirement</span></span>| <span data-ttu-id="b117e-106">Valor</span><span class="sxs-lookup"><span data-stu-id="b117e-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="b117e-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b117e-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b117e-108">1.0</span><span class="sxs-lookup"><span data-stu-id="b117e-108">1.0</span></span>|
|[<span data-ttu-id="b117e-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b117e-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b117e-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b117e-110">ReadItem</span></span>|
|[<span data-ttu-id="b117e-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b117e-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b117e-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b117e-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b117e-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="b117e-113">Members and methods</span></span>

| <span data-ttu-id="b117e-114">Membro</span><span class="sxs-lookup"><span data-stu-id="b117e-114">Member</span></span> | <span data-ttu-id="b117e-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="b117e-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b117e-116">displayName</span><span class="sxs-lookup"><span data-stu-id="b117e-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="b117e-117">Membro</span><span class="sxs-lookup"><span data-stu-id="b117e-117">Member</span></span> |
| [<span data-ttu-id="b117e-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="b117e-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="b117e-119">Membro</span><span class="sxs-lookup"><span data-stu-id="b117e-119">Member</span></span> |
| [<span data-ttu-id="b117e-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="b117e-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="b117e-121">Membro</span><span class="sxs-lookup"><span data-stu-id="b117e-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="b117e-122">Membros</span><span class="sxs-lookup"><span data-stu-id="b117e-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="b117e-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="b117e-123">displayName :String</span></span>

<span data-ttu-id="b117e-124">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="b117e-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="b117e-125">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b117e-125">Type:</span></span>

*   <span data-ttu-id="b117e-126">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b117e-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b117e-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b117e-127">Requirements</span></span>

|<span data-ttu-id="b117e-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="b117e-128">Requirement</span></span>| <span data-ttu-id="b117e-129">Valor</span><span class="sxs-lookup"><span data-stu-id="b117e-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="b117e-130">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b117e-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b117e-131">1.0</span><span class="sxs-lookup"><span data-stu-id="b117e-131">1.0</span></span>|
|[<span data-ttu-id="b117e-132">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b117e-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b117e-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b117e-133">ReadItem</span></span>|
|[<span data-ttu-id="b117e-134">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b117e-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b117e-135">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b117e-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b117e-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b117e-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="b117e-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="b117e-137">emailAddress :String</span></span>

<span data-ttu-id="b117e-138">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="b117e-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="b117e-139">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b117e-139">Type:</span></span>

*   <span data-ttu-id="b117e-140">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b117e-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b117e-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b117e-141">Requirements</span></span>

|<span data-ttu-id="b117e-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="b117e-142">Requirement</span></span>| <span data-ttu-id="b117e-143">Valor</span><span class="sxs-lookup"><span data-stu-id="b117e-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="b117e-144">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b117e-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b117e-145">1.0</span><span class="sxs-lookup"><span data-stu-id="b117e-145">1.0</span></span>|
|[<span data-ttu-id="b117e-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b117e-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b117e-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b117e-147">ReadItem</span></span>|
|[<span data-ttu-id="b117e-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b117e-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b117e-149">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b117e-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b117e-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b117e-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="b117e-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="b117e-151">timeZone :String</span></span>

<span data-ttu-id="b117e-152">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="b117e-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="b117e-153">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b117e-153">Type:</span></span>

*   <span data-ttu-id="b117e-154">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b117e-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b117e-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b117e-155">Requirements</span></span>

|<span data-ttu-id="b117e-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="b117e-156">Requirement</span></span>| <span data-ttu-id="b117e-157">Valor</span><span class="sxs-lookup"><span data-stu-id="b117e-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="b117e-158">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b117e-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b117e-159">1.0</span><span class="sxs-lookup"><span data-stu-id="b117e-159">1.0</span></span>|
|[<span data-ttu-id="b117e-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b117e-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b117e-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b117e-161">ReadItem</span></span>|
|[<span data-ttu-id="b117e-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b117e-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b117e-163">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b117e-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b117e-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b117e-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```