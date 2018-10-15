# <a name="userprofile"></a><span data-ttu-id="4375c-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="4375c-101">userProfile</span></span>

### <span data-ttu-id="4375c-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="4375c-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="4375c-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4375c-104">Requirements</span></span>

|<span data-ttu-id="4375c-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="4375c-105">Requirement</span></span>| <span data-ttu-id="4375c-106">Valor</span><span class="sxs-lookup"><span data-stu-id="4375c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="4375c-107">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4375c-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4375c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="4375c-108">1.0</span></span>|
|[<span data-ttu-id="4375c-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4375c-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4375c-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4375c-110">ReadItem</span></span>|
|[<span data-ttu-id="4375c-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4375c-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4375c-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="4375c-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4375c-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="4375c-113">Members and methods</span></span>

| <span data-ttu-id="4375c-114">Membro</span><span class="sxs-lookup"><span data-stu-id="4375c-114">Member</span></span> | <span data-ttu-id="4375c-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="4375c-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4375c-116">displayName</span><span class="sxs-lookup"><span data-stu-id="4375c-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="4375c-117">Membro</span><span class="sxs-lookup"><span data-stu-id="4375c-117">Member</span></span> |
| [<span data-ttu-id="4375c-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="4375c-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="4375c-119">Membro</span><span class="sxs-lookup"><span data-stu-id="4375c-119">Member</span></span> |
| [<span data-ttu-id="4375c-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="4375c-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="4375c-121">Membro</span><span class="sxs-lookup"><span data-stu-id="4375c-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="4375c-122">Membros</span><span class="sxs-lookup"><span data-stu-id="4375c-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="4375c-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="4375c-123">displayName :String</span></span>

<span data-ttu-id="4375c-124">Obtém o nome para exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="4375c-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="4375c-125">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4375c-125">Type:</span></span>

*   <span data-ttu-id="4375c-126">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="4375c-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4375c-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4375c-127">Requirements</span></span>

|<span data-ttu-id="4375c-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="4375c-128">Requirement</span></span>| <span data-ttu-id="4375c-129">Valor</span><span class="sxs-lookup"><span data-stu-id="4375c-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="4375c-130">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4375c-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4375c-131">1.0</span><span class="sxs-lookup"><span data-stu-id="4375c-131">1.0</span></span>|
|[<span data-ttu-id="4375c-132">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4375c-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4375c-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4375c-133">ReadItem</span></span>|
|[<span data-ttu-id="4375c-134">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4375c-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4375c-135">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="4375c-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4375c-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4375c-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="4375c-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="4375c-137">emailAddress :String</span></span>

<span data-ttu-id="4375c-138">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="4375c-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="4375c-139">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4375c-139">Type:</span></span>

*   <span data-ttu-id="4375c-140">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="4375c-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4375c-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4375c-141">Requirements</span></span>

|<span data-ttu-id="4375c-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="4375c-142">Requirement</span></span>| <span data-ttu-id="4375c-143">Valor</span><span class="sxs-lookup"><span data-stu-id="4375c-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="4375c-144">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4375c-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4375c-145">1.0</span><span class="sxs-lookup"><span data-stu-id="4375c-145">1.0</span></span>|
|[<span data-ttu-id="4375c-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4375c-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4375c-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4375c-147">ReadItem</span></span>|
|[<span data-ttu-id="4375c-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4375c-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4375c-149">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="4375c-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4375c-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4375c-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="4375c-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="4375c-151">timeZone :String</span></span>

<span data-ttu-id="4375c-152">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="4375c-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="4375c-153">Tipo:</span><span class="sxs-lookup"><span data-stu-id="4375c-153">Type:</span></span>

*   <span data-ttu-id="4375c-154">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="4375c-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4375c-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4375c-155">Requirements</span></span>

|<span data-ttu-id="4375c-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="4375c-156">Requirement</span></span>| <span data-ttu-id="4375c-157">Valor</span><span class="sxs-lookup"><span data-stu-id="4375c-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="4375c-158">Versão mínima do conjunto de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4375c-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4375c-159">1.0</span><span class="sxs-lookup"><span data-stu-id="4375c-159">1.0</span></span>|
|[<span data-ttu-id="4375c-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4375c-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4375c-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4375c-161">ReadItem</span></span>|
|[<span data-ttu-id="4375c-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4375c-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4375c-163">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="4375c-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4375c-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4375c-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```