# <a name="userprofile"></a><span data-ttu-id="90481-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="90481-101">userProfile</span></span>

### <span data-ttu-id="90481-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="90481-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="90481-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="90481-104">Requirements</span></span>

|<span data-ttu-id="90481-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="90481-105">Requirement</span></span>| <span data-ttu-id="90481-106">Valor</span><span class="sxs-lookup"><span data-stu-id="90481-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="90481-107">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="90481-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="90481-108">1.0</span><span class="sxs-lookup"><span data-stu-id="90481-108">1.0</span></span>|
|[<span data-ttu-id="90481-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="90481-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="90481-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="90481-110">ReadItem</span></span>|
|[<span data-ttu-id="90481-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="90481-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="90481-112">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="90481-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="90481-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="90481-113">Members and methods</span></span>

| <span data-ttu-id="90481-114">Membro</span><span class="sxs-lookup"><span data-stu-id="90481-114">Member</span></span> | <span data-ttu-id="90481-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="90481-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="90481-116">displayName</span><span class="sxs-lookup"><span data-stu-id="90481-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="90481-117">Membro</span><span class="sxs-lookup"><span data-stu-id="90481-117">Member</span></span> |
| [<span data-ttu-id="90481-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="90481-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="90481-119">Membro</span><span class="sxs-lookup"><span data-stu-id="90481-119">Member</span></span> |
| [<span data-ttu-id="90481-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="90481-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="90481-121">Membro</span><span class="sxs-lookup"><span data-stu-id="90481-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="90481-122">Membros</span><span class="sxs-lookup"><span data-stu-id="90481-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="90481-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="90481-123">displayName :String</span></span>

<span data-ttu-id="90481-124">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="90481-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="90481-125">Tipo:</span><span class="sxs-lookup"><span data-stu-id="90481-125">Type:</span></span>

*   <span data-ttu-id="90481-126">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="90481-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="90481-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="90481-127">Requirements</span></span>

|<span data-ttu-id="90481-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="90481-128">Requirement</span></span>| <span data-ttu-id="90481-129">Valor</span><span class="sxs-lookup"><span data-stu-id="90481-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="90481-130">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="90481-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="90481-131">1.0</span><span class="sxs-lookup"><span data-stu-id="90481-131">1.0</span></span>|
|[<span data-ttu-id="90481-132">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="90481-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="90481-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="90481-133">ReadItem</span></span>|
|[<span data-ttu-id="90481-134">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="90481-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="90481-135">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="90481-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="90481-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="90481-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="90481-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="90481-137">emailAddress :String</span></span>

<span data-ttu-id="90481-138">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="90481-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="90481-139">Tipo:</span><span class="sxs-lookup"><span data-stu-id="90481-139">Type:</span></span>

*   <span data-ttu-id="90481-140">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="90481-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="90481-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="90481-141">Requirements</span></span>

|<span data-ttu-id="90481-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="90481-142">Requirement</span></span>| <span data-ttu-id="90481-143">Valor</span><span class="sxs-lookup"><span data-stu-id="90481-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="90481-144">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="90481-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="90481-145">1.0</span><span class="sxs-lookup"><span data-stu-id="90481-145">1.0</span></span>|
|[<span data-ttu-id="90481-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="90481-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="90481-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="90481-147">ReadItem</span></span>|
|[<span data-ttu-id="90481-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="90481-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="90481-149">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="90481-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="90481-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="90481-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="90481-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="90481-151">timeZone :String</span></span>

<span data-ttu-id="90481-152">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="90481-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="90481-153">Tipo:</span><span class="sxs-lookup"><span data-stu-id="90481-153">Type:</span></span>

*   <span data-ttu-id="90481-154">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="90481-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="90481-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="90481-155">Requirements</span></span>

|<span data-ttu-id="90481-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="90481-156">Requirement</span></span>| <span data-ttu-id="90481-157">Valor</span><span class="sxs-lookup"><span data-stu-id="90481-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="90481-158">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="90481-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="90481-159">1.0</span><span class="sxs-lookup"><span data-stu-id="90481-159">1.0</span></span>|
|[<span data-ttu-id="90481-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="90481-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="90481-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="90481-161">ReadItem</span></span>|
|[<span data-ttu-id="90481-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="90481-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="90481-163">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="90481-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="90481-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="90481-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```