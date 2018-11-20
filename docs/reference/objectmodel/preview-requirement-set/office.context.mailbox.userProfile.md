
# <a name="userprofile"></a><span data-ttu-id="32787-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="32787-101">userProfile</span></span>

### <span data-ttu-id="32787-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="32787-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="32787-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="32787-104">Requirements</span></span>

|<span data-ttu-id="32787-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="32787-105">Requirement</span></span>| <span data-ttu-id="32787-106">Valor</span><span class="sxs-lookup"><span data-stu-id="32787-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="32787-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="32787-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="32787-108">1.0</span><span class="sxs-lookup"><span data-stu-id="32787-108">1.0</span></span>|
|[<span data-ttu-id="32787-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="32787-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="32787-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="32787-110">ReadItem</span></span>|
|[<span data-ttu-id="32787-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="32787-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="32787-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="32787-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="32787-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="32787-113">Members and methods</span></span>

| <span data-ttu-id="32787-114">Membro</span><span class="sxs-lookup"><span data-stu-id="32787-114">Member</span></span> | <span data-ttu-id="32787-115">Type</span><span class="sxs-lookup"><span data-stu-id="32787-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="32787-116">accountType</span><span class="sxs-lookup"><span data-stu-id="32787-116">AccountType</span></span>](#accounttype-string) | <span data-ttu-id="32787-117">Member</span><span class="sxs-lookup"><span data-stu-id="32787-117">Member</span></span> |
| [<span data-ttu-id="32787-118">displayName</span><span class="sxs-lookup"><span data-stu-id="32787-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="32787-119">Membro</span><span class="sxs-lookup"><span data-stu-id="32787-119">Member</span></span> |
| [<span data-ttu-id="32787-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="32787-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="32787-121">Membro</span><span class="sxs-lookup"><span data-stu-id="32787-121">Member</span></span> |
| [<span data-ttu-id="32787-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="32787-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="32787-123">Membro</span><span class="sxs-lookup"><span data-stu-id="32787-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="32787-124">Members</span><span class="sxs-lookup"><span data-stu-id="32787-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="32787-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="32787-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="32787-126">No momento, esse membro só tem suporte no Outlook 2016 ou posterior para Mac (build 16.9.1212 e posterior).</span><span class="sxs-lookup"><span data-stu-id="32787-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="32787-127">Obtém o tipo de conta do usuário associado à caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="32787-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="32787-128">Os valores possíveis são listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="32787-128">The values are listed in the following table.</span></span>

| <span data-ttu-id="32787-129">Value</span><span class="sxs-lookup"><span data-stu-id="32787-129">Value</span></span> | <span data-ttu-id="32787-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="32787-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="32787-131">A caixa de correio está em um servidor local do Exchange.</span><span class="sxs-lookup"><span data-stu-id="32787-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="32787-132">A caixa de correio está associada a uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="32787-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="32787-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="32787-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="32787-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="32787-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="32787-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="32787-135">Type:</span></span>

*   <span data-ttu-id="32787-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="32787-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="32787-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="32787-137">Requirements</span></span>

|<span data-ttu-id="32787-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="32787-138">Requirement</span></span>| <span data-ttu-id="32787-139">Valor</span><span class="sxs-lookup"><span data-stu-id="32787-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="32787-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="32787-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="32787-141">1.6</span><span class="sxs-lookup"><span data-stu-id="32787-141">-16</span></span> |
|[<span data-ttu-id="32787-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="32787-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="32787-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="32787-143">ReadItem</span></span>|
|[<span data-ttu-id="32787-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="32787-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="32787-145">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="32787-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="32787-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="32787-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="32787-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="32787-147">displayName :String</span></span>

<span data-ttu-id="32787-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="32787-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="32787-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="32787-149">Type:</span></span>

*   <span data-ttu-id="32787-150">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="32787-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="32787-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="32787-151">Requirements</span></span>

|<span data-ttu-id="32787-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="32787-152">Requirement</span></span>| <span data-ttu-id="32787-153">Valor</span><span class="sxs-lookup"><span data-stu-id="32787-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="32787-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="32787-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="32787-155">1.0</span><span class="sxs-lookup"><span data-stu-id="32787-155">1.0</span></span>|
|[<span data-ttu-id="32787-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="32787-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="32787-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="32787-157">ReadItem</span></span>|
|[<span data-ttu-id="32787-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="32787-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="32787-159">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="32787-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="32787-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="32787-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="32787-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="32787-161">emailAddress :String</span></span>

<span data-ttu-id="32787-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="32787-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="32787-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="32787-163">Type:</span></span>

*   <span data-ttu-id="32787-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="32787-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="32787-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="32787-165">Requirements</span></span>

|<span data-ttu-id="32787-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="32787-166">Requirement</span></span>| <span data-ttu-id="32787-167">Valor</span><span class="sxs-lookup"><span data-stu-id="32787-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="32787-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="32787-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="32787-169">1.0</span><span class="sxs-lookup"><span data-stu-id="32787-169">1.0</span></span>|
|[<span data-ttu-id="32787-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="32787-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="32787-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="32787-171">ReadItem</span></span>|
|[<span data-ttu-id="32787-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="32787-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="32787-173">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="32787-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="32787-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="32787-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="32787-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="32787-175">timeZone :String</span></span>

<span data-ttu-id="32787-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="32787-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="32787-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="32787-177">Type:</span></span>

*   <span data-ttu-id="32787-178">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="32787-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="32787-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="32787-179">Requirements</span></span>

|<span data-ttu-id="32787-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="32787-180">Requirement</span></span>| <span data-ttu-id="32787-181">Valor</span><span class="sxs-lookup"><span data-stu-id="32787-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="32787-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="32787-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="32787-183">1.0</span><span class="sxs-lookup"><span data-stu-id="32787-183">1.0</span></span>|
|[<span data-ttu-id="32787-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="32787-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="32787-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="32787-185">ReadItem</span></span>|
|[<span data-ttu-id="32787-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="32787-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="32787-187">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="32787-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="32787-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="32787-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```