
# <a name="userprofile"></a><span data-ttu-id="1f25a-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="1f25a-101">userProfile</span></span>

### <span data-ttu-id="1f25a-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="1f25a-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f25a-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f25a-104">Requirements</span></span>

|<span data-ttu-id="1f25a-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f25a-105">Requirement</span></span>| <span data-ttu-id="1f25a-106">Valor</span><span class="sxs-lookup"><span data-stu-id="1f25a-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f25a-107">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f25a-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1f25a-108">1.0</span><span class="sxs-lookup"><span data-stu-id="1f25a-108">1.0</span></span>|
|[<span data-ttu-id="1f25a-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f25a-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1f25a-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f25a-110">ReadItem</span></span>|
|[<span data-ttu-id="1f25a-111">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="1f25a-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1f25a-112">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f25a-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1f25a-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="1f25a-113">Members and methods</span></span>

| <span data-ttu-id="1f25a-114">Membro</span><span class="sxs-lookup"><span data-stu-id="1f25a-114">Member</span></span> | <span data-ttu-id="1f25a-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="1f25a-115">Type</span></span> |
|--------|------|
| <span data-ttu-id="1f25a-116">[accountType](#accounttype-string)</span><span class="sxs-lookup"><span data-stu-id="1f25a-116">[](#accounttype-string)account_type=...</span></span> | <span data-ttu-id="1f25a-117">Membro</span><span class="sxs-lookup"><span data-stu-id="1f25a-117">Member</span></span> |
| [<span data-ttu-id="1f25a-118">displayName</span><span class="sxs-lookup"><span data-stu-id="1f25a-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="1f25a-119">Membro</span><span class="sxs-lookup"><span data-stu-id="1f25a-119">Member</span></span> |
| [<span data-ttu-id="1f25a-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="1f25a-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="1f25a-121">Membro</span><span class="sxs-lookup"><span data-stu-id="1f25a-121">Member</span></span> |
| [<span data-ttu-id="1f25a-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="1f25a-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="1f25a-123">Membro</span><span class="sxs-lookup"><span data-stu-id="1f25a-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="1f25a-124">Membros</span><span class="sxs-lookup"><span data-stu-id="1f25a-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="1f25a-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="1f25a-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="1f25a-126">Atualmente, este membro só é suportado no Outlook 2016 para Mac, build 16.9.1212 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="1f25a-126">This member is currently only supported in Outlook 2016 for Mac, build 16.9.1212 and greater.</span></span>

<span data-ttu-id="1f25a-127">Obtém o tipo de conta do usuário associada à caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="1f25a-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="1f25a-128">Os valores possíveis são listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="1f25a-128">The values are listed in the following table.</span></span>

| <span data-ttu-id="1f25a-129">Valor</span><span class="sxs-lookup"><span data-stu-id="1f25a-129">Value</span></span> | <span data-ttu-id="1f25a-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="1f25a-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="1f25a-131">A caixa de correio está em um servidor local do Exchange.</span><span class="sxs-lookup"><span data-stu-id="1f25a-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="1f25a-132">A caixa de correio está associada a uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="1f25a-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="1f25a-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="1f25a-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="1f25a-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="1f25a-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="1f25a-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f25a-135">Type:</span></span>

*   <span data-ttu-id="1f25a-136">String</span><span class="sxs-lookup"><span data-stu-id="1f25a-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f25a-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f25a-137">Requirements</span></span>

|<span data-ttu-id="1f25a-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f25a-138">Requirement</span></span>| <span data-ttu-id="1f25a-139">Valor</span><span class="sxs-lookup"><span data-stu-id="1f25a-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f25a-140">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f25a-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1f25a-141">1.6</span><span class="sxs-lookup"><span data-stu-id="1f25a-141">-16</span></span> |
|[<span data-ttu-id="1f25a-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f25a-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1f25a-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f25a-143">ReadItem</span></span>|
|[<span data-ttu-id="1f25a-144">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="1f25a-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1f25a-145">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f25a-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f25a-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f25a-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="1f25a-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="1f25a-147">displayName :String</span></span>

<span data-ttu-id="1f25a-148">Obtém o nome para exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="1f25a-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="1f25a-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f25a-149">Type:</span></span>

*   <span data-ttu-id="1f25a-150">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f25a-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f25a-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f25a-151">Requirements</span></span>

|<span data-ttu-id="1f25a-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f25a-152">Requirement</span></span>| <span data-ttu-id="1f25a-153">Valor</span><span class="sxs-lookup"><span data-stu-id="1f25a-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f25a-154">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f25a-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1f25a-155">1.0</span><span class="sxs-lookup"><span data-stu-id="1f25a-155">1.0</span></span>|
|[<span data-ttu-id="1f25a-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f25a-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1f25a-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f25a-157">ReadItem</span></span>|
|[<span data-ttu-id="1f25a-158">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="1f25a-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1f25a-159">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f25a-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f25a-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f25a-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="1f25a-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="1f25a-161">emailAddress :String</span></span>

<span data-ttu-id="1f25a-162">Obtém o endereço de e-mail SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="1f25a-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="1f25a-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f25a-163">Type:</span></span>

*   <span data-ttu-id="1f25a-164">String</span><span class="sxs-lookup"><span data-stu-id="1f25a-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f25a-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f25a-165">Requirements</span></span>

|<span data-ttu-id="1f25a-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f25a-166">Requirement</span></span>| <span data-ttu-id="1f25a-167">Valor</span><span class="sxs-lookup"><span data-stu-id="1f25a-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f25a-168">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f25a-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1f25a-169">1.0</span><span class="sxs-lookup"><span data-stu-id="1f25a-169">1.0</span></span>|
|[<span data-ttu-id="1f25a-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f25a-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1f25a-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f25a-171">ReadItem</span></span>|
|[<span data-ttu-id="1f25a-172">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="1f25a-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1f25a-173">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f25a-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f25a-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f25a-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="1f25a-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="1f25a-175">timeZone :String</span></span>

<span data-ttu-id="1f25a-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="1f25a-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="1f25a-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="1f25a-177">Type:</span></span>

*   <span data-ttu-id="1f25a-178">Sequência de caracteres</span><span class="sxs-lookup"><span data-stu-id="1f25a-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1f25a-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1f25a-179">Requirements</span></span>

|<span data-ttu-id="1f25a-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="1f25a-180">Requirement</span></span>| <span data-ttu-id="1f25a-181">Valor</span><span class="sxs-lookup"><span data-stu-id="1f25a-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="1f25a-182">Versão do conjunto mínimo de requisitos de caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1f25a-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1f25a-183">1.0</span><span class="sxs-lookup"><span data-stu-id="1f25a-183">1.0</span></span>|
|[<span data-ttu-id="1f25a-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1f25a-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1f25a-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1f25a-185">ReadItem</span></span>|
|[<span data-ttu-id="1f25a-186">Modo aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="1f25a-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1f25a-187">Redigir ou ler</span><span class="sxs-lookup"><span data-stu-id="1f25a-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1f25a-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1f25a-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```